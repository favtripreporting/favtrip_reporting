from __future__ import annotations
import csv
import io
import re
import requests
from dataclasses import dataclass
from datetime import datetime
from zoneinfo import ZoneInfo
from email.message import EmailMessage

from io import BytesIO
from openpyxl import load_workbook, Workbook


from .config import Config
from .google_client import get_credentials, services
from .sheets_utils import (
    delete_sheet, copy_sheet_as, copy_first_sheet_as, refresh_sheets_with_prefix,
    get_value, first_gid
)
from .drive_utils import find_latest_sheet, upload_to_drive
from .gmail_utils import send_email, email_manager_report

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def clean_tag(s: str) -> str:
    return re.sub(r"[^A-Za-z0-9._-]+", "-", s.strip()).strip("-") or "UNKNOWN"


def export_sheet(creds, spreadsheet_id: str, gid: str | int, fmt: str) -> bytes:
    url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export?format={fmt}&gid={gid}"
    headers = {"Authorization": f"Bearer {creds.token}"}
    r = requests.get(url, headers=headers)
    r.raise_for_status()

    content = r.content

    # Freeze formulas if exporting Excel
    if fmt == "xlsx":
        df = pd.read_excel(BytesIO(content))
        output = BytesIO()
        df.to_excel(output, index=False)
        return output.getvalue()

    return content


def timestamp_now(tz: str, fmt: str) -> str:
    return datetime.now(ZoneInfo(tz)).strftime(fmt)

import re

_EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")

def _clean_emails(items):
    """
    Accepts a list or a comma-separated string and returns a list of valid emails.
    Trailing commas and blanks are removed. Invalid tokens are dropped silently.
    """
    if items is None:
        return []
    if isinstance(items, str):
        items = [p.strip() for p in items.split(",")]
    return [e for e in (p.strip() for p in items) if e and _EMAIL_RE.match(e)]

def _fallback_recipients(hint, *candidates):
    """
    Return the first non-empty, valid recipient list from the provided candidates.
    If all candidates are empty/invalid, raise a friendly error.
    """
    for c in candidates:
        cleaned = _clean_emails(c)
        if cleaned:
            return cleaned
    # Nothing usable found:
    raise ValueError(
        f"No valid recipients available for: {hint}. "
        f"Please provide at least one email in the UI or .env "
        f"(TO_RECIPIENTS, DEFAULT_ORDER_RECIPIENTS, or per-report-key)."
    )


@dataclass
class RunResult:
    ok: bool
    elapsed_seconds: int
    location: str
    timestamp: str
    manager_pdf_link: str | None
    full_order_link: str | None


def run_pipeline(cfg: Config, logger=None) -> RunResult:
    import time
    start = time.perf_counter()

    if logger:
        logger.info("Authorizing with Google APIs…")
    creds = get_credentials(cfg.SCOPES, cfg.REDIRECT_PORT, cfg.FORCE_REAUTH)
    sheets_svc, drive_svc, gmail_svc = services(creds, cfg.HTTP_TIMEOUT_SECONDS)
    if logger:
        logger.info("Google services ready")

    # Step 1: latest incoming
    if logger:
        logger.info("Finding latest incoming spreadsheet…")
    latest = find_latest_sheet(drive_svc, cfg.INCOMING_FOLDER_ID)
    if not latest:
        raise RuntimeError("No incoming report found.")
    new_report_id = latest["id"]
    if logger:
        logger.info(f"Latest incoming: {latest['name']} ({new_report_id})")

    # Step 2: prep calculations workbook
    if logger:
        logger.info("Preparing calculations workbook…")
    delete_sheet(sheets_svc, cfg.CALC_SPREADSHEET_ID, "Last Week")
    try:
        copy_sheet_as(sheets_svc, cfg.CALC_SPREADSHEET_ID, "Current Week", "Last Week")
        if logger:
            logger.info("Copied old 'Current Week' to 'Last Week'")
    except Exception:
        if logger:
            logger.warn("No 'Current Week' sheet exists to copy")
    delete_sheet(sheets_svc, cfg.CALC_SPREADSHEET_ID, "Current Week")
    copy_first_sheet_as(sheets_svc, new_report_id, cfg.CALC_SPREADSHEET_ID, "Current Week")
    if logger:
        logger.info("Inserted new 'Current Week' from latest incoming report")

    if logger:
        logger.info("Refreshing reference sheets (prefix 'REFR: ')…")
    refresh_sheets_with_prefix(sheets_svc, cfg.CALC_SPREADSHEET_ID, prefix="REFR: ", logger=logger)

    # Step 3: read location code
    location = get_value(sheets_svc, cfg.CALC_SPREADSHEET_ID, cfg.LOCATION_SHEET_TITLE, cfg.LOCATION_NAMED_RANGE)
    ts = timestamp_now(cfg.TIMESTAMP_TZ, cfg.TIMESTAMP_FMT)
    if logger:
        logger.info(f"Location: {location}; Timestamp: {ts}")

    # Step 4A: Manager Report PDF
    if logger:
        logger.info("Exporting Manager Report (PDF)…")
    pdf_bytes = export_sheet(creds, cfg.CALC_SPREADSHEET_ID, cfg.GID_MANAGER_PDF, "pdf")
    pdf_name = f"Manager_Report_{ts}_{location}.pdf"
    uploaded_pdf = upload_to_drive(drive_svc, pdf_bytes, pdf_name, "application/pdf", cfg.MANAGER_REPORT_FOLDER_ID, to_sheet=False)
    manager_link = uploaded_pdf.get("webViewLink")
    if logger:
        logger.info(f"Uploaded Manager PDF: {manager_link}")

    # Step 4B: Master Order XLSX
    if logger:
        logger.info("Exporting Master Order (XLSX)…")
    master_xlsx_bytes = export_sheet(creds, cfg.CALC_SPREADSHEET_ID, cfg.GID_ORDER_CSV, "xlsx")

    # Step 4C: Full order upload (XLSX) and export (PDF)
    full_xlsx_name = f"Order_Report_{ts}_{location}_FULL.xlsx"
    full_created = upload_to_drive(drive_svc, master_xlsx_bytes, full_xlsx_name, XLSX_MIME, cfg.ORDER_REPORT_FOLDER_ID, to_sheet=True)
    full_file_id = full_created["id"]
    full_gid = first_gid(sheets_svc, full_file_id)
    full_pdf = export_sheet(creds, full_file_id, full_gid, "pdf")
    full_pdf_name = f"Order_Report_{ts}_{location}_FULL.pdf"
    if logger:
        logger.info(f"Uploaded FULL sheet: {full_created.get('webViewLink')}")

    # Step 4D: Create per-report-key outputs (XLSX) and email

    # --- Parse the master XLSX into rows of dicts ---
    wb = load_workbook(filename=BytesIO(master_xlsx_bytes), read_only=True, data_only=True)
    ws = wb.active  # or specify a sheet name if needed
    rows_iter = ws.iter_rows(values_only=True)

    # Header row
    headers = next(rows_iter, None)
    if not headers:
        raise RuntimeError("XLSX has no header.")

    headers = [str(h).strip() if h is not None else "" for h in headers]
    report_col = next((h for h in headers if h and h.lower() == "report_key"), None)
    if not report_col:
        raise RuntimeError("Report_Key column missing.")

    # Materialize rows as list[dict]
    rows = []
    for r in rows_iter:
        rows.append({headers[i]: (r[i] if i < len(r) else None) for i in range(len(headers))})

    # Group by report key
    groups: dict[str, list[dict]] = {}
    for r in rows:
        key = (str(r.get(report_col) or "").strip()) or "UNASSIGNED"
        groups.setdefault(key, []).append(r)

    for key, key_rows in groups.items():
        if not cfg.USE_ALL_REPORT_KEYS and key.upper() not in (cfg.REPORT_KEY_RUN_LIST or []):
            continue

        # Build a per-key XLSX in memory
        out_wb = Workbook(write_only=True)
        out_ws = out_wb.create_sheet("Sheet1")
        out_ws.append(headers)
        for rr in key_rows:
            out_ws.append([rr.get(h) for h in headers])

        bio = BytesIO()
        out_wb.save(bio)
        key_xlsx_bytes = bio.getvalue()

        tag = clean_tag(key.upper())
        xlsx_name = f"Order_Report_{ts}_{location}_{tag}.xlsx"

        # Upload XLSX to Drive; set to_sheet=True so Drive converts to Google Sheet (needed for your PDF export + link)
        created = upload_to_drive(drive_svc, key_xlsx_bytes, xlsx_name, XLSX_MIME, cfg.ORDER_REPORT_FOLDER_ID, to_sheet=True) 
        file_id = created["id"]
        gid = first_gid(sheets_svc, file_id)

        # Export the Google Sheet as PDF (unchanged behavior)
        pdf = export_sheet(creds, file_id, gid, "pdf")
        pdfname = f"Order_Report_{ts}_{location}_{tag}.pdf"

        # Prefer per‑key recipients; else default; else TO
        candidates = None
        if cfg.REPORT_KEY_RECIPIENTS:
            # The keys in UI are upper-cased; normalize here too
            candidates = cfg.REPORT_KEY_RECIPIENTS.get(tag)

        recipients = _fallback_recipients(
            f"REPORT_KEY {tag}",
            candidates,
            cfg.DEFAULT_ORDER_RECIPIENTS,
            cfg.TO_RECIPIENTS,
        )

        msg = EmailMessage()
        msg["Subject"] = f"Order Report – {ts} – {location} – {tag}"
        msg["From"] = "me"
        msg["To"] = ", ".join(recipients)
        if cfg.CC_RECIPIENTS:
            msg["Cc"] = ", ".join(cfg.CC_RECIPIENTS)

        # Keep the Google Sheet link; PDF remains the attached artifact
        msg.set_content(
            f"Hi {key} team,\nYour order report is ready. \n"
            f"Google Sheet: {created.get('webViewLink')}\n"
            f"Attached: {pdfname}\n—Automated"
        )
        msg.add_attachment(pdf, maintype="application", subtype="pdf", filename=pdfname)
        if cfg.INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL:
            msg.add_attachment(full_pdf, maintype="application", subtype="pdf", filename=full_pdf_name)

        send_email(gmail_svc, "me", msg)
        if logger:
            logger.info(f"Emailed {tag}")

    # Step 4E: Send Manager Report (guarded by cfg.EMAIL_MANAGER_REPORT)
    if getattr(cfg, "EMAIL_MANAGER_REPORT", True):
        to_list = _fallback_recipients("Manager Report (TO_RECIPIENTS)", cfg.TO_RECIPIENTS)
        cc_list = _clean_emails(cfg.CC_RECIPIENTS)
        email_manager_report(
            gmail_svc, "me", to_list, cc_list,
            pdf_name, pdf_bytes, manager_link, ts, location
        )
        if logger:
            logger.info("Manager email sent")
    else:
        if logger:
            logger.info("Manager email skipped by configuration (EMAIL_MANAGER_REPORT = False)")

    

    # Step 4F: Send Full Order if needed
    full_link = full_created.get('webViewLink')
    if cfg.SEND_SEPARATE_FULL_ORDER_EMAIL:
        to_full = _fallback_recipients("FULL order", cfg.DEFAULT_ORDER_RECIPIENTS, cfg.TO_RECIPIENTS)
        msg = EmailMessage()
        msg["Subject"] = f"Order Report – {ts} – {location} – FULL"
        msg["From"] = "me"
        msg["To"] = ", ".join(to_full)
        if cfg.CC_RECIPIENTS:
            msg["Cc"] = ", ".join(cfg.CC_RECIPIENTS)
        msg.set_content(
            f"Hi team,\nFULL order report is ready.\nSheet: {full_link}\nAttached: {full_pdf_name}\n—Automated")
        msg.add_attachment(full_pdf, maintype="application", subtype="pdf", filename=full_pdf_name)
        send_email(gmail_svc, "me", msg)
        if logger:
            logger.info("FULL order email sent")
    else:
        if logger:
            logger.info("Separate full order email disabled")

    elapsed = int(time.perf_counter() - start)
    if logger:
        h = elapsed // 3600
        m = (elapsed % 3600) // 60
        s = elapsed % 60
        logger.info(f"Run completed in {h:02d}:{m:02d}:{s:02d}")

    return RunResult(True, elapsed, location, ts, manager_link, full_link)
