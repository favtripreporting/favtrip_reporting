from __future__ import annotations
import pandas
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
    get_value, first_gid,
    get_first_sheet_meta, get_values_2d, add_blank_sheet,
    add_or_replace_sheet, put_values_2d, _force_column_as_text, delete_row_indices, delete_rows_range, copy_sheet_to_another_spreadsheet
)
from .drive_utils import find_latest_sheet, upload_to_drive, _rfc3339, trash_file, cleanup_folder_by_age, find_sheet_by_name, copy_file_to_folder, rename_file
from .gmail_utils import send_email, email_manager_report

CSV_MIME = "text/csv"


def clean_tag(s: str) -> str:
    return re.sub(r"[^A-Za-z0-9._-]+", "-", s.strip()).strip("-") or "UNKNOWN"


import requests
from io import BytesIO
from openpyxl import Workbook


def export_sheet(creds, spreadsheet_id: str, gid: str | int, fmt: str) -> bytes:
    url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export?format={fmt}&gid={gid}"
    headers = {"Authorization": f"Bearer {creds.token}"}
    r = requests.get(url, headers=headers)
    r.raise_for_status()
    return r.content


def timestamp_now(tz: str, fmt: str) -> str:
    return datetime.now(ZoneInfo(tz)).strftime(fmt)

class IncomingDataValidationError(Exception):
    """Raised when the incoming report is not 1 or 2 full weeks as configured."""
    pass

_DOW_MAP = {
    "Monday": 0, "Tuesday": 1, "Wednesday": 2, "Thursday": 3,
    "Friday": 4, "Saturday": 5, "Sunday": 6, "Any": None,
}

def _parse_sheet_date(cell: str | int | float):
    """
    Parse a Google Sheets date cell into a date (drops time if present).
    Accepts:
      - Google serial numbers (days since 1899-12-30)
      - Date strings: YYYY-MM-DD, MM/DD/YYYY, MM/DD/YY
      - DateTime strings: with 12h or 24h time, with or without seconds, AM/PM
      - ISO strings (date or datetime)

    Returns: datetime.date or None if unparseable.
    """
    from datetime import datetime, timedelta

    if cell is None or cell == "":
        return None

    # --- 1) Numeric serial (Google Sheets) ---
    try:
        if isinstance(cell, (int, float)) or (isinstance(cell, str) and cell.replace(".", "", 1).isdigit()):
            serial = float(cell)
            base = datetime(1899, 12, 30)
            return (base + timedelta(days=serial)).date()
    except Exception:
        pass

    s = str(cell).strip()

    # Quick strip for weird whitespace
    s = " ".join(s.split())

    # If a timezone suffix or trailing text exists, try to isolate the datetime token
    # (We keep it simple: split on two spaces or take first token that contains '/')
    if " " in s and "/" in s:
        # Nothing fancy; the format tries below will accept the full string if they match
        pass

    # --- 2) Try common datetime formats (12h and 24h), with or without seconds ---
    dt_formats = [
        "%m/%d/%Y %I:%M:%S %p",  # 03/01/2026 12:03:45 AM
        "%m/%d/%Y %I:%M %p",     # 03/01/2026 12:03 AM
        "%m/%d/%Y %H:%M:%S",     # 03/01/2026 00:03:45
        "%m/%d/%Y %H:%M",        # 03/01/2026 00:03
        "%Y-%m-%d %H:%M:%S",     # 2026-03-01 00:03:45
        "%Y-%m-%d %H:%M",        # 2026-03-01 00:03
        "%Y-%m-%dT%H:%M:%S",     # 2026-03-01T00:03:45
        "%Y-%m-%dT%H:%M",        # 2026-03-01T00:03
    ]
    for fmt in dt_formats:
        try:
            return datetime.strptime(s, fmt).date()
        except Exception:
            continue

    # --- 3) Try date-only formats ---
    date_formats = [
        "%Y-%m-%d",   # 2026-03-01
        "%m/%d/%Y",   # 03/01/2026
        "%m/%d/%y",   # 03/01/26
    ]
    for fmt in date_formats:
        try:
            return datetime.strptime(s, fmt).date()
        except Exception:
            continue

    # --- 4) Last resort: Python ISO parser (handles 'YYYY-MM-DD' and full ISO datetimes) ---
    try:
        return datetime.fromisoformat(s).date()
    except Exception:
        pass

    # --- 5) If still not parsed, try taking only the date token before a space ---
    try:
        token = s.split(" ")[0]
        for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y"):
            try:
                return datetime.strptime(token, fmt).date()
            except Exception:
                continue
    except Exception:
        pass

    return None

def _find_header_and_date_col(values2d):
    """
    Find the header row whose first cell == 'Store', and the 'Date' column index.
    Returns (header_row_ix, date_col_ix) or (None, None).
    """
    header_ix = None
    for r, row in enumerate(values2d):
        c0 = (row[0].strip() if row and isinstance(row[0], str) else row[0] if row else "")
        if str(c0).strip().lower() == "store":
            header_ix = r
            break
    if header_ix is None:
        return None, None
    headers = [str(h).strip() for h in values2d[header_ix]]
    date_col_ix = None
    for c, h in enumerate(headers):
        if h.lower() == "date":
            date_col_ix = c
            break
    return header_ix, date_col_ix

def _collect_unique_dates(values2d, header_ix, date_cix):
    dates = []
    for r in range(header_ix + 1, len(values2d)):
        row = values2d[r]
        if date_cix >= len(row):
            continue
        d = _parse_sheet_date(row[date_cix])
        if d:
            dates.append(d)
    return sorted(set(dates))

def _check_week_boundaries(unique_dates, start_dow, end_dow):
    """Validate first/last weekday (unless set to Any). Return (earliest, latest)."""
    if not unique_dates:
        raise IncomingDataValidationError("No dates found in incoming report.")
    earliest, latest = unique_dates[0], unique_dates[-1]
    s_ok = (_DOW_MAP[start_dow] is None) or (earliest.weekday() == _DOW_MAP[start_dow])
    e_ok = (_DOW_MAP[end_dow]   is None) or (latest.weekday()   == _DOW_MAP[end_dow])
    error_text = None
    if not (s_ok and e_ok):
        error_text = f"Please only upload 1 or 2 full weeks of data. The first day of week included in the report should be {start_dow} and the last day of week included in the report should be {end_dow}"
        raise IncomingDataValidationError(
            error_text
        )
    return earliest, latest, error_text

def _plan_weeks(unique_dates):
    """
    Decide if we have one or two weeks by count of unique calendar days.
    Returns ('one', set7) or ('two', (set7_oldest, set7_newest)).
    """
    if len(unique_dates) == 7:
        return "one", set(unique_dates)
    if len(unique_dates) == 14:
        return "two", (set(unique_dates[:7]), set(unique_dates[7:]))
    # Not 7 or 14
    raise IncomingDataValidationError(
        "Please only upload 1 or 2 full weeks of data. The first day of week included in the report should be XXX and the last day of week included in the report should be YYY"
    )

def _trim_header_if_needed(svc, spreadsheet_id: str, sheet_id: int, values2d, header_ix):
    """Ensure header is at row 0 by deleting rows above it."""
    if header_ix and header_ix > 0:
        delete_rows_range(svc, spreadsheet_id, sheet_id, 0, header_ix)

def _filter_rows_to_dates(svc, spreadsheet_id: str, sheet_id: int, values2d, header_ix, date_cix, keep_dates_set):
    """Delete all non-header rows whose Date is not in keep_dates_set."""
    bad_rows = []
    for r in range(header_ix + 1, len(values2d)):
        row = values2d[r]
        d = _parse_sheet_date(row[date_cix] if date_cix < len(row) else None)
        if (d is None) or (d not in keep_dates_set):
            bad_rows.append(r)
    delete_row_indices(svc, spreadsheet_id, sheet_id, bad_rows)


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
    user_calc_sheet_id: str | None = None


def run_pipeline(cfg: Config, logger=None) -> RunResult:
    import time
    start = time.perf_counter()

    if logger:
        logger.info("Authorizing with Google APIs…")
    creds = get_credentials(cfg.SCOPES, cfg.REDIRECT_PORT, cfg.FORCE_REAUTH)
    sheets_svc, drive_svc, gmail_svc = services(creds, cfg.HTTP_TIMEOUT_SECONDS)
    if logger:
        logger.info("Google services ready")

    
    user_calc_sheet_id = None
    master_update_time = _parse_sheet_date(get_value(sheets_svc, cfg.CALC_SPREADSHEET_ID, cfg.LOCATION_SHEET_TITLE, cfg.TEMPLATE_UPDATE_RANGE))
    if logger:
        logger.info(f"Master update time: {master_update_time}")
    calc_ss_id = cfg.CALC_SPREADSHEET_ID  # default/fallback
    try:
        me = drive_svc.about().get(fields="user(emailAddress,permissionId,displayName)").execute().get("user", {})
        user_email = (me or {}).get("emailAddress") or "UNKNOWN_USER"
        # If you prefer a stable opaque id instead of email for file names:
        # user_id_for_name = (me or {}).get("permissionId") or user_email
        user_id_for_name = user_email

        if cfg.USER_FOLDER_ID:
            if logger:
                logger.info(f"Looking for per-user calc sheet in USER_FOLDER_ID for: {user_id_for_name}")
            found = find_sheet_by_name(drive_svc, cfg.USER_FOLDER_ID, user_id_for_name)
            if found:
                user_calc_sheet_id = found["id"]
                if logger:
                    logger.info(f"Found existing per-user workbook: {found.get('webViewLink')}")
                
                user_update_time = _parse_sheet_date(get_value(sheets_svc, user_calc_sheet_id, cfg.LOCATION_SHEET_TITLE, cfg.TEMPLATE_UPDATE_RANGE))
                if logger:
                    logger.info(f"User Update Time: {user_update_time}")

                if master_update_time > user_update_time:
                    if logger:
                        logger.info("Per-user workbook found but out of date; duplicating master into USER_FOLDER_ID…")
                    created = copy_file_to_folder(
                        drive_svc,
                        cfg.CALC_SPREADSHEET_ID,
                        cfg.USER_FOLDER_ID,
                        new_name=f"{user_id_for_name}_temp",
                    )
                    user_calc_sheet_id_temp = created["id"]
                    if logger:
                        logger.info(f"Created new per-user workbook: {created.get('webViewLink')}")

                    delete_sheet(sheets_svc, user_calc_sheet_id_temp, "Current Week")
                    delete_sheet(sheets_svc, user_calc_sheet_id_temp, "Last Week")

                    if logger:
                        logger.info(f"Deleted data sheets in new user file.")

                    copy_sheet_to_another_spreadsheet(sheets_svc, user_calc_sheet_id, "Current Week", user_calc_sheet_id_temp, "Current Week")
                    copy_sheet_to_another_spreadsheet(sheets_svc, user_calc_sheet_id, "Last Week", user_calc_sheet_id_temp, "Last Week")

                    if logger:
                        logger.info(f"Copied old data sheets to new user file.")

                    trash_file(drive_svc, user_calc_sheet_id)

                    if logger:
                        logger.info(f"Deleted old user file.")

                    rename_file(drive_svc, user_calc_sheet_id_temp, user_id_for_name)

                    if logger:
                        logger.info(f"Renamed new user file for continued use.")
                    
                    user_calc_sheet_id = user_calc_sheet_id_temp

            else:
                if logger:
                    logger.info("No per-user workbook found; duplicating master into USER_FOLDER_ID…")
                created = copy_file_to_folder(
                    drive_svc,
                    cfg.CALC_SPREADSHEET_ID,
                    cfg.USER_FOLDER_ID,
                    new_name=user_id_for_name,
                )
                user_calc_sheet_id = created["id"]
                if logger:
                    logger.info(f"Created per-user workbook: {created.get('webViewLink')}")

            # From here on, operate on the per-user workbook
            calc_ss_id = user_calc_sheet_id
        else:
            if logger:
                logger.info("USER_FOLDER_ID not configured; using CALC_SPREADSHEET_ID directly.")
    except Exception as e:
        if logger:
            logger.warn(f"Could not resolve per-user workbook (continuing with CALC_SPREADSHEET_ID): {e}")
    

    # Step 1: latest incoming
    if logger:
        logger.info("Finding latest incoming spreadsheet…")
    latest = find_latest_sheet(drive_svc, cfg.INCOMING_FOLDER_ID)
    if not latest:
        raise SystemExit("No incoming report found.")
    new_report_id = latest["id"]
    if logger:
        logger.info(f"Latest incoming: {latest['name']} ({new_report_id})")

    # ---- NEW: Validate incoming weeks & plan actions (no workbook changes yet) ----
    if logger:
        logger.info("Validating incoming report (header, dates, week boundaries)…")
    first_title, first_sid = get_first_sheet_meta(sheets_svc, new_report_id)
    values = get_values_2d(sheets_svc, new_report_id, first_title, "A:Z")

    h_ix, d_cix = _find_header_and_date_col(values)
    if h_ix is None or d_cix is None:
        raise IncomingDataValidationError(
            "Unable to locate header ('Store' in A1) and/or 'Date' column in the incoming report."
        )

    unique_dates = _collect_unique_dates(values, h_ix, d_cix)

    if logger:
        logger.info(f"Found {len(unique_dates)} unique date(s) in incoming report")

    check_outputs = _check_week_boundaries(unique_dates, cfg.START_DAY_OF_WEEK, cfg.END_DAY_OF_WEEK)
    plan_kind, plan_payload = _plan_weeks(unique_dates)

    # Step 2: prep calculations workbook (branch by plan)
    if logger:
        logger.info("Preparing calculations workbook…")

    if plan_kind == "two":
        # Two weeks → build values in memory and write each in a single call
        if logger:
            logger.info("Detected 2 weeks; writing 'Last Week' (oldest 7) and 'Current Week' (newest 7) without row deletions")

        # Source header & body (we already loaded 'values' from the first sheet)
        header = [str(h) for h in values[h_ix]]
        body_rows = values[h_ix + 1 :]

        def _slice_rows(rows, date_cix, keep_dates: set):
            out = []
            for row in rows:
                d = _parse_sheet_date(row[date_cix] if date_cix < len(row) else None)
                if d and d in keep_dates:
                    out.append(row)
            return out

        keep_oldest7, keep_newest7 = plan_payload  # sets of dates from _plan_weeks
        last_week_rows = _slice_rows(body_rows, d_cix, keep_oldest7)
        current_week_rows = _slice_rows(body_rows, d_cix, keep_newest7)

        # Create fresh target sheets
        add_or_replace_sheet(sheets_svc, calc_ss_id, "Last Week")
        add_or_replace_sheet(sheets_svc, calc_ss_id, "Current Week")

        # Force column 'Scan Code' to be text with a prefixed apostrophe
        last_week_rows = _force_column_as_text(header, last_week_rows, "Scan Code")
        current_week_rows = _force_column_as_text(header, current_week_rows, "Scan Code")

        # Bulk write (header + rows) → 1 write per sheet
        put_values_2d(sheets_svc, calc_ss_id, "Last Week", [header] + last_week_rows)
        put_values_2d(sheets_svc, calc_ss_id, "Current Week", [header] + current_week_rows)

    elif plan_kind == "one" and cfg.USE_AUTO_ROLLOVER_IF_ONE_WEEK:
        # One week + rollover ON → current behavior
        if logger:
            logger.info("Detected 1 week; auto-rollover enabled → copying old Current→Last and inserting new Current")
        delete_sheet(sheets_svc, calc_ss_id, "Last Week")
        try:
            copy_sheet_as(sheets_svc, calc_ss_id, "Current Week", "Last Week")
            if logger:
                logger.info("Copied old 'Current Week' to 'Last Week'")
        except Exception:
            if logger:
                logger.warn("No 'Current Week' sheet exists to copy")
        delete_sheet(sheets_svc, calc_ss_id, "Current Week")
        copy_first_sheet_as(sheets_svc, new_report_id, calc_ss_id, "Current Week")

        # Trim header for Current Week
        meta = sheets_svc.spreadsheets().get(spreadsheetId=calc_ss_id).execute()
        cw_sid = next(s["properties"]["sheetId"] for s in meta["sheets"] if s["properties"]["title"] == "Current Week")
        _trim_header_if_needed(sheets_svc, calc_ss_id, cw_sid, values, h_ix)

    else:
        # One week + rollover OFF → Current Week only; Last Week blank
        if logger:
            logger.info("Detected 1 week; auto-rollover disabled → Current only, Last Week blank")
        delete_sheet(sheets_svc, calc_ss_id, "Last Week")
        delete_sheet(sheets_svc, calc_ss_id, "Current Week")
        add_blank_sheet(sheets_svc, calc_ss_id, "Last Week")
        copy_first_sheet_as(sheets_svc, new_report_id, calc_ss_id, "Current Week")

        meta = sheets_svc.spreadsheets().get(spreadsheetId=calc_ss_id).execute()
        cw_sid = next(s["properties"]["sheetId"] for s in meta["sheets"] if s["properties"]["title"] == "Current Week")
        _trim_header_if_needed(sheets_svc, calc_ss_id, cw_sid, values, h_ix)

    # Refresh reference sheets (unchanged)
    if logger:
        logger.info("Refreshing reference sheets (prefix 'REFR: ')…")
    refresh_sheets_with_prefix(sheets_svc, calc_ss_id, prefix="REFR: ", logger=logger)

    # Step 3: read location code
    location = get_value(sheets_svc, calc_ss_id, cfg.LOCATION_SHEET_TITLE, cfg.LOCATION_NAMED_RANGE)
    ts = timestamp_now(cfg.TIMESTAMP_TZ, cfg.TIMESTAMP_FMT)
    if logger:
        logger.info(f"Location: {location}; Timestamp: {ts}")

    # Step 4A: Manager Report PDF
    if logger:
        logger.info("Exporting Manager Report (PDF)…")
    pdf_bytes = export_sheet(creds, calc_ss_id, cfg.GID_MANAGER_PDF, "pdf")
    pdf_name = f"Manager_Report_{ts}_{location}.pdf"
    uploaded_pdf = upload_to_drive(drive_svc, pdf_bytes, pdf_name, "application/pdf", cfg.MANAGER_REPORT_FOLDER_ID, to_sheet=False)
    manager_link = uploaded_pdf.get("webViewLink")
    if logger:
        logger.info(f"Uploaded Manager PDF: {manager_link}")

    # Step 4B: Master Order CSV
    if logger:
        logger.info("Exporting Master Order (CSV)…")
    master_csv_bytes = export_sheet(creds, calc_ss_id, cfg.GID_ORDER_CSV, "csv")

    # Step 4C: Full order upload (CSV) and export (PDF)
    full_csv_name = f"Order_Report_{ts}_{location}_FULL.csv"
    full_created = upload_to_drive(drive_svc, master_csv_bytes, full_csv_name, CSV_MIME, cfg.ORDER_REPORT_FOLDER_ID, to_sheet=True)
    full_file_id = full_created["id"]
    full_gid = first_gid(sheets_svc, full_file_id)
    full_pdf = export_sheet(creds, full_file_id, full_gid, "pdf")
    full_pdf_name = f"Order_Report_{ts}_{location}_FULL.pdf"
    if logger:
        logger.info(f"Uploaded FULL sheet: {full_created.get('webViewLink')}")

    # Step 4D: Create per-report-key outputs (CSV) and email

    # --- Parse the master CSV into rows of dicts ---
    
    text = master_csv_bytes.decode("utf-8-sig", errors="replace")
    reader = csv.reader(io.StringIO(text))
    
    rows_list = list(reader)
    if not rows_list:
        raise RuntimeError("CSV has no rows.")
    
    headers = [h.strip() for h in rows_list[0]]
    if not headers:
        raise RuntimeError("CSV has no header.")
    
    # Find required columns (case-insensitive)
    lower_idx = {h.lower(): i for i, h in enumerate(headers)}
    
    if "report_key" not in lower_idx:
        raise RuntimeError("Report_Key column missing.")
    if "store" not in lower_idx:
        raise RuntimeError("Store column missing.")
    
    report_idx = lower_idx["report_key"]
    store_idx = lower_idx["store"]
    
    # Headers to export (exclude report_key)
    export_headers = [h for i, h in enumerate(headers) if i != report_idx]
    
    # Materialize rows as list[dict]
    rows = []
    for row in rows_list[1:]:
        rows.append({headers[i]: (row[i] if i < len(row) else "") for i in range(len(headers))})
    
    # Group by (report_key, store)
    groups: dict[tuple[str, str], list[dict]] = {}
    for r in rows:
        report_key = (str(r.get(headers[report_idx]) or "").strip()) or "UNASSIGNED"
        store = (str(r.get(headers[store_idx]) or "").strip()) or "UNKNOWN"
        groups.setdefault((store.upper(), report_key.upper()), []).append(r)
    
    
    for (store, key), key_rows in groups.items():
    
        if not cfg.USE_ALL_REPORT_KEYS and key not in (cfg.REPORT_KEY_RUN_LIST or []):
            continue
    
        # Build CSV text in memory
        sio = io.StringIO()
        w = csv.writer(sio, lineterminator="\n")
    
        w.writerow(export_headers)
    
        for rr in key_rows:
            w.writerow([rr.get(h, "") for h in export_headers])
    
        key_csv_bytes = sio.getvalue().encode("utf-8")
    
        tag = clean_tag(key)
        store_tag = clean_tag(store)
    
        csv_name = f"Order_Report_{ts}_{location}_{tag}_{store_tag}.csv"
    
        # Upload CSV to Drive; conversion to Google Sheet happens via to_sheet=True
        created = upload_to_drive(
            drive_svc, key_csv_bytes, csv_name,
            CSV_MIME, cfg.ORDER_REPORT_FOLDER_ID, to_sheet=True
        )
    
        file_id = created["id"]
        gid = first_gid(sheets_svc, file_id)
    
        # Export the Google Sheet as PDF
        pdf = export_sheet(creds, file_id, gid, "pdf")
        pdfname = f"Order_Report_{ts}_{location}_{tag}_{store_tag}.pdf"
    
        # Prefer Store+Key; else Key; else Store; else To; else Default
        candidates = None
        if cfg.REPORT_KEY_RECIPIENTS:
            store_key = (store_tag, tag)
            key_only = (None, tag)
            store_only = (store_tag, None)
        
            if store_key in cfg.REPORT_KEY_RECIPIENTS:
                candidates = cfg.REPORT_KEY_RECIPIENTS[store_key]
            elif key_only in cfg.REPORT_KEY_RECIPIENTS:
                candidates = cfg.REPORT_KEY_RECIPIENTS[key_only]
            elif store_only in cfg.REPORT_KEY_RECIPIENTS:
                candidates = cfg.REPORT_KEY_RECIPIENTS[store_only]
    
        recipients = _fallback_recipients(
            f"REPORT_KEY {tag}",
            candidates,
            cfg.TO_RECIPIENTS,
            cfg.DEFAULT_ORDER_RECIPIENTS
        )
    
        msg = EmailMessage()
        msg["Subject"] = f"Order Report – {ts} – {location} – {tag} – {store}"
        msg["From"] = "me"
        msg["To"] = ", ".join(recipients)
    
        if cfg.CC_RECIPIENTS:
            msg["Cc"] = ", ".join(cfg.CC_RECIPIENTS)
    
        msg.set_content(
            f"Hi {key} team,\nYour order report for store {store} is ready.\n"
            f"Google Sheet: {created.get('webViewLink')}\n"
            f"Attached: {pdfname}\n—Automated"
        )
    
        msg.add_attachment(pdf, maintype="application", subtype="pdf", filename=pdfname)
    
        if cfg.INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL:
            msg.add_attachment(full_pdf, maintype="application", subtype="pdf", filename=full_pdf_name)
    
        send_email(gmail_svc, "me", msg)
    
        if logger:
            logger.info(f"Emailed {tag} - {store}")
    
        
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
        to_full = _fallback_recipients("FULL order", cfg.TO_RECIPIENTS, cfg.DEFAULT_ORDER_RECIPIENTS)
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

    
    # Step 4G: File Cleanup

    try:
        if logger:
            logger.info("Cleaning up used incoming file…")
        trash_file(drive_svc, new_report_id)

        if logger:
            logger.info("Cleaning old incoming files…")
        cleanup_folder_by_age(
            drive_svc,
            cfg.INCOMING_FOLDER_ID,
            cfg.FAILED_INPUT_TIME_TO_LIFE,
            logger
        )

        if logger:
            logger.info("Cleaning old output files…")
        for folder in [
            cfg.MANAGER_REPORT_FOLDER_ID,
            cfg.ORDER_REPORT_FOLDER_ID
        ]:
            cleanup_folder_by_age(
                drive_svc,
                folder,
                cfg.OUTPUT_TIME_TO_LIFE,
                logger
            )
        
        if logger:
            logger.info("Cleaning old calculation files…")
            cleanup_folder_by_age(
                drive_svc,
                cfg.USER_FOLDER_ID,
                cfg.USER_TIME_TO_LIFE,
                logger
            )

    except Exception as e:
        if logger:
            logger.warn(f"Housekeeping failed: {e}")

    elapsed = int(time.perf_counter() - start)
    if logger:
        h = elapsed // 3600
        m = (elapsed % 3600) // 60
        s = elapsed % 60
        logger.info(f"Run completed in {h:02d}:{m:02d}:{s:02d}")

    return RunResult(True, elapsed, location, ts, manager_link, full_link)
