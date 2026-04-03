"""
Pipeline
======================================

Overview
--------
This is the main workhorse file that the user interface runs. This pipeline automates a weekly reporting workflow around Google Workspace
(Drive, Sheets, and Gmail) for store ordering. At a high level it:

1. Authenticates to Google APIs and locates the latest incoming spreadsheet in
   a designated Drive folder.
2. Validates the data contains **one or two full weeks** of daily records and
   that the first/last days match your configured week boundaries.
3. Prepares (or rolls) a per-user **Calculations** workbook, then populates the
   **Current Week** and (optionally) **Last Week** sheets using the incoming
   data.
4. Refreshes reference sheets by prefix (e.g., `REFR: `, `REFC: `).
5. Exports and uploads:
   - Manager report (**PDF**)
   - Full order (**CSV** → Google Sheet) and a **PDF** rendition
   - Per **report key** CSVs (converted to Sheets) and their PDFs
6. Emails the manager report and per-report-key packages to the appropriate
   recipients (with configurable CCs and an option to include the Full order PDF
   in each email).
7. Performs Drive housekeeping (trash the consumed incoming file and prune old
   items from configured folders).

Key Components
--------------
- **Configuration (`Config`)**: Centralizes IDs, options, and behavior toggles
  consumed throughout the pipeline (folder IDs, spreadsheet IDs, GIDs, named
  ranges, week boundary settings, time-to-live values, and email recipient
  settings).
- **Google Clients**: `get_credentials()` and `services()` establish authorized
  clients for Sheets, Drive, and Gmail using the configured scopes and timeouts.
- **Sheets Utilities**: Helpers to copy, add, delete, and write sheets; retrieve
  values; and coerce specific columns as text (e.g., `Scan Code`).
- **Drive Utilities**: Locate the latest file, upload byte content as Drive
  files (with optional conversion to Sheets), rename, copy between folders,
  trash, and clean folders by age.
- **Gmail Utilities**: Compose and send emails with attachments and Drive links.

Validation & Planning
---------------------
The pipeline inspects the first tab of the incoming report and:
- Locates the header row where the first cell equals **"Store"** and the
  **Date** column.
- Parses dates (string, serial, ISO) and collects the unique calendar days.
- Ensures the first and last dates align with configured week boundaries
  (e.g., Monday–Sunday), raising `IncomingDataValidationError` if not.
- Determines whether the upload covers **one** or **two** weeks (7 or 14 unique
  days) and plans sheet operations accordingly.

Per‑User Workbook Behavior
--------------------------
If `USER_FOLDER_ID` is set, the pipeline attempts to locate (by the user's
email) a dedicated Calculations workbook in that folder; if absent or outdated
compared to the master template, it duplicates/refreshes it while preserving the
`Current Week` and `Last Week` data tabs from the user's prior workbook.

Email Routing & Fallbacks
-------------------------
Recipients are selected in the following order (first non-empty wins):
1. A store+report‑key specific list (from `REPORT_KEY_RECIPIENTS`), then
   key‑only, then store‑only
2. `TO_RECIPIENTS`
3. `DEFAULT_ORDER_RECIPIENTS`

Invalid emails and stray commas are sanitized. Missing recipients lead to a
friendly `ValueError` that explains how to supply valid addresses.

"""


from __future__ import annotations
import pandas
import csv
import io
import re
import requests
import time
from dataclasses import dataclass
from datetime import datetime
from zoneinfo import ZoneInfo
from email.message import EmailMessage

from io import BytesIO
from openpyxl import load_workbook, Workbook


from .config import Config
from .google_client import get_credentials, services
from .sheets_utils import (
    delete_sheet, copy_sheet_as, copy_first_sheet_as, refresh_sheets_with_prefix, refresh_sheets_with_prefix_chunked,
    get_value, first_gid,
    get_first_sheet_meta, get_values_2d, add_blank_sheet,
    add_or_replace_sheet, put_values_2d, _force_column_as_text, delete_row_indices, delete_rows_range, copy_sheet_to_another_spreadsheet
)
from .drive_utils import find_latest_sheet, upload_to_drive, _rfc3339, trash_file, cleanup_folder_by_age, find_sheet_by_name, copy_file_to_folder, rename_file, get_or_create_subfolder
from .gmail_utils import send_email, email_manager_report, email_order_report, email_error_report

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

class VendorPriceBookError(Exception):
    """Raised when one or more items to be ordered are not found on the Vendor Price Book."""
    pass

_DOW_MAP = {
    "Monday": 0, "Tuesday": 1, "Wednesday": 2, "Thursday": 3,
    "Friday": 4, "Saturday": 5, "Sunday": 6, "Any": None,
}

def _parse_sheet_date(cell: str | int | float, include_time: bool = False) -> datetime | date | None:
    """
    Parse a Google Sheets date/time cell.

    Args:
        cell: Google Sheets date value (serial number, date string, datetime string, or ISO string)
        include_time: If True, return datetime with time. If False (default), return date only.

    Returns:
        datetime.datetime (if include_time=True) or datetime.date (if include_time=False), or None if unparseable.
    """

    if cell is None or cell == "":
        return None

    # --- 1) Numeric serial (Google Sheets) ---
    try:
        if isinstance(cell, (int, float)) or (isinstance(cell, str) and cell.replace(".", "", 1).isdigit()):
            serial = float(cell)
            base = datetime(1899, 12, 30)
            dt = base + timedelta(days=serial)
            return dt if include_time else dt.date()
    except Exception:
        pass

    s = str(cell).strip()
    s = " ".join(s.split())  # remove extra whitespace

    # --- 2) Try common datetime formats (with time) ---
    dt_formats = [
        "%m/%d/%Y %I:%M:%S %p",
        "%m/%d/%Y %I:%M %p",
        "%m/%d/%Y %H:%M:%S",
        "%m/%d/%Y %H:%M",
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M",
        "%Y-%m-%dT%H:%M:%S",
        "%Y-%m-%dT%H:%M",
    ]
    for fmt in dt_formats:
        try:
            dt = datetime.strptime(s, fmt)
            return dt if include_time else dt.date()
        except Exception:
            continue

    # --- 3) Try date-only formats ---
    date_formats = ["%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y"]
    for fmt in date_formats:
        try:
            dt = datetime.strptime(s, fmt)
            return dt if include_time else dt.date()
        except Exception:
            continue

    # --- 4) ISO format fallback ---
    try:
        dt = datetime.fromisoformat(s)
        return dt if include_time else dt.date()
    except Exception:
        pass

    # --- 5) Last resort: first token before space ---
    try:
        token = s.split(" ")[0]
        for fmt in date_formats:
            try:
                dt = datetime.strptime(token, fmt)
                return dt if include_time else dt.date()
            except Exception:
                continue
    except Exception:
        pass

    return None

def _find_header_and_date_col(values2d, firstheader, col=""):
    """
    Find the header row whose first cell == 'Store', and the 'Date' column index.
    Returns (header_row_ix, date_col_ix) or (None, None).
    """
    header_ix = None
    for r, row in enumerate(values2d):
        c0 = (row[0].strip() if row and isinstance(row[0], str) else row[0] if row else "")
        if str(c0).strip().lower() == str(firstheader).strip().lower():
            header_ix = r
            break
    if header_ix is None:
        return None, None
    headers = [str(h).strip() for h in values2d[header_ix]]
    date_col_ix = None
    for c, h in enumerate(headers):
        if h.lower() == str(col).strip().lower():
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


def csv_has_data_rows(csv_bytes: bytes) -> bool:
    if not csv_bytes:
        return False

    text = csv_bytes.decode("utf-8-sig")  # handles BOM if present
    reader = csv.reader(io.StringIO(text))

    rows = list(reader)

    # More than just the header row
    return len(rows) > 1



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

def should_run(cfg, report_key, sub_key):
    allowed = set(cfg.REPORT_KEY_RUN_LIST or [])

    fmt_sub_key = f"{report_key}-{sub_key}"

    if cfg.USE_ALL_REPORT_KEYS:
        return True

    # explicit sub-report key
    if sub_key:
        if sub_key in allowed:
            return True
        if fmt_sub_key in allowed:
            return True
        if report_key in allowed:
            return True
        return False

    # no sub key
    return report_key in allowed


@dataclass
class RunResult:
    ok: bool
    elapsed_seconds: int
    location: str
    timestamp: str
    manager_pdf_link: str | None
    full_order_link: str | None
    user_calc_sheet_id: str | None = None    
    err_exist: bool = False
    err_link: str | None = None



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
    master_update_time = _parse_sheet_date(get_value(sheets_svc, cfg.CALC_SPREADSHEET_ID, cfg.LOCATION_SHEET_TITLE, cfg.TEMPLATE_UPDATE_RANGE), True)
    if logger:
        logger.info(f"Master update time: {master_update_time}")
    calc_ss_id = cfg.CALC_SPREADSHEET_ID  # default/fallback
    user_sales_folder_id = None
    try:
        me = drive_svc.about().get(fields="user(emailAddress,permissionId,displayName)").execute().get("user", {})
        user_email = (me or {}).get("emailAddress") or "UNKNOWN_USER"
        # If you prefer a stable opaque id instead of email for file names:
        # user_id_for_name = (me or {}).get("permissionId") or user_email
        user_id_for_name = user_email

        # Resolve per-user Incoming subfolder
        if logger:
            logger.info(f"Resolving per-user incoming folder for {user_id_for_name}")

        incoming_folder = get_or_create_subfolder(
            drive_svc,
            cfg.INCOMING_FOLDER_ID,
            user_id_for_name
        )

        user_level_folder_id = incoming_folder["id"]
        user_sales_folder_id = get_or_create_subfolder(drive_svc, user_level_folder_id, "01 Sales Data Inputs")["id"]
        user_vendor_folder_id = get_or_create_subfolder(drive_svc, user_level_folder_id, "02 Vendor Price Data Inputs")["id"]


        if logger:
            logger.info(
                f"Using incoming folder: {incoming_folder.get('webViewLink')}"
            )

        if cfg.USER_FOLDER_ID:
            if logger:
                logger.info(
                    f"Looking for per-user calc sheet in {cfg.USER_FOLDER_ID} for: {user_id_for_name}"
                )

            found = find_sheet_by_name(
                drive_svc,
                cfg.USER_FOLDER_ID,
                user_id_for_name
            )

            if found:
                user_calc_sheet_id = found["id"]
                if logger:
                    logger.info(f"Found existing per-user workbook: {found.get('webViewLink')}")
                
                user_update_time = _parse_sheet_date(get_value(sheets_svc, user_calc_sheet_id, cfg.LOCATION_SHEET_TITLE, cfg.TEMPLATE_UPDATE_RANGE), True)
                if logger:
                    logger.info(f"User Update Time: {user_update_time}")

                if master_update_time > user_update_time:
                    if logger:
                        logger.info(f"Per-user workbook found but out of date; duplicating master into {cfg.USER_FOLDER_ID}…")
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
                    logger.info(f"No per-user workbook found; duplicating master into {cfg.USER_FOLDER_ID}…")
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
                logger.info(f"USER_FOLDER_ID not configured; using {cfg.CALC_SPREADSHEET_ID} directly.")
    except Exception as e:
        if logger:
            logger.warn(f"Could not resolve per-user workbook (continuing with {cfg.CALC_SPREADSHEET_ID}): {e}")
    
    # Fallback: if per-user incoming folder could not be resolved,
    # use the shared incoming folder
    if not user_sales_folder_id:
        if logger:
            logger.warn(
                "Per-user incoming folder not resolved; "
                f"falling back to shared {cfg.INCOMING_FOLDER_ID}"
            )
        user_sales_folder_id = cfg.INCOMING_FOLDER_ID

    # Step 1: latest incoming
    if logger:
        logger.info(f"Finding latest incoming sales spreadsheet in {user_sales_folder_id}…")

    latest_sales = None
    n = 10
    for attempt in range(n):
        latest_sales = find_latest_sheet(drive_svc, user_sales_folder_id)
        if latest_sales:
            break

        if logger:
            logger.info(
                f"No incoming sheet in {user_sales_folder_id} yet (attempt {attempt + 1}/{n}); retrying..."
            )
        time.sleep(2)

    if not latest_sales:
        raise SystemExit(
            "No incoming sales report found in per-user incoming folder."
        )
    

    if logger:
        logger.info(f"Finding latest incoming vendor spreadsheet in {user_vendor_folder_id}…")

    latest_vendor = None
    n = 10
    for attempt in range(n):
        latest_vendor = find_latest_sheet(drive_svc, user_vendor_folder_id)
        if latest_vendor:
            break

        if logger:
            logger.info(
                f"No incoming sheet in {user_vendor_folder_id} yet (attempt {attempt + 1}/{n}); retrying..."
            )
        time.sleep(2)

    if not latest_vendor:
        raise SystemExit(
            "No incoming vendor report found in per-user incoming folder."
        )
    
    new_sales_report_id = latest_sales["id"]
    new_vendor_report_id = latest_vendor["id"]

    # ---- NEW: Validate incoming weeks & plan actions (no workbook changes yet) ----
    if logger:
        logger.info("Validating incoming report (header, dates, week boundaries)…")
    sales_first_title, sales_first_sid = get_first_sheet_meta(sheets_svc, new_sales_report_id)
    sales_values = get_values_2d(sheets_svc, new_sales_report_id, sales_first_title, "A:Z")

    vendor_first_title, vendor_first_sid = get_first_sheet_meta(sheets_svc, new_vendor_report_id)
    vendor_values = get_values_2d(sheets_svc, new_vendor_report_id, vendor_first_title, "A:Z")

    sales_h_ix, sales_d_cix = _find_header_and_date_col(sales_values, 'Store', 'Date')
    if sales_h_ix is None or sales_d_cix is None:
        raise IncomingDataValidationError(
            "Unable to locate header ('Store' in A1) and/or 'Date' column in the incoming sales report."
        )
    
    vendor_h_ix, vendor_d_cix = _find_header_and_date_col(vendor_values, 'Scan Code', 'Scan Code')
    if vendor_h_ix is None:
        raise IncomingDataValidationError(
            "Unable to locate header ('Scan Code' in A1) in the incoming vendor price book report."
        )

    unique_dates = _collect_unique_dates(sales_values, sales_h_ix, sales_d_cix)

    if logger:
        logger.info(f"Found {len(unique_dates)} unique date(s) in incoming report")

    check_outputs = _check_week_boundaries(unique_dates, cfg.START_DAY_OF_WEEK, cfg.END_DAY_OF_WEEK)
    plan_kind, plan_payload = _plan_weeks(unique_dates)

    # Step 2: prep calculations workbook (branch by plan)
    if logger:
        logger.info("Preparing calculations workbook…")

    # Source header & body (we already loaded 'values' from the first sheet)
    sales_header = [str(h) for h in sales_values[sales_h_ix]]
    sales_body_rows = sales_values[sales_h_ix + 1 :]

    vendor_header = [str(h) for h in vendor_values[vendor_h_ix]]
    vendor_body_rows = vendor_values[vendor_h_ix + 1 :]
    
    if plan_kind == "two":
        # Two weeks → build values in memory and write each in a single call
        if logger:
            logger.info("Detected 2 weeks; writing 'Last Week' (oldest 7) and 'Current Week' (newest 7) without row deletions")

        def _slice_rows(rows, date_cix, keep_dates: set):
            out = []
            for row in rows:
                d = _parse_sheet_date(row[date_cix] if date_cix < len(row) else None)
                if d and d in keep_dates:
                    out.append(row)
            return out

        keep_oldest7, keep_newest7 = plan_payload  # sets of dates from _plan_weeks
        last_week_rows = _slice_rows(sales_body_rows, sales_d_cix, keep_oldest7)
        current_week_rows = _slice_rows(sales_body_rows, sales_d_cix, keep_newest7)

        # Create fresh target sheets
        add_or_replace_sheet(sheets_svc, calc_ss_id, "Last Week")
        add_or_replace_sheet(sheets_svc, calc_ss_id, "Current Week")
        add_or_replace_sheet(sheets_svc, calc_ss_id, "Vendor Price Book")

        # Force column 'Scan Code' to be text with a prefixed apostrophe
        last_week_rows = _force_column_as_text(sales_header, last_week_rows, "Scan Code")
        current_week_rows = _force_column_as_text(sales_header, current_week_rows, "Scan Code")
        vendor_body_rows = _force_column_as_text(vendor_header, vendor_body_rows, "Scan Code")

        # Bulk write (header + rows) → 1 write per sheet
        put_values_2d(sheets_svc, calc_ss_id, "Last Week", [sales_header] + last_week_rows)
        put_values_2d(sheets_svc, calc_ss_id, "Current Week", [sales_header] + current_week_rows)
        put_values_2d(sheets_svc, calc_ss_id, "Vendor Price Book", [vendor_header] + vendor_body_rows)

    elif plan_kind == "one" and cfg.USE_AUTO_ROLLOVER_IF_ONE_WEEK:
        # One week + rollover ON → current behavior
        if logger:
            logger.info("Detected 1 week; auto-rollover enabled → copying old Current→Last and inserting new Current")

        delete_sheet(sheets_svc, calc_ss_id, "Last Week")
        add_or_replace_sheet(sheets_svc, calc_ss_id, "Vendor Price Book")

        try:
            copy_sheet_as(sheets_svc, calc_ss_id, "Current Week", "Last Week")
            if logger:
                logger.info("Copied old 'Current Week' to 'Last Week'")
        except Exception:
            if logger:
                logger.warn("No 'Current Week' sheet exists to copy")
        
        add_or_replace_sheet(sheets_svc, calc_ss_id, "Current Week")

        current_week_rows = _force_column_as_text(sales_header, sales_body_rows, "Scan Code")
        vendor_body_rows = _force_column_as_text(vendor_header, vendor_body_rows, "Scan Code")

        put_values_2d(sheets_svc, calc_ss_id, "Current Week", [sales_header] + current_week_rows)
        put_values_2d(sheets_svc, calc_ss_id, "Vendor Price Book", [vendor_header] + vendor_body_rows)

        # Trim header for Current Week
        meta = sheets_svc.spreadsheets().get(spreadsheetId=calc_ss_id).execute()
        cw_sid = next(s["properties"]["sheetId"] for s in meta["sheets"] if s["properties"]["title"] == "Current Week")
        _trim_header_if_needed(sheets_svc, calc_ss_id, cw_sid, sales_values, sales_h_ix)

    else:
        # One week + rollover OFF → Current Week only; Last Week blank
        if logger:
            logger.info("Detected 1 week; auto-rollover disabled → Current only, Last Week blank")
        
        add_or_replace_sheet(sheets_svc, calc_ss_id, 'Last Week')
        add_or_replace_sheet(sheets_svc, calc_ss_id, 'Current Week')
        add_or_replace_sheet(sheets_svc, calc_ss_id, 'Vendor Price Book')

        current_week_rows = _force_column_as_text(sales_header, sales_body_rows, "Scan Code")
        vendor_body_rows = _force_column_as_text(vendor_header, vendor_body_rows, "Scan Code")

        put_values_2d(sheets_svc, calc_ss_id, "Current Week", [sales_header] + current_week_rows)
        put_values_2d(sheets_svc, calc_ss_id, "Vendor Price Book", [vendor_header] + vendor_body_rows)

        meta = sheets_svc.spreadsheets().get(spreadsheetId=calc_ss_id).execute()
        cw_sid = next(s["properties"]["sheetId"] for s in meta["sheets"] if s["properties"]["title"] == "Current Week")
        _trim_header_if_needed(sheets_svc, calc_ss_id, cw_sid, sales_values, sales_h_ix)

    # Refresh reference sheets (unchanged)
    if logger:
        logger.info("Refreshing reference sheets (prefix 'REFR: ' or 'REFC ')…")
        
    refresh_sheets_with_prefix(sheets_svc, calc_ss_id, prefix = "REFA: ", logger=logger)

    time.sleep(5)

    refresh_sheets_with_prefix(sheets_svc, calc_ss_id, prefix = "REFR: ", logger=logger)
    
    refresh_sheets_with_prefix_chunked(
        sheets_svc,
        calc_ss_id,
        prefix = "REFC: ",
        logger=logger
    )

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

    # Step 4C: Error Report CSV, Upload, Export PDF
    if logger:
        logger.info("Exporting Error Report (CSV)…")

    err_csv_bytes = export_sheet(creds, calc_ss_id, cfg.GID_ERROR_REPORT, "csv")

    err_text = err_csv_bytes.decode("utf-8-sig", errors="replace")
    reader = csv.reader(io.StringIO(err_text))
    rows_list = list(reader)

    if not rows_list or len(rows_list) <= 1:
        err_exist = False
    else:
        headers = [h.strip() for h in rows_list[0]]
        lower_idx = {h.lower(): i for i, h in enumerate(headers)}
        sub_idx = lower_idx.get("sub_report_key")

        if "report_key" not in lower_idx:
            raise RuntimeError("Error report missing Report_Key column")

        report_idx = lower_idx["report_key"]
    
    if cfg.USE_ALL_REPORT_KEYS:
        allowed_keys = None  # no filtering
    else:
        allowed_keys = {k.upper() for k in (cfg.REPORT_KEY_RUN_LIST or [])}

    filtered_err_rows = []

    for r in rows_list[1:]:
        key = (r[report_idx] if report_idx < len(r) else "").strip().upper()

        if not key:
            continue

        if allowed_keys is None or key in allowed_keys:
            filtered_err_rows.append(r)

    err_exist = bool(filtered_err_rows)

    err_link = None

    if err_exist:
        err_csv_name = f"Error_Report_{ts}.csv"

        
        output = io.StringIO()
        writer = csv.writer(output)
        writer.writerow(rows_list[0])
        writer.writerows(filtered_err_rows)

        filtered_err_csv_bytes = output.getvalue().encode("utf-8-sig")

        err_created = upload_to_drive(
            drive_svc,
            filtered_err_csv_bytes,
            err_csv_name,
            CSV_MIME,
            cfg.ERROR_REPORT_FOLDER_ID,
            to_sheet=True
        )


        err_file_id = err_created["id"]
        err_link = err_created.get("webViewLink")

        err_gid = first_gid(sheets_svc, err_file_id)
        err_pdf = export_sheet(creds, err_file_id, err_gid, "pdf")
        err_pdf_name = f"Error_Report_{ts}.pdf"

        if logger:
            logger.info(f"Uploaded filtered Error Sheet: {err_link}")

        # Step 4C.1: Send Error Report if Needed

        to_err = _fallback_recipients(
            "ERROR REPORT",
            cfg.ERROR_RECIPIENTS,
            cfg.TO_RECIPIENTS,
            cfg.DEFAULT_ORDER_RECIPIENTS,
        )

                
        err_cc_list = list(dict.fromkeys(
            set(_clean_emails(cfg.TO_RECIPIENTS))
            | set(_clean_emails(cfg.CC_RECIPIENTS))
            - set(to_err)
        ))

        to_err = sorted(set(to_err) | set(err_cc_list))


        email_error_report(gmail_svc=gmail_svc, sender="me", to_list=to_err, cc_list=None, ts=ts, pdf_name=err_pdf_name, pdf_bytes=err_pdf, sheet_link=err_link)
        if logger:
            logger.info("Error report email sent")
        
        raise VendorPriceBookError(
            f"""One or more items were not found in the Vendor Price Book. The list of missing items has been sent to the technical support email.\n
            Once those items are added to the Vendor Price Book, please rerun the pipeline.\n
            Error Report: {err_link}
            """
        )




    # Step 4D: Full order upload (CSV) and export (PDF)
    full_csv_name = f"Order_Report_FULL_{location}_{ts}.csv"
    full_created = upload_to_drive(drive_svc, master_csv_bytes, full_csv_name, CSV_MIME, cfg.ORDER_REPORT_FOLDER_ID, to_sheet=True)
    full_file_id = full_created["id"]
    full_link = full_created.get('webViewLink')
    full_gid = first_gid(sheets_svc, full_file_id)
    full_pdf = export_sheet(creds, full_file_id, full_gid, "pdf")
    full_pdf_name = f"Order_Report_FULL_{location}_{ts}.pdf"
    if logger:
        logger.info(f"Uploaded FULL sheet: {full_created.get('webViewLink')}")

    # Step 4E: Create per-report-key outputs (CSV) and email

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
    sub_idx = lower_idx.get("sub_report_key")
    
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

        sub_key = None
        if sub_idx is not None:
            sub_key = (str(r.get(headers[sub_idx]) or "").strip().upper()) or None

        groups.setdefault((store.upper(), report_key.upper(), sub_key), []).append(r)
    
    
    for (store, key, sub_key), key_rows in groups.items():

        if not should_run(cfg, key, sub_key):
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
        sub_tag = clean_tag(sub_key)

        name_parts = []
        name_parts.append(store_tag)
        name_parts.append(tag)
        if sub_tag:
            name_parts.append(sub_tag)

        csv_name = f"Order_Report_{'_'.join(name_parts)}_{ts}.csv"
    
        # Upload CSV to Drive; conversion to Google Sheet happens via to_sheet=True
        created = upload_to_drive(
            drive_svc, key_csv_bytes, csv_name,
            CSV_MIME, cfg.ORDER_REPORT_FOLDER_ID, to_sheet=True
        )
    
        file_id = created["id"]
        gid = first_gid(sheets_svc, file_id)
    
        # Export the Google Sheet as PDF
        pdf = export_sheet(creds, file_id, gid, "pdf")
        pdfname = f"Order_Report_{'_'.join(name_parts)}_{ts}.pdf"
    
        # Prefer Store+Key; else Key; else Store; else To; else Default
        candidates = None
        
        candidates = None
        lookup_order = [
            (store_tag, tag, sub_tag),
            (store_tag, tag, None),
            (None, tag, sub_tag),
            (None, tag, None),
            (store_tag, None, None),
        ]

        for lk in lookup_order:
            if lk in cfg.REPORT_KEY_RECIPIENTS:
                candidates = cfg.REPORT_KEY_RECIPIENTS[lk]
                break
    
        recipients = _fallback_recipients(
            f"REPORT_KEY {tag}",
            candidates,
            cfg.TO_RECIPIENTS,
            cfg.DEFAULT_ORDER_RECIPIENTS
        )

        email_tag_parts = [tag, sub_tag]
        email_tag = ' - '.join(email_tag_parts)


        email_order_report(
            gmail_svc=gmail_svc,
            sender="me",
            to_list=recipients,
            cc_list=cfg.CC_RECIPIENTS,
            key=key,
            tag=email_tag,
            ts=ts,
            location=store,
            pdf_name=pdfname,
            pdf_bytes=pdf,
            sheet_link=created.get("webViewLink"),
            include_full_order=cfg.INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL,
            full_pdf_bytes=full_pdf,
            full_pdf_name=full_pdf_name,
        )
    
        if logger:
            logger.info(f"Emailed {store} - {tag} to {recipients}")
    
        
    # Step 4F: Send Manager Report (guarded by cfg.EMAIL_MANAGER_REPORT)
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

    

    # Step 4G: Send Full Order if needed
    if cfg.SEND_SEPARATE_FULL_ORDER_EMAIL:
        to_full = _fallback_recipients(
            "FULL ORDER",
            cfg.TO_RECIPIENTS,
            cfg.DEFAULT_ORDER_RECIPIENTS,
        )

        email_order_report(
            gmail_svc=gmail_svc,
            sender="me",
            to_list=to_full,
            cc_list=cfg.CC_RECIPIENTS,
            key='', # or a specific key if your function requires it
            tag="FULL",
            ts=ts,
            location=location,
            pdf_name=full_pdf_name,
            pdf_bytes=full_pdf,
            sheet_link=full_created.get("webViewLink"),
            include_full_order=False,  # already a full-only email
            full_pdf_bytes=None,
            full_pdf_name=None,
        )

        if logger:
            logger.info("FULL order email sent")
    else:
        if logger:
            logger.info("Separate full order email disabled")

    # Step 4H: File Cleanup

    try:
        if logger:
            logger.info("Cleaning up used incoming file…")
        trash_file(drive_svc, new_sales_report_id)
        trash_file(drive_svc, new_vendor_report_id)

        if logger:
            logger.info("Cleaning old incoming files…")
        for folder in [
            user_sales_folder_id,
            user_vendor_folder_id
        ]:
            cleanup_folder_by_age(
                drive_svc,
                folder,
                cfg.OUTPUT_TIME_TO_LIFE,
                logger
            )
        


        if logger:
            logger.info("Cleaning old output files…")
        for folder in [
            cfg.MANAGER_REPORT_FOLDER_ID,
            cfg.ORDER_REPORT_FOLDER_ID,
            cfg.ERROR_REPORT_FOLDER_ID
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

    return RunResult(
        ok=True,
        elapsed_seconds=elapsed,
        location=location,
        timestamp=ts,
        manager_pdf_link=manager_link,
        full_order_link=full_link,
        err_exist=err_exist,
        err_link=err_link
        )
