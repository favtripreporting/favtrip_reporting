"""
sheet_utils
======================================

Google Sheets utility helpers for copying sheets, refreshing formulas, and basic row/values ops.

This module wraps common tasks against the Google Sheets API v4 (via an authenticated
`svc = googleapiclient.discovery.build("sheets", "v4", ...)` service), including:

- Discovering and selecting sheets within a spreadsheet:
  • list_sheets() – list sheet metadata
  • get_sheet() – find a sheet's properties by title
  • first_gid(), get_first_sheet_meta() – convenient access to the first sheet

- Sheet lifecycle utilities:
  • delete_sheet() – remove a sheet by title
  • add_blank_sheet() – create a blank sheet with a title and grid size
  • add_or_replace_sheet() – delete a sheet if it exists, then add a fresh one
  • copy_sheet_as() – duplicate a sheet within a spreadsheet and rename it
  • copy_first_sheet_as() – copy the first sheet to another spreadsheet and rename it
  • copy_sheet_to_another_spreadsheet() – copy a sheet by title across spreadsheets with optional rename

- Values and range helpers:
  • get_values_2d() – read a 2D values range from a sheet
  • put_values_2d() – write a 2D values matrix starting at A1
  • get_value() – try a named range first, then fall back to the sheet’s first column

- Row manipulation:
  • delete_rows_range() – delete a contiguous 0-based row range (end exclusive)
  • delete_row_indices() – delete multiple absolute row indices (descending order)

- Formula recomputation workarounds:
  • refresh_sheets_with_prefix() – trigger recalc on all sheets whose titles start with a prefix
  • refresh_sheets_with_prefix_chunked() – same, but in column chunks (useful for large sheets)
  • _force_column_as_text() – coerce a column (matched by header name) to text by prefixing values with "'"

------------------------------------------------------------------------------
Requirements & assumptions
------------------------------------------------------------------------------
- Authentication: All functions expect a pre-authenticated Sheets API service object
  (`svc`) with permissions to read/update the target spreadsheet(s).
- Access: The caller (service account or user) needs editor access to any
  spreadsheet being modified or receiving copies.
- API: These helpers use the Sheets API v4 `spreadsheets` and `values` methods,
  including `get`, `batchUpdate`, and `copyTo`.
- Error handling: Most functions surface API errors as exceptions from the client
  library. Select functions include simple retry loops (with jitter) on write
  operations to reduce transient failures.
- Idempotency: Destructive operations (e.g., delete) are NOT idempotent. Use with care.
- Indexing: Row/column indices in batchUpdate ranges are 0-based and end-exclusive,
  mirroring the Sheets API.

------------------------------------------------------------------------------
Key behaviors & caveats
------------------------------------------------------------------------------
- copy_sheet_as() and copy_sheet_to_another_spreadsheet():
  - Return the new sheetId (int) on success, or None if the source sheet isn't found
    or the API returns an unexpected structure.
  - If you pass a `new_title` that collides with an existing sheet title, the request
    only attempts to update title; it does not resolve conflicts.
- refresh_sheets_with_prefix*():
  - These functions "poke" formulas by performing a find/replace of "=" -> "="
    (no visible change), prompting recalculation.
  - The chunked variant determines the number of used columns based on a header row.
    Adjust `header_row` and `chunk_cols` to control scope and batching.
- get_value():
  - First attempts to read a named range. If not found or empty, falls back to
    the first column (A) of the provided `sheet_title`. Returns "UNKNOWN" if empty.

------------------------------------------------------------------------------
Function reference (selected)
------------------------------------------------------------------------------
list_sheets(svc, spreadsheet_id) -> List[Dict[str, Any]]:
    Fetch metadata for all sheets in a spreadsheet.

get_sheet(sheets, title) -> Optional[Dict[str, Any]]:
    Return the `properties` of the sheet whose title matches `title`, else None.

delete_sheet(svc, spreadsheet_id, title) -> None:
    Delete the sheet with the provided title if it exists.

copy_sheet_as(svc, spreadsheet_id, src_title, new_title) -> Optional[int]:
    Copy a sheet (by title) within the same spreadsheet, rename it, and return its sheetId.

copy_sheet_to_another_spreadsheet(
    svc, src_spreadsheet_id, src_title, dest_spreadsheet_id, new_title=None
) -> Optional[int]:
    Copy a sheet (by title) from one spreadsheet to another, optionally renaming it.

copy_first_sheet_as(svc, src_spreadsheet, dest_spreadsheet, new_title) -> int:
    Copy the first sheet of the source into the destination and rename it. Returns new sheetId.

get_values_2d(svc, spreadsheet_id, sheet_title, a1_range="A:Z") -> list[list]:
    Return a 2D array of values for the A1 range within the specified sheet.

put_values_2d(svc, spreadsheet_id, sheet_title, values) -> None:
    Write a 2D matrix to the sheet starting at A1 using USER_ENTERED semantics.

delete_rows_range(svc, spreadsheet_id, sheet_id, start_row_index, end_row_index) -> None:
    Delete 0-based rows in [start_row_index, end_row_index).

delete_row_indices(svc, spreadsheet_id, sheet_id, row_indices_desc) -> None:
    Delete multiple absolute row indices (0-based). Internally sorts in descending order.

refresh_sheets_with_prefix(
    svc, spreadsheet_id, prefix="REFR: ", retries=5, logger=None
) -> None:
    For each sheet whose title starts with prefix, forces formula recalc with retries.

refresh_sheets_with_prefix_chunked(
    svc, spreadsheet_id, prefix="REFR: ", retries=5, chunk_cols=3, header_row=1, logger=None
) -> None:
    As above, but operates on small column ranges per attempt to reduce request size/timeouts.

_force_column_as_text(header, rows, header_name) -> list[list]:
    Return a new rows array where the column matching `header_name` is coerced to text by
    prefixing non-blank values with a single apostrophe.

------------------------------------------------------------------------------
Usage examples
------------------------------------------------------------------------------
# 1) Copy a sheet within the same spreadsheet and rename it
new_id = copy_sheet_as(svc, spreadsheet_id="AAA...", src_title="Template", new_title="Run 2026-03-21")

# 2) Copy a sheet from one spreadsheet to another and rename it
new_id = copy_sheet_to_another_spreadsheet(
    svc,
    src_spreadsheet_id="SRC_ID",
    src_title="Report",
    dest_spreadsheet_id="DEST_ID",
    new_title="Report (Copy)"
)

# 3) Force formula recalculation on all sheets prefixed with "REFR: "
refresh_sheets_with_prefix(svc, spreadsheet_id="AAA...", prefix="REFR: ", retries=3)

# 4) Write a 2D table to a sheet starting at A1
put_values_2d(svc, spreadsheet_id="AAA...", sheet_title="Data", values=[["A","B"], [1,2], [3,4]])

# 5) Delete rows 10..20 (0-based, end-exclusive)
delete_rows_range(svc, spreadsheet_id="AAA...", sheet_id=123456789, start_row_index=10, end_row_index=21)

------------------------------------------------------------------------------
Logging & retries
------------------------------------------------------------------------------
Some functions accept an optional `logger` (any object exposing `.info`, `.warning`, or `.warn`)
to receive progress messages. Retry loops use a simple exponential-ish backoff with random jitter
(`time.sleep(1 + random.random())`) up to `retries` attempts.

------------------------------------------------------------------------------
Safety notes
------------------------------------------------------------------------------
- Destructive operations (delete/replace) cannot be undone by this module. Make sure you
  have backups and required permissions before running them in production.
- Title-based targeting assumes unique sheet titles. Name collisions can lead to unexpected results.
- For very large sheets, consider the chunked refresh function to avoid request size/timeouts.

"""


from __future__ import annotations
import random
import time
from typing import Any, Dict, List


def list_sheets(svc, spreadsheet_id: str) -> List[Dict[str, Any]]:
    return svc.spreadsheets().get(spreadsheetId=spreadsheet_id).execute().get("sheets", [])


def get_sheet(sheets, title: str):
    for s in sheets:
        if s["properties"]["title"] == title:
            return s["properties"]
    return None


def delete_sheet(svc, spreadsheet_id: str, title: str):
    s = get_sheet(list_sheets(svc, spreadsheet_id), title)
    if s:
        body = {"requests": [{"deleteSheet": {"sheetId": s["sheetId"]}}]}
        svc.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()


def copy_sheet_as(svc, spreadsheet_id: str, src_title: str, new_title: str):
    s = get_sheet(list_sheets(svc, spreadsheet_id), src_title)
    if not s:
        return None
    copied = svc.spreadsheets().sheets().copyTo(
        spreadsheetId=spreadsheet_id,
        sheetId=s["sheetId"],
        body={"destinationSpreadsheetId": spreadsheet_id}
    ).execute()
    new_id = copied["sheetId"]
    body = {"requests": [{
        "updateSheetProperties": {
            "properties": {"sheetId": new_id, "title": new_title},
            "fields": "title"
        }
    }]}
    svc.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
    return new_id


def copy_sheet_to_another_spreadsheet(
    svc,
    src_spreadsheet_id: str,
    src_title: str,
    dest_spreadsheet_id: str,
    new_title: str | None = None
) -> int | None:
    """
    Copy a sheet (by title) from one Google Sheets spreadsheet to another.

    Args:
        svc: An authenticated Google Sheets API service (from googleapiclient.discovery.build('sheets','v4', ...)).
        src_spreadsheet_id: The ID of the source spreadsheet (the file that currently contains the sheet).
        src_title: The title of the sheet in the source spreadsheet to copy.
        dest_spreadsheet_id: The ID of the destination spreadsheet (the file to receive the copied sheet).
        new_title: Optional new title to apply to the copied sheet in the destination.

    Returns:
        The new sheetId in the destination spreadsheet, or None if the source sheet wasn't found.

    Notes:
        - The service account or authenticated user must have at least editor access to both spreadsheets.
        - If new_title is provided and a sheet with that title already exists in the destination,
          this function will attempt to rename the new sheet to new_title and will not resolve title conflicts.
    """
    # Find the source sheet by title
    src_sheet = get_sheet(list_sheets(svc, src_spreadsheet_id), src_title)
    if not src_sheet:
        return None

    # Copy the sheet into the destination spreadsheet
    copied = (
        svc.spreadsheets()
        .sheets()
        .copyTo(
            spreadsheetId=src_spreadsheet_id,
            sheetId=src_sheet["sheetId"],
            body={"destinationSpreadsheetId": dest_spreadsheet_id}
        )
        .execute()
    )

    new_id = copied.get("sheetId")
    if not new_id:
        # Unexpected, but guard just in case
        return None

    # Optionally rename the newly copied sheet in the destination
    if new_title:
        body = {
            "requests": [
                {
                    "updateSheetProperties": {
                        "properties": {"sheetId": new_id, "title": new_title},
                        "fields": "title",
                    }
                }
            ]
        }
        svc.spreadsheets().batchUpdate(
            spreadsheetId=dest_spreadsheet_id, body=body
        ).execute()

    return new_id



def copy_first_sheet_as(svc, src_spreadsheet: str, dest_spreadsheet: str, new_title: str):
    meta = svc.spreadsheets().get(spreadsheetId=src_spreadsheet).execute()
    first_id = meta["sheets"][0]["properties"]["sheetId"]
    copied = svc.spreadsheets().sheets().copyTo(
        spreadsheetId=src_spreadsheet,
        sheetId=first_id,
        body={"destinationSpreadsheetId": dest_spreadsheet}
    ).execute()
    new_id = copied["sheetId"]
    body = {"requests": [{
        "updateSheetProperties": {
            "properties": {"sheetId": new_id, "title": new_title},
            "fields": "title"
        }
    }]}
    svc.spreadsheets().batchUpdate(spreadsheetId=dest_spreadsheet, body=body).execute()
    return new_id

def refresh_sheets_with_prefix(svc, spreadsheet_id: str, prefix: str = "REFR: ", retries: int = 5, logger=None):
    sheets = list_sheets(svc, spreadsheet_id)
    targets = [s["properties"] for s in sheets if s["properties"]["title"].startswith(prefix)]
    for idx, t in enumerate(targets, start=1):
        body = {"requests": [{
            "findReplace": {
                "find": "=",
                "replacement": "=",
                "includeFormulas": True,
                "sheetId": t["sheetId"]
            }
        }]}
        attempt = 0
        while True:
            try:
                svc.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
                if logger:
                    logger.info(f"[{idx}/{len(targets)}] Recalc OK: {t['title']}")
                break
            except Exception:
                attempt += 1
                if attempt > retries:
                    if logger:
                        logger.warn(f"FAILED recalc for {t['title']}")
                    break
                time.sleep(1 + random.random())


def refresh_sheets_with_prefix_chunked(
    svc,
    spreadsheet_id: str,
    prefix: str = "REFR: ",
    retries: int = 5,
    chunk_cols: int = 3,
    header_row: int = 1,
    logger=None,
):
    sheets = list_sheets(svc, spreadsheet_id)
    targets = [s["properties"] for s in sheets if s["properties"]["title"].startswith(prefix)]

    for idx, t in enumerate(targets, start=1):
        sheet_id = t["sheetId"]
        title = t["title"]

        # Get header row to detect used columns
        resp = svc.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=f"'{title}'!{header_row}:{header_row}"
        ).execute()

        row = resp.get("values", [[]])[0]
        col_count = len(row)

        if col_count == 0:
            continue

        for start_col in range(0, col_count, chunk_cols):
            end_col = min(start_col + chunk_cols, col_count)

            body = {
                "requests": [{
                    "findReplace": {
                        "find": "=",
                        "replacement": "=",
                        "includeFormulas": True,
                        "range": {
                            "sheetId": sheet_id,
                            "startColumnIndex": start_col,
                            "endColumnIndex": end_col,
                        },
                    }
                }]
            }

            attempt = 0
            while True:
                try:
                    svc.spreadsheets().batchUpdate(
                        spreadsheetId=spreadsheet_id,
                        body=body
                    ).execute()

                    if logger:
                        logger.info(
                            f"[{idx}/{len(targets)}] {title} cols {start_col}-{end_col} recalculated"
                        )
                    break

                except Exception:
                    attempt += 1
                    if attempt > retries:
                        if logger:
                            logger.warning(f"FAILED recalc {title} cols {start_col}-{end_col}")
                        break
                    time.sleep(1 + random.random())


def get_value(svc, spreadsheet_id: str, sheet_title: str, named_range: str) -> str:
    try:
        vals = svc.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=named_range
        ).execute().get("values", [])
    except Exception:
        vals = []
    if not vals:
        try:
            vals = svc.spreadsheets().values().get(
                spreadsheetId=spreadsheet_id,
                range=f"'{sheet_title}'!A1:A"
            ).execute().get("values", [])
        except Exception:
            vals = []
    return vals[0][0] if vals and vals[0] else "UNKNOWN"


def first_gid(svc, spreadsheet_id: str) -> int:
    meta = svc.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    return meta["sheets"][0]["properties"]["sheetId"]

# --- Additional helpers for row inspection/edits ---

def get_first_sheet_meta(svc, spreadsheet_id: str):
    """Return (first_sheet_title, first_sheet_id)."""
    meta = svc.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    first = meta["sheets"][0]["properties"]
    return first["title"], first["sheetId"]

def get_values_2d(svc, spreadsheet_id: str, sheet_title: str, a1_range: str = "A:Z"):
    """Fetch a 2D values array from a sheet title + A1 range."""
    rng = f"'{sheet_title}'!{a1_range}"
    res = svc.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=rng).execute()
    return res.get("values", [])

def delete_rows_range(svc, spreadsheet_id: str, sheet_id: int, start_row_index: int, end_row_index: int):
    """Delete [start_row_index, end_row_index) (0‑based; end exclusive)."""
    if end_row_index <= start_row_index:
        return
    body = {"requests": [{
        "deleteDimension": {
            "range": {
                "sheetId": sheet_id,
                "dimension": "ROWS",
                "startIndex": start_row_index,
                "endIndex": end_row_index,
            }
        }
    }]}
    svc.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()

def delete_row_indices(svc, spreadsheet_id: str, sheet_id: int, row_indices_desc: list[int]):
    """Delete multiple absolute row indices (0‑based) in descending order."""
    for r in sorted(row_indices_desc, reverse=True):
        delete_rows_range(svc, spreadsheet_id, sheet_id, r, r+1)

def add_blank_sheet(svc, spreadsheet_id: str, title: str, rows: int = 1000, cols: int = 26):
    """Create a blank sheet with a given title."""
    body = {"requests": [{
        "addSheet": {"properties": {"title": title, "gridProperties": {"rowCount": rows, "columnCount": cols}}}
    }]}
    svc.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()

def add_or_replace_sheet(svc, spreadsheet_id: str, title: str, rows: int = 2000, cols: int = 50):
    """
    Remove any existing sheet with 'title' and add a blank one.
    """
    try:
        delete_sheet(svc, spreadsheet_id, title)
    except Exception:
        # if not present, ignore
        pass
    add_blank_sheet(svc, spreadsheet_id, title, rows, cols)

def put_values_2d(svc, spreadsheet_id: str, sheet_title: str, values: list[list]):
    """
    Write a 2D array to 'A1' of 'sheet_title' in a single update.
    """
    rng = f"'{sheet_title}'!A1"
    svc.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=rng,
        valueInputOption="USER_ENTERED",
        body={"values": values}
    ).execute()

def _force_column_as_text(header: list[str], rows: list[list], header_name: str) -> list[list]:
    """
    For the column matching header_name, coerce every non-blank value to a string
    prefixed with a single apostrophe, so Google Sheets stores it as text.
    """
    idx = None
    for i, h in enumerate(header):
        if str(h).strip().lower() == header_name.strip().lower():
            idx = i
            break
    if idx is None:
        return rows  # header not found; nothing to do

    out = []
    for r in rows:
        r2 = list(r)
        if idx < len(r2) and r2[idx] not in (None, ""):
            # ensure string and prefix with apostrophe
            r2[idx] = "'" + str(r2[idx])
        out.append(r2)
    return out
