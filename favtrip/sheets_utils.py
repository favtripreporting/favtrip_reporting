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


def refresh_sheets_with_prefix(
    svc,
    spreadsheet_id: str,
    prefix: str = "REFR: ",
    retries: int = 5,
    chunk_cols: int = 2,
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
