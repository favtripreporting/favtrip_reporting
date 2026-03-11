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