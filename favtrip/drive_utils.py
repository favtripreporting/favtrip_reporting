""" 
drive_utils
======================================

Google Drive helper utilities for working with files and Google Sheets (Drive v3).

This module provides small, focused helpers to:
  • Upload arbitrary bytes (optionally as a native Google Sheet) to a folder.
  • Find the most recently created Google Sheet in a folder (optionally by exact name).
  • Copy or rename Drive files.
  • Soft-delete (trash) files in a folder that are older than a given age.
  • Safely escape literals for Drive v3 `q` search strings.
  • Format datetimes as RFC 3339 (UTC) for Drive queries.

It is designed to be used with an authenticated Google Drive v3 client from
`googleapiclient.discovery.build("drive", "v3", ...)`. All functions expect a
Drive service instance (here named `drive_svc` or `drive`) that is already
authorized for the necessary scopes.

-------------------------------------------------------------------------------
Key Functions
-------------------------------------------------------------------------------
_drive_q_escape(value: str) -> str
    Escape a literal for inclusion in the Drive v3 Files: list `q` parameter.
    This function ensures backslashes and single quotes are escaped in the
    correct order to avoid malformed query strings.

find_latest_sheet(drive_svc, folder_id: str) -> Optional[dict]
    Return the most recently created Google Sheet in the specified folder, or
    None if no spreadsheets exist. The returned object is a Drive file resource
    with fields: id, name, createdTime.

upload_to_drive(drive_svc, data: bytes, name: str, mime: str, folder_id: str, to_sheet: bool=False) -> dict
    Upload bytes as a Drive file into a folder. If `to_sheet=True`, the file is
    converted to a native Google Sheet (mimeType set to
    application/vnd.google-apps.spreadsheet). Returns the created file resource
    with fields: id, name, mimeType, webViewLink.

_rfc3339(dt: datetime) -> str
    Convert a datetime to RFC 3339 in UTC (e.g., "2024-01-01T00:00:00Z") for use
    in Drive queries such as `createdTime < '...'`.

trash_file(drive, file_id: str) -> dict
    Soft-delete (move to trash) a Drive file by ID. Uses `supportsAllDrives=True`
    so it also works with shared drives. Returns the updated file resource.

cleanup_folder_by_age(drive, folder_id: str, days: int, logger=None) -> int
    Find and trash all files in the folder whose `createdTime` is older than
    `now - days`. Returns the number of files trashed. When provided, `logger`
    is used to log info/warn messages for each file trashed or error encountered.

find_sheet_by_name(drive_svc, folder_id: str, name: str) -> Optional[dict]
    Return the most recently created Google Sheet in the folder that matches the
    given name exactly (case-sensitive), or None if not found. The returned
    object includes: id, name, createdTime, webViewLink.

copy_file_to_folder(drive_svc, src_file_id: str, dest_folder_id: str, new_name: str) -> dict
    Copy a Drive file (including native Docs/Sheets/Slides) into a destination
    folder and give it a new name. Returns the created file resource with:
    id, name, mimeType, webViewLink.

rename_file(drive_svc, file_id: str, new_name: str) -> dict
    Rename an existing Drive file by ID. Returns the updated file resource with:
    id, name, mimeType, webViewLink.

-------------------------------------------------------------------------------
Inputs, Outputs, and Contracts
-------------------------------------------------------------------------------
• All Drive service parameters (`drive_svc` / `drive`) must be a valid
  `Resource` from `googleapiclient.discovery.build("drive", "v3", ...)`.
• Folder/file identifiers must be the opaque Drive IDs (not paths).
• `upload_to_drive(..., to_sheet=True)` will attempt server-side conversion to a
  Google Sheet; this is appropriate for tabular formats (e.g., CSV). If you
  supply a non-tabular format with `to_sheet=True`, Google may reject the
  conversion.
• Functions return minimal file resources constrained by the `fields` parameter
  for efficiency. If you need additional fields, adjust the `fields` in the
  function(s) or perform a subsequent `files().get(...)`.

-------------------------------------------------------------------------------
Date/Time Handling
-------------------------------------------------------------------------------
• All time comparisons use UTC. `_rfc3339` normalizes input datetimes to UTC and
  formats them as `"YYYY-MM-DDTHH:MM:SSZ"`. When supplying your own datetimes,
  prefer timezone-aware objects.

"""

from __future__ import annotations
import io
from datetime import datetime, timedelta, timezone
from googleapiclient.http import MediaIoBaseUpload


def _drive_q_escape(value: str) -> str:
    """Escape a literal for Google Drive v3 'q' strings."""
    # Order matters: escape backslashes first, then single quotes.
    return value.replace("\\", "\\\\").replace("'", "\\'")

def find_latest_sheet(drive_svc, folder_id: str):
    q = (
        f"'{folder_id}' in parents and "
        "mimeType='application/vnd.google-apps.spreadsheet' and trashed=false"
    )
    resp = drive_svc.files().list(
        q=q, orderBy="createdTime desc", pageSize=1,
        fields="files(id,name,createdTime)"
    ).execute()
    files = resp.get("files", [])
    return files[0] if files else None


def upload_to_drive(drive_svc, data: bytes, name: str, mime: str, folder_id: str, to_sheet: bool=False):
    meta = {"name": name, "parents": [folder_id]}
    if to_sheet:
        meta["mimeType"] = "application/vnd.google-apps.spreadsheet"
    media = MediaIoBaseUpload(io.BytesIO(data), mimetype=mime, resumable=True)
    return drive_svc.files().create(
        body=meta, media_body=media, fields="id,name,mimeType,webViewLink"
    ).execute()

def _rfc3339(dt: datetime) -> str:
    return dt.astimezone(timezone.utc).isoformat(timespec="seconds").replace("+00:00", "Z")

def trash_file(drive, file_id: str):
    return drive.files().update(fileId=file_id, body={"trashed": True}, supportsAllDrives=True).execute()

def cleanup_folder_by_age(drive, folder_id: str, days: int, logger=None):
    if days <= 0:
        return 0

    cutoff = datetime.now(timezone.utc) - timedelta(days=days)
    cutoff_str = _rfc3339(cutoff)

    q = (
        f"'{folder_id}' in parents and trashed=false "
        f"and createdTime < '{cutoff_str}'"
    )

    trashed = 0
    page_token = None

    while True:
        resp = drive.files().list(
            q=q,
            pageSize=1000,
            orderBy="createdTime asc",
            fields="nextPageToken, files(id,name,createdTime)",
            pageToken=page_token
        ).execute() or {}

        for f in resp.get("files", []):
            try:
                trash_file(drive, f["id"])
                trashed += 1
                if logger:
                    logger.info(f"Trashed file: {f['name']} ({f['id']})")
            except Exception as e:
                if logger:
                    logger.warn(f"Failed to trash {f['id']}: {e}")

        page_token = resp.get("nextPageToken")
        if not page_token:
            break

    return trashed


def find_sheet_by_name(drive_svc, folder_id: str, name: str):
    """
    Return the most-recently-created Google Sheet in folder_id with exact name, or None.
    """
    
    q = (
        f"'{folder_id}' in parents and "
        f"name = '{_drive_q_escape(name)}' and "
        "mimeType = 'application/vnd.google-apps.spreadsheet' and trashed = false"
    )

    resp = drive_svc.files().list(
        q=q,
        orderBy="createdTime desc",
        pageSize=1,
        fields="files(id,name,createdTime,webViewLink)"
    ).execute()
    files = resp.get("files", [])
    return files[0] if files else None

def copy_file_to_folder(drive_svc, src_file_id: str, dest_folder_id: str, new_name: str):
    """
    Copy a Drive file (e.g., Google Spreadsheet) into a folder with a new name.
    Returns the created file resource (id, name, webViewLink).
    """
    body = {"name": new_name, "parents": [dest_folder_id]}
    return drive_svc.files().copy(
        fileId=src_file_id,
        body=body,
        fields="id,name,mimeType,webViewLink"
    ).execute()

def rename_file(drive_svc, file_id: str, new_name: str):
    """
    Rename a Google Drive file by its fileId.
    Returns the updated file resource (id, name, mimeType, webViewLink).
    """
    body = {"name": new_name}
    return drive_svc.files().update(
        fileId=file_id,
        body=body,
        fields="id,name,mimeType,webViewLink"
    ).execute()
