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