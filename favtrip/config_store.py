# favtrip/config_store.py
from __future__ import annotations
import io
import json
from typing import Any, Dict, Optional
from googleapiclient.http import MediaIoBaseUpload
from googleapiclient.discovery import Resource

# Minimal, explicit scope: we already request Drive scope in your SCOPES list.

DEFAULT_CONFIG_FILENAME = "favtrip_config.json"
DEFAULT_MIMETYPE = "application/json"

def load_config_from_drive(drive: Resource, file_id: Optional[str] = None) -> Dict[str, Any]:
    """
    Read the JSON config stored in Google Drive.
    If file_id is None, try to discover the newest file named DEFAULT_CONFIG_FILENAME.
    Returns {} if the file doesn't exist or is empty.
    """
    if not file_id:
        # Discover by name (newest wins)
        resp = drive.files().list(
            q=f"name='{DEFAULT_CONFIG_FILENAME}' and mimeType='{DEFAULT_MIMETYPE}' and trashed=false",
            orderBy="modifiedTime desc",
            pageSize=1,
            fields="files(id,name,modifiedTime)"
        ).execute() or {}
        files = resp.get("files", [])
        if not files:
            return {}
        file_id = files[0]["id"]

    # Download file content
    buf = io.BytesIO()
    request = drive.files().get_media(fileId=file_id)
    downloader = __import__("googleapiclient.http").apiclient.http.MediaIoBaseDownload(buf, request)
    done = False
    while not done:
        status, done = downloader.next_chunk()
    raw = buf.getvalue().decode("utf-8", errors="replace").strip()
    if not raw:
        return {}
    try:
        return json.loads(raw)
    except Exception:
        return {}

def save_config_to_drive(drive: Resource, data: Dict[str, Any], file_id: Optional[str] = None, parent_folder_id: Optional[str] = None) -> str:
    """
    Write JSON config to Google Drive.
    - If file_id provided, update that file.
    - Else create (or replace by name) DEFAULT_CONFIG_FILENAME in optional parent folder.
    Returns the Drive file ID.
    """
    payload = json.dumps(data, ensure_ascii=False, indent=2).encode("utf-8")
    media = MediaIoBaseUpload(io.BytesIO(payload), mimetype=DEFAULT_MIMETYPE, resumable=True)

    if file_id:
        updated = drive.files().update(
            fileId=file_id,
            media_body=media
        ).execute()
        return updated["id"]

    # Try to find an existing file by name to update
    resp = drive.files().list(
        q=f"name='{DEFAULT_CONFIG_FILENAME}' and mimeType='{DEFAULT_MIMETYPE}' and trashed=false",
        orderBy="modifiedTime desc",
        pageSize=1,
        fields="files(id,name)"
    ).execute() or {}
    files = resp.get("files", [])
    if files:
        fid = files[0]["id"]
        updated = drive.files().update(fileId=fid, media_body=media).execute()
        return updated["id"]

    # Create a new file
    meta = {"name": DEFAULT_CONFIG_FILENAME}
    if parent_folder_id:
        meta["parents"] = [parent_folder_id]

    created = drive.files().create(
        body=meta,
        media_body=media,
        fields="id,name"
    ).execute()
    return created["id"]