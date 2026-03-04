from __future__ import annotations
import io
from googleapiclient.http import MediaIoBaseUpload


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
