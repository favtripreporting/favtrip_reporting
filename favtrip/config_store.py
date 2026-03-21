""" 
config_store
======================================
This module provides small, focused helpers for reading and writing a JSON
configuration file stored in Google Drive using the `googleapiclient` (a.k.a.
Google API Python Client). It supports both direct file-ID addressing and a
convention-based "find by name" workflow using the constants
`DEFAULT_CONFIG_FILENAME` and `DEFAULT_MIMETYPE`.

Primary capabilities
--------------------
- **load_config_from_drive(...)**: Fetches and parses JSON from a Drive file.
  If no `file_id` is provided, the newest (by `modifiedTime`) non-trashed file
  named `DEFAULT_CONFIG_FILENAME` with MIME type `DEFAULT_MIMETYPE` is used.
  Returns an empty dict `{}` if the file does not exist, is empty, or contains
  invalid JSON.

- **save_config_to_drive(...)**: Writes JSON to Drive, either updating an
  existing file (by `file_id` or the latest matching name) or creating a new
  file. Returns the Drive file ID of the written resource. Supports optionally
  placing newly created files into a specific parent folder.

Design notes
------------
- **Non-throwing reads**: `load_config_from_drive` is intentionally resilient:
  it catches JSON parsing errors and returns `{}` for "not found" or invalid
  content scenarios to simplify caller logic.
- **Upsert semantics on save**: If `file_id` is not given, `save_config_to_drive`
  attempts to update the newest matching file by name and MIME type; if none is
  found, it creates a new one (optionally under `parent_folder_id`).
- **Streaming I/O**: Uses `MediaIoBaseDownload`/`MediaIoBaseUpload` for
  efficient transfer and compatibility with large files (even though configs
  are typically small).


Functions
---------
def load_config_from_drive(
    drive: googleapiclient.discovery.Resource,
    file_id: Optional[str] = None
) -> Dict[str, Any]:
    
    Read a JSON config from Google Drive.

    Behavior:
      - If `file_id` is provided, reads that exact file.
      - Otherwise, discovers the newest non-trashed file with:
           name == DEFAULT_CONFIG_FILENAME and mimeType == DEFAULT_MIMETYPE.
      - Returns `{}` if the file is not found, empty, or contains invalid JSON.

    Parameters:
      drive: An authenticated Google Drive v3 `Resource` client.
      file_id: Optional Drive file ID to read directly.

    Returns:
      A `dict` representing the parsed JSON configuration, or `{}` on failure.
    

save_config_to_drive(
    drive: googleapiclient.discovery.Resource,
    data: Dict[str, Any],
    file_id: Optional[str] = None,
    parent_folder_id: Optional[str] = None
) -> str:
    
    Write a JSON config to Google Drive (update or create).

    Behavior:
      - If `file_id` is provided, updates that file's content.
      - Else, attempts to find the newest matching file by name/mimeType and
        updates it.
      - If no matching file exists, creates a new file named
        `DEFAULT_CONFIG_FILENAME` (optionally under `parent_folder_id`).

    Parameters:
      drive: An authenticated Google Drive v3 `Resource` client.
      data: A JSON-serializable dictionary to write.
      file_id: Optional Drive file ID to update directly.
      parent_folder_id: Optional parent folder ID to place a newly created file.

    Returns:
      The Drive file ID (`str`) of the updated or created file.
    

Error handling & edge cases
---------------------------
- **Network/API errors**: This module defers to `googleapiclient` exceptions
  for request/transport failures. Callers may wish to wrap calls with retry
  logic (e.g., exponential backoff) or central error handling.
- **Invalid JSON on read**: Returns `{}` rather than raising, to keep consumers
  simple and robust to manual edits or empty files.
- **Encoding**: Files are read as UTF-8 (with replacement for invalid bytes)
  and written as UTF-8 with `ensure_ascii=False` to preserve Unicode.
- **Trashed files**: Explicitly filtered out during "discover by name".


"""


from __future__ import annotations
import io
import json
from typing import Any, Dict, Optional
from googleapiclient.discovery import Resource
from googleapiclient.http import MediaIoBaseUpload, MediaIoBaseDownload

DEFAULT_CONFIG_FILENAME = "favtrip_config.json"
DEFAULT_MIMETYPE = "application/json"

def load_config_from_drive(drive: Resource, file_id: Optional[str] = None) -> Dict[str, Any]:
    """
    Read the JSON config stored in Google Drive.
    If file_id is None, try to discover the newest file named DEFAULT_CONFIG_FILENAME.
    Returns {} if the file doesn't exist or is empty/invalid JSON.
    """
    # Discover by name if a specific id wasn't provided
    if not file_id:
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

    # Stream download the file
    request = drive.files().get_media(fileId=file_id)
    buf = io.BytesIO()
    downloader = MediaIoBaseDownload(buf, request)

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

def save_config_to_drive(
    drive: Resource,
    data: Dict[str, Any],
    file_id: Optional[str] = None,
    parent_folder_id: Optional[str] = None
) -> str:
    """
    Write JSON config to Google Drive.
    - If file_id provided, update that file.
    - Else upsert (update if found by name, otherwise create) DEFAULT_CONFIG_FILENAME.
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
