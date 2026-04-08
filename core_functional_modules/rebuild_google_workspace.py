from typing import Dict
import re

from googleapiclient.discovery import Resource

from core_functional_modules.config import Config
from core_functional_modules.google_client import load_valid_token, services
from core_functional_modules.config_store import save_config_to_drive
from core_functional_modules.logger import StatusLogger

def is_drive_link(value: str) -> bool:
    return value.startswith("http")


def extract_drive_id(value: str) -> str:
    """
    Accepts a Google Drive file ID or URL and returns the file ID.
    Raises if no valid ID can be extracted.
    """

    if not value:
        raise ValueError("Empty file ID or URL")

    value = value.strip()

    # If it already looks like a Drive ID, return it
    if re.fullmatch(r"[a-zA-Z0-9_-]{20,}", value):
        return value

    # Try extracting ID from common Drive URL formats
    patterns = [
        r"/d/([a-zA-Z0-9_-]+)",          # /d/<id>
        r"[?&]id=([a-zA-Z0-9_-]+)",      # ?id=<id>
    ]

    for pattern in patterns:
        match = re.search(pattern, value)
        if match:
            return match.group(1)

    raise ValueError(f"Could not extract Google Drive file ID from: {value}")

def create_folder_tree(drive, root_id):
    def mk(parent, name):
        return drive.files().create(
            body={
                "name": name,
                "mimeType": "application/vnd.google-apps.folder",
                "parents": [parent],
            },
            fields="id",
        ).execute()["id"]

    folders = {}

    folders["00_DOCS"] = mk(root_id, "00 Documentation")
    folders["01_MASTER"] = mk(root_id, "01 Master Files")
    folders["02_INPUT"] = mk(root_id, "02 Input Files")
    folders["03_WORK"] = mk(root_id, "03 Workhorse Files")
    folders["04_OUTPUT"] = mk(root_id, "04 Output Files")
    folders["99_UTIL"] = mk(root_id, "99 Utilities")

    folders["ORDER_REPORT"] = mk(folders["04_OUTPUT"], "01 Order Reports")
    folders["MANAGER_REPORT"] = mk(folders["04_OUTPUT"], "02 Manager Reports")
    folders["ERROR_REPORT"] = mk(folders["04_OUTPUT"], "03 Error Reports")

    return folders

def copy_master_files(drive, folders, cfg):

    def handle_master(attr_name):
        value = getattr(cfg, attr_name, None)
        if not value:
            return

        # If link → keep link, do NOT copy
        if is_drive_link(value):
            cfg_value = extract_drive_id(value)
            setattr(cfg, attr_name, cfg_value)
            return

        # If raw ID → copy into Master folder
        source_id = extract_drive_id(value)
        meta = drive.files().get(
            fileId=source_id,
            fields="name",
        ).execute()

        copied = drive.files().copy(
            fileId=source_id,
            body={
                "name": meta["name"],
                "parents": [folders["01_MASTER"]],
            },
            fields="id",
        ).execute()

        setattr(cfg, attr_name, copied["id"])

    handle_master("CALC_SPREADSHEET_ID")
    handle_master("BEV_MAPPING_LINK")


def find_single_folder_by_name(
    drive,
    name: str,
    parent_id: str,
):
    """
    Find exactly one non-trashed Google Drive folder by name
    that is a direct child of parent_id.
    Raises if none or more than one are found.
    """
    resp = drive.files().list(
        q=(
            "mimeType='application/vnd.google-apps.folder' "
            f"and name='{name}' "
            f"and '{parent_id}' in parents "
            "and trashed=false"
        ),
        fields="files(id,name,parents)",
    ).execute()

    files = resp.get("files", [])

    if not files:
        raise RuntimeError(
            f"Folder '{name}' not found under parent {parent_id}"
        )

    if len(files) > 1:
        raise RuntimeError(
            f"Multiple folders named '{name}' found under parent {parent_id}"
        )

    return files[0]["id"]

def copy_documentation(drive, folders, cfg, main_id):
    OLD_DOCUMENTATION_FOLDER_NAME = "00 Documentation"

    old_docs_id = find_single_folder_by_name(
        drive,
        OLD_DOCUMENTATION_FOLDER_NAME,
        parent_id=main_id,
    )

    resp = drive.files().list(
        q=f"'{old_docs_id}' in parents and trashed=false",
        fields="files(id,name)",
    ).execute()

    for f in resp.get("files", []):
        meta = drive.files().get(
            fileId=f["id"],
            fields="name",
        ).execute()

        drive.files().copy(
            fileId=f["id"],
            body={
                "name": meta["name"],
                "parents": [folders["00_DOCS"]],
            },
            fields="id",
        ).execute()


def handle_utilities(drive, folders, cfg, main_id):
    """
    Locate the old Utilities folder by name, copy its contents
    into the new Utilities folder, and update any referenced IDs.
    """

    def create_config_files_in_utilities(drive, utilities_folder_id, cfg):
        """
        Create a base config file, rename it to *_dev,
        then copy it and rename the copy to *_prod.
        """
        payload = cfg.to_drive_defaults()

        # 1️⃣ Create initial config file (no suffix yet)
        base_config_id = save_config_to_drive(
            drive,
            payload,
            file_id=None,
            parent_folder_id=utilities_folder_id,
        )

        # 2️⃣ Fetch its actual name from Drive
        meta = drive.files().get(
            fileId=base_config_id,
            fields="name",
        ).execute()

        original_name = meta["name"]
        stem = original_name.rsplit(".json", 1)[0]

        dev_name = f"{stem}_dev.json"
        prod_name = f"{stem}_prod.json"

        # 3️⃣ Rename the original file → *_dev.json
        drive.files().update(
            fileId=base_config_id,
            body={"name": dev_name},
        ).execute()

        dev_config_id = base_config_id

        # 4️⃣ Copy DEV → PROD
        prod_config_id = drive.files().copy(
            fileId=dev_config_id,
            body={"name": prod_name},
            fields="id",
        ).execute()["id"]

        return {
            "dev_config_file_id": dev_config_id,
            "prod_config_file_id": prod_config_id,
        }
    
    UTILITIES_FOLDER_NAME = "99 Utilities"
    EXCLUDED_NAME = "google_client_secret_copy_paste_in_app_secrets.json"

    # 1️⃣ Locate OLD utilities folder
    old_utilities_folder_id = find_single_folder_by_name(
        drive,
        UTILITIES_FOLDER_NAME,
        parent_id=main_id,
    )


    # 2️⃣ Copy files (excluding secrets)
    resp = drive.files().list(
        q=f"'{old_utilities_folder_id}' in parents and trashed=false",
        fields="files(id,name)",
    ).execute()

    old_to_new_id = {}

    for f in resp.get("files", []):
        if f["name"] == EXCLUDED_NAME:
            continue

        meta = drive.files().get(
            fileId=f['id'],
            fields="name",
        ).execute()

        copied = drive.files().copy(
            fileId=f['id'],
            body={
                "name": meta["name"],
                "parents": [old_utilities_folder_id],
            },
            fields="id",
        ).execute()

        old_to_new_id[f["id"]] = copied["id"]

    # 3️⃣ Update utility file IDs referenced in cfg
    if hasattr(cfg, "UTIL_FILE_IDS") and isinstance(cfg.UTIL_FILE_IDS, dict):
        for key, old_id in cfg.UTIL_FILE_IDS.items():
            if old_id in old_to_new_id:
                cfg.UTIL_FILE_IDS[key] = old_to_new_id[old_id]

    # 4️⃣ Create NEW DEV + PROD config files in Utilities
    return create_config_files_in_utilities(
        drive,
        folders["99_UTIL"],
        cfg
    )


def apply_new_folder_ids_to_cfg(drive, cfg, folders):
    """
    Writes new folder IDs into the DEV config JSON file.
    """

    # Update folder bindings
    cfg.MANAGER_REPORT_FOLDER_ID = folders["MANAGER_REPORT"]
    cfg.ORDER_REPORT_FOLDER_ID = folders["ORDER_REPORT"]
    cfg.ERROR_REPORT_FOLDER_ID = folders["ERROR_REPORT"]

    cfg.INPUT_FOLDER_ID = folders["02_INPUT"]
    cfg.WORKHORSE_FOLDER_ID = folders["03_WORK"]
    cfg.UTILITIES_FOLDER_ID = folders["99_UTIL"]

    DEV_CONFIG_FILE_ID = (
        cfg.DEV_CONFIG_FILE_ID
        if hasattr(cfg, "DEV_CONFIG_FILE_ID")
        else None
    )

    if not DEV_CONFIG_FILE_ID:
        raise RuntimeError("DEV_CONFIG_FILE_ID not set")

    save_config_to_drive(
        drive,
        cfg.to_drive_defaults(),
        file_id=DEV_CONFIG_FILE_ID
    )


def rebuild_google_workspace(cfg: Config):
    creds = load_valid_token(cfg.SCOPES)
    if not creds:
        raise RuntimeError("Google authentication required.")

    _, drive, _ = services(creds, cfg.HTTP_TIMEOUT_SECONDS)

    logger = StatusLogger()
    logger.info("Starting Google Workspace rebuild")

    # 1️⃣ Create main folder
    main_folder = drive.files().create(
        body={
            "name": "Reporting Pipeline Workspace",
            "mimeType": "application/vnd.google-apps.folder",
        },
        fields="id",
    ).execute()

    main_id = main_folder["id"]

    # 2️⃣ Create subfolders
    folders = create_folder_tree(drive, main_id)

    # 3️⃣ Copy content
    copy_documentation(drive, folders, cfg, main_id)
    copy_master_files(drive, folders, cfg)

    config_ids = handle_utilities(drive, folders, cfg, main_id)

    # 4️⃣ Apply new folder IDs to cfg (in‑memory)
    apply_new_folder_ids_to_cfg(drive, cfg, folders)

    return {
        "main_folder_id": main_id,
        "folders": folders,
        "dev_config_file_id": config_ids["dev_config_file_id"],
        "prod_config_file_id": config_ids["prod_config_file_id"],
    }