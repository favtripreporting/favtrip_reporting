# Project Snapshot (CodeBundle)

                This single Markdown file contains a **self-contained snapshot** of your project so another AI/engineer can review or modify it without needing the original folder.

                **How to use this file with an AI**
                1. Upload or paste this file as a single attachment.
                2. Ask for changes; the AI can reference specific `file:` sections below.
                3. Copy updated blocks back into the corresponding files in your project.

                > Notes: secrets like `token.json` are intentionally excluded. Virtual envs and build artifacts are omitted to keep this readable.

                ## Directory tree (filtered)
                .env [skipped: secret]
.env.example
.gitignore
FavTripPipeline.spec
FavTripPipelineUI.spec
_user_interface_.py
cli.py
credentials.json [skipped: secret]
last_run.log
launcher_streamlit.py
requirements.txt
setup_py2app.py
token.json [skipped: secret]
web_url_credentials.json [skipped: secret]
  core_functional_modules/
    __init__.py
    config.py
    config_store.py
    drive_utils.py
    gmail_utils.py
    google_client.py
    logger.py
    pipeline.py
    pipeline_bus.py
    rebuild_google_workspace.py
    sheets_utils.py
  documentation/
    PROJECT_SNAPSHOT_CODEBUNDLE.md
    README.md
    generate_code_bundle.py
    git_workflow.txt
    requirements.txt
  __dev_input_sales_file/
    1 Store - 1 Week.xlsx
    1 Store - 2 Weeks.xlsx  [skipped: too large]
    1 Store - Bad End.xlsx  [skipped: too large]
    1 Store - Bad Start.xlsx
    2 Stores - 1 Week.xlsx  [skipped: too large]
    2 Stores - 2 Weeks.xlsx  [skipped: too large]
    VPB Error - BEV.xlsx  [skipped: too large]
  __dev_input_vendor_file/
    Vendors Price Book.xlsx
  __executable/
    run_windows.bat
---
### file: .env.example

```
# --- Required IDs ---
CALC_SPREADSHEET_ID=
INCOMING_FOLDER_ID=
MANAGER_REPORT_FOLDER_ID=
ORDER_REPORT_FOLDER_ID=

# --- Optional IDs / settings ---
GID_MANAGER_PDF=1921812573
GID_ORDER_CSV=1875928148
LOCATION_SHEET_TITLE=REFR: Values
LOCATION_NAMED_RANGE=_locations

TIMESTAMP_TZ=America/Chicago
TIMESTAMP_FMT=%Y-%m-%d-%I-%M-%p

# Recipients
TO_RECIPIENTS=
CC_RECIPIENTS=
DEFAULT_ORDER_RECIPIENTS=

# Report keys
USE_ALL_REPORT_KEYS=false
REPORT_KEY_RUN_LIST=GROCERY,COFFEE
# JSON mapping: {"GROCERY":["a@b.com","c@d.com"],"COFFEE":["x@y.com"]}
REPORT_KEY_RECIPIENTS={}

INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL=false
SEND_SEPARATE_FULL_ORDER_EMAIL=true

# Google API scopes (normally leave as-is)
SCOPES=https://www.googleapis.com/auth/drive,https://www.googleapis.com/auth/spreadsheets,https://www.googleapis.com/auth/gmail.send

FORCE_REAUTH=false
REDIRECT_PORT=58285
HTTP_TIMEOUT_SECONDS=300

```

---
### file: .gitignore

```
*.env
*credentials.json
token.json
web_url_credentials.json

```

---
### file: FavTripPipeline.spec

```
# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['cli.py'],
    pathex=[],
    binaries=[],
    datas=[('.env', '.'), ('credentials.json', '.')],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='FavTripPipeline',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

```

---
### file: FavTripPipelineUI.spec

```
# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['launcher_streamlit.py'],
    pathex=[],
    binaries=[],
    datas=[('.env', '.'), ('credentials.json', '.'), ('ui_streamlit.py', '.'), ('favtrip', 'favtrip')],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='FavTripPipelineUI',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

```

---
### file: __executable/run_windows.bat

```bat
@echo off
setlocal
REM ---------------------------------------------------------------------------
REM Run Streamlit UI without a persistent console window.
REM Location: __executable\run_web_windows_silent.bat
REM Behavior: brief flash at launch, then only the browser tab remains.
REM ---------------------------------------------------------------------------

REM Move into the folder of this .bat
pushd "%~dp0"

REM Go to the project root (one level up from __executable)
cd ..

REM Choose Python: prefer venv's interpreter if present
set "PY_VENV=.\.venv\Scripts\python.exe"
set "PY="
if exist "%PY_VENV%" (
  set "PY=%PY_VENV%"
) else (
  for %%P in (python.exe py.exe) do (
    where %%P >nul 2>&1 && (set "PY=%%P" & goto :gotpy)
  )
)
:gotpy
if not defined PY (
  echo [Launcher] Python was not found. Install Python or create .\.venv and try again.
  popd
  exit /b 1
)

REM Streamlit prefs: ensure it opens the browser and stays local
set "STREAMLIT_SERVER_HEADLESS=false"
set "STREAMLIT_BROWSER_GATHER_USAGE_STATS=false"

REM Start Streamlit hidden and detach from this console (which then closes)
REM - We invoke PowerShell only to spawn the hidden child process.
start "" /MIN powershell -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -Command ^
  "Start-Process -FilePath '%PY%' -ArgumentList '-m','streamlit','run','ui_streamlit.py' -WindowStyle Hidden"

popd
exit /b 0
```

---
### file: _user_interface_.py

```python
"""
FavTrip Reporting Streamlit UI
================================

Overview
--------
This Streamlit app is the front-end for the FavTrip Reporting pipeline. It lets an authenticated
Google user upload a Modisoft "Live Items Report", configure who receives emails for each report
(or fallback recipients), tune advanced IDs/GIDs/time settings, and then orchestrate the
`core_functional_modules.pipeline.run_pipeline` execution while streaming status updates and a timer.

Design goals
------------
* **No local secrets in code**: OAuth client JSON is read from `st.secrets["GOOGLE_CREDENTIALS"]` and
  the app base URL from `st.secrets["APP_BASE_URL"]`. Optional `CONFIG_FILE_ID` pins a Drive JSON that
  stores editable defaults.
* **Constrained, clear UX**: A two-step flow (Upload ➜ Run). The Run button is enabled only after a
  successful upload, reducing accidental runs on stale input.
* **Robust OAuth (PKCE)**: Uses an explicit code verifier/challenge and encodes the verifier inside
  the `state` payload to remain stateless across redirects.
* **Operational safety**: Detects common mistakes (e.g., invalid email inputs, duplicate keys,
  wrong-week uploads) and surfaces warnings or blocks execution accordingly.

Key concepts
------------
* **Incoming file**: The uploaded Modisoft report (CSV/XLSX). It is pushed to a configured Google
  Drive folder and optionally converted to a Google Sheet for the downstream pipeline.
* **Report Keys**: Categories/tags used by the pipeline to partition output and email recipients.
  You can either process *all* keys present in the data or restrict to a comma-separated subset.
* **Per-Report-Key Recipients**: Optional overrides that map `(Store, Report Key)` pairs to recipient
  lists. Fallback recipients apply where no specific mapping exists.
* **Drive-backed defaults**: The app can persist your current UI settings to a JSON in Drive. Supplying
  `CONFIG_FILE_ID` in Streamlit secrets will cause subsequent sessions to update that exact file.

Security model
--------------
* OAuth scopes are supplied by `Config` and used to mint a user token saved locally as `token.json`.
* The app opens the Google consent screen in a **new tab**, and that tab becomes the main app after
  redirect. Tokens are not sent back to the opener page; they are stored only in the process serving
  the tab that completed OAuth.

Operational notes
-----------------
* If a run fails with the message "Please only upload 1 or 2 full weeks of data", the UI locks to
  prevent immediate re-runs. Use **Retry** to clear the lock and upload a correct file.
* Set **Offer full log download** (sidebar) to expose a download button for `last_run.log` after a run.
* The **green Run button** appears once a fresh upload succeeds, indicating the pipeline is ready to run.

Dependencies & integration points
---------------------------------
* `core_functional_modules.google_client`: token loading/clearing and service factories (Drive, Sheets, Gmail)
* `core_functional_modules.config`: the central configuration object. `Config.load()` merges defaults, secrets, and any
  Drive-stored overrides.
* `core_functional_modules.drive_utils.upload_to_drive`: uploads the incoming report and (optionally) converts to Sheet.
* `core_functional_modules.pipeline.run_pipeline`: the orchestrated processing step; returns an object with links and
  timing information used to render the result panel.

This file intentionally includes **documentation-only** additions (module docstring and inline comments)
without modifying the executable logic.
"""

# ------------------------------
# Quick-start for maintainers
# ------------------------------
# 1) Configure Streamlit secrets:
#    - APP_BASE_URL: The exact external base URL of your deployed app (with trailing slash normalization).
#    - GOOGLE_CREDENTIALS: A JSON string containing your OAuth client configuration.
#    - CONFIG_FILE_ID (optional): The Drive file ID for persisted UI defaults.
# 2) Grant your Google Cloud OAuth Client access to the app origin and redirect URI.
# 3) Run: `streamlit run ui_streamlit.py` (ensure the backend `favtrip` package is importable).
# 4) Upload a Modisoft report, adjust recipients and options, then click **Run Pipeline**.


import os
import time
import threading
import json
import base64
import hashlib
import secrets
import re
import queue
import uuid
import traceback
import requests

import streamlit as st
from streamlit.components.v1 import html
from streamlit.components.v1 import html as _html_listener
from streamlit_autorefresh import st_autorefresh

from google_auth_oauthlib.flow import Flow
from google.oauth2.credentials import Credentials

from core_functional_modules.google_client import load_valid_token, services, clear_token
from core_functional_modules.config_store import save_config_to_drive
from core_functional_modules.config import Config
from core_functional_modules.logger import StatusLogger
from core_functional_modules.pipeline import run_pipeline
from core_functional_modules.drive_utils import upload_to_drive, get_or_create_subfolder
from core_functional_modules.pipeline_bus import get_pipeline_queue
from core_functional_modules.rebuild_google_workspace import rebuild_google_workspace


# =========================
# Constants & Simple Helpers
# =========================

EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")

err = None

UI_UPLOAD  = "UPLOAD"
UI_READY   = "READY"
UI_RUNNING = "RUNNING"
UI_RESULT = "RESULT"
UI_RESULT_ERROR = "RESULT_ERROR"     # run_pipeline failed
UI_UPLOAD_ERROR = "UPLOAD_ERROR"     # invalid input (1–2 weeks)


PIPE_STATUS_IDLE = "idle"
PIPE_STATUS_RUNNING = "running"
PIPE_STATUS_DONE = "done"
PIPE_STATUS_ERROR = "error"



UI_REBUILD_RUNNING = "REBUILD_RUNNING"
UI_REBUILD_DONE = "REBUILD_DONE"



class UIError(Exception):
    """
    Base class for errors that should show a friendly message
    plus optional technical details.
    """
    user_message: str
    title: str = "Error"

    def __init__(self, user_message: str, *, title: str | None = None):
        super().__init__(user_message)
        self.user_message = user_message
        if title:
            self.title = title


def _split_emails(csv_str: str):
    return [e.strip() for e in (csv_str or "").split(",") if e.strip()]


def _parse_emails(csv_str: str):
    return _split_emails(csv_str)


def _invalid_emails(csv_str: str):
    return [e for e in _parse_emails(csv_str) if not EMAIL_RE.match(e)]


def _analyze_rk_rows(rows):
    """
    Validate the 'Per-Report-Key Recipients' editor rows.
    Returns (issues: List[str], preview_lines: List[str], rk_map: Dict[str, List[str]])
    """
    issues, preview, rk_map = [], [], {}
    seen, dupes = set(), set()

    for idx, r in enumerate(rows or [], start=1):
        raw_key = (r.get("REPORT KEY (ALL CAPS)") or "").strip()
        emails_csv = r.get("Emails (comma)") or ""
        if not raw_key and not emails_csv:
            # allow a blank template row
            continue

        # uppercase flag
        if raw_key != raw_key.upper():
            issues.append(f"Row {idx}: key '{raw_key}' is not ALL CAPS.")

        # duplicate detection
        if raw_key:
            if raw_key in seen:
                dupes.add(raw_key)
            else:
                seen.add(raw_key)

        # email validation
        bads = _invalid_emails(emails_csv)
        if bads:
            issues.append(f"Row {idx}: invalid emails → {', '.join(bads)}")

        # mapping + preview
        if raw_key:
            emails = _parse_emails(emails_csv)
            if emails:
                rk_map[raw_key] = emails
            preview.append(f"{raw_key} → {', '.join(emails) if emails else emails_csv}")

    if dupes:
        issues.append(f"Duplicate keys detected: {', '.join(sorted(dupes))}")
    return issues, preview, rk_map


def _b64url(b: bytes) -> str:
    return base64.urlsafe_b64encode(b).rstrip(b"=").decode("ascii")


def _pkce_pair() -> tuple[str, str]:
    """
    Generate a high-entropy PKCE code_verifier and its S256 code_challenge.
    RFC 7636 requires 43–128 chars; this approach yields a URL-safe value.
    """
    verifier = _b64url(secrets.token_bytes(64))        # ~86 chars, URL-safe, no padding
    challenge = _b64url(hashlib.sha256(verifier.encode("ascii")).digest())
    return verifier, challenge


def _redirect_base() -> str:
    """
    Always return a non-empty redirect base that exactly matches your OAuth client's
    Authorized redirect URI. Prefer Secrets; normalize to one trailing slash.
    """
    base = (st.secrets.get("APP_BASE_URL", "") or "").strip()
    if not base:
        # Fallback to request (often available), still normalized
        try:
            base = (st.request.url_root or "").strip()
        except Exception:
            base = ""
    if not base:
        st.error("OAuth redirect base is not set. Define APP_BASE_URL in Secrets.")
        st.stop()
    return base.rstrip("/") + "/"


def _parse_state(state_b64: str) -> dict:
    # Add padding back for base64 decoding if needed
    padding = "=" * ((4 - len(state_b64) % 4) % 4)
    raw = base64.urlsafe_b64decode(state_b64 + padding)
    return json.loads(raw.decode("utf-8"))


def _infer_media_mime(name: str) -> str:
    n = (name or "").lower()
    if n.endswith(".csv"):
        return "text/csv"
    if n.endswith(".xlsx"):
        return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    return "application/octet-stream"


def _get_drive_service_or_raise(cfg):
    creds = load_valid_token(cfg.SCOPES)
    if not creds:
        raise RuntimeError("Google authorization required. Please sign in first.")
    _sheets, drive, _gmail = services(creds, cfg.HTTP_TIMEOUT_SECONDS)
    return drive


def _rerun():
    try:
        st.rerun()
    except AttributeError:
        st.experimental_rerun()

def reset_to_upload():
    st.session_state.sales_uploaded_ok = False
    st.session_state.vendor_uploaded_ok = False

    st.session_state.sales_selected_name = None
    st.session_state.vendor_selected_name = None

    st.session_state.reset_generation += 1

    st.session_state.sales_selection_generation = None
    st.session_state.vendor_selection_generation = None


    # 🔑 Increment the upload epoch
    st.session_state.upload_epoch += 1
    st.session_state.sales_selected_epoch = None
    st.session_state.vendor_selected_epoch = None

    st.session_state.running_ui_initialized = False
    st.session_state.uploader_version += 1
    st.session_state.ui_phase = UI_UPLOAD

    # Fully clear uploader widget state
    for k in list(st.session_state.keys()):
        if k.startswith("sales_upload_") or k.startswith("vendor_upload_"):
            st.session_state.pop(k, None)



def init_thread_state():
    if "pipeline_thread_started" not in st.session_state:
        st.session_state.pipeline_thread_started = False
    if "pipeline_done" not in st.session_state:
        st.session_state.pipeline_done = False
    if "pipeline_error" not in st.session_state:
        st.session_state.pipeline_error = None
    if "pipeline_thread" not in st.session_state:
        st.session_state.pipeline_thread = None

def init_pipeline_state():
    st.session_state.setdefault("pipe_status", PIPE_STATUS_IDLE)
    st.session_state.setdefault("pipe_result", None)
    st.session_state.setdefault("pipe_finished", False)
    st.session_state.setdefault("pipe_error", None)
    st.session_state.setdefault("pipe_run_id", None)

def reset_pipeline_state():
    # Thread control
    st.session_state.pipeline_thread_started = False
    st.session_state.pipeline_done = False
    st.session_state.pipeline_error = None
    st.session_state.pipeline_thread = None

    # Pipeline result & lifecycle
    st.session_state.pipe_status = PIPE_STATUS_IDLE
    st.session_state.pipe_finished = False
    st.session_state.pipe_result = None
    st.session_state.pipe_error = None
    st.session_state.pipe_run_id = None

    # Timer
    st.session_state._run_start_time = None


def start_run():
    st.session_state.pipe_run_id = str(uuid.uuid4())

    st.session_state.pipe_status = PIPE_STATUS_RUNNING
    st.session_state.pipeline_thread_started = True
    st.session_state.pipeline_done = False
    st.session_state.pipeline_error = None


    st.session_state.pipeline_refresh_key = f"pipeline_refresh_{time.time()}"



def _both_uploads_ok():
    epoch = st.session_state.upload_epoch

    return (
        st.session_state.sales_selected_epoch == epoch
        and st.session_state.vendor_selected_epoch == epoch
    )



def _validate_pipeline_result(result):
    required_attrs = ("location", "timestamp", "elapsed_seconds")
    return (
        result is not None
        and all(hasattr(result, attr) for attr in required_attrs)
    )

def apply_per_run_config(
    *,
    cfg,
    to,
    cc,
    error_recipients,
    use_all,
    report_keys,
    include_full,
    send_full,
    email_mgr,
    calc_id,
    incoming_id,
    mgr_folder,
    order_folder,
    error_folder,
    user_folder,
    redirect_port,
    gid_mgr,
    gid_order,
    gid_err,
    gid_bev_err,
    loc_sheet,
    loc_range,
    update_range,
    tz,
    tfmt,
    output_ttl,
    failed_input_ttl,
    user_ttl,
    use_rollover,
    start_dow,
    end_dow,
    soft_cases_enabled,
    soft_cases_threshold,
    BEV_MAPPING_LINK,
    edited_rows,
):
    # ------------------------------------------------------------
    # Basic email & behavior flags
    # ------------------------------------------------------------
    cfg.TO_RECIPIENTS = _split_emails(to)
    cfg.CC_RECIPIENTS = _split_emails(cc)
    cfg.ERROR_RECIPIENTS = _split_emails(error_recipients)

    cfg.USE_ALL_REPORT_KEYS = use_all
    cfg.REPORT_KEY_RUN_LIST = [
        s.strip().upper()
        for s in (report_keys or "").split(",")
        if s.strip()
    ]

    cfg.INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL = bool(include_full)
    cfg.SEND_SEPARATE_FULL_ORDER_EMAIL = bool(send_full)
    cfg.EMAIL_MANAGER_REPORT = bool(email_mgr)

    # ------------------------------------------------------------
    # IDs, folders, sheets
    # ------------------------------------------------------------
    cfg.CALC_SPREADSHEET_ID = calc_id
    cfg.INCOMING_FOLDER_ID = incoming_id
    cfg.MANAGER_REPORT_FOLDER_ID = mgr_folder
    cfg.ORDER_REPORT_FOLDER_ID = order_folder
    cfg.ERROR_REPORT_FOLDER_ID = error_folder
    cfg.USER_FOLDER_ID = user_folder
    cfg.REDIRECT_PORT = int(redirect_port)

    cfg.GID_MANAGER_PDF = gid_mgr
    cfg.GID_ORDER_CSV = gid_order
    cfg.GID_ERROR_REPORT = gid_err
    cfg.GID_BEV_ERRORS = gid_bev_err

    cfg.LOCATION_SHEET_TITLE = loc_sheet
    cfg.LOCATION_NAMED_RANGE = loc_range
    cfg.TEMPLATE_UPDATE_RANGE = update_range

    cfg.TIMESTAMP_TZ = tz
    cfg.TIMESTAMP_FMT = tfmt

    # ------------------------------------------------------------
    # Lifecycle / TTL
    # ------------------------------------------------------------
    cfg.OUTPUT_TIME_TO_LIFE = int(output_ttl)
    cfg.FAILED_INPUT_TIME_TO_LIFE = int(failed_input_ttl)
    cfg.USER_TIME_TO_LIFE = int(user_ttl)

    # ------------------------------------------------------------
    # Date & integrity controls
    # ------------------------------------------------------------
    cfg.USE_AUTO_ROLLOVER_IF_ONE_WEEK = bool(use_rollover)
    cfg.START_DAY_OF_WEEK = start_dow
    cfg.END_DAY_OF_WEEK = end_dow

    # ------------------------------------------------------------
    # Soft cases alerting
    # ------------------------------------------------------------
    cfg.SOFT_CASES_ALERT_ENABLED = bool(soft_cases_enabled)
    cfg.SOFT_CASES_ALERT_THRESHOLD = int(soft_cases_threshold)

    # ------------------------------------------------------------
    # External links
    # ------------------------------------------------------------
    cfg.BEV_MAPPING_LINK = BEV_MAPPING_LINK

    # ------------------------------------------------------------
    # Per‑report‑key recipients
    # ------------------------------------------------------------
    rk_map: dict[tuple[str | None, str | None, str | None], list[str]] = {}

    for r in edited_rows or []:
        store = (r.get("Store (optional)") or "").strip().upper() or None
        key = (r.get("Report Key (optional)") or "").strip().upper() or None
        sub_key = (r.get("Sub-Report Key (optional)") or "").strip().upper() or None

        emails = [
            e.strip()
            for e in (r.get("Emails (comma)") or "").split(",")
            if e.strip()
        ]

        if not emails or not (store or key or sub_key):
            continue

        rk_map[
            (
                clean_tag(store) if store else None,
                clean_tag(key) if key else None,
                clean_tag(sub_key) if sub_key else None,
            )
        ] = emails

    cfg.REPORT_KEY_RECIPIENTS = rk_map

    return rk_map


def github_merge(from_branch: str, to_branch: str) -> bool:
    """
    Merge `from_branch` into `to_branch` using GitHub's API.

    Returns True on success, False on failure.
    All secrets are fetched internally with safe failure.
    """

    try:
        token = (st.secrets.get("GITHUB_TOKEN", "") or "").strip()
        owner = (st.secrets.get("GITHUB_OWNER", "") or "").strip()
        repo = (st.secrets.get("GITHUB_REPO", "") or "").strip()

        missing = []
        if not token:
            missing.append("GITHUB_TOKEN")
        if not owner:
            missing.append("GITHUB_OWNER")
        if not repo:
            missing.append("GITHUB_REPO")

        if missing:
            st.error(f"Missing GitHub secrets: {', '.join(missing)}")
            return False

        url = f"https://api.github.com/repos/{owner}/{repo}/merges"
        headers = {
            "Authorization": f"Bearer {token}",
            "Accept": "application/vnd.github+json",
        }
        payload = {
            "base": to_branch,
            "head": from_branch,
            "commit_message": f"Merge {from_branch} → {to_branch} (via Streamlit)",
        }

        resp = requests.post(url, headers=headers, json=payload, timeout=20)

        if resp.status_code == 201:
            return True

        if resp.status_code == 409:
            st.error("❌ Merge conflict detected. Resolve manually.")
            return False

        st.error(f"GitHub merge failed ({resp.status_code}): {resp.text}")
        return False

    except Exception as e:
        st.error(f"Unexpected merge error: {e}")
        return False


# =========================
# OAuth (Web / PKCE)
# =========================

def start_web_oauth(scopes):
    """
    Build an authorization URL that:
      - uses a stable redirect_uri (from Secrets)
      - uses explicit PKCE (S256)
      - embeds the code_verifier inside the state (base64url(JSON))
    """
    cfg = json.loads(st.secrets["GOOGLE_CREDENTIALS"])
    redirect = _redirect_base()

    # Explicit PKCE (stateless across redirect)
    code_verifier, code_challenge = _pkce_pair()

    # CSRF token + verifier encoded into state that Google will return unchanged.
    state_obj = {
        "csrf": _b64url(secrets.token_bytes(16)),
        "v": code_verifier,
        "r": redirect,
    }
    state_b64 = _b64url(json.dumps(state_obj).encode("utf-8"))

    flow = Flow.from_client_config(cfg, scopes=scopes, redirect_uri=redirect)
    auth_url, _ = flow.authorization_url(
        access_type="offline",
        include_granted_scopes="true",
        prompt="consent",
        state=state_b64,
        code_challenge=code_challenge,
        code_challenge_method="S256",
    )

    # Keep only minimal context; state carries the verifier.
    st.session_state["_oauth_redirect"] = redirect
    return auth_url


def finish_web_oauth(code: str, state_b64: str, scopes):
    """
    Recreate a Flow with the same redirect_uri and exchange code + code_verifier for tokens.
    (No UI side effects here.)
    """
    cfg = json.loads(st.secrets["GOOGLE_CREDENTIALS"])
    state_obj = _parse_state(state_b64)
    code_verifier = state_obj.get("v")
    redirect = state_obj.get("r") or st.session_state.get("_oauth_redirect") or _redirect_base()

    if not code_verifier:
        st.error("OAuth state did not include a PKCE code_verifier.")
        st.stop()

    flow = Flow.from_client_config(cfg, scopes=scopes, redirect_uri=redirect)
    flow.fetch_token(code=code, code_verifier=code_verifier)

    creds = flow.credentials
    with open("token.json", "w") as f:
        f.write(creds.to_json())
    return creds

def clean_tag(s: str) -> str:
    import re
    s = (s or "").strip()
    s = re.sub(r"[^A-Za-z0-9._-]+", "-", s)
    return s.strip("-") or "UNKNOWN"

# --- OAuth Redirect Handler (Web/PKCE only) ---


# =========================
# UI Sections
# =========================

def render_upload_card(cfg):
    with st.container(border=True):
        st.subheader("Upload Required Input Files")
        st.caption("Both Sales Data and Vendor Price Data are required. Sales Data must be 1 or 2 complete weeks.")

        up_col, _, upbtn_col = st.columns([4, 1, 1])

        current_gen = st.session_state.reset_generation

        with up_col:
            sales_key = f"sales_upload_v{st.session_state.uploader_version}"
            sales_file = st.file_uploader(
                "Upload Sales Data",
                type=["xlsx", "csv"],
                key=sales_key,
                help="Go to Modisoft -> Sales -> Live Items, Select Stores & Dates, Download as Excel"
            )
            
            if sales_file:
                if st.session_state.sales_selection_generation != current_gen:
                    st.session_state.sales_selected_name = sales_file.name
                    st.session_state.sales_selected_epoch = st.session_state.upload_epoch
                    st.session_state.sales_selection_generation = current_gen


            vendor_key = f"vendor_upload_v{st.session_state.uploader_version}"
            vendor_file = st.file_uploader(
                "Upload Vendor Price Data",
                type=["xlsx", "csv"],
                key=vendor_key,
                help="Go to Modisoft -> Products -> Price Book , Download as Excel"
            )
            
            if vendor_file:
                if st.session_state.vendor_selection_generation != current_gen:
                    st.session_state.vendor_selected_name = vendor_file.name
                    st.session_state.vendor_selected_epoch = st.session_state.upload_epoch
                    st.session_state.vendor_selection_generation = current_gen

        
        with upbtn_col:
            st.markdown('<div class="ft-right-btn">', unsafe_allow_html=True)

            upload_clicked = st.button(
                "⬆️ Upload Now",
                width="stretch",
                type="primary",
                disabled= (not _both_uploads_ok()),
                key="upload_submit",
            )

            st.markdown('</div>', unsafe_allow_html=True)


        # --- Handle the upload action immediately ---
        if upload_clicked:
            if not cfg.INCOMING_FOLDER_ID:
                st.session_state.upload_error = "Incoming Folder ID is empty."
                st.session_state.ui_phase = UI_UPLOAD_ERROR
                _rerun()

            if sales_file is None or vendor_file is None:
                st.session_state.upload_error = (
                    "Both Sales Data and Vendor Price Data are required."
                )
                st.session_state.ui_phase = UI_UPLOAD_ERROR
                _rerun()

            try:
                st.warning("Uploading files to google drive...")
                drive = _get_drive_service_or_raise(cfg)
                

                # --- Resolve user ---
                me = drive.about().get(
                    fields="user(emailAddress,permissionId,displayName)"
                ).execute().get("user", {})

                user_email = (me or {}).get("emailAddress") or "UNKNOWN_USER"

                # --- Per-user folder ---
                user_folder = get_or_create_subfolder(
                    drive,
                    cfg.INCOMING_FOLDER_ID,
                    user_email,
                )

                sales_folder = get_or_create_subfolder(
                    drive,
                    user_folder["id"],
                    "01 Sales Data Inputs",
                )

                vendor_folder = get_or_create_subfolder(
                    drive,
                    user_folder["id"],
                    "02 Vendor Price Data Inputs",
                )

                # --- Upload SALES ---
                sales_created = upload_to_drive(
                    drive,
                    data=sales_file.getvalue(),
                    name=f"{os.path.splitext(sales_file.name)[0]} (Sales Data via UI)",
                    mime=_infer_media_mime(sales_file.name),
                    folder_id=sales_folder["id"],
                    to_sheet=True,
                )

                # --- Upload VENDOR ---
                vendor_created = upload_to_drive(
                    drive,
                    data=vendor_file.getvalue(),
                    name=f"{os.path.splitext(vendor_file.name)[0]} (Vendor Price Data via UI)",
                    mime=_infer_media_mime(vendor_file.name),
                    folder_id=vendor_folder["id"],
                    to_sheet=True,
                )

                st.session_state.ui_phase = UI_READY
                _rerun()

            except Exception as e:
                st.session_state.upload_error = f"Upload failed: {e}"
                st.session_state.ui_phase = UI_UPLOAD_ERROR
                _rerun()
    
def render_run_options(cfg):
    run_form_wrapper_classes = "ft-card ft-row"

    # A file is "dirty" if the user has selected something not yet uploaded
    files_dirty = (
        st.session_state.sales_selected_name is not None
        or st.session_state.vendor_selected_name is not None
    )

    # Have we successfully uploaded both files?

    # OPEN the wrapper with real HTML (no entities)
    st.markdown(f'<div class="{run_form_wrapper_classes}">', unsafe_allow_html=True)



    with st.form("run_form"):
        # Header row uses the same columns to align the Run button with Upload button above
        tl, _, col_run = st.columns([4, 1, 1])
        with tl:
            st.subheader("Run Options")
            st.caption("Configure email behavior and report keys. Use **Advanced** for IDs/GIDs/timezone.")

        # --- Unified gating logic ---
        # A) If a file is currently selected but not uploaded -> disable Run
        # B) If no file selected and we have a prior successful upload -> enable Run
        # C) Otherwise (no prior upload or ambiguous state) -> disable Run

        with col_run:
            # Right-align and full-width, matching Upload Now
            st.markdown('<div class="ft-right-btn">', unsafe_allow_html=True)
            submitted = st.form_submit_button(
                "▶️ Run Pipeline",
                width='stretch',
                disabled=False,
                type="primary",
                key="run_submit"
            )
            st.markdown('</div>', unsafe_allow_html=True)

        # ----- Main options -----

        # Recipients
        st.markdown("##### Recipients")
        col1, col2 = st.columns([1, 1])
        with col1:
            to = st.text_input(
                "To (comma)", value=",".join(cfg.TO_RECIPIENTS or []),
                help="Fallback recipients for Manager & Order emails."
            )
        with col2:
            cc = st.text_input(
                "CC (comma)", value=",".join(cfg.CC_RECIPIENTS or []),
                help="Optional CC added to all emails."
            )

        # Report Keys
        st.markdown("##### Report Keys")
        colk1, colk2, colk3 = st.columns([1, 1, 2])
        with colk1:
            use_all = st.toggle(
                "Use all keys from CSV",
                value=cfg.USE_ALL_REPORT_KEYS,
                help="ON: process every key found. OFF: only the keys you list."
            )

        with colk2:
            pass

        with colk3:
            report_keys = st.text_input(
                "Keys to run (comma)",
                value=",".join(cfg.REPORT_KEY_RUN_LIST or []),
                help="Used when 'Use all keys' is OFF. For Sub_Report_Keys use Report_Key-Sub_Report_Key. Example: COFFEE,GROCERY,BEV-7UP"
            )

        # General Behavior
        st.markdown("##### General Behavior")
        cole1, cole2, cole3, cole4 = st.columns([1, 1, 1, 1])
        with cole1:
            include_full = st.toggle(
                "Attach FULL order in each email",
                value=cfg.INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL
            )
        with cole2:
            send_full = st.toggle(
                "Send separate FULL order email",
                value=cfg.SEND_SEPARATE_FULL_ORDER_EMAIL
            )
        with cole3:
            email_mgr = st.toggle(
                "Email Manager Report",
                value=getattr(cfg, "EMAIL_MANAGER_REPORT", True),
                help="When ON, the Manager Report email is sent. When OFF, it is skipped."
            )
        with cole4:
            use_rollover = st.toggle(
                    'Use auto-rollover for single week uploads',
                    value=cfg.USE_AUTO_ROLLOVER_IF_ONE_WEEK,
                    help='If this is on, when only 1 week is uploaded, the most recent previously uploaded data will become the "Last Week" data; If this is off then the "Last Week" data will be left blank'
                )

        # Per-Report-Key Recipients
        with st.expander("Per-Report-Key Recipients (optional)", expanded=False):

            st.caption("""
              Map **Store, Report Key → Emails (comma)**.
              
              **Email Delivery Priority:**  
              - `(Store, Key, Sub-Key)` → 1st priority set of emails  
              - `(Store, , Sub-Key)`    → 2nd priority set of emails  
              - `(, , Sub-Key)`         → 3rd priority set of emails  
              - `(Store, Key, )`        → 4th priority set of emails  
              - `(, Key, )`             → 5th priority set of emails  
              - `(Store, , )`           → 6th priority set of emails  
              - If not defined, it will use the default set of emails in `To (comma)` field above
              """)
        
            rows = []
        
            if cfg.REPORT_KEY_RECIPIENTS:
                for (store, key, sub_key), emails in cfg.REPORT_KEY_RECIPIENTS.items():
                    rows.append({
                        "Store (optional)": store or "",
                        "Report Key (optional)": key or "",
                        "Sub-Report Key (optional)": sub_key or "",
                        "Emails (comma)": ",".join(emails or [])
                    })
            else:
                rows = [{
                    "Store (optional)": "",
                    "Report Key (optional)": "",
                    "Sub-Report Key (optional)": "",
                    "Emails (comma)": ""
                }]
        
            edited_rows = st.data_editor(
                rows,
                num_rows="dynamic",
                width='stretch',
                key="rk_editor",
            )
        
            rk_map = {}
            rk_preview = []
            rk_issues = []
        
            for i, r in enumerate(edited_rows):
        
                store = (r.get("Store (optional)") or "").strip().upper()
                key = (r.get("Report Key (optional)") or "").strip().upper()
                sub_key = (r.get("Sub-Report Key (optional)") or "").strip().upper() or None
                emails_raw = (r.get("Emails (comma)") or "").strip()
        
                emails = [e.strip() for e in emails_raw.split(",") if e.strip()]
        
                store_val = store if store else None
                key_val = key if key else None
                sub_val = sub_key if sub_key else None
        
                if emails and not (store_val or key_val or sub_key):
                    rk_issues.append(f"Row {i+1}: Must include Store, Key, or both.")
                    continue
        
                if (store_val or key_val or sub_key) and not emails:
                    rk_issues.append(f"Row {i+1}: Missing email(s).")
                    continue

                store_tag = clean_tag(store_val)
                key_tag = clean_tag(key_val)
                sub_tag = clean_tag(sub_val)
        
                rk_map[(store_tag, key_tag, sub_tag)] = emails
        
                rk_preview.append(f"{(store_val, key_val, sub_val)} -> {emails}")
        
            #if rk_preview:
            #    with st.expander("Recipient mapping preview"):
            #        st.code("\n".join(rk_preview), language="text")
        
            #if rk_issues:
            #    st.warning("Recipient configuration issues:\n\n- " + "\n- ".join(rk_issues))

        # Advanced
        with st.expander("Advanced", expanded=False):
            tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(["Folders", "Files & Links", "Ranges", "Timing", "Lifecycle", "Technical"])
            with tab1:
                ga1, ga2 = st.columns([1, 1])
                with ga1:
                    st.markdown("###### Input Folders")
                    user_folder = st.text_input("User Calculations Folder ID", value=cfg.USER_FOLDER_ID,
                        help="The file ID of the google drive folder that user workhorse files should be stored.")
                    incoming_id = st.text_input("Incoming Folder ID", value=cfg.INCOMING_FOLDER_ID,
                        help="The file ID of the google drive folder that user input folders & files should be stored.")
                    
                with ga2:
                    st.markdown("###### Output Folders")
                    order_folder = st.text_input("Order Report Folder ID", value=cfg.ORDER_REPORT_FOLDER_ID,
                        help="The file ID of the google drive folder that order report csv output files should be stored.")
                    mgr_folder = st.text_input("Manager Report Folder ID", value=cfg.MANAGER_REPORT_FOLDER_ID,
                        help="The file ID of the google drive folder that manager report pdf output files should be stored.")
                    error_folder = st.text_input("Error Report Folder ID", value=cfg.ERROR_REPORT_FOLDER_ID,
                        help="The file ID of the google drive folder that error report csv output files should be stored.")

            with tab2:
                gb1, gb2 = st.columns([1, 1])
                with gb1:
                    calc_id = st.text_input("Master Calculations Spreadsheet ID", value=cfg.CALC_SPREADSHEET_ID,
                        help="The file ID of the Master Calculations google sheets file that user workhorse files should be based off of.")
                with gb2:
                    BEV_MAPPING_LINK = st.text_input("BEV Mapping Link", value=cfg.BEV_MAPPING_LINK,
                        help="The url to the live, editable BEV Sub-Key Mapping google sheets file.")
            
            with tab3:
                gc1, gc2 = st.columns([1, 1])
                with gc1:
                    st.markdown("###### GIDs")
                    gid_mgr = st.text_input("Manager Report gid", value=str(cfg.GID_MANAGER_PDF),
                        help="The GID of the Manager Report Tab within the Master Calculations Sheet that should be used for outputs.")
                    gid_err = st.text_input("Error Report gid", value=str(cfg.GID_ERROR_REPORT),
                        help="The GID of the Error Report Tab within the Master Calculations Sheet that should be used for outputs.")
                    gid_order = st.text_input("Order Report gid", value=str(cfg.GID_ORDER_CSV),
                        help="The GID of the Order Report Tab within the Master Calculations Sheet that should be used for outputs.")
                    gid_bev_err = st.text_input("Unassigned Beverages Report gid", value=str(cfg.GID_ORDER_CSV),
                        help="The GID of the UB Report Tab within the Master Calculations Sheet that should be used for outputs.")
                with gc2:
                    st.markdown("###### Titles")
                    loc_sheet = st.text_input("Named Range Sheet Title", value=cfg.LOCATION_SHEET_TITLE,
                        help="The Sheet Title of the tab within the Master Calculations Sheet where the below named ranges exist.")
                    loc_range = st.text_input("Location Named Range", value=cfg.LOCATION_NAMED_RANGE,
                        help="The named range within the Master Calculations Sheet that refrences the cleaned location(s) name.")
                    update_range = st.text_input("Update Timestamp Range", value=cfg.TEMPLATE_UPDATE_RANGE,
                        help="The named range within the Master Calculations Sheet that refrences the last time the template was updated.")

            with tab4:
                gd1, gd2 = st.columns([1, 1])
                with gd1:
                    st.markdown("###### Data Integrity Controls")
                    _days = ["Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Any"]
                    start_dow = st.selectbox(
                        "Start day of week", _days, index=_days.index(cfg.START_DAY_OF_WEEK),
                        help="The day of week that the uploaded data should start at, any other day will raise an error."
                    )
                    end_dow = st.selectbox(
                        "End day of week", _days, index=_days.index(cfg.END_DAY_OF_WEEK),
                        help="The day of week that the uploaded data should end at, any other day will raise an error."
                    )
                with gd2:
                    st.markdown("###### Formatting")                                
                    tz = st.text_input("Timestamp Timezone", value=cfg.TIMESTAMP_TZ,
                        help="The timezone that should be used in all timestamps.")
                    tfmt = st.text_input("Timestamp Format", value=cfg.TIMESTAMP_FMT,
                        help="The format that should be used in all timestamps.")
            with tab5:
                ge1, ge2 = st.columns([1, 1])
                with ge1:
                    st.markdown("###### One-Time Use Files")
                    failed_input_ttl = st.number_input(
                        "Failed Input Time-To-Life (days)",
                        min_value=0,
                        max_value=3650,
                        value=int(cfg.FAILED_INPUT_TIME_TO_LIFE),
                        help="Delete old unused incoming files older than this many days after a successful run."
                        )
                    output_ttl = st.number_input(
                        "Output Time-To-Life (days)",
                        min_value=0,
                        max_value=3650,
                        value=int(cfg.OUTPUT_TIME_TO_LIFE),
                        help="Delete output files older than this many days after a successful run."
                        )
                with ge2:
                    st.markdown("###### Recurring Use Files")
                    user_ttl = st.number_input(
                        "User Calculations Time-To-Life (days)",
                        min_value=0,
                        max_value=3650,
                        value=int(cfg.USER_TIME_TO_LIFE),
                        help="Delete old unused user calculations files older than this many days after a successful run."
                        )
            
            with tab6:

                #row1
                gf1_1, gf2_1 = st.columns([1, 1])
                with gf1_1:                
                    soft_cases_enabled = st.toggle(
                        "Alert on large case quantities",
                        value=cfg.SOFT_CASES_ALERT_ENABLED,
                        help="Send a technical alert if any FULL order rows exceed the cases threshold"
                    )

                with gf2_1:
                    error_recipients = st.text_input(
                        "Technical Support Email(s) (comma)",
                        value=",".join(cfg.ERROR_RECIPIENTS or []),
                        help="If errors arise such as missing items in the Vendor Price Book, the error report will be sent here."
                    )

                #row2
                gf1_2, gf2_2 = st.columns([1, 1])
                with gf1_2:                
                    soft_cases_threshold = st.number_input(
                        "Cases-to-order alert threshold",
                        min_value=1,
                        max_value=1000,
                        value=int(cfg.SOFT_CASES_ALERT_THRESHOLD),
                        help="Any FULL order line above this number will trigger a soft alert"
                    )

                with gf2_2:
                    raw_redirect_port = int(cfg.REDIRECT_PORT) if str(cfg.REDIRECT_PORT).isdigit() else 0
                    redirect_port = st.number_input(
                        "Redirect Port (0 = auto)",
                        min_value=0, max_value=65535,
                        value=raw_redirect_port if raw_redirect_port in (0, *range(1024, 65536)) else 0,
                        help="Use 0 to auto-pick a free port. Otherwise choose 1024–65535."
                    )

                
        save_defaults_clicked = st.form_submit_button("💾 Save as defaults", type="secondary", help="Persist current settings for future sessions")


        # ----- Submission handling -----

        if save_defaults_clicked:
            try:
                rk_map = apply_per_run_config(
                    cfg=cfg,
                    to=to,
                    cc=cc,
                    error_recipients=error_recipients,
                    use_all=use_all,
                    report_keys=report_keys,
                    include_full=include_full,
                    send_full=send_full,
                    email_mgr=email_mgr,
                    calc_id=calc_id,
                    incoming_id=incoming_id,
                    mgr_folder=mgr_folder,
                    order_folder=order_folder,
                    error_folder=error_folder,
                    user_folder=user_folder,
                    redirect_port=redirect_port,
                    gid_mgr=gid_mgr,
                    gid_order=gid_order,
                    gid_err=gid_err,
                    gid_bev_err=gid_bev_err,
                    loc_sheet=loc_sheet,
                    loc_range=loc_range,
                    update_range=update_range,
                    tz=tz,
                    tfmt=tfmt,
                    output_ttl=output_ttl,
                    failed_input_ttl=failed_input_ttl,
                    user_ttl=user_ttl,
                    use_rollover=use_rollover,
                    start_dow=start_dow,
                    end_dow=end_dow,
                    soft_cases_enabled=soft_cases_enabled,
                    soft_cases_threshold=soft_cases_threshold,
                    BEV_MAPPING_LINK=BEV_MAPPING_LINK,
                    edited_rows=edited_rows,
                )

                # Ensure we have a user token first
                creds = load_valid_token(cfg.SCOPES)
                if not creds:
                    st.error("Not authenticated. Please complete Google sign‑in first (top of page).")
                else:

                    # Drive service
                    _sheets, drive, _gmail = services(creds, cfg.HTTP_TIMEOUT_SECONDS)

                    # What we persist
                    drive_defaults = cfg.to_drive_defaults()

                    DEV_ENVIRONMENT = st.secrets.get("DEV_ENVIRONMENT", False)
                    DEV_CONFIG_FILE_ID = (st.secrets.get("DEV_CONFIG_FILE_ID", "") or "").strip()
                    CONFIG_FILE_ID = (st.secrets.get("CONFIG_FILE_ID", "") or "").strip()

                    # Decide where to SAVE
                    if DEV_ENVIRONMENT:
                        save_target_id = DEV_CONFIG_FILE_ID or None
                    else:
                        save_target_id = CONFIG_FILE_ID or None

                    new_id = save_config_to_drive(
                        drive,
                        drive_defaults,
                        file_id=save_target_id
                    )

                    # DEV auto-bootstrap case
                    if DEV_ENVIRONMENT and not DEV_CONFIG_FILE_ID:
                        st.success("✅ Created new DEV config file.")
                        st.info(
                            "Add this to Streamlit secrets as DEV_CONFIG_FILE_ID:\n\n"
                            f"`{new_id}`"
                        )
                    else:
                        st.success(f"✅ Defaults saved (file id: {new_id})")

            except Exception as e:
                st.error(f"Failed to save defaults to Drive: {e}")

        if submitted:
            # Apply per-run config
            rk_map = apply_per_run_config(
                cfg=cfg,
                to=to,
                cc=cc,
                error_recipients=error_recipients,
                use_all=use_all,
                report_keys=report_keys,
                include_full=include_full,
                send_full=send_full,
                email_mgr=email_mgr,
                calc_id=calc_id,
                incoming_id=incoming_id,
                mgr_folder=mgr_folder,
                order_folder=order_folder,
                error_folder=error_folder,
                user_folder=user_folder,
                redirect_port=redirect_port,
                gid_mgr=gid_mgr,
                gid_order=gid_order,
                gid_err=gid_err,
                gid_bev_err=gid_bev_err,
                loc_sheet=loc_sheet,
                loc_range=loc_range,
                update_range=update_range,
                tz=tz,
                tfmt=tfmt,
                output_ttl=output_ttl,
                failed_input_ttl=failed_input_ttl,
                user_ttl=user_ttl,
                use_rollover=use_rollover,
                start_dow=start_dow,
                end_dow=end_dow,
                soft_cases_enabled=soft_cases_enabled,
                soft_cases_threshold=soft_cases_threshold,
                BEV_MAPPING_LINK=BEV_MAPPING_LINK,
                edited_rows=edited_rows,
            )

            # --- ADD: warnings before kicking off the run ---
            if not cfg.USE_ALL_REPORT_KEYS and not cfg.REPORT_KEY_RUN_LIST:
                st.session_state.upload_error = (
                    "No report keys selected. Either enable 'Use all keys' "
                    "or provide explicit report keys."
                )
                st.session_state.ui_phase = UI_UPLOAD_ERROR
                _rerun()

            if not cfg.TO_RECIPIENTS and not cfg.DEFAULT_TO_RECIPIENTS and not rk_map:
                st.session_state.run_error = (
                    "No email recipients defined. At least one recipient is required."
                )
                st.session_state.ui_phase = UI_RESULT_ERROR
                _rerun()

            if rk_issues:
                st.session_state.run_error = (
                    "Invalid per‑report‑key recipient configuration:\n\n"
                    + "\n".join(rk_issues)
                )
                st.session_state.ui_phase = UI_RESULT_ERROR
                _rerun()

            
            # All validation must already be done
            st.session_state._run_start_time = None
            st.session_state.ui_phase = UI_RUNNING

            # Start pipeline in background (one time)
            if not st.session_state.pipeline_thread_started:
                start_run()

                t = threading.Thread(
                    target=run_pipeline_controller,
                    args=(cfg, st.session_state.pipe_run_id),
                    daemon=True
                )
                st.session_state.pipeline_thread = t
                t.start()

            _rerun()
            # --- END ADD ---

def run_pipeline_controller(cfg, run_id):

    logger = StatusLogger(
        print_to_console=True,
        file_path="last_run.log",
        overwrite=True,
    )

    try:
        result = run_pipeline(cfg, logger=logger)
        get_pipeline_queue().put(
            (run_id, PIPE_STATUS_DONE, result)
        )
    except Exception as e:
        get_pipeline_queue().put(
            (
                run_id,
                PIPE_STATUS_ERROR,
                {
                    "type": type(e).__name__,
                    "user_message": str(e),
                    "traceback": traceback.format_exc(),
                },
            )
        )

    finally:
        logger.close()


def render_running_status(cfg):
    import time
    import os
    import queue
    import streamlit as st
    from streamlit_autorefresh import st_autorefresh

    # ------------------------------------------------------------
    # Poll queue FIRST (edge-triggered)
    # ------------------------------------------------------------
    
    q = get_pipeline_queue()
    
    while True:
        try:
            run_id, status, payload = q.get_nowait()
        except queue.Empty:
            break

        if run_id != st.session_state.get("pipe_run_id"):
            continue

        if status == PIPE_STATUS_DONE:
            st.session_state.pipe_result = payload
            st.session_state.pipe_status = PIPE_STATUS_DONE
            st.session_state.pipe_finished = True

        elif status == PIPE_STATUS_ERROR:
            st.session_state.pipe_error = payload
            st.session_state.run_error = payload
            st.session_state.pipe_status = PIPE_STATUS_ERROR
            st.session_state.pipe_finished = True
    
    if st.session_state.pipe_finished:
        if st.session_state.pipe_status == PIPE_STATUS_DONE:
            st.session_state.ui_phase = UI_RESULT
        elif st.session_state.pipe_status == PIPE_STATUS_ERROR:
            st.session_state.ui_phase = UI_RESULT_ERROR
        _rerun()

    # ------------------------------------------------------------
    # ALWAYS render something
    # ------------------------------------------------------------
    with st.status("Running pipeline…", expanded=True):

        # ----- Bulletproof timer -----
        start_time = st.session_state.get("_run_start_time")
        if not isinstance(start_time, (int, float)):
            start_time = time.perf_counter()
            st.session_state._run_start_time = start_time

        elapsed = int(time.perf_counter() - start_time)
        h, m, s = elapsed // 3600, (elapsed % 3600) // 60, elapsed % 60
        st.markdown(f"**Elapsed:** `{h:02d}:{m:02d}:{s:02d}`")

        # ----- Log tail -----
        if os.path.exists("last_run.log"):
            try:
                with open("last_run.log", "r", encoding="utf-8") as f:
                    st.code("".join(f.readlines()[-8:]), language="text")
            except Exception:
                st.markdown("*Waiting for logs…*")
        else:
            st.markdown("*Waiting for logs…*")

    # ------------------------------------------------------------
    # LIVE status check (DO NOT CACHE THIS)
    # ------------------------------------------------------------
    status = st.session_state.get("pipe_status")

    # ✅ This is what keeps the UI alive
    
    if (
        st.session_state.pipe_status == PIPE_STATUS_RUNNING
        and not st.session_state.pipe_finished
    ):
        st_autorefresh(
            interval=1000,
            key=f"pipeline_tick_{st.session_state.pipe_run_id}",
        )



def render_results(cfg):
    result = st.session_state.get("pipe_result")

    
    if not _validate_pipeline_result(result):
        st.error(
            "Run completed, but the pipeline did not return a valid result object."
        )
        if os.path.exists("last_run.log"):
            with open("last_run.log", "rb") as f:
                st.download_button(
                    "⬇️ Download log",
                    f.read(),
                    file_name="last_run.log",
                    mime="text/plain",
                )


    with st.container(border=True):
        st.subheader("✅ Run Complete")

        st.write("### Outputs")
        col1, col2, col3 = st.columns(3)

        col1.metric("Location", result.location)
        col2.metric("Timestamp", result.timestamp)
        col3.metric(
            "Elapsed",
            f"{result.elapsed_seconds//3600:02d}:"
            f"{(result.elapsed_seconds%3600)//60:02d}:"
            f"{result.elapsed_seconds%60:02d}"
        )

        if getattr(result, "manager_pdf_link", None):
            st.success(f"Manager PDF: {result.manager_pdf_link}")
        if getattr(result, "full_order_link", None):
            st.success(f"Full Order Sheet: {result.full_order_link}")

    
        if os.path.exists("last_run.log"):
            with open("last_run.log", "rb") as f:
                st.session_state["last_run_log"] = f.read()
                st.session_state["last_run_timestamp"] = result.timestamp


        if "last_run_log" in st.session_state:
            st.download_button(
                "⬇️ Download full log (last_run.log)",
                st.session_state["last_run_log"],
                file_name=f"last_run_{st.session_state['last_run_timestamp']}.log",
                mime="text/plain",
                width='stretch'
                )

def render_rebuild_status():
    if st.session_state.ui_phase == UI_REBUILD_RUNNING:
        with st.container(border=True):
            st.subheader("🔄 Rebuilding Google Workspace")
            st.info("Workspace rebuild is in progress. Please wait…")

    elif st.session_state.ui_phase == UI_REBUILD_DONE:
        with st.container(border=True):
            if st.session_state.rebuild_error:
                st.error("❌ Workspace rebuild failed")
                st.code(st.session_state.rebuild_error)
            else:
                st.success("✅ Workspace rebuilt successfully")

                if st.session_state.rebuild_result:
                    st.json(st.session_state.rebuild_result)

            if st.button("⬅️ Return to app"):
                st.session_state.rebuild_result = None
                st.session_state.rebuild_error = None
                st.session_state.ui_phase = UI_UPLOAD


def trigger_rebuild(cfg):
    try:
        st.session_state.ui_phase = UI_REBUILD_RUNNING

        result = rebuild_google_workspace(cfg)

        st.session_state.rebuild_result = result
        st.session_state.rebuild_error = None
    except Exception as e:
        st.session_state.rebuild_error = traceback.format_exc()
        st.session_state.rebuild_result = None
    finally:
        st.session_state.ui_phase = UI_REBUILD_DONE
        _rerun()

def render_sidebar(cfg):
    with st.sidebar:
        st.header("Utilities")

        # --- Existing buttons ---
        if st.button("Google Sign Out", type="secondary", width='stretch'):
            clear_token()
            for key in ["auth_required", "oauth_flow", "oauth_url", "auth_checked"]:
                if key in st.session_state:
                    del st.session_state[key]
            _rerun()

        st.link_button("Add Users to App", "https://console.cloud.google.com/auth/audience?project=favtripdev", width='stretch')
        st.link_button("Open Google Drive", "https://drive.google.com/drive/u/6/folders/1fhzbq0r8iugIJb9t-EQOdHGNvlr9gLT5", width='stretch')
        st.link_button("Open Modisoft", "https://insights.modisoft.com/account/logon", width='stretch')
        st.link_button("Open Bev Mapping File", cfg.BEV_MAPPING_LINK, width='stretch')

        if False:
            st.checkbox(
                "Offer full log download",
                key="offer_log_download",
                help="If enabled, a 'Download last_run.log' button appears when a run finishes."
            )


        # =============================================================
        # DEV-ONLY: Push DEV Defaults → PROD Defaults
        # =============================================================
        DEV_ENVIRONMENT = bool(st.secrets.get("DEV_ENVIRONMENT", False))

        if DEV_ENVIRONMENT:
            st.divider()
            st.subheader("DEV Tools")

            if st.button(
                "🚀 Push Dev Defaults to Prod",
                type="primary",
                width="stretch",
                help="Overwrite the PROD defaults JSON with the current DEV defaults",
            ):
                st.session_state["confirm_push_dev_to_prod"] = True
            
            
            st.divider()
            st.subheader("🧨 Dangerous DEV Tools")


            if st.button(
                    "🚀 Push Code Changes to Prod",
                    type="primary",
                    width="stretch",
                    help="Merge the dev branch directly into main via GitHub",
                ):
                    st.session_state["confirm_merge_dev_to_main"] = True            
        
            if st.button(
                    "🛠️ Rebuild Google Workspace",
                    type="primary",
                    width="stretch",
                    help="Creates a brand-new folder tree and rebinds all DEV config IDs"
                ):
                    st.session_state["confirm_rebuild_workspace"] = True




        @st.dialog("⚠️ Confirm Default Push to Production")
        def confirm_push_dev_to_prod():
            st.markdown(
                """
                **You are about to overwrite the PROD defaults configuration.**

                - ✅ PROD file ID will remain unchanged  
                - ✅ DEV defaults will completely replace PROD defaults  
                - ❌ This action **cannot be undone**

                Please confirm you want to continue.
                """
            )

            col_confirm, col_cancel = st.columns(2)

            with col_confirm:
                if st.button("✅ Yes — Push to PROD", type="primary", width="stretch"):
                    try:
                        DEV_CONFIG_FILE_ID = (st.secrets.get("DEV_CONFIG_FILE_ID", "") or "").strip()
                        PROD_CONFIG_FILE_ID = (st.secrets.get("CONFIG_FILE_ID", "") or "").strip()

                        if not DEV_CONFIG_FILE_ID or not PROD_CONFIG_FILE_ID:
                            st.error("Missing DEV_CONFIG_FILE_ID or CONFIG_FILE_ID in secrets.")
                        else:
                            creds = load_valid_token(cfg.SCOPES)
                            if not creds:
                                st.error("Google authentication required.")
                            else:
                                _, drive, _ = services(creds, cfg.HTTP_TIMEOUT_SECONDS)

                                # Load DEV defaults
                                dev_blob = drive.files().get_media(
                                    fileId=DEV_CONFIG_FILE_ID
                                ).execute()
                                dev_defaults = json.loads(dev_blob.decode("utf-8"))

                                # Overwrite PROD defaults (same file ID)
                                save_config_to_drive(
                                    drive,
                                    dev_defaults,
                                    file_id=PROD_CONFIG_FILE_ID
                                )

                                st.success("✅ DEV defaults successfully pushed to PROD.")

                    except Exception as e:
                        st.error(f"Push failed: {e}")

                    finally:
                        st.session_state.pop("confirm_push_dev_to_prod", None)
                        _rerun()

            with col_cancel:
                if st.button("❌ Cancel", width="stretch"):
                    st.session_state.pop("confirm_push_dev_to_prod", None)
                    _rerun()

        @st.dialog("⚠️ Confirm Code Push to Production")
        def confirm_merge_dev_to_main():
            base = 'dev'
            target = 'main'
            st.markdown(
                f"""
                **You are about to merge `{base}` into `{target}`.**

                - ✅ GitHub history will be preserved  
                - ✅ Branch protections still apply  
                - ❌ This action **may deploy to production**
                - ❌ This action **cannot be undone**

                Please confirm you want to continue.
                """
            )

            col_confirm, col_cancel = st.columns(2)

            with col_confirm:
                if st.button(
                    f"✅ Yes — Merge {base} → {target}",
                    type="primary",
                    width="stretch",
                ):
                    try:
                        with st.spinner("Merging branches…"):
                            success = github_merge(base, target)

                        if success:
                            st.success(f"✅ {base} successfully merged into {target}.")
                        else:
                            st.error("❌ Merge did not complete.")

                    finally:
                        st.session_state.pop("confirm_merge_dev_to_main", None)
                        st.rerun()

            with col_cancel:
                if st.button("❌ Cancel", width="stretch"):
                    st.session_state.pop("confirm_merge_dev_to_main", None)
                    st.rerun()
        
        @st.dialog("⚠️ Confirm Google Workspace Rebuild")
        def confirm_rebuild_workspace():
            st.markdown("""
            **This will create an entirely new Google Drive workspace.**

            ✅ A new main folder will be created  
            ✅ All folder IDs will be replaced  
            ❌ This action cannot be undone  

            **DEV ONLY**
            """)

            confirm = st.checkbox("I understand this is destructive")
            password = st.text_input("Enter admin password", type="password")

            if st.session_state.rebuild_error:
                st.error(st.session_state.rebuild_error)

            col1, col2 = st.columns(2)

            with col1:
                if st.button("❌ Cancel", width="stretch"):
                    st.session_state.confirm_rebuild_workspace = False
                    st.session_state.rebuild_error = None

            with col2:
                if st.button("✅ Rebuild Workspace", type="primary", width="stretch"):
                    if not confirm:
                        st.session_state.rebuild_error = "You must confirm the action."
                        return

                    if password != "admin":
                        st.session_state.rebuild_error = "Incorrect password."
                        return
                                        
                    
                    
                    st.session_state.confirm_rebuild_workspace = False
                    st.session_state.rebuild_requested = True
                    st.rerun()



                    


        # Trigger dialog
        if st.session_state.get("confirm_push_dev_to_prod"):
            confirm_push_dev_to_prod()
        
        if st.session_state.get("confirm_merge_dev_to_main"):
            confirm_merge_dev_to_main()
        
        if st.session_state.get("confirm_rebuild_workspace"):
            confirm_rebuild_workspace()
        

def render_upload_different_button(cfg):
    if st.button("🔁 Upload different files", width="stretch"):
        reset_to_upload()
        reset_pipeline_state()

        st.session_state.ui_phase = UI_UPLOAD
        _rerun()



def render_result_error(cfg):
    payload = st.session_state.get("pipe_error")

    with st.container(border=True):
        st.subheader("❌ Run Failed")

        if isinstance(payload, dict):
            # ✅ Friendly message (wrapped, readable)
            st.error(f"{payload['type']}: {payload['user_message']}")

            # ✅ Technical details hidden by default
            with st.expander("Technical details"):
                st.text(payload["traceback"])

        else:
            st.error("Unknown error occurred.")
        
        if st.button("🔁 Upload different files", type="primary"):
            st.session_state.pop("run_error", None)
            reset_to_upload()
            reset_pipeline_state()
            _rerun()


    if os.path.exists("last_run.log") and "last_run_log" not in st.session_state:
        with open("last_run.log", "rb") as f:
            st.session_state["last_run_log"] = f.read()
        st.session_state.setdefault("last_run_timestamp", "error")
    
    if "last_run_log" in st.session_state:
        st.download_button(
            "⬇️ Download full log (last_run.log)",
            st.session_state["last_run_log"],
            file_name=f"last_run_{st.session_state['last_run_timestamp']}.log",
            mime="text/plain",
            width='stretch'
            )



def render_upload_error(cfg):
    with st.container(border=True):
        st.subheader("❌ Invalid Upload")

        st.error(
            "Your uploaded file is invalid.\n\n"
            "Please upload **1 or 2 full weeks of data only**."
        )

        st.warning(st.session_state.get("upload_error", ""))

        if st.button("🔁 Upload different files", type="primary"):
            st.session_state.pop("run_error", None)
            reset_to_upload()
            reset_pipeline_state()
            _rerun()


def render_app(cfg):
    phase = st.session_state.ui_phase

    if phase == UI_UPLOAD:
        render_sidebar(cfg)
        render_upload_card(cfg)

    elif phase == UI_READY:
        render_sidebar(cfg)
        render_run_options(cfg)
        render_upload_different_button(cfg)

    elif phase == UI_RUNNING:
        render_running_status(cfg)

    elif phase == UI_RESULT:
        render_sidebar(cfg)
        render_results(cfg)
        render_upload_different_button(cfg)
    
    elif phase == UI_RESULT_ERROR:
        render_sidebar(cfg)
        render_result_error(cfg)

    elif phase == UI_UPLOAD_ERROR:
        render_sidebar(cfg)
        render_upload_error(cfg)

    elif phase in (UI_REBUILD_RUNNING, UI_REBUILD_DONE):
        render_rebuild_status()

        


# =========================
# App Entrypoint
# =========================

#st.title("🧾 FavTrip Reporting Pipeline")


st.set_page_config(
    page_title="FT Reporting",
    page_icon="🧾",          # emoji or path/URL to an image
    layout="wide",           # "centered" or "wide"
    initial_sidebar_state="collapsed",  # "auto", "expanded", "collapsed"
    menu_items={
        "Get Help": "mailto:ryan-morrow@uiowa.edu",
        "Report a bug": "https://github.com/ryan-j-morrow/favtrip_reporting/issues",
        "About": "FavTrip Reporting Pipeline",
    },
)

defaults = {
    "sales_selected_name": None,
    "vendor_selected_name": None,
    "sales_uploaded_ok": False,
    "vendor_uploaded_ok": False,
    "offer_log_download": False,
    "uploader_version": 0,
    "ui_phase": UI_UPLOAD,
    "auth_required": True,
    "running_ui_initialized": False,    
    "upload_epoch": 0,                 # increments on “Upload different files”
    "sales_selected_epoch": None,       # epoch when sales file was selected
    "vendor_selected_epoch": None,    
    "reset_generation": 0,
    "sales_selection_generation": None,
    "vendor_selection_generation": None,
    "sidebar_hint_seen": True,
    "rebuild_error": None,
    "rebuild_result": None,
    "confirm_rebuild_workspace": False,
    "rebuild_requested": False
}

    
for key, default in defaults.items():
    if key not in st.session_state:
        st.session_state[key] = default

cfg = Config.load()

creds = load_valid_token(cfg.SCOPES)
st.session_state.auth_required = creds is None


# --- STATE INIT ---
init_thread_state()
init_pipeline_state()

if st.session_state.get("rebuild_requested"):
    st.session_state.rebuild_requested = False
    trigger_rebuild(cfg)

# --- Finish OAuth inline when redirect comes back (this is in the NEW TAB) ---
params = st.query_params
if "code" in params and "state" in params:
    try:
        finish_web_oauth(params["code"], params["state"], cfg.SCOPES)
        # Token is saved locally in this new tab's app process
        st.success("✅ Google authentication complete.")
        
        has_token = (load_valid_token(cfg.SCOPES) is not None)
        st.session_state.auth_required = not has_token

        # Remove code/state from URL
        st.query_params.clear()

        # No messaging back to opener and NO window.close().
        # This tab becomes the main app; just rerun to flip UI.
        st.toast("Signed in. Loading the app…")
        _rerun()
    except Exception as e:
        st.error(f"OAuth error: {e}")

if (not st.session_state.auth_required) and ("sidebar_hint_seen" not in st.session_state):
    col_msg, col_btn = st.columns([6, 1], vertical_alignment="center")

    with col_msg:
        st.info(
            "⬅️ **Open the sidebar** for Utilities, Google auth, and DEV tools.",
            icon="👈",
        )

    with col_btn:
        if st.button("Got it", type="secondary"):
            st.session_state["sidebar_hint_seen"] = True
            _rerun()

# Auth gate
if st.session_state.auth_required:
    # ----------------------------
    # Authentication panel (shown only if auth required)
    # ----------------------------
    if st.session_state.auth_required:
        with st.expander("Google Authentication", expanded=True):
            st.caption(
                "Authentication is required before running. "
                "Click **Sign in with Google** to open the consent screen (it will open in a new tab)."
            )

            sign_in_ph = st.empty()
            clicked = sign_in_ph.button("Sign in with Google", type="primary", width='stretch')

            if clicked:
                try:
                    auth_url = start_web_oauth(cfg.SCOPES)
                    sign_in_ph.empty()

                    # Friendly message in this (original) tab
                    st.markdown(
                        """
                        <div style="
                            display:flex;align-items:center;justify-content:center;
                            height:55vh;text-align:center;
                            font-family: system-ui, Segoe UI, Roboto, Helvetica, Arial, sans-serif;">
                        <div>
                            <h2 style="margin-bottom:0.5rem;">You're being signed in…</h2>
                            <p style="font-size:1.05rem;opacity:.9;">
                            A new browser tab was opened for Google sign‑in.<br/>
                            <strong>After it completes, continue in that tab.</strong>
                            </p>
                        </div>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )

                    # Optional: refresh this tab when user returns (not required)
                    html(
                        """
                        <script>
                        document.addEventListener("visibilitychange", function() {
                            if (!document.hidden) { location.reload(); }
                        });
                        </script>
                        """,
                        height=0,
                    )

                    # Open Google auth in a NEW tab (this will ultimately become the main app)
                    html(
                        f"""
                        <script>
                        window.open({json.dumps(auth_url)}, "_blank", "noopener");
                        </script>
                        """,
                        height=0,
                    )

                    st.stop()
                except Exception as e:
                    st.error(f"Failed to start OAuth: {e}")

            with st.expander("Having trouble?", expanded=False):
                st.write(
                    "- The Google authorization page opens in a **new browser tab**.\n"
                    "- After completing consent, the **new tab** will load the app.\n"
                    "- If you renamed your Streamlit app or URL, ensure the Google OAuth "
                    "Authorized redirect URI matches exactly (including trailing slash)."
                )
            st.stop()

render_app(cfg)

```

---
### file: cli.py

```python
import argparse
from favtrip.config import Config
from favtrip.logger import StatusLogger
from favtrip.pipeline import run_pipeline


def main():
    parser = argparse.ArgumentParser(description="FavTrip Reporting Pipeline")
    parser.add_argument("--env", help="Path to .env file", default=None)

    # Per-run overrides (subset)
    parser.add_argument("--to", help="Comma-separated recipients", default=None)
    parser.add_argument("--cc", help="Comma-separated cc", default=None)
    parser.add_argument("--use-all-keys", action="store_true")
    parser.add_argument("--report-keys", help="Comma-separated report keys to run", default=None)
    parser.add_argument("--force-reauth", action="store_true")

    args = parser.parse_args()
    cfg = Config.load(args.env)

    if args.to:
        cfg.TO_RECIPIENTS = [s.strip() for s in args.to.split(',') if s.strip()]
    if args.cc:
        cfg.CC_RECIPIENTS = [s.strip() for s in args.cc.split(',') if s.strip()]
    if args.use_all_keys:
        cfg.USE_ALL_REPORT_KEYS = True
    if args.report_keys:
        cfg.REPORT_KEY_RUN_LIST = [s.strip().upper() for s in args.report_keys.split(',') if s.strip()]
    if args.force_reauth:
        cfg.FORCE_REAUTH = True

    logger = StatusLogger()
    result = run_pipeline(cfg, logger=logger)

    print("===== SUMMARY =====")
    print(logger.as_text())
    print("===================")


if __name__ == "__main__":
    main()

```

---
### file: core_functional_modules/__init__.py

```python
__all__ = [
    "config",
    "google_client",
    "sheets_utils",
    "drive_utils",
    "gmail_utils",
    "pipeline",
    "logger",
]

```

---
### file: core_functional_modules/config.py

```python
"""
config
======================================

Configuration loader and serializer for FavTrip reporting apps.

This module centralizes all runtime configuration for both local development
and cloud deployments (e.g., Streamlit Community Cloud). It provides a single,
typed `Config` dataclass plus helper functions that safely read from multiple
sources, coerce values to the expected Python types, and (optionally) overlay
a remote, Google Drive–hosted JSON configuration at runtime.

The loader is designed to be:
- **Layered**: Values are pulled from three tiers, in this order:
  1) Streamlit `st.secrets` (preferred in cloud; values may already be typed)
  2) Process environment and/or a local `.env` file (string-based; coerced) #Sandbox use only
  3) A Google Drive JSON override, applied last if credentials and
     a config file are available
- **Safe**: Missing keys never raise; reasonable defaults are used instead.
- **Type-aware**: Bools, lists, and dicts are parsed/coerced consistently so the
  same code works with typed TOML (in `st.secrets`) and string-based `.env`.

-------------------------------------------------------------------------------
Core API
-------------------------------------------------------------------------------

- `_get_secret(key: str, default: Any = None) -> Any`
  Attempts to read `key` from `streamlit.secrets` (if Streamlit is present and
  has `secrets`), else falls back to `os.getenv(key, default)`. Never raises for
  missing keys; always returns a value (possibly `default`). Streamlit import is
  lazy to avoid a hard dependency for non-Streamlit contexts.

- `_coerce_bool(v: Any, default: bool = False) -> bool`
  Accepts `bool | str | int | None` and returns a Python `bool`.
  Truthy strings (case-insensitive, trimmed) include:
  `{"1", "true", "yes", "on", "y", "t"}`. Non-parseable inputs fall back to
  `default`.

- `_coerce_csv(v: Any) -> List[str]`
  Accepts a list/tuple (already structured) or a comma-separated string and
  yields a list of **trimmed** strings. `None`/empty returns `[]`.

- `_coerce_json(v: Any) -> Dict[str, Any]`
  Accepts a `dict` or a JSON string. Returns a `dict`; parse failures yield `{}`.

- `@dataclass class Config`
  A top-level dataclass holding all tunable settings for the application:
  * **Drive/Sheets IDs**:
    - `CALC_SPREADSHEET_ID`, `INCOMING_FOLDER_ID`, `MANAGER_REPORT_FOLDER_ID`,
      `ORDER_REPORT_FOLDER_ID`, `USER_FOLDER_ID`
  * **GIDs, sheet metadata, timestamps**:
    - `GID_MANAGER_PDF`, `GID_ORDER_CSV`, `LOCATION_SHEET_TITLE`,
      `LOCATION_NAMED_RANGE`, `TEMPLATE_UPDATE_RANGE`
    - `TIMESTAMP_TZ` (e.g., "America/Chicago")
    - `TIMESTAMP_FMT` (default "%Y-%m-%d-%I-%M-%p")
  * **Email & distribution**:
    - `TO_RECIPIENTS`, `CC_RECIPIENTS`, `USE_ALL_REPORT_KEYS`,
      `REPORT_KEY_RUN_LIST`, `REPORT_KEY_RECIPIENTS`,
      `DEFAULT_ORDER_RECIPIENTS`
    - `INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL`,
      `SEND_SEPARATE_FULL_ORDER_EMAIL`, `EMAIL_MANAGER_REPORT`
  * **Google API**:
    - `SCOPES` (Drive/Sheets/Gmail), `FORCE_REAUTH`,
      `REDIRECT_PORT`, `HTTP_TIMEOUT_SECONDS`
  * **Advanced intake**:
    - `USE_AUTO_ROLLOVER_IF_ONE_WEEK`,
      `START_DAY_OF_WEEK`, `END_DAY_OF_WEEK`
      (Accepted values include: Sunday, Monday, Tuesday, Wednesday, Thursday,
      Friday, Saturday, Any)
  * **Cleanup (days)**:
    - `OUTPUT_TIME_TO_LIFE`, `FAILED_INPUT_TIME_TO_LIFE`, `USER_TIME_TO_LIFE`

  Defaults are provided for all fields. When loading from secrets or `.env`,
  values are coerced into the correct types; lists and dicts are parsed as
  necessary. `REPORT_KEY_RUN_LIST` values are uppercased to reduce downstream
  casing issues.

- `Config.load(env_path: Optional[pathlib.Path] = None) -> Config`
  Loads the final, effective configuration via a layered merge:
  1) Loads a local `.env` file from `env_path` (default: `cwd/.env`) using
     `python-dotenv` with `override=False` (so existing process env vars win).
  2) Reads settings from `st.secrets` if available; otherwise from environment.
     Values are passed through the coercers defined above.
  3) Attempts to overlay a Google Drive–hosted JSON config:
     - Uses `core_functional_modules.google_client.load_valid_token` and `services` to obtain a
       Drive client (respecting `HTTP_TIMEOUT_SECONDS`).
     - Reads a JSON dict via `core_functional_modules.config_store.load_config_from_drive`,
       optionally using `CONFIG_FILE_ID` from `st.secrets` if present.
     - Keys in the override dict that match `Config` attributes replace
       previously loaded values.
     - On any failure (no token, network error, file missing, etc.), the loader
       **fails open** and returns the base config without raising (best-effort).

- `Config.to_env() -> str`
  Serializes the current configuration to a string in `.env` format. Collections
  are flattened—lists are joined with commas, and dicts are JSON-encoded—so the
  output can be written to disk and re-read later in a purely string-based env.

- `Config.save(env_path: Optional[pathlib.Path] = None) -> None`
  Convenience wrapper around `to_env()` that writes the serialized configuration
  to `env_path` (default: `cwd/.env`, UTF-8).

-------------------------------------------------------------------------------
Environment / Secrets Reference (all optional; sensible defaults apply)
-------------------------------------------------------------------------------

Drive / Sheets IDs:
- `CALC_SPREADSHEET_ID`, `INCOMING_FOLDER_ID`, `MANAGER_REPORT_FOLDER_ID`,
  `ORDER_REPORT_FOLDER_ID`, `USER_FOLDER_ID`

Sheet metadata & timestamps:
- `GID_MANAGER_PDF`, `GID_ORDER_CSV`, `LOCATION_SHEET_TITLE`,
  `LOCATION_NAMED_RANGE`, `TEMPLATE_UPDATE_RANGE`
- `TIMESTAMP_TZ`, `TIMESTAMP_FMT`

Email & distribution:
- `TO_RECIPIENTS` (CSV), `CC_RECIPIENTS` (CSV)
- `USE_ALL_REPORT_KEYS` (bool)
- `REPORT_KEY_RUN_LIST` (CSV; uppercased during load)
- `REPORT_KEY_RECIPIENTS` (JSON dict)
- `DEFAULT_ORDER_RECIPIENTS` (CSV)
- `INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL` (bool)
- `SEND_SEPARATE_FULL_ORDER_EMAIL` (bool)
- `EMAIL_MANAGER_REPORT` (bool)

Google API:
- `SCOPES` (CSV; typical: Drive/Sheets/Gmail send)
- `FORCE_REAUTH` (bool)
- `REDIRECT_PORT` (int)
- `HTTP_TIMEOUT_SECONDS` (int)

Advanced intake / rollover:
- `USE_AUTO_ROLLOVER_IF_ONE_WEEK` (bool)
- `START_DAY_OF_WEEK`, `END_DAY_OF_WEEK`
  (Sunday|Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Any)

Cleanup (days):
- `OUTPUT_TIME_TO_LIFE`, `FAILED_INPUT_TIME_TO_LIFE`, `USER_TIME_TO_LIFE`

Drive override:
- `CONFIG_FILE_ID` (usually provided via `st.secrets`, if using a specific file)

-------------------------------------------------------------------------------
Operational Notes
-------------------------------------------------------------------------------

- **Lazy imports**: `streamlit` and Google client utilities are imported inside
  the loader so the module remains usable in non-Streamlit or headless contexts.
- **Fail-open Drive overrides**: If Drive credentials are unavailable or an
  override file cannot be retrieved/parsed, the loader returns the base config
  without raising (best-effort behavior).
- **Deterministic parsing**: Coercers are idempotent for already-typed values.
  For example, booleans in TOML remain booleans; CSV strings are split and
  trimmed; JSON strings are parsed into dicts.
- **Case normalization**: `REPORT_KEY_RUN_LIST` is uppercased at load time to
  minimize case-related mismatches elsewhere in the app.

Import this module early in your app to construct a single, consistent
`Config` instance and pass it through to components that require configuration.

"""


from __future__ import annotations

import os
import json
from dataclasses import dataclass, asdict, field
from typing import Any, Dict, List, Optional
from pathlib import Path
from dotenv import load_dotenv

# -----------------------------------------------------------------------------
# Helpers: read from Streamlit secrets (typed) or .env (strings) and coerce
# -----------------------------------------------------------------------------

def _get_secret(key: str, default: Any = None) -> Any:
    """
    Read from Streamlit secrets if present, else env var, else default.
    Does not raise if key missing; returns `default`.
    """
    try:
        import streamlit as st  # imported lazily to avoid hard dependency
        if hasattr(st, "secrets") and key in st.secrets:
            return st.secrets.get(key, default)
    except Exception:
        pass
    return os.getenv(key, default)

_TRUE = {"1", "true", "yes", "on", "y", "t"}

def _coerce_bool(v: Any, default: bool = False) -> bool:
    """
    Accept bool | str | int | None and return a Python bool.
    Works for typed TOML (bool) and .env strings.
    """
    if isinstance(v, bool):
        return v
    if v is None:
        return default
    try:
        return str(v).strip().lower() in _TRUE
    except Exception:
        return default

def _coerce_csv(v: Any) -> List[str]:
    """
    Accept list/tuple (already structured) or a comma-separated string.
    Returns a list of trimmed strings.
    """
    if v is None or v == "":
        return []
    if isinstance(v, (list, tuple)):
        return [str(x).strip() for x in v if str(x).strip()]
    return [p.strip() for p in str(v).split(",") if p.strip()]

def _coerce_json(v: Any) -> Dict[str, Any]:
    """
    Accept dict (already structured) or a JSON string.
    Returns a dict; falls back to {} on parse issues.
    """
    if v is None or v == "":
        return {}
    if isinstance(v, dict):
        return v
    try:
        return json.loads(v)
    except Exception:
        return {}

# -----------------------------------------------------------------------------
# Config dataclass (TOP-LEVEL — must start at column 0)
# -----------------------------------------------------------------------------

@dataclass
class Config:
    NON_PERSISTED_FIELDS = set()  
  
    # IDs and basic settings
    CALC_SPREADSHEET_ID: str = "1ibkGkQ2khYMJydeenJkTzC4KoLQAyBZW_esQrbjSHXs"
    INCOMING_FOLDER_ID: str = "1jJE3r9DOHXwBdd94E6ZhxBBH9xvSjI-b"
    MANAGER_REPORT_FOLDER_ID: str = "17Nqwo6HYe30JP0wnZYoLRG0F1s-X-IVZ"
    ORDER_REPORT_FOLDER_ID: str = "171dqzMim-IdpB_kzjYQnzoSbW89uJTfP"
    ERROR_REPORT_FOLDER_ID: str = "1T-rnyXmPD1eFcxi-s8i4b1EP6-pW5ETW"
    USER_FOLDER_ID: str = "1JBHBcnS6397ka2ITW6Wbuu2aKjbgCCHj"

    # GIDs, sheet metadata, timestamp settings
    GID_MANAGER_PDF: str = "1921812573"
    GID_ORDER_CSV: str = "1875928148"
    GID_ERROR_REPORT: str = "1581903111"
    GID_BEV_ERRORS: str = "72711538"
    LOCATION_SHEET_TITLE: str = "REFR: Values"
    LOCATION_NAMED_RANGE: str = "_locations"
    TIMESTAMP_TZ: str = "America/Chicago"
    TIMESTAMP_FMT: str = "%Y-%m-%d-%I-%M-%p"
    TEMPLATE_UPDATE_RANGE: str = "_update"

    # Email config
    TO_RECIPIENTS: List[str] = field(default_factory=lambda: ["FavtripReporting@gmail.com"])
    CC_RECIPIENTS: List[str] = None
    ERROR_RECIPIENTS: List[str] = field(default_factory=lambda: ["FavtripReporting@gmail.com"])
    USE_ALL_REPORT_KEYS: bool = False
    REPORT_KEY_RUN_LIST: List[str] = field(default_factory=lambda: ["COFFEE"])
    REPORT_KEY_RECIPIENTS: Dict[str, List[str]] = None
    DEFAULT_ORDER_RECIPIENTS: List[str] = None
    INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL: bool = False
    SEND_SEPARATE_FULL_ORDER_EMAIL: bool = False
    EMAIL_MANAGER_REPORT: bool = True

    # Google API
    SCOPES: List[str] = None
    FORCE_REAUTH: bool = False
    REDIRECT_PORT: int = 58285
    HTTP_TIMEOUT_SECONDS: int = 300

    # Advanced intake settings
    USE_AUTO_ROLLOVER_IF_ONE_WEEK: bool = True
    START_DAY_OF_WEEK: str = "Sunday"    # Sunday, Monday, Tuesday, Wednesday, Thursday, Friday, Saturday, Any
    END_DAY_OF_WEEK: str = "Saturday"    # Sunday, Monday, Tuesday, Wednesday, Thursday, Friday, Saturday, Any
    
    SOFT_CASES_ALERT_ENABLED: bool = True
    SOFT_CASES_ALERT_THRESHOLD: int = 10


    # Cleanup
    OUTPUT_TIME_TO_LIFE: int = 30
    FAILED_INPUT_TIME_TO_LIFE: int = 1
    USER_TIME_TO_LIFE: int = 90

    #Other
    BEV_MAPPING_LINK: str = "https://docs.google.com/spreadsheets/d/1O6MtF-GM0VayqMr_v3oJC5PnRK5yv6biiDtA_qw-Z3g/"

    @staticmethod
    def load(env_path: Optional[Path] = None) -> "Config":
        """
        Load config from Streamlit secrets (preferred on cloud) or from env/.env (local dev),
        then overlay any values found in a Drive-backed JSON config (optional).
        Secrets may be typed (bool/list/dict), so we coerce safely.
        """
        if env_path is None:
            env_path = Path.cwd() / ".env"
        load_dotenv(dotenv_path=env_path, override=False)

        cfg = Config(
            CALC_SPREADSHEET_ID=str(_get_secret("CALC_SPREADSHEET_ID", "")),
            INCOMING_FOLDER_ID=str(_get_secret("INCOMING_FOLDER_ID", "")),
            MANAGER_REPORT_FOLDER_ID=str(_get_secret("MANAGER_REPORT_FOLDER_ID", "")),
            ORDER_REPORT_FOLDER_ID=str(_get_secret("ORDER_REPORT_FOLDER_ID", "")),
            ERROR_REPORT_FOLDER_ID=str(_get_secret("ERROR_REPORT_FOLDER_ID", "")),
            USER_FOLDER_ID=str(_get_secret("USER_FOLDER_ID", "")),

            GID_MANAGER_PDF=str(_get_secret("GID_MANAGER_PDF", "1921812573")),
            GID_ORDER_CSV=str(_get_secret("GID_ORDER_CSV", "1875928148")),
            GID_ERROR_REPORT=str(_get_secret("GID_ERROR_REPORT", "1581903111")),
            GID_BEV_ERRORS=str(_get_secret("GID_BEV_ERRORS", "72711538")),
            LOCATION_SHEET_TITLE=str(_get_secret("LOCATION_SHEET_TITLE", "REFR: Values")),
            LOCATION_NAMED_RANGE=str(_get_secret("LOCATION_NAMED_RANGE", "_locations")),
            TEMPLATE_UPDATE_RANGE=str(_get_secret("TEMPLATE_UPDATE_RANGE", "_update")),
            TIMESTAMP_TZ=str(_get_secret("TIMESTAMP_TZ", "America/Chicago")),
            TIMESTAMP_FMT=str(_get_secret("TIMESTAMP_FMT", "%Y-%m-%d-%I-%M-%p")),

            OUTPUT_TIME_TO_LIFE=int(_get_secret("OUTPUT_TIME_TO_LIFE", 30)),
            FAILED_INPUT_TIME_TO_LIFE=int(_get_secret("FAILED_INPUT_TIME_TO_LIFE", 1)),
            USER_TIME_TO_LIFE=int(_get_secret("USER_TIME_TO_LIFE", 1)),

            TO_RECIPIENTS=_coerce_csv(_get_secret("TO_RECIPIENTS", "")),
            CC_RECIPIENTS=_coerce_csv(_get_secret("CC_RECIPIENTS", "")),
            ERROR_RECIPIENTS=_coerce_csv(_get_secret("ERROR_RECIPIENTS", "")),
            USE_ALL_REPORT_KEYS=_coerce_bool(_get_secret("USE_ALL_REPORT_KEYS", "false")),
            REPORT_KEY_RUN_LIST=[s.upper() for s in _coerce_csv(_get_secret("REPORT_KEY_RUN_LIST", ""))],
            REPORT_KEY_RECIPIENTS=_coerce_json(_get_secret("REPORT_KEY_RECIPIENTS", "{}")),
            DEFAULT_ORDER_RECIPIENTS=_coerce_csv(_get_secret("DEFAULT_ORDER_RECIPIENTS", "")),
            INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL=_coerce_bool(
                _get_secret("INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL", "false")
            ),
            SEND_SEPARATE_FULL_ORDER_EMAIL=_coerce_bool(
                _get_secret("SEND_SEPARATE_FULL_ORDER_EMAIL", "true")
            ),
            EMAIL_MANAGER_REPORT=_coerce_bool(_get_secret("EMAIL_MANAGER_REPORT", "true")),

            SCOPES=_coerce_csv(
                _get_secret(
                    "SCOPES",
                    "https://www.googleapis.com/auth/drive,"
                    "https://www.googleapis.com/auth/spreadsheets,"
                    "https://www.googleapis.com/auth/gmail.send",
                )
            ),
            FORCE_REAUTH=_coerce_bool(_get_secret("FORCE_REAUTH", "false")),
            REDIRECT_PORT=int(str(_get_secret("REDIRECT_PORT", "58285")) or "58285"),
            HTTP_TIMEOUT_SECONDS=int(str(_get_secret("HTTP_TIMEOUT_SECONDS", "300")) or "300"),

            USE_AUTO_ROLLOVER_IF_ONE_WEEK=_coerce_bool(_get_secret("USE_AUTO_ROLLOVER_IF_ONE_WEEK", "true")),
            START_DAY_OF_WEEK=str(_get_secret("START_DAY_OF_WEEK", "Sunday")),
            END_DAY_OF_WEEK=str(_get_secret("END_DAY_OF_WEEK", "Saturday")),

            
            SOFT_CASES_ALERT_ENABLED=_coerce_bool(_get_secret("SOFT_CASES_ALERT_ENABLED", "true")),
            SOFT_CASES_ALERT_THRESHOLD=int(_get_secret("SOFT_CASES_ALERT_THRESHOLD", 10)),

            BEV_MAPPING_LINK=str(_get_secret("BEV_MAPPING_LINK", "https://docs.google.com/spreadsheets/d/1O6MtF-GM0VayqMr_v3oJC5PnRK5yv6biiDtA_qw-Z3g/")),
        )

        normalized = {}
        for k, v in cfg.REPORT_KEY_RECIPIENTS.items():
            if isinstance(k, (list, tuple)):
                if len(k) == 2:
                    normalized[(k[0], k[1], None)] = v
                elif len(k) == 3:
                    normalized[tuple(k)] = v
            else:
                # defensive fallback
                normalized[(None, k, None)] = v

        cfg.REPORT_KEY_RECIPIENTS = normalized


        # Optional overlay from Drive JSON config (if creds + file present)
        # ---------------- Drive-backed config overlay ----------------
        try:
            import streamlit as st
            from core_functional_modules.google_client import load_valid_token, services
            from core_functional_modules.config_store import load_config_from_drive

            DEV_ENVIRONMENT = _coerce_bool(_get_secret("DEV_ENVIRONMENT", False))
            DEV_CONFIG_FILE_ID = str(_get_secret("DEV_CONFIG_FILE_ID", "") or "").strip()
            CONFIG_FILE_ID = str(_get_secret("CONFIG_FILE_ID", "") or "").strip()

            # Select which config file ID to READ from
            if DEV_ENVIRONMENT and DEV_CONFIG_FILE_ID:
                active_config_file_id = DEV_CONFIG_FILE_ID
            else:
                active_config_file_id = CONFIG_FILE_ID or None

            creds = load_valid_token(cfg.SCOPES)
            if creds:
                _sheets, drive, _gmail = services(creds, cfg.HTTP_TIMEOUT_SECONDS)

                overrides = {}
                
                # 1️⃣ Try DEV config first (if enabled)
                if DEV_ENVIRONMENT and DEV_CONFIG_FILE_ID:
                    overrides = load_config_from_drive(drive, DEV_CONFIG_FILE_ID)

                # 2️⃣ Fallback to PROD config if DEV missing/empty
                if not overrides and CONFIG_FILE_ID:
                    overrides = load_config_from_drive(drive, CONFIG_FILE_ID)

                # 3️⃣ Apply overrides if any
                if isinstance(overrides, dict):
                    for k, v in overrides.items():
                        if hasattr(cfg, k):
                            setattr(cfg, k, v)

        except Exception:
            # Fail-open by design
            
            import traceback
            print("[Config] Drive overlay failed:", e)
            traceback.print_exc()

            #pass
        
        return cfg


    # -------------------------------------------------------------------------
    # .env serialization (optional helper)
    # -------------------------------------------------------------------------
    def to_env(self) -> str:
        """Serialize to .env format (simple, string-based)."""
        data = asdict(self)
        as_env = {
            **data,
            "TO_RECIPIENTS": ",".join(self.TO_RECIPIENTS or []),
            "CC_RECIPIENTS": ",".join(self.CC_RECIPIENTS or []),
            "REPORT_KEY_RUN_LIST": ",".join(self.REPORT_KEY_RUN_LIST or []),
            "REPORT_KEY_RECIPIENTS": json.dumps(self.REPORT_KEY_RECIPIENTS or {}),
            "DEFAULT_ORDER_RECIPIENTS": ",".join(self.DEFAULT_ORDER_RECIPIENTS or []),
            "SCOPES": ",".join(self.SCOPES or []),
        }
        lines = [f"{k}={v}" for k, v in as_env.items()]
        return "\n".join(lines) + "\n"

    def save(self, env_path: Optional[Path] = None):
        if env_path is None:
            env_path = Path.cwd() / ".env"
        env_path.write_text(self.to_env(), encoding="utf-8")
    

    
    def to_drive_defaults(self) -> dict:
        return {
            k: v
            for k, v in vars(self).items()
            if not k.startswith("_")
            and k not in self.NON_PERSISTED_FIELDS
        }


```

---
### file: core_functional_modules/config_store.py

```python
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

```

---
### file: core_functional_modules/drive_utils.py

```python
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

def get_or_create_subfolder(drive_svc, parent_folder_id: str, name: str):
    """
    Return a Drive folder with the given name under parent_folder_id.
    Create it if it does not already exist.
    """
    q = (
        f"mimeType='application/vnd.google-apps.folder' "
        f"and name='{name}' "
        f"and '{parent_folder_id}' in parents "
        f"and trashed=false"
    )

    res = drive_svc.files().list(
        q=q,
        fields="files(id, name, webViewLink)",
        pageSize=1
    ).execute()

    files = res.get("files", [])
    if files:
        return files[0]

    metadata = {
        "name": name,
        "mimeType": "application/vnd.google-apps.folder",
        "parents": [parent_folder_id],
    }

    return drive_svc.files().create(
        body=metadata,
        fields="id, name, webViewLink"
    ).execute()


```

---
### file: core_functional_modules/gmail_utils.py

```python
""" 
gmail_utils
======================================
Email utilities for sending Gmail messages with PDF attachments via the Gmail API.

This module provides small, focused helpers for sending two types of emails through
the Gmail API using a pre-authorized `gmail_svc` client (e.g., returned by
`googleapiclient.discovery.build("gmail", "v1", ...)`). It includes:

- `send_email(...)`: Low-level helper that accepts an `email.message.EmailMessage`,
  base64-url encodes it as required by Gmail, and dispatches it via
  `users.messages.send`.

- `email_manager_report(...)`: Composes and sends a standardized "Manager Report"
  email with a primary PDF attachment and a backup link. Supports optional CC.

- `email_order_report(...)`: Composes and sends an "Order Report" email for a
  given vendor or category key, including a primary PDF attachment and an optional
  "full order" PDF. Also includes links to a backing Google Sheet and supports CC.

The functions here intentionally perform minimal validation and assume that callers
supply valid addresses, attachments, and links. Authentication, token refresh, and
error handling policy (e.g., retries, backoff, alerting) should be implemented by
the caller.

---
Key Behaviors
-------------
- **MIME construction**: Uses Python's stdlib `email.message.EmailMessage` to build
  multipart emails with both plain-text and HTML alternatives, and PDF attachments.
- **Gmail API compliance**: Serializes the email to bytes and encodes it with
  URL-safe Base64 as required by Gmail's `users.messages.send` endpoint.
- **Idempotency**: Sending is not idempotent; calling functions repeatedly may
  result in duplicate emails. Callers should implement their own guardrails if
  needed (e.g., deduplication keys, sent-flagging).
- **Internationalization**: The functions do not localize content; callers can adapt
  the text if i18n is required.
- **HTML content**: Simple HTML bodies are included via `add_alternative(..., subtype="html")`.
  The HTML snippets intentionally avoid external assets for reliable delivery.

---
Functions
---------
send_email(gmail_svc, user, msg)
    Low-level send helper. Encodes the `EmailMessage` and dispatches via the Gmail API.

email_manager_report(gmail_svc, sender, to_list, cc_list, pdf_name, pdf_bytes, pdf_link, ts, location)
    Sends a standardized "Manager Report" email with a PDF attachment and a backup link.

email_order_report(
    gmail_svc,
    sender,
    to_list,
    cc_list,
    key,
    tag,
    ts,
    location,
    pdf_name,
    pdf_bytes,
    sheet_link,
    include_full_order=False,
    full_pdf_bytes=None,
    full_pdf_name=None,
)
    Sends an "Order Report" email targeted to a `{key}` team with a primary PDF,
    optional full-order PDF, and a link to the backing Google Sheet.

---
Parameters (Shared Concepts)
----------------------------
gmail_svc : Any
    An authenticated Gmail API service client (e.g., from `googleapiclient.discovery.build`).

sender : str
    The "From" email address to display in the message header. The authenticated
    Gmail account must be authorized to send from this address.

to_list : Iterable[str]
    Recipient email addresses for the `To` field. Must contain at least one valid address.

cc_list : Optional[Iterable[str]]
    Optional CC recipient addresses. If empty or `None`, the `Cc` header is omitted.

pdf_name : str
    Filename for the attached PDF (e.g., `"report_2026-03-21.pdf"`).

pdf_bytes : bytes
    Raw bytes of the primary PDF attachment.

ts : str
    A timestamp string suitable for inclusion in the subject (e.g., `"2026-03-21"` or
    `"2026-03-21 18:25"`).

location : str
    A human-readable location name included in the subject/body (e.g., store or site).

pdf_link : str
    (Manager Report) A backup URL users can access if attachments are blocked.

key : str
    (Order Report) An identifier for the receiving team or vendor (e.g., `"Dairy"`, `"VendorX"`).

tag : str
    (Order Report) A secondary descriptor (e.g., `"Weekly"`, `"Overstock"`, `"Emergency"`).

sheet_link : str
    (Order Report) URL to the backing Google Sheet with order details.

include_full_order : bool
    (Order Report) Whether to attach an additional "full order" PDF.

full_pdf_bytes : Optional[bytes]
    (Order Report) Raw bytes of the full order PDF (required when `include_full_order=True`).

full_pdf_name : Optional[str]
    (Order Report) Filename for the full order PDF (required when `include_full_order=True`).

user : str
    (send_email) Gmail user identifier for the API call. Typically `"me"` to refer
    to the authenticated account.

msg : EmailMessage
    (send_email) A fully-constructed email message to be sent.

---
Returns
-------
dict
    The Gmail API response payload from `users.messages.send()` (e.g., includes `id`, `threadId`).

---
Raises
------
googleapiclient.errors.HttpError
    If the Gmail API call fails (e.g., quota exceeded, invalid permissions, bad request).
ValueError / TypeError
    If provided inputs (addresses, bytes, filenames) are invalid (may be raised by stdlib or caller validations).

---
"""


from __future__ import annotations
import base64
from email.message import EmailMessage


def send_email(gmail_svc, user: str, msg: EmailMessage):
    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
    return gmail_svc.users().messages().send(userId=user, body={"raw": raw}).execute()


def email_manager_report(gmail_svc, sender: str, to_list, cc_list, pdf_name, pdf_bytes, pdf_link, ts, location):
    msg = EmailMessage()
    msg["Subject"] = f"Manager Report – {location} – {ts}"
    msg["From"] = sender
    msg["To"] = ", ".join(to_list)
    if cc_list:
        msg["Cc"] = ", ".join(cc_list)
    msg.set_content(f"Hi team,\nAttached is the Manager Report ({location}).\nBackup link: {pdf_link}\n—Sent from an automated reporting pipeline")

    msg.add_alternative(
        f"""
        <p>Hi team,</p>
        <p>Your manager report for store <b>{location}</b> is ready.</p>
        <p><a href='{pdf_link}'>Backup Link</a></p>
        <p>Attached: {pdf_name}</p>
        <p>—Sent from an automated reporting pipeline</p>
        """,
        subtype="html",
    )

    msg.add_attachment(pdf_bytes, maintype="application", subtype="pdf", filename=pdf_name)
    return send_email(gmail_svc, sender, msg)


def email_order_report(
    gmail_svc,
    sender: str,
    to_list,
    cc_list,
    key: str,
    tag: str,
    ts: str,
    location: str,
    pdf_name: str,
    pdf_bytes: bytes,
    sheet_link: str,
    include_full_order: bool = False,
    full_pdf_bytes: bytes | None = None,
    full_pdf_name: str | None = None,
):
    msg = EmailMessage()

    msg["Subject"] = f"Order Report – {location} – {tag} – {ts}"
    msg["From"] = sender
    msg["To"] = ", ".join(to_list)

    if cc_list:
        msg["Cc"] = ", ".join(cc_list)

    msg.set_content(
        f"Hi {tag} team,\n"
        f"Your order report for {location} - {tag} is ready.\n"
        f"Google Sheet: {sheet_link}\n"
        f"Attached: {pdf_name}\n"
        "—Sent from an automated reporting pipeline"
    )

    msg.add_alternative(
        f"""
        <p>Hi {tag} team,</p>
        <p>Your order report for store <b>{location}</b> is ready.</p>
        <p><a href="{sheet_link}">Open Google Sheet</a></p>
        <p>Attached: {pdf_name}</p>
        <p>—Sent from an automated reporting pipeline</p>
        """,
        subtype="html",
    )

    msg.add_attachment(
        pdf_bytes,
        maintype="application",
        subtype="pdf",
        filename=pdf_name,
    )

    if include_full_order and full_pdf_bytes:
        msg.add_attachment(
            full_pdf_bytes,
            maintype="application",
            subtype="pdf",
            filename=full_pdf_name,
        )

    return send_email(gmail_svc, sender, msg)


def email_error_report(
    gmail_svc,
    sender: str,
    to_list,
    cc_list,
    ts: str,
    pdf_name: str,
    pdf_bytes: bytes,
    sheet_link: str
    ):
    msg = EmailMessage()

    msg["Subject"] = f"Error Report – {ts}"
    msg["From"] = sender
    msg["To"] = ", ".join(to_list)

    if cc_list:
        msg["Cc"] = ", ".join(cc_list)

    msg.set_content(
        f"Hi Technical Support Team,\n"
        f"A user tried using the reporting pipeline, however some of the items that were uploaded are not listed on the Vendor Price Book.\n"
        f"The default recipient of the pipeline run is CC'd on this email for visibility and communication purposes.\n"
        f"Please reply to this email once the Vendor Price Book is updated so that the user knows they can rerun the pipeline.\n\n"
        f"Google Sheet: {sheet_link}\n"
        f"Attached: {pdf_name}\n"
        "—Sent from an automated reporting pipeline"
    )

    msg.add_alternative(
        f"""
        <p>Hi Technical Support Team,</p>
        <p>A user tried using the reporting pipeline, however some of the items that were uploaded are not listed on the Vendor Price Book.</p>
        <p>The default recipient of the pipeline run is CC'd on this email for visibility and communication purposes.</p>
        <p>Please reply to this email once the Vendor Price Book is updated so that the user knows they can rerun the pipeline.</p>
        <p></p>
        <p><a href="{sheet_link}">Open Error Report in Google Sheets</a></p>
        <p>Attached: {pdf_name}</p>
        <p>—Sent from an automated reporting pipeline</p>
        """,
        subtype="html",
    )

    msg.add_attachment(
        pdf_bytes,
        maintype="application",
        subtype="pdf",
        filename=pdf_name,
    )

    return send_email(gmail_svc, sender, msg)

def email_bev_error_report(
    gmail_svc,
    sender: str,
    to_list,
    cc_list,
    ts: str,
    pdf_name: str,
    pdf_bytes: bytes,
    sheet_link: str,
    mapping_link: str
    ):
    msg = EmailMessage()

    msg["Subject"] = f"Soft Alert – Unassigned Beverages Report – {ts}"
    msg["From"] = sender
    msg["To"] = ", ".join(to_list)

    if cc_list:
        msg["Cc"] = ", ".join(cc_list)

    msg.set_content(
        f"Hi Technical Support Team,\n"
        f"A user just ran the ordering pipeline successfully, however some beverages were not mapped to a sub-report-key / vendor.\n"
        f"The BEV orders were still sent, however these unassigned beverages were included on their own order under BEV - UNASSIGNED.\n\n"
        f"The default recipient of the pipeline run is CC'd on this email for visibility and communication purposes.\n"
        f"To fix this error for future runs look at the Unassigned Beverages Report and add those Scan Codes to the Mapping File.\n"
        f"Please reply to this email once the Mapping File is updated so that the user knows they can rerun the pipeline if needed.\n\n"
        f"Beverage Mapping File: {mapping_link}\n"
        f"Unassigned Beverages Google Sheet: {sheet_link}\n"
        f"Attached: {pdf_name}\n"
        "—Sent from an automated reporting pipeline"
    )

    msg.add_alternative(
        f"""
        <p>Hi Technical Support Team,</p>
        <p>A user just ran the ordering pipeline successfully, however some beverages were not mapped to a sub-report-key / vendor.</p>
        <p>The BEV orders were still sent; however, these unassigned beverages were included on their own order under <strong>BEV - UNASSIGNED</strong>.</p><p></p>
        <p>The default recipient of the pipeline run is CC'd on this email for visibility and communication purposes.</p>
        <p>To fix this error for future runs, please review the Unassigned Beverages Report and add those Scan Codes to the Mapping File.</p>
        <p>Please reply to this email once the Mapping File is updated so that the user knows they can rerun the pipeline if needed.</p><p></p>
        <p><a href="{mapping_link}">Open Beverage Mapping File</a></p>
        <p><a href="{sheet_link}">Open Unassigned Beverages Google Sheet</a></p>
        <p>Attached: {pdf_name}</p>
        <p>—Sent from an automated reporting pipeline</p>
        """,
        subtype="html",
    )

    msg.add_attachment(
        pdf_bytes,
        maintype="application",
        subtype="pdf",
        filename=pdf_name,
    )

    return send_email(gmail_svc, sender, msg)

def email_large_case_alert_report(
    gmail_svc,
    sender: str,
    to_list,
    cc_list,
    ts: str,
    location: str,
    threshold: int,
    pdf_name: str,
    pdf_bytes: bytes,
    sheet_link: str,
):
    """
    Send a soft-alert email when FULL order lines exceed a case threshold.

    This is a NON-blocking informational alert intended for
    technical review, not end users.
    """

    msg = EmailMessage()

    msg["Subject"] = (
        f"Soft Alert – High Case Quantities – {location} – {ts}"
    )
    msg["From"] = sender
    msg["To"] = ", ".join(to_list)

    if cc_list:
        msg["Cc"] = ", ".join(cc_list)

    msg.set_content(
        f"""Hi Technical Support Team,

        This is a soft alert generated by the ordering pipeline.

        One or more items exceeded the configured cases-to-order
        threshold of {threshold}. This usually is a result of Units Per Case being set to 1 in the Vendor Price Book by mistake.

        This alert does NOT block the pipeline and the order reports were still sent.

        Google Sheet:
        {sheet_link}

        Attached: {pdf_name}

        — Sent from an automated reporting pipeline
        """
    )

    msg.add_alternative(
        f"""
        <p><strong>Soft Alert – High Case Quantities</strong></p>
        <p>One or more items exceeded the configured cases-to-order threshold of {threshold}. This usually is a result of Units Per Case being set to 1 in the Vendor Price Book by mistake.</p>
        <p>This alert does NOT block the pipeline and the order reports were still sent.</p>
        <p><a href="{sheet_link}">Open Alert Sheet in Google Sheets</a></p>
        <p>Attached: {pdf_name}</p>
        <p>— Sent from an automated reporting pipeline</p>
        """,
        subtype="html",
    )

    msg.add_attachment(
        pdf_bytes,
        maintype="application",
        subtype="pdf",
        filename=pdf_name,
    )

    return send_email(gmail_svc, sender, msg)
```

---
### file: core_functional_modules/google_client.py

```python
"""
google_client
======================================
This module centralizes Google OAuth 2.0 sign‑in for local Python applications,
supporting both:

1) **Classic CLI flow** – prints an authorization URL to the console and accepts
   a pasted redirect URL or auth code (for headless shells, remote SSH, or when
   opening a browser is impractical).

2) **Streamlit-/Desktop-friendly local-server flow** – opens the user's default
   browser and spins up a temporary HTTP listener on ``127.0.0.1`` to complete
   the OAuth redirect without any console copy/paste.

It also provides small helpers to manage a persistent token cache
(``token.json``), refresh expired tokens when possible, and construct Google API
service clients (Sheets, Drive, Gmail) using the official ``google-api-python-client``.

-------------------------------------------------------------------------------
Key Features
-------------------------------------------------------------------------------
- **Token cache**: Reads/writes ``token.json`` to persist credentials between runs.
  Includes best-effort refresh of expired tokens when a refresh token exists.
- **Two auth paths**:
  - *CLI/manual path*: URL is printed; user pastes back the full redirect URL or
    just the ``code`` parameter.
  - *Local server/one-click path*: Automatically opens browser and listens on a
    local port (tries an OS-chosen free port first, then a configured fallback).
- **Graceful fallbacks**: If automated browser auth fails, raises a descriptive
  error suggesting the manual method.
- **Service builders**: Convenience helpers to create Sheets v4, Drive v3, and
  Gmail v1 service clients with the provided credentials.

-------------------------------------------------------------------------------
Files Used
-------------------------------------------------------------------------------
- ``credentials.json`` (required):
  The OAuth 2.0 client secrets file downloaded from Google Cloud Console.

- ``token.json`` (optional, auto-created):
  The persisted user credentials (access/refresh tokens). If present and valid,
  it is reused to avoid re-authentication. If expired but refreshable, it is
  refreshed automatically and re-written.

-------------------------------------------------------------------------------
Function Overview
-------------------------------------------------------------------------------
- ``clear_token()``:
    Deletes ``token.json`` if present (best-effort). Useful to force a
    re-authentication scenario.

- ``load_valid_token(scopes) -> Optional[Credentials]``:
    Loads credentials from ``token.json`` for the given scopes. If expired but
    refreshable, refreshes and persists the updated token. Returns a valid
    ``Credentials`` or ``None``.

- ``get_credentials(scopes, redirect_port, force_reauth=False) -> Credentials``:
    **CLI-friendly** method. If no valid token exists, prints an auth URL and
    prompts for a pasted redirect URL or code. Persists the resulting token to
    ``token.json``.

- ``login_via_local_server(scopes, redirect_port) -> Credentials``:
    **Streamlit-/desktop-friendly** one-click OAuth that opens a browser and
    listens on ``127.0.0.1``. Tries an OS-chosen free port first (``port=0``),
    then the provided ``redirect_port``. Uses a 120s timeout for safety.

- ``start_oauth(scopes, redirect_port) -> (InstalledAppFlow, auth_url)``:
    Starts the manual flow by creating an ``InstalledAppFlow`` with a configured
    redirect URI and returns the authorization URL to display in your own UI.

- ``finish_oauth(flow, pasted) -> Credentials``:
    Completes the manual flow using the pasted redirect URL (or raw ``code``),
    fetches tokens, writes ``token.json``, and returns ``Credentials``.

- ``_service(api, version, creds)``:
    Internal helper to construct a Google API service for the given
    ``api``/``version`` using the supplied ``Credentials``.

- ``services(creds, _http_timeout_seconds)``:
    Convenience function returning a tuple of ready-to-use clients:
    ``(sheets, drive, gmail)``. The ``_http_timeout_seconds`` parameter is
    currently reserved for future use.

-------------------------------------------------------------------------------
Error Handling & Edge Cases
-------------------------------------------------------------------------------
- If ``credentials.json`` is missing, a ``FileNotFoundError`` is raised early.
- Token refresh failures fall back to a fresh login.
- The local-server path uses a 120-second timeout to avoid hanging the process.
- If both automatic local-server attempts fail, a ``RuntimeError`` is raised
  advising the manual copy/paste method with detailed error messages from both
  attempts.
-------------------------------------------------------------------------------
Maintainer Tips
-------------------------------------------------------------------------------
- If you add new Google APIs, extend ``services(...)`` or call ``_service(...)``
  directly with the desired API name/version.
- Consider surfacing the timeout and host/port as user configuration if your app
  needs more control in diverse environments.

"""

from __future__ import annotations
import os
from urllib.parse import urlparse, parse_qs

from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build


# ---------- Token helpers ----------

def clear_token():
    """Delete token.json if present."""
    try:
        if os.path.exists("token.json"):
            os.remove("token.json")
    except Exception:
        pass


def load_valid_token(scopes):
    """
    Try to load token.json. If expired but refreshable, refresh it and persist.
    Returns valid Credentials or None.
    """
    if not os.path.exists("token.json"):
        return None
    try:
        creds = Credentials.from_authorized_user_file("token.json", scopes)
    except Exception:
        return None

    if creds and creds.valid:
        return creds

    if creds and creds.expired and creds.refresh_token:
        try:
            creds.refresh(Request())
            with open("token.json", "w") as f:
                f.write(creds.to_json())
            return creds
        except Exception:
            return None

    return None


# ---------- Classic CLI path (kept for completeness) ----------

def get_credentials(scopes, redirect_port: int, force_reauth: bool = False) -> Credentials:
    """
    CLI-friendly: prints URL and waits for input() if token is missing/invalid.
    The Streamlit UI uses the in-UI functions below instead.
    """
    if force_reauth:
        clear_token()

    creds = load_valid_token(scopes)
    if creds:
        return creds

    if not os.path.exists("credentials.json"):
        raise FileNotFoundError("Missing credentials.json in working directory")

    flow = InstalledAppFlow.from_client_secrets_file("credentials.json", scopes)
    flow.redirect_uri = f"http://127.0.0.1:{redirect_port}/"
    auth_url, _ = flow.authorization_url(
        access_type="offline",
        include_granted_scopes="true",
        prompt="consent",
    )
    print("Open this URL and complete the login:\n", auth_url)
    pasted = input("Paste full redirect URL or auth code here: ").strip()
    code = pasted
    if pasted.startswith("http"):
        qs = parse_qs(urlparse(pasted).query)
        if "code" in qs:
            code = qs["code"][0]
    flow.fetch_token(code=code)
    creds = flow.credentials
    with open("token.json", "w") as f:
        f.write(creds.to_json())
    return creds


# ---------- Streamlit-friendly OAuth (no console) ----------

# favtrip/google_client.py

def login_via_local_server(scopes, redirect_port: int) -> Credentials:
    """
    One-click OAuth: open browser and listen on 127.0.0.1.
    Tries OS-chosen port first, then the configured port.
    Uses a timeout to avoid hanging indefinitely.
    NOTE: No optional text parameters are passed, for compatibility with older google-auth-oauthlib.
    """
    if not os.path.exists("credentials.json"):
        raise FileNotFoundError("Missing credentials.json in working directory")

    # Attempt 1: OS-chosen free port (port=0)
    flow = InstalledAppFlow.from_client_secrets_file("credentials.json", scopes)
    try:
        creds = flow.run_local_server(
            host="127.0.0.1",
            port=0,                 # let OS choose a free port
            open_browser=True,
            timeout_seconds=120,    # bail out after 2 minutes
        )
        with open("token.json", "w") as f:
            f.write(creds.to_json())
        return creds
    except Exception as first_err:
        # Attempt 2: user-configured port (from .env)
        flow = InstalledAppFlow.from_client_secrets_file("credentials.json", scopes)
        try:
            creds = flow.run_local_server(
                host="127.0.0.1",
                port=int(redirect_port),
                open_browser=True,
                timeout_seconds=120,
            )
            with open("token.json", "w") as f:
                f.write(creds.to_json())
            return creds
        except Exception as second_err:
            raise RuntimeError(
                "Automatic browser auth failed both on a random port and on your configured REDIRECT_PORT. "
                "Please use the manual method (copy/paste URL). "
                f"Details: first={first_err}; second={second_err}"
            )


def start_oauth(scopes, redirect_port: int):
    """
    Manual fallback: returns (flow, auth_url) for paste-based completion.
    """
    if not os.path.exists("credentials.json"):
        raise FileNotFoundError("Missing credentials.json in working directory")
    flow = InstalledAppFlow.from_client_secrets_file("credentials.json", scopes)
    flow.redirect_uri = f"http://127.0.0.1:{redirect_port}/"
    auth_url, _ = flow.authorization_url(
        access_type="offline",
        include_granted_scopes="true",
        prompt="consent",
    )
    return flow, auth_url


def finish_oauth(flow: InstalledAppFlow, pasted: str) -> Credentials:
    """
    Manual fallback: accepts the pasted redirect URL or the code; returns Credentials and writes token.json.
    """
    code = pasted.strip()
    if pasted.startswith("http"):
        qs = parse_qs(urlparse(pasted).query)
        if "code" in qs:
            code = qs["code"][0]
    flow.fetch_token(code=code)
    creds = flow.credentials
    with open("token.json", "w") as f:
        f.write(creds.to_json())
    return creds


# ---------- Google services ----------

def _service(api: str, version: str, creds: Credentials):
    # Pass credentials directly (no google_auth_httplib2 dependency)
    return build(api, version, credentials=creds, cache_discovery=False)


def services(creds: Credentials, _http_timeout_seconds: int):
    sheets = _service("sheets", "v4", creds)
    drive = _service("drive", "v3", creds)
    gmail = _service("gmail", "v1", creds)
    return sheets, drive, gmail

```

---
### file: core_functional_modules/logger.py

```python
"""
logger
======================================
This module provides two dataclasses—`LogEvent` and `StatusLogger`—to record simple,
human-readable status messages during a process or script run. It is designed to be:

- **Simple**: minimal API (`info`, `warn`, `error`) and a small in-memory log.
- **Immediate**: console prints occur synchronously; file writes are line-buffered and flushed.
- **Fail-open**: if a log file cannot be opened or written, logging proceeds to console and memory.
- **Portable**: standard library only (dataclasses, datetime, typing).

-------------------------------------------------------------------------------
Data Model
-------------------------------------------------------------------------------
- LogEvent
    - ts (datetime.datetime): Timestamp captured via `datetime.now()` when the event is recorded.
      Note: this is a **naive** datetime in local time.
    - level (str): Log level label (e.g., "INFO", "WARN", "ERROR").
    - message (str): The event text.

- StatusLogger
    - events (list[LogEvent]): In-memory event history in append order.
    - print_to_console (bool): If True (default), each log line is printed to stdout.
    - file_path (str | None): If set, lines are also written to this file. If `None`, file logging
      is disabled. Default is "last_run.log".
    - overwrite (bool): If True (default), the log file is opened in write mode on instantiation;
      otherwise it is appended to.

-------------------------------------------------------------------------------
Output Format
-------------------------------------------------------------------------------
- Console/file lines: `[YYYY-MM-DD HH:MM:SS] LEVEL: message`
- `as_text()`:         `[HH:MM:SS] LEVEL: message` per line (no date, suitable for compact display)
- `last_line()`:       Returns the most recent line in `as_text()` format, or `"Starting…"` if empty.

-------------------------------------------------------------------------------
Behavior & Guarantees
-------------------------------------------------------------------------------
- **File handling**: On initialization, if `file_path` is provided, the file is opened once in
  line-buffered text mode (`buffering=1`) and UTF-8 encoding. If opening fails, the logger
  continues without a file handle.
- **Atomicity**: Each `_emit` call attempts to write a single line and then flush. Any file write
  errors are swallowed; console output and in-memory storage are unaffected.
- **Timestamps**: Timestamps are captured at call time (`datetime.now()`), local time, naive datetimes.
- **Memory growth**: All events are retained in `events`; for long-running processes, consider
  pruning or exporting periodically.
- **Thread-safety**: Not thread-safe. If you need concurrent logging, protect calls with a lock or
  adapt the implementation for multi-thread/process usage.
- **No rotation**: No file rotation or size limiting. Use external tools or extend as needed.
"""

from dataclasses import dataclass, field
from datetime import datetime
from typing import List, Optional

@dataclass
class LogEvent:
    ts: datetime
    level: str
    message: str

@dataclass
class StatusLogger:
    events: List[LogEvent] = field(default_factory=list)
    print_to_console: bool = True
    file_path: Optional[str] = "last_run.log"
    overwrite: bool = True

    def __post_init__(self):
        # Prepare the file on first use
        self._fh = None
        if self.file_path:
            mode = "w" if self.overwrite else "a"
            try:
                self._fh = open(self.file_path, mode, encoding="utf-8", buffering=1)  # line-buffered
            except Exception:
                # If we cannot open a file, we keep running without file logging
                self._fh = None

    def _emit(self, line: str):
        if self.print_to_console:
            print(line)
        if self._fh:
            try:
                self._fh.write(line + "\n")
                self._fh.flush()  # ensure immediate persistence
            except Exception:
                pass

    def _log(self, level: str, message: str):
        evt = LogEvent(datetime.now(), level, message)
        self.events.append(evt)
        self._emit(f"[{evt.ts:%Y-%m-%d %H:%M:%S}] {level}: {message}")

    def info(self, message: str):
        self._log("INFO", message)

    def warn(self, message: str):
        self._log("WARN", message)

    def error(self, message: str):
        self._log("ERROR", message)

    def as_text(self) -> str:
        return "\n".join(f"[{e.ts:%H:%M:%S}] {e.level}: {e.message}" for e in self.events)

    def last_line(self) -> str:
        if not self.events:
            return "Starting…"
        e = self.events[-1]
        return f"[{e.ts:%H:%M:%S}] {e.level}: {e.message}"

    def close(self):
        try:
            if self._fh:
                self._fh.close()
        except Exception:
            pass

```

---
### file: core_functional_modules/pipeline.py

```python
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
    add_or_replace_sheet, put_values_2d, _force_column_as_text, delete_row_indices, delete_rows_range, copy_sheet_to_another_spreadsheet, autoresize_columns, export_sheet
)
from .drive_utils import find_latest_sheet, upload_to_drive, _rfc3339, trash_file, cleanup_folder_by_age, find_sheet_by_name, copy_file_to_folder, rename_file, get_or_create_subfolder
from .gmail_utils import send_email, email_manager_report, email_order_report, email_error_report, email_bev_error_report, email_large_case_alert_report

CSV_MIME = "text/csv"


def clean_tag(s: str | None) -> str | None:
    if s is None:
        return None
    s = str(s).strip()
    if not s:
        return None
    return re.sub(r"[^A-Za-z0-9._-]+", "-", s).strip("-")



import requests
from io import BytesIO
from openpyxl import Workbook


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


def filter_master_csv_to_ran(master_csv_bytes, cfg):
    text = master_csv_bytes.decode("utf-8-sig", errors="replace")
    reader = csv.reader(io.StringIO(text))
    rows = list(reader)

    if not rows:
        return master_csv_bytes  # nothing to do

    headers = [h.strip() for h in rows[0]]
    lower_idx = {h.lower(): i for i, h in enumerate(headers)}

    report_idx = lower_idx.get("report_key")
    sub_idx = lower_idx.get("sub_report_key")

    if report_idx is None:
        raise RuntimeError("Master CSV missing Report_Key column")

    output = io.StringIO()
    writer = csv.writer(output, lineterminator="\n")
    writer.writerow(headers)

    for r in rows[1:]:
        key = (r[report_idx] if report_idx < len(r) else "").strip().upper()
        sub = None
        if sub_idx is not None:
            sub = (r[sub_idx] if sub_idx < len(r) else "").strip().upper() or None

        if should_run(cfg, key, sub):
            writer.writerow(r)

    return output.getvalue().encode("utf-8-sig")

def sort_master_csv(csv_bytes: bytes) -> bytes:
    df = pandas.read_csv(io.BytesIO(csv_bytes))

    # Normalize column names
    cols = {c.lower(): c for c in df.columns}

    report_col = cols.get("report_key")
    sub_col = cols.get("sub_report_key")
    cases_col = cols.get("cases to order")

    if not report_col:
        raise RuntimeError("Report_Key column not found for sorting")

    # Fill missing sub-keys so they sort last
    if sub_col:
        df[sub_col] = df[sub_col].fillna("ZZZ")

    # Ensure numeric sorting for cases
    if cases_col:
        df[cases_col] = pandas.to_numeric(df[cases_col], errors="coerce").fillna(0)

    sort_cols = [report_col]
    sort_ascending = [True]

    if sub_col:
        sort_cols.append(sub_col)
        sort_ascending.append(True)

    if cases_col:
        sort_cols.append(cases_col)
        sort_ascending.append(False)  # DESCENDING

    df = df.sort_values(
        by=sort_cols,
        ascending=sort_ascending,
        kind="mergesort"  # stable sort
    )

    # Restore blanks if we filled them
    if sub_col:
        df[sub_col] = df[sub_col].replace("ZZZ", "")

    out = io.StringIO()
    df.to_csv(out, index=False)
    return out.getvalue().encode("utf-8-sig")


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

    
    ran_report_keys = set()
    ran_sub_report_keys = set()
    ran_reports = {}


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

    # Step 4: Manager Report PDF
    if logger:
        logger.info("Exporting Manager Report (PDF)…")
    pdf_bytes = export_sheet(creds, calc_ss_id, cfg.GID_MANAGER_PDF, "pdf", True)
    pdf_name = f"Manager_Report_{ts}_{location}.pdf"
    uploaded_pdf = upload_to_drive(drive_svc, pdf_bytes, pdf_name, "application/pdf", cfg.MANAGER_REPORT_FOLDER_ID, to_sheet=False)
    manager_link = uploaded_pdf.get("webViewLink")
    if logger:
        logger.info(f"Uploaded Manager PDF: {manager_link}")

    # Step 5: Master Order CSV
    if logger:
        logger.info("Exporting Master Order (CSV)…")
    master_csv_bytes = export_sheet(creds, calc_ss_id, cfg.GID_ORDER_CSV, "csv")
    master_csv_bytes = filter_master_csv_to_ran(master_csv_bytes, cfg)
    master_csv_bytes = sort_master_csv(master_csv_bytes)

    # Step 6: Error Report CSV, Upload, Export PDF
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
        
        autoresize_columns(sheets_svc, err_file_id, err_gid)

        err_pdf = export_sheet(creds, err_file_id, err_gid, "pdf", False)
        err_pdf_name = f"Error_Report_{ts}.pdf"

        if logger:
            logger.info(f"Uploaded filtered Error Sheet: {err_link}")

        # Step 6.1: Send Error Report if Needed

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




    # Step 7: Full order upload (CSV) and export (PDF)
    full_csv_name = f"Order_Report_FULL_{location}_{ts}.csv"
    full_created = upload_to_drive(drive_svc, master_csv_bytes, full_csv_name, CSV_MIME, cfg.ORDER_REPORT_FOLDER_ID, to_sheet=True)
    full_file_id = full_created["id"]
    full_link = full_created.get('webViewLink')
    full_gid = first_gid(sheets_svc, full_file_id)
    autoresize_columns(sheets_svc, full_file_id, full_gid)
    full_pdf = export_sheet(creds, full_file_id, full_gid, "pdf", False)
    full_pdf_name = f"Order_Report_FULL_{location}_{ts}.pdf"
    if logger:
        logger.info(f"Uploaded FULL sheet: {full_created.get('webViewLink')}")


    #Step 7.1: Large Case Alert
    try:
        if cfg.SOFT_CASES_ALERT_ENABLED:
            if logger:
                logger.info(
                    f"Checking FULL order for case quantities > "
                    f"{cfg.SOFT_CASES_ALERT_THRESHOLD}"
                )

            # Load FULL order CSV into DataFrame
            df = pandas.read_csv(io.BytesIO(master_csv_bytes))

            # Normalize column lookup (case-insensitive)
            lower_cols = {c.lower(): c for c in df.columns}
            cases_col = lower_cols.get("cases to order")

            if not cases_col:
                if logger:
                    logger.warn(
                        "Large case alert skipped — 'Cases to Order' "
                        "column not found in FULL order CSV"
                    )
            else:
                # Ensure numeric comparison
                df[cases_col] = pandas.to_numeric(
                    df[cases_col], errors="coerce"
                ).fillna(0)

                flagged = df[
                    df[cases_col] > cfg.SOFT_CASES_ALERT_THRESHOLD
                ]

                if flagged.empty:
                    if logger:
                        logger.info(
                            "No FULL order rows exceed case threshold"
                        )
                else:
                    if logger:
                        logger.warn(
                            f"Soft alert triggered: {len(flagged)} "
                            f"rows exceed case threshold"
                        )

                    # --------------------------------------------------
                    # Create filtered CSV (only flagged rows)
                    # --------------------------------------------------
                    buf = io.StringIO()
                    flagged.to_csv(buf, index=False)
                    alert_csv_bytes = buf.getvalue().encode("utf-8-sig")

                    alert_csv_name = (
                        f"Large_Case_Alert_{location}_{ts}.csv"
                    )

                    # Upload alert CSV → Google Sheet
                    created = upload_to_drive(
                        drive_svc,
                        alert_csv_bytes,
                        alert_csv_name,
                        CSV_MIME,
                        cfg.ERROR_REPORT_FOLDER_ID,
                        to_sheet=True,
                    )

                    alert_sheet_id = created["id"]
                    alert_sheet_link = created.get("webViewLink")

                    alert_gid = first_gid(sheets_svc, alert_sheet_id)
                    autoresize_columns(
                        sheets_svc,
                        alert_sheet_id,
                        alert_gid,
                    )

                    # Export alert PDF
                    alert_pdf_bytes = export_sheet(
                        creds,
                        alert_sheet_id,
                        alert_gid,
                        "pdf",
                        False,
                    )
                    alert_pdf_name = (
                        f"Large_Case_Alert_{location}_{ts}.pdf"
                    )

                    # --------------------------------------------------
                    # Resolve recipients (technical first)
                    # --------------------------------------------------
                    to_list = _fallback_recipients(
                        "LARGE CASE ALERT",
                        cfg.ERROR_RECIPIENTS,
                        cfg.TO_RECIPIENTS,
                        cfg.DEFAULT_ORDER_RECIPIENTS,
                    )

                    cc_list = [
                        e for e in _clean_emails(cfg.CC_RECIPIENTS)
                        if e not in to_list
                    ]

                    # --------------------------------------------------
                    # Send email (SOFT ALERT)
                    # --------------------------------------------------
                    email_large_case_alert_report(
                        gmail_svc=gmail_svc,
                        sender="me",
                        to_list=to_list,
                        cc_list=cc_list,
                        ts=ts,
                        location=location,
                        threshold=cfg.SOFT_CASES_ALERT_THRESHOLD,
                        pdf_name=alert_pdf_name,
                        pdf_bytes=alert_pdf_bytes,
                        sheet_link=alert_sheet_link,
                    )

                    if logger:
                        logger.info(
                            "Large case quantity alert email sent"
                        )

    except Exception as e:
        # 🔐 Soft failure only — never block pipeline
        if logger:
            logger.warn(
                f"Large case quantity alert failed (soft): {e}"
            )

    # Step 8: Create per-report-key outputs (CSV) and email

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
    

    bev_order_sent = False
    
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
        autoresize_columns(sheets_svc, file_id, gid)
        pdf = export_sheet(creds, file_id, gid, "pdf", False)
        pdfname = f"Order_Report_{'_'.join(name_parts)}_{ts}.pdf"
    
        # Prefer Store+Key; else Key; else Store; else To; else Default
        candidates = None
        
        candidates = None
        lookup_order = [
            (store_tag, tag, sub_tag), #Independence, BEV, 7UP
            (store_tag, None, sub_tag),  #Independence, 7UP
            (None, None, sub_tag),  #7UP
            (store_tag, tag, None),  #Independence, BEV
            (None, tag, None),  #BEV
            (store_tag, None, None),  #Independence
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

        
        email_tag_parts = [tag]
        if sub_tag:
            email_tag_parts.append(sub_tag)

        email_tag = " - ".join(email_tag_parts)



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

        
        if key.upper() == "BEV":
            bev_order_sent = True

    
        if logger:
            logger.info(f"Emailed {store} - {email_tag} to {recipients}")
    
    # Step 9: Unassigned Beverages Report (Soft Error)
    try:
        # Must have successfully sent a BEV order
        if not bev_order_sent:
            if logger:
                logger.info("Skipping Unassigned Beverages Report — no BEV order was sent.")
        else:
            if logger:
                logger.info("Exporting Unassigned Beverages Report (CSV)…")

            unassigned_csv_bytes = export_sheet(
                creds,
                calc_ss_id,
                cfg.GID_BEV_ERRORS,
                "csv"
            )

            # Inspect CSV to ensure it actually has data
            text = unassigned_csv_bytes.decode("utf-8-sig", errors="replace")
            reader = csv.reader(io.StringIO(text))
            rows = list(reader)

            if not rows or len(rows) <= 1:
                if logger:
                    logger.info("No unassigned beverages found — report will not be sent.")
            else:
                headers = [h.strip() for h in rows[0]]
                lower_idx = {h.lower(): i for i, h in enumerate(headers)}

                if "report_key" not in lower_idx:
                    raise RuntimeError("Unassigned BEV report missing Report_Key column")

                if "sub_report_key" not in lower_idx:
                    raise RuntimeError("Unassigned BEV report missing Sub_Report_Key column")

                report_idx = lower_idx["report_key"]
                sub_idx = lower_idx["sub_report_key"]

                # Filter to BEV + BEV_UNASSIGNED only
                unassigned_rows = [
                    r for r in rows[1:]
                    if r[report_idx].strip().upper() == "BEV"
                    and r[sub_idx].strip().upper() == "UNASSIGNED"
                ]

                if not unassigned_rows:
                    if logger:
                        logger.info("No BEV_UNASSIGNED rows found — skipping email.")
                else:
                    # Upload filtered sheet
                    output = io.StringIO()
                    writer = csv.writer(output)
                    writer.writerow(rows[0])
                    writer.writerows(unassigned_rows)

                    filtered_bytes = output.getvalue().encode("utf-8-sig")

                    csv_name = f"Unassigned_Beverages_Report_{ts}.csv"

                    created = upload_to_drive(
                        drive_svc,
                        filtered_bytes,
                        csv_name,
                        CSV_MIME,
                        cfg.ERROR_REPORT_FOLDER_ID,
                        to_sheet=True
                    )

                    sheet_id = created["id"]
                    sheet_link = created.get("webViewLink")
                    gid = first_gid(sheets_svc, sheet_id)

                    autoresize_columns(sheets_svc, sheet_id, gid)
                    pdf_bytes = export_sheet(creds, sheet_id, gid, "pdf", False)
                    pdf_name = f"Unassigned_Beverages_Report_{ts}.pdf"

                    # Resolve recipients
                    to_list = _fallback_recipients(
                        "UNASSIGNED BEVERAGES REPORT",
                        cfg.ERROR_RECIPIENTS,
                        cfg.TO_RECIPIENTS,
                        cfg.DEFAULT_ORDER_RECIPIENTS,
                    )

                    cc_list = list(dict.fromkeys(
                        set(_clean_emails(cfg.TO_RECIPIENTS))
                        | set(_clean_emails(cfg.CC_RECIPIENTS))
                        - set(to_list)
                    ))

                    email_bev_error_report(
                        gmail_svc=gmail_svc,
                        sender="me",
                        to_list=to_list,
                        cc_list=cc_list,
                        ts=ts,
                        pdf_name=pdf_name,
                        pdf_bytes=pdf_bytes,
                        sheet_link=sheet_link,
                        mapping_link=cfg.BEV_MAPPING_LINK,
                    )

                    if logger:
                        logger.info("Unassigned Beverages Report email sent")

    except Exception as e:
        # Soft error — log and continue
        if logger:
            logger.warn(f"Unassigned Beverages Report failed (soft): {e}")
        
    # Step 10: Send Manager Report (guarded by cfg.EMAIL_MANAGER_REPORT)
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

    

    # Step 11: Send Full Order if needed
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

    # Step 12: File Cleanup

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

```

---
### file: core_functional_modules/pipeline_bus.py

```python
# pipeline_bus.py
import queue

_PIPELINE_QUEUE = None

def get_pipeline_queue():
    global _PIPELINE_QUEUE
    if _PIPELINE_QUEUE is None:
        _PIPELINE_QUEUE = queue.Queue()
    return _PIPELINE_QUEUE
```

---
### file: core_functional_modules/rebuild_google_workspace.py

```python
import streamlit as st
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
        st.secrets.get("DEV_CONFIG_FILE_ID")
        or None
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
        "message": (
            "Please replace the following secrets in Streamlit:\n\n"
            f"CONFIG_FILE_ID = {config_ids["prod_config_file_id"]}\n"
            f"DEV_CONFIG_FILE_ID = {config_ids["dev_config_file_id"]}"
        )
    }

```

---
### file: core_functional_modules/sheets_utils.py

```python
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
import requests


def list_sheets(svc, spreadsheet_id: str) -> List[Dict[str, Any]]:
    return svc.spreadsheets().get(spreadsheetId=spreadsheet_id).execute().get("sheets", [])


def get_sheet(sheets, title: str):
    for s in sheets:
        if s["properties"]["title"] == title:
            return s["properties"]
    return None


def export_sheet(creds, spreadsheet_id: str, gid: str | int, fmt: str, portrait: bool = True,) -> bytes:
    params = {
        "format": fmt,
        "gid": gid,
    }

    # PDF-only layout options
    if fmt.lower() == "pdf":
        params.update({
            "portrait": "true" if portrait else "false",
            "fitw": "true",   # fit to width
        })

    # Build query string
    query = "&".join(f"{k}={v}" for k, v in params.items())

    url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export?{query}"

    headers = {"Authorization": f"Bearer {creds.token}"}
    r = requests.get(url, headers=headers)
    r.raise_for_status()
    return r.content



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

def autoresize_columns(sheets_svc, spreadsheet_id, sheet_id):
    sheets_svc.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={
            "requests": [{
                "autoResizeDimensions": {
                    "dimensions": {
                        "sheetId": sheet_id,
                        "dimension": "COLUMNS",
                        "startIndex": 0,
                        "endIndex": 26  # A–Z (adjust if needed)
                    }
                }
            }]
        }
    ).execute()

```

---
### file: documentation/PROJECT_SNAPSHOT_CODEBUNDLE.md

```markdown
# Project Snapshot (CodeBundle)

                This single Markdown file contains a **self-contained snapshot** of your project so another AI/engineer can review or modify it without needing the original folder.

                **How to use this file with an AI**
                1. Upload or paste this file as a single attachment.
                2. Ask for changes; the AI can reference specific `file:` sections below.
                3. Copy updated blocks back into the corresponding files in your project.

                > Notes: secrets like `token.json` are intentionally excluded. Virtual envs and build artifacts are omitted to keep this readable.

                ## Directory tree (filtered)
                .env [skipped: secret]
.env.example
.gitignore
FavTripPipeline.spec
FavTripPipelineUI.spec
_user_interface_.py
cli.py
credentials.json [skipped: secret]
last_run.log
launcher_streamlit.py
requirements.txt
setup_py2app.py
token.json [skipped: secret]
web_url_credentials.json [skipped: secret]
  core_functional_modules/
    __init__.py
    config.py
    config_store.py
    drive_utils.py
    gmail_utils.py
    google_client.py
    logger.py
    pipeline.py
    pipeline_bus.py
    sheets_utils.py
  documentation/
    PROJECT_SNAPSHOT_CODEBUNDLE.md
    README.md
    generate_code_bundle.py
    git_workflow.txt
    requirements.txt
  __dev_input_sales_file/
    1 Store - 1 Week.xlsx
    1 Store - 2 Weeks.xlsx  [skipped: too large]
    1 Store - Bad End.xlsx  [skipped: too large]
    1 Store - Bad Start.xlsx
    2 Stores - 1 Week.xlsx  [skipped: too large]
    2 Stores - 2 Weeks.xlsx  [skipped: too large]
    VPB Error - BEV.xlsx  [skipped: too large]
  __dev_input_vendor_file/
    Vendors Price Book.xlsx
  __executable/
    run_windows.bat
---
### file: .env.example

```
# --- Required IDs ---
CALC_SPREADSHEET_ID=
INCOMING_FOLDER_ID=
MANAGER_REPORT_FOLDER_ID=
ORDER_REPORT_FOLDER_ID=

# --- Optional IDs / settings ---
GID_MANAGER_PDF=1921812573
GID_ORDER_CSV=1875928148
LOCATION_SHEET_TITLE=REFR: Values
LOCATION_NAMED_RANGE=_locations

TIMESTAMP_TZ=America/Chicago
TIMESTAMP_FMT=%Y-%m-%d-%I-%M-%p

# Recipients
TO_RECIPIENTS=
CC_RECIPIENTS=
DEFAULT_ORDER_RECIPIENTS=

# Report keys
USE_ALL_REPORT_KEYS=false
REPORT_KEY_RUN_LIST=GROCERY,COFFEE
# JSON mapping: {"GROCERY":["a@b.com","c@d.com"],"COFFEE":["x@y.com"]}
REPORT_KEY_RECIPIENTS={}

INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL=false
SEND_SEPARATE_FULL_ORDER_EMAIL=true

# Google API scopes (normally leave as-is)
SCOPES=https://www.googleapis.com/auth/drive,https://www.googleapis.com/auth/spreadsheets,https://www.googleapis.com/auth/gmail.send

FORCE_REAUTH=false
REDIRECT_PORT=58285
HTTP_TIMEOUT_SECONDS=300

```

---
### file: .gitignore

```
*.env
*credentials.json
token.json
web_url_credentials.json

```

---
### file: FavTripPipeline.spec

```
# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['cli.py'],
    pathex=[],
    binaries=[],
    datas=[('.env', '.'), ('credentials.json', '.')],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='FavTripPipeline',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

```

---
### file: FavTripPipelineUI.spec

```
# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['launcher_streamlit.py'],
    pathex=[],
    binaries=[],
    datas=[('.env', '.'), ('credentials.json', '.'), ('ui_streamlit.py', '.'), ('favtrip', 'favtrip')],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='FavTripPipelineUI',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

```

---
### file: __executable/run_windows.bat

```bat
@echo off
setlocal
REM ---------------------------------------------------------------------------
REM Run Streamlit UI without a persistent console window.
REM Location: __executable\run_web_windows_silent.bat
REM Behavior: brief flash at launch, then only the browser tab remains.
REM ---------------------------------------------------------------------------

REM Move into the folder of this .bat
pushd "%~dp0"

REM Go to the project root (one level up from __executable)
cd ..

REM Choose Python: prefer venv's interpreter if present
set "PY_VENV=.\.venv\Scripts\python.exe"
set "PY="
if exist "%PY_VENV%" (
  set "PY=%PY_VENV%"
) else (
  for %%P in (python.exe py.exe) do (
    where %%P >nul 2>&1 && (set "PY=%%P" & goto :gotpy)
  )
)
:gotpy
if not defined PY (
  echo [Launcher] Python was not found. Install Python or create .\.venv and try again.
  popd
  exit /b 1
)

REM Streamlit prefs: ensure it opens the browser and stays local
set "STREAMLIT_SERVER_HEADLESS=false"
set "STREAMLIT_BROWSER_GATHER_USAGE_STATS=false"

REM Start Streamlit hidden and detach from this console (which then closes)
REM - We invoke PowerShell only to spawn the hidden child process.
start "" /MIN powershell -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -Command ^
  "Start-Process -FilePath '%PY%' -ArgumentList '-m','streamlit','run','ui_streamlit.py' -WindowStyle Hidden"

popd
exit /b 0
```

---
### file: _user_interface_.py

```python
"""
FavTrip Reporting Streamlit UI
================================

Overview
--------
This Streamlit app is the front-end for the FavTrip Reporting pipeline. It lets an authenticated
Google user upload a Modisoft "Live Items Report", configure who receives emails for each report
(or fallback recipients), tune advanced IDs/GIDs/time settings, and then orchestrate the
`core_functional_modules.pipeline.run_pipeline` execution while streaming status updates and a timer.

Design goals
------------
* **No local secrets in code**: OAuth client JSON is read from `st.secrets["GOOGLE_CREDENTIALS"]` and
  the app base URL from `st.secrets["APP_BASE_URL"]`. Optional `CONFIG_FILE_ID` pins a Drive JSON that
  stores editable defaults.
* **Constrained, clear UX**: A two-step flow (Upload ➜ Run). The Run button is enabled only after a
  successful upload, reducing accidental runs on stale input.
* **Robust OAuth (PKCE)**: Uses an explicit code verifier/challenge and encodes the verifier inside
  the `state` payload to remain stateless across redirects.
* **Operational safety**: Detects common mistakes (e.g., invalid email inputs, duplicate keys,
  wrong-week uploads) and surfaces warnings or blocks execution accordingly.

Key concepts
------------
* **Incoming file**: The uploaded Modisoft report (CSV/XLSX). It is pushed to a configured Google
  Drive folder and optionally converted to a Google Sheet for the downstream pipeline.
* **Report Keys**: Categories/tags used by the pipeline to partition output and email recipients.
  You can either process *all* keys present in the data or restrict to a comma-separated subset.
* **Per-Report-Key Recipients**: Optional overrides that map `(Store, Report Key)` pairs to recipient
  lists. Fallback recipients apply where no specific mapping exists.
* **Drive-backed defaults**: The app can persist your current UI settings to a JSON in Drive. Supplying
  `CONFIG_FILE_ID` in Streamlit secrets will cause subsequent sessions to update that exact file.

Security model
--------------
* OAuth scopes are supplied by `Config` and used to mint a user token saved locally as `token.json`.
* The app opens the Google consent screen in a **new tab**, and that tab becomes the main app after
  redirect. Tokens are not sent back to the opener page; they are stored only in the process serving
  the tab that completed OAuth.

Operational notes
-----------------
* If a run fails with the message "Please only upload 1 or 2 full weeks of data", the UI locks to
  prevent immediate re-runs. Use **Retry** to clear the lock and upload a correct file.
* Set **Offer full log download** (sidebar) to expose a download button for `last_run.log` after a run.
* The **green Run button** appears once a fresh upload succeeds, indicating the pipeline is ready to run.

Dependencies & integration points
---------------------------------
* `core_functional_modules.google_client`: token loading/clearing and service factories (Drive, Sheets, Gmail)
* `core_functional_modules.config`: the central configuration object. `Config.load()` merges defaults, secrets, and any
  Drive-stored overrides.
* `core_functional_modules.drive_utils.upload_to_drive`: uploads the incoming report and (optionally) converts to Sheet.
* `core_functional_modules.pipeline.run_pipeline`: the orchestrated processing step; returns an object with links and
  timing information used to render the result panel.

This file intentionally includes **documentation-only** additions (module docstring and inline comments)
without modifying the executable logic.
"""

# ------------------------------
# Quick-start for maintainers
# ------------------------------
# 1) Configure Streamlit secrets:
#    - APP_BASE_URL: The exact external base URL of your deployed app (with trailing slash normalization).
#    - GOOGLE_CREDENTIALS: A JSON string containing your OAuth client configuration.
#    - CONFIG_FILE_ID (optional): The Drive file ID for persisted UI defaults.
# 2) Grant your Google Cloud OAuth Client access to the app origin and redirect URI.
# 3) Run: `streamlit run ui_streamlit.py` (ensure the backend `favtrip` package is importable).
# 4) Upload a Modisoft report, adjust recipients and options, then click **Run Pipeline**.


import os
import time
import threading
import json
import base64
import hashlib
import secrets
import re
import queue
import uuid
import traceback
import requests

import streamlit as st
from streamlit.components.v1 import html
from streamlit.components.v1 import html as _html_listener
from streamlit_autorefresh import st_autorefresh

from google_auth_oauthlib.flow import Flow
from google.oauth2.credentials import Credentials

from core_functional_modules.google_client import load_valid_token, services, clear_token
from core_functional_modules.config_store import save_config_to_drive
from core_functional_modules.config import Config
from core_functional_modules.logger import StatusLogger
from core_functional_modules.pipeline import run_pipeline
from core_functional_modules.drive_utils import upload_to_drive, get_or_create_subfolder
from core_functional_modules.pipeline_bus import get_pipeline_queue


# =========================
# Constants & Simple Helpers
# =========================

EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")

err = None

UI_UPLOAD  = "UPLOAD"
UI_READY   = "READY"
UI_RUNNING = "RUNNING"
UI_RESULT = "RESULT"
UI_RESULT_ERROR = "RESULT_ERROR"     # run_pipeline failed
UI_UPLOAD_ERROR = "UPLOAD_ERROR"     # invalid input (1–2 weeks)


PIPE_STATUS_IDLE = "idle"
PIPE_STATUS_RUNNING = "running"
PIPE_STATUS_DONE = "done"
PIPE_STATUS_ERROR = "error"


class UIError(Exception):
    """
    Base class for errors that should show a friendly message
    plus optional technical details.
    """
    user_message: str
    title: str = "Error"

    def __init__(self, user_message: str, *, title: str | None = None):
        super().__init__(user_message)
        self.user_message = user_message
        if title:
            self.title = title


def _split_emails(csv_str: str):
    return [e.strip() for e in (csv_str or "").split(",") if e.strip()]


def _parse_emails(csv_str: str):
    return _split_emails(csv_str)


def _invalid_emails(csv_str: str):
    return [e for e in _parse_emails(csv_str) if not EMAIL_RE.match(e)]


def _analyze_rk_rows(rows):
    """
    Validate the 'Per-Report-Key Recipients' editor rows.
    Returns (issues: List[str], preview_lines: List[str], rk_map: Dict[str, List[str]])
    """
    issues, preview, rk_map = [], [], {}
    seen, dupes = set(), set()

    for idx, r in enumerate(rows or [], start=1):
        raw_key = (r.get("REPORT KEY (ALL CAPS)") or "").strip()
        emails_csv = r.get("Emails (comma)") or ""
        if not raw_key and not emails_csv:
            # allow a blank template row
            continue

        # uppercase flag
        if raw_key != raw_key.upper():
            issues.append(f"Row {idx}: key '{raw_key}' is not ALL CAPS.")

        # duplicate detection
        if raw_key:
            if raw_key in seen:
                dupes.add(raw_key)
            else:
                seen.add(raw_key)

        # email validation
        bads = _invalid_emails(emails_csv)
        if bads:
            issues.append(f"Row {idx}: invalid emails → {', '.join(bads)}")

        # mapping + preview
        if raw_key:
            emails = _parse_emails(emails_csv)
            if emails:
                rk_map[raw_key] = emails
            preview.append(f"{raw_key} → {', '.join(emails) if emails else emails_csv}")

    if dupes:
        issues.append(f"Duplicate keys detected: {', '.join(sorted(dupes))}")
    return issues, preview, rk_map


def _b64url(b: bytes) -> str:
    return base64.urlsafe_b64encode(b).rstrip(b"=").decode("ascii")


def _pkce_pair() -> tuple[str, str]:
    """
    Generate a high-entropy PKCE code_verifier and its S256 code_challenge.
    RFC 7636 requires 43–128 chars; this approach yields a URL-safe value.
    """
    verifier = _b64url(secrets.token_bytes(64))        # ~86 chars, URL-safe, no padding
    challenge = _b64url(hashlib.sha256(verifier.encode("ascii")).digest())
    return verifier, challenge


def _redirect_base() -> str:
    """
    Always return a non-empty redirect base that exactly matches your OAuth client's
    Authorized redirect URI. Prefer Secrets; normalize to one trailing slash.
    """
    base = (st.secrets.get("APP_BASE_URL", "") or "").strip()
    if not base:
        # Fallback to request (often available), still normalized
        try:
            base = (st.request.url_root or "").strip()
        except Exception:
            base = ""
    if not base:
        st.error("OAuth redirect base is not set. Define APP_BASE_URL in Secrets.")
        st.stop()
    return base.rstrip("/") + "/"


def _parse_state(state_b64: str) -> dict:
    # Add padding back for base64 decoding if needed
    padding = "=" * ((4 - len(state_b64) % 4) % 4)
    raw = base64.urlsafe_b64decode(state_b64 + padding)
    return json.loads(raw.decode("utf-8"))


def _infer_media_mime(name: str) -> str:
    n = (name or "").lower()
    if n.endswith(".csv"):
        return "text/csv"
    if n.endswith(".xlsx"):
        return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    return "application/octet-stream"


def _get_drive_service_or_raise(cfg):
    creds = load_valid_token(cfg.SCOPES)
    if not creds:
        raise RuntimeError("Google authorization required. Please sign in first.")
    _sheets, drive, _gmail = services(creds, cfg.HTTP_TIMEOUT_SECONDS)
    return drive


def _rerun():
    try:
        st.rerun()
    except AttributeError:
        st.experimental_rerun()

def reset_to_upload():
    st.session_state.sales_uploaded_ok = False
    st.session_state.vendor_uploaded_ok = False

    st.session_state.sales_selected_name = None
    st.session_state.vendor_selected_name = None

    st.session_state.reset_generation += 1

    st.session_state.sales_selection_generation = None
    st.session_state.vendor_selection_generation = None


    # 🔑 Increment the upload epoch
    st.session_state.upload_epoch += 1
    st.session_state.sales_selected_epoch = None
    st.session_state.vendor_selected_epoch = None

    st.session_state.running_ui_initialized = False
    st.session_state.uploader_version += 1
    st.session_state.ui_phase = UI_UPLOAD

    # Fully clear uploader widget state
    for k in list(st.session_state.keys()):
        if k.startswith("sales_upload_") or k.startswith("vendor_upload_"):
            st.session_state.pop(k, None)



def init_thread_state():
    if "pipeline_thread_started" not in st.session_state:
        st.session_state.pipeline_thread_started = False
    if "pipeline_done" not in st.session_state:
        st.session_state.pipeline_done = False
    if "pipeline_error" not in st.session_state:
        st.session_state.pipeline_error = None
    if "pipeline_thread" not in st.session_state:
        st.session_state.pipeline_thread = None

def init_pipeline_state():
    st.session_state.setdefault("pipe_status", PIPE_STATUS_IDLE)
    st.session_state.setdefault("pipe_result", None)
    st.session_state.setdefault("pipe_finished", False)
    st.session_state.setdefault("pipe_error", None)
    st.session_state.setdefault("pipe_run_id", None)

def reset_pipeline_state():
    # Thread control
    st.session_state.pipeline_thread_started = False
    st.session_state.pipeline_done = False
    st.session_state.pipeline_error = None
    st.session_state.pipeline_thread = None

    # Pipeline result & lifecycle
    st.session_state.pipe_status = PIPE_STATUS_IDLE
    st.session_state.pipe_finished = False
    st.session_state.pipe_result = None
    st.session_state.pipe_error = None
    st.session_state.pipe_run_id = None

    # Timer
    st.session_state._run_start_time = None


def start_run():
    st.session_state.pipe_run_id = str(uuid.uuid4())

    st.session_state.pipe_status = PIPE_STATUS_RUNNING
    st.session_state.pipeline_thread_started = True
    st.session_state.pipeline_done = False
    st.session_state.pipeline_error = None


    st.session_state.pipeline_refresh_key = f"pipeline_refresh_{time.time()}"



def _both_uploads_ok():
    epoch = st.session_state.upload_epoch

    return (
        st.session_state.sales_selected_epoch == epoch
        and st.session_state.vendor_selected_epoch == epoch
    )



def _validate_pipeline_result(result):
    required_attrs = ("location", "timestamp", "elapsed_seconds")
    return (
        result is not None
        and all(hasattr(result, attr) for attr in required_attrs)
    )

def apply_per_run_config(
    *,
    cfg,
    to,
    cc,
    error_recipients,
    use_all,
    report_keys,
    include_full,
    send_full,
    email_mgr,
    calc_id,
    incoming_id,
    mgr_folder,
    order_folder,
    error_folder,
    user_folder,
    redirect_port,
    gid_mgr,
    gid_order,
    gid_err,
    gid_bev_err,
    loc_sheet,
    loc_range,
    update_range,
    tz,
    tfmt,
    output_ttl,
    failed_input_ttl,
    user_ttl,
    use_rollover,
    start_dow,
    end_dow,
    soft_cases_enabled,
    soft_cases_threshold,
    BEV_MAPPING_LINK,
    edited_rows,
):
    # ------------------------------------------------------------
    # Basic email & behavior flags
    # ------------------------------------------------------------
    cfg.TO_RECIPIENTS = _split_emails(to)
    cfg.CC_RECIPIENTS = _split_emails(cc)
    cfg.ERROR_RECIPIENTS = _split_emails(error_recipients)

    cfg.USE_ALL_REPORT_KEYS = use_all
    cfg.REPORT_KEY_RUN_LIST = [
        s.strip().upper()
        for s in (report_keys or "").split(",")
        if s.strip()
    ]

    cfg.INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL = bool(include_full)
    cfg.SEND_SEPARATE_FULL_ORDER_EMAIL = bool(send_full)
    cfg.EMAIL_MANAGER_REPORT = bool(email_mgr)

    # ------------------------------------------------------------
    # IDs, folders, sheets
    # ------------------------------------------------------------
    cfg.CALC_SPREADSHEET_ID = calc_id
    cfg.INCOMING_FOLDER_ID = incoming_id
    cfg.MANAGER_REPORT_FOLDER_ID = mgr_folder
    cfg.ORDER_REPORT_FOLDER_ID = order_folder
    cfg.ERROR_REPORT_FOLDER_ID = error_folder
    cfg.USER_FOLDER_ID = user_folder
    cfg.REDIRECT_PORT = int(redirect_port)

    cfg.GID_MANAGER_PDF = gid_mgr
    cfg.GID_ORDER_CSV = gid_order
    cfg.GID_ERROR_REPORT = gid_err
    cfg.GID_BEV_ERRORS = gid_bev_err

    cfg.LOCATION_SHEET_TITLE = loc_sheet
    cfg.LOCATION_NAMED_RANGE = loc_range
    cfg.TEMPLATE_UPDATE_RANGE = update_range

    cfg.TIMESTAMP_TZ = tz
    cfg.TIMESTAMP_FMT = tfmt

    # ------------------------------------------------------------
    # Lifecycle / TTL
    # ------------------------------------------------------------
    cfg.OUTPUT_TIME_TO_LIFE = int(output_ttl)
    cfg.FAILED_INPUT_TIME_TO_LIFE = int(failed_input_ttl)
    cfg.USER_TIME_TO_LIFE = int(user_ttl)

    # ------------------------------------------------------------
    # Date & integrity controls
    # ------------------------------------------------------------
    cfg.USE_AUTO_ROLLOVER_IF_ONE_WEEK = bool(use_rollover)
    cfg.START_DAY_OF_WEEK = start_dow
    cfg.END_DAY_OF_WEEK = end_dow

    # ------------------------------------------------------------
    # Soft cases alerting
    # ------------------------------------------------------------
    cfg.SOFT_CASES_ALERT_ENABLED = bool(soft_cases_enabled)
    cfg.SOFT_CASES_ALERT_THRESHOLD = int(soft_cases_threshold)

    # ------------------------------------------------------------
    # External links
    # ------------------------------------------------------------
    cfg.BEV_MAPPING_LINK = BEV_MAPPING_LINK

    # ------------------------------------------------------------
    # Per‑report‑key recipients
    # ------------------------------------------------------------
    rk_map: dict[tuple[str | None, str | None, str | None], list[str]] = {}

    for r in edited_rows or []:
        store = (r.get("Store (optional)") or "").strip().upper() or None
        key = (r.get("Report Key (optional)") or "").strip().upper() or None
        sub_key = (r.get("Sub-Report Key (optional)") or "").strip().upper() or None

        emails = [
            e.strip()
            for e in (r.get("Emails (comma)") or "").split(",")
            if e.strip()
        ]

        if not emails or not (store or key or sub_key):
            continue

        rk_map[
            (
                clean_tag(store) if store else None,
                clean_tag(key) if key else None,
                clean_tag(sub_key) if sub_key else None,
            )
        ] = emails

    cfg.REPORT_KEY_RECIPIENTS = rk_map

    return rk_map


def github_merge(from_branch: str, to_branch: str) -> bool:
    """
    Merge `from_branch` into `to_branch` using GitHub's API.

    Returns True on success, False on failure.
    All secrets are fetched internally with safe failure.
    """

    try:
        token = (st.secrets.get("GITHUB_TOKEN", "") or "").strip()
        owner = (st.secrets.get("GITHUB_OWNER", "") or "").strip()
        repo = (st.secrets.get("GITHUB_REPO", "") or "").strip()

        missing = []
        if not token:
            missing.append("GITHUB_TOKEN")
        if not owner:
            missing.append("GITHUB_OWNER")
        if not repo:
            missing.append("GITHUB_REPO")

        if missing:
            st.error(f"Missing GitHub secrets: {', '.join(missing)}")
            return False

        url = f"https://api.github.com/repos/{owner}/{repo}/merges"
        headers = {
            "Authorization": f"Bearer {token}",
            "Accept": "application/vnd.github+json",
        }
        payload = {
            "base": to_branch,
            "head": from_branch,
            "commit_message": f"Merge {from_branch} → {to_branch} (via Streamlit)",
        }

        resp = requests.post(url, headers=headers, json=payload, timeout=20)

        if resp.status_code == 201:
            return True

        if resp.status_code == 409:
            st.error("❌ Merge conflict detected. Resolve manually.")
            return False

        st.error(f"GitHub merge failed ({resp.status_code}): {resp.text}")
        return False

    except Exception as e:
        st.error(f"Unexpected merge error: {e}")
        return False


# =========================
# OAuth (Web / PKCE)
# =========================

def start_web_oauth(scopes):
    """
    Build an authorization URL that:
      - uses a stable redirect_uri (from Secrets)
      - uses explicit PKCE (S256)
      - embeds the code_verifier inside the state (base64url(JSON))
    """
    cfg = json.loads(st.secrets["GOOGLE_CREDENTIALS"])
    redirect = _redirect_base()

    # Explicit PKCE (stateless across redirect)
    code_verifier, code_challenge = _pkce_pair()

    # CSRF token + verifier encoded into state that Google will return unchanged.
    state_obj = {
        "csrf": _b64url(secrets.token_bytes(16)),
        "v": code_verifier,
        "r": redirect,
    }
    state_b64 = _b64url(json.dumps(state_obj).encode("utf-8"))

    flow = Flow.from_client_config(cfg, scopes=scopes, redirect_uri=redirect)
    auth_url, _ = flow.authorization_url(
        access_type="offline",
        include_granted_scopes="true",
        prompt="consent",
        state=state_b64,
        code_challenge=code_challenge,
        code_challenge_method="S256",
    )

    # Keep only minimal context; state carries the verifier.
    st.session_state["_oauth_redirect"] = redirect
    return auth_url


def finish_web_oauth(code: str, state_b64: str, scopes):
    """
    Recreate a Flow with the same redirect_uri and exchange code + code_verifier for tokens.
    (No UI side effects here.)
    """
    cfg = json.loads(st.secrets["GOOGLE_CREDENTIALS"])
    state_obj = _parse_state(state_b64)
    code_verifier = state_obj.get("v")
    redirect = state_obj.get("r") or st.session_state.get("_oauth_redirect") or _redirect_base()

    if not code_verifier:
        st.error("OAuth state did not include a PKCE code_verifier.")
        st.stop()

    flow = Flow.from_client_config(cfg, scopes=scopes, redirect_uri=redirect)
    flow.fetch_token(code=code, code_verifier=code_verifier)

    creds = flow.credentials
    with open("token.json", "w") as f:
        f.write(creds.to_json())
    return creds

def clean_tag(s: str) -> str:
    import re
    s = (s or "").strip()
    s = re.sub(r"[^A-Za-z0-9._-]+", "-", s)
    return s.strip("-") or "UNKNOWN"

# --- OAuth Redirect Handler (Web/PKCE only) ---


# =========================
# UI Sections
# =========================

def render_upload_card(cfg):
    with st.container(border=True):
        st.subheader("Upload Required Input Files")
        st.caption("Both Sales Data and Vendor Price Data are required. Sales Data must be 1 or 2 complete weeks.")

        up_col, _, upbtn_col = st.columns([4, 1, 1])

        current_gen = st.session_state.reset_generation

        with up_col:
            sales_key = f"sales_upload_v{st.session_state.uploader_version}"
            sales_file = st.file_uploader(
                "Upload Sales Data",
                type=["xlsx", "csv"],
                key=sales_key,
                help="Go to Modisoft -> Sales -> Live Items, Select Stores & Dates, Download as Excel"
            )
            
            if sales_file:
                if st.session_state.sales_selection_generation != current_gen:
                    st.session_state.sales_selected_name = sales_file.name
                    st.session_state.sales_selected_epoch = st.session_state.upload_epoch
                    st.session_state.sales_selection_generation = current_gen


            vendor_key = f"vendor_upload_v{st.session_state.uploader_version}"
            vendor_file = st.file_uploader(
                "Upload Vendor Price Data",
                type=["xlsx", "csv"],
                key=vendor_key,
                help="Go to Modisoft -> Products -> Price Book , Download as Excel"
            )
            
            if vendor_file:
                if st.session_state.vendor_selection_generation != current_gen:
                    st.session_state.vendor_selected_name = vendor_file.name
                    st.session_state.vendor_selected_epoch = st.session_state.upload_epoch
                    st.session_state.vendor_selection_generation = current_gen

        
        with upbtn_col:
            st.markdown('<div class="ft-right-btn">', unsafe_allow_html=True)

            upload_clicked = st.button(
                "⬆️ Upload Now",
                width="stretch",
                type="primary",
                disabled= (not _both_uploads_ok()),
                key="upload_submit",
            )

            st.markdown('</div>', unsafe_allow_html=True)


        # --- Handle the upload action immediately ---
        if upload_clicked:
            if not cfg.INCOMING_FOLDER_ID:
                st.session_state.upload_error = "Incoming Folder ID is empty."
                st.session_state.ui_phase = UI_UPLOAD_ERROR
                _rerun()

            if sales_file is None or vendor_file is None:
                st.session_state.upload_error = (
                    "Both Sales Data and Vendor Price Data are required."
                )
                st.session_state.ui_phase = UI_UPLOAD_ERROR
                _rerun()

            try:
                st.warning("Uploading files to google drive...")
                drive = _get_drive_service_or_raise(cfg)
                

                # --- Resolve user ---
                me = drive.about().get(
                    fields="user(emailAddress,permissionId,displayName)"
                ).execute().get("user", {})

                user_email = (me or {}).get("emailAddress") or "UNKNOWN_USER"

                # --- Per-user folder ---
                user_folder = get_or_create_subfolder(
                    drive,
                    cfg.INCOMING_FOLDER_ID,
                    user_email,
                )

                sales_folder = get_or_create_subfolder(
                    drive,
                    user_folder["id"],
                    "01 Sales Data Inputs",
                )

                vendor_folder = get_or_create_subfolder(
                    drive,
                    user_folder["id"],
                    "02 Vendor Price Data Inputs",
                )

                # --- Upload SALES ---
                sales_created = upload_to_drive(
                    drive,
                    data=sales_file.getvalue(),
                    name=f"{os.path.splitext(sales_file.name)[0]} (Sales Data via UI)",
                    mime=_infer_media_mime(sales_file.name),
                    folder_id=sales_folder["id"],
                    to_sheet=True,
                )

                # --- Upload VENDOR ---
                vendor_created = upload_to_drive(
                    drive,
                    data=vendor_file.getvalue(),
                    name=f"{os.path.splitext(vendor_file.name)[0]} (Vendor Price Data via UI)",
                    mime=_infer_media_mime(vendor_file.name),
                    folder_id=vendor_folder["id"],
                    to_sheet=True,
                )

                st.session_state.ui_phase = UI_READY
                _rerun()

            except Exception as e:
                st.session_state.upload_error = f"Upload failed: {e}"
                st.session_state.ui_phase = UI_UPLOAD_ERROR
                _rerun()
    
def render_run_options(cfg):
    run_form_wrapper_classes = "ft-card ft-row"

    # A file is "dirty" if the user has selected something not yet uploaded
    files_dirty = (
        st.session_state.sales_selected_name is not None
        or st.session_state.vendor_selected_name is not None
    )

    # Have we successfully uploaded both files?

    # OPEN the wrapper with real HTML (no entities)
    st.markdown(f'<div class="{run_form_wrapper_classes}">', unsafe_allow_html=True)



    with st.form("run_form"):
        # Header row uses the same columns to align the Run button with Upload button above
        tl, _, col_run = st.columns([4, 1, 1])
        with tl:
            st.subheader("Run Options")
            st.caption("Configure email behavior and report keys. Use **Advanced** for IDs/GIDs/timezone.")

        # --- Unified gating logic ---
        # A) If a file is currently selected but not uploaded -> disable Run
        # B) If no file selected and we have a prior successful upload -> enable Run
        # C) Otherwise (no prior upload or ambiguous state) -> disable Run

        with col_run:
            # Right-align and full-width, matching Upload Now
            st.markdown('<div class="ft-right-btn">', unsafe_allow_html=True)
            submitted = st.form_submit_button(
                "▶️ Run Pipeline",
                width='stretch',
                disabled=False,
                type="primary",
                key="run_submit"
            )
            st.markdown('</div>', unsafe_allow_html=True)

        # ----- Main options -----

        # Recipients
        st.markdown("##### Recipients")
        col1, col2 = st.columns([1, 1])
        with col1:
            to = st.text_input(
                "To (comma)", value=",".join(cfg.TO_RECIPIENTS or []),
                help="Fallback recipients for Manager & Order emails."
            )
        with col2:
            cc = st.text_input(
                "CC (comma)", value=",".join(cfg.CC_RECIPIENTS or []),
                help="Optional CC added to all emails."
            )

        # Report Keys
        st.markdown("##### Report Keys")
        colk1, colk2, colk3 = st.columns([1, 1, 2])
        with colk1:
            use_all = st.toggle(
                "Use all keys from CSV",
                value=cfg.USE_ALL_REPORT_KEYS,
                help="ON: process every key found. OFF: only the keys you list."
            )

        with colk2:
            pass

        with colk3:
            report_keys = st.text_input(
                "Keys to run (comma)",
                value=",".join(cfg.REPORT_KEY_RUN_LIST or []),
                help="Used when 'Use all keys' is OFF. For Sub_Report_Keys use Report_Key-Sub_Report_Key. Example: COFFEE,GROCERY,BEV-7UP"
            )

        # General Behavior
        st.markdown("##### General Behavior")
        cole1, cole2, cole3, cole4 = st.columns([1, 1, 1, 1])
        with cole1:
            include_full = st.toggle(
                "Attach FULL order in each email",
                value=cfg.INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL
            )
        with cole2:
            send_full = st.toggle(
                "Send separate FULL order email",
                value=cfg.SEND_SEPARATE_FULL_ORDER_EMAIL
            )
        with cole3:
            email_mgr = st.toggle(
                "Email Manager Report",
                value=getattr(cfg, "EMAIL_MANAGER_REPORT", True),
                help="When ON, the Manager Report email is sent. When OFF, it is skipped."
            )
        with cole4:
            use_rollover = st.toggle(
                    'Use auto-rollover for single week uploads',
                    value=cfg.USE_AUTO_ROLLOVER_IF_ONE_WEEK,
                    help='If this is on, when only 1 week is uploaded, the most recent previously uploaded data will become the "Last Week" data; If this is off then the "Last Week" data will be left blank'
                )

        # Per-Report-Key Recipients
        with st.expander("Per-Report-Key Recipients (optional)", expanded=False):

            st.caption("""
              Map **Store, Report Key → Emails (comma)**.
              
              **Email Delivery Priority:**  
              - `(Store, Key)` → First priority set of emails  
              - `(, Key)` → Second priority set of emails  
              - `(Store,)` → Third priority set of emails
              - If not defined, it will use the default set of emails in `To (comma)` field above
              """)
        
            rows = []
        
            if cfg.REPORT_KEY_RECIPIENTS:
                for (store, key, sub_key), emails in cfg.REPORT_KEY_RECIPIENTS.items():
                    rows.append({
                        "Store (optional)": store or "",
                        "Report Key (optional)": key or "",
                        "Sub-Report Key (optional)": sub_key or "",
                        "Emails (comma)": ",".join(emails or [])
                    })
            else:
                rows = [{
                    "Store (optional)": "",
                    "Report Key (optional)": "",
                    "Sub-Report Key (optional)": "",
                    "Emails (comma)": ""
                }]
        
            edited_rows = st.data_editor(
                rows,
                num_rows="dynamic",
                width='stretch',
                key="rk_editor",
            )
        
            rk_map = {}
            rk_preview = []
            rk_issues = []
        
            for i, r in enumerate(edited_rows):
        
                store = (r.get("Store (optional)") or "").strip().upper()
                key = (r.get("Report Key (optional)") or "").strip().upper()
                sub_key = (r.get("Sub-Report Key (optional)") or "").strip().upper() or None
                emails_raw = (r.get("Emails (comma)") or "").strip()
        
                emails = [e.strip() for e in emails_raw.split(",") if e.strip()]
        
                store_val = store if store else None
                key_val = key if key else None
                sub_val = sub_key if sub_key else None
        
                if emails and not (store_val or key_val or sub_key):
                    rk_issues.append(f"Row {i+1}: Must include Store, Key, or both.")
                    continue
        
                if (store_val or key_val or sub_key) and not emails:
                    rk_issues.append(f"Row {i+1}: Missing email(s).")
                    continue

                store_tag = clean_tag(store_val)
                key_tag = clean_tag(key_val)
                sub_tag = clean_tag(sub_val)
        
                rk_map[(store_tag, key_tag, sub_tag)] = emails
        
                rk_preview.append(f"{(store_val, key_val, sub_val)} -> {emails}")
        
            #if rk_preview:
            #    with st.expander("Recipient mapping preview"):
            #        st.code("\n".join(rk_preview), language="text")
        
            #if rk_issues:
            #    st.warning("Recipient configuration issues:\n\n- " + "\n- ".join(rk_issues))

        # Advanced
        with st.expander("Advanced", expanded=False):
            tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(["Folders", "Files & Links", "Ranges", "Timing", "Lifecycle", "Technical"])
            with tab1:
                ga1, ga2 = st.columns([1, 1])
                with ga1:
                    st.markdown("###### Input Folders")
                    user_folder = st.text_input("User Calculations Folder ID", value=cfg.USER_FOLDER_ID,
                        help="The file ID of the google drive folder that user workhorse files should be stored.")
                    incoming_id = st.text_input("Incoming Folder ID", value=cfg.INCOMING_FOLDER_ID,
                        help="The file ID of the google drive folder that user input folders & files should be stored.")
                    
                with ga2:
                    st.markdown("###### Output Folders")
                    order_folder = st.text_input("Order Report Folder ID", value=cfg.ORDER_REPORT_FOLDER_ID,
                        help="The file ID of the google drive folder that order report csv output files should be stored.")
                    mgr_folder = st.text_input("Manager Report Folder ID", value=cfg.MANAGER_REPORT_FOLDER_ID,
                        help="The file ID of the google drive folder that manager report pdf output files should be stored.")
                    error_folder = st.text_input("Error Report Folder ID", value=cfg.ERROR_REPORT_FOLDER_ID,
                        help="The file ID of the google drive folder that error report csv output files should be stored.")

            with tab2:
                gb1, gb2 = st.columns([1, 1])
                with gb1:
                    calc_id = st.text_input("Master Calculations Spreadsheet ID", value=cfg.CALC_SPREADSHEET_ID,
                        help="The file ID of the Master Calculations google sheets file that user workhorse files should be based off of.")
                with gb2:
                    BEV_MAPPING_LINK = st.text_input("BEV Mapping Link", value=cfg.BEV_MAPPING_LINK,
                        help="The url to the live, editable BEV Sub-Key Mapping google sheets file.")
            
            with tab3:
                gc1, gc2 = st.columns([1, 1])
                with gc1:
                    st.markdown("###### GIDs")
                    gid_mgr = st.text_input("Manager Report gid", value=str(cfg.GID_MANAGER_PDF),
                        help="The GID of the Manager Report Tab within the Master Calculations Sheet that should be used for outputs.")
                    gid_err = st.text_input("Error Report gid", value=str(cfg.GID_ERROR_REPORT),
                        help="The GID of the Error Report Tab within the Master Calculations Sheet that should be used for outputs.")
                    gid_order = st.text_input("Order Report gid", value=str(cfg.GID_ORDER_CSV),
                        help="The GID of the Order Report Tab within the Master Calculations Sheet that should be used for outputs.")
                    gid_bev_err = st.text_input("Unassigned Beverages Report gid", value=str(cfg.GID_ORDER_CSV),
                        help="The GID of the UB Report Tab within the Master Calculations Sheet that should be used for outputs.")
                with gc2:
                    st.markdown("###### Titles")
                    loc_sheet = st.text_input("Named Range Sheet Title", value=cfg.LOCATION_SHEET_TITLE,
                        help="The Sheet Title of the tab within the Master Calculations Sheet where the below named ranges exist.")
                    loc_range = st.text_input("Location Named Range", value=cfg.LOCATION_NAMED_RANGE,
                        help="The named range within the Master Calculations Sheet that refrences the cleaned location(s) name.")
                    update_range = st.text_input("Update Timestamp Range", value=cfg.TEMPLATE_UPDATE_RANGE,
                        help="The named range within the Master Calculations Sheet that refrences the last time the template was updated.")

            with tab4:
                gd1, gd2 = st.columns([1, 1])
                with gd1:
                    st.markdown("###### Data Integrity Controls")
                    _days = ["Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Any"]
                    start_dow = st.selectbox(
                        "Start day of week", _days, index=_days.index(cfg.START_DAY_OF_WEEK),
                        help="The day of week that the uploaded data should start at, any other day will raise an error."
                    )
                    end_dow = st.selectbox(
                        "End day of week", _days, index=_days.index(cfg.END_DAY_OF_WEEK),
                        help="The day of week that the uploaded data should end at, any other day will raise an error."
                    )
                with gd2:
                    st.markdown("###### Formatting")                                
                    tz = st.text_input("Timestamp Timezone", value=cfg.TIMESTAMP_TZ,
                        help="The timezone that should be used in all timestamps.")
                    tfmt = st.text_input("Timestamp Format", value=cfg.TIMESTAMP_FMT,
                        help="The format that should be used in all timestamps.")
            with tab5:
                ge1, ge2 = st.columns([1, 1])
                with ge1:
                    st.markdown("###### One-Time Use Files")
                    failed_input_ttl = st.number_input(
                        "Failed Input Time-To-Life (days)",
                        min_value=0,
                        max_value=3650,
                        value=int(cfg.FAILED_INPUT_TIME_TO_LIFE),
                        help="Delete old unused incoming files older than this many days after a successful run."
                        )
                    output_ttl = st.number_input(
                        "Output Time-To-Life (days)",
                        min_value=0,
                        max_value=3650,
                        value=int(cfg.OUTPUT_TIME_TO_LIFE),
                        help="Delete output files older than this many days after a successful run."
                        )
                with ge2:
                    st.markdown("###### Recurring Use Files")
                    user_ttl = st.number_input(
                        "User Calculations Time-To-Life (days)",
                        min_value=0,
                        max_value=3650,
                        value=int(cfg.USER_TIME_TO_LIFE),
                        help="Delete old unused user calculations files older than this many days after a successful run."
                        )
            
            with tab6:

                #row1
                gf1_1, gf2_1 = st.columns([1, 1])
                with gf1_1:                
                    soft_cases_enabled = st.toggle(
                        "Alert on large case quantities",
                        value=cfg.SOFT_CASES_ALERT_ENABLED,
                        help="Send a technical alert if any FULL order rows exceed the cases threshold"
                    )

                with gf2_1:
                    error_recipients = st.text_input(
                        "Technical Support Email(s) (comma)",
                        value=",".join(cfg.ERROR_RECIPIENTS or []),
                        help="If errors arise such as missing items in the Vendor Price Book, the error report will be sent here."
                    )

                #row2
                gf1_2, gf2_2 = st.columns([1, 1])
                with gf1_2:                
                    soft_cases_threshold = st.number_input(
                        "Cases-to-order alert threshold",
                        min_value=1,
                        max_value=1000,
                        value=int(cfg.SOFT_CASES_ALERT_THRESHOLD),
                        help="Any FULL order line above this number will trigger a soft alert"
                    )

                with gf2_2:
                    raw_redirect_port = int(cfg.REDIRECT_PORT) if str(cfg.REDIRECT_PORT).isdigit() else 0
                    redirect_port = st.number_input(
                        "Redirect Port (0 = auto)",
                        min_value=0, max_value=65535,
                        value=raw_redirect_port if raw_redirect_port in (0, *range(1024, 65536)) else 0,
                        help="Use 0 to auto-pick a free port. Otherwise choose 1024–65535."
                    )

                
        save_defaults_clicked = st.form_submit_button("💾 Save as defaults", type="secondary", help="Persist current settings for future sessions")


        # ----- Submission handling -----

        if save_defaults_clicked:
            try:
                rk_map = apply_per_run_config(
                    cfg=cfg,
                    to=to,
                    cc=cc,
                    error_recipients=error_recipients,
                    use_all=use_all,
                    report_keys=report_keys,
                    include_full=include_full,
                    send_full=send_full,
                    email_mgr=email_mgr,
                    calc_id=calc_id,
                    incoming_id=incoming_id,
                    mgr_folder=mgr_folder,
                    order_folder=order_folder,
                    error_folder=error_folder,
                    user_folder=user_folder,
                    redirect_port=redirect_port,
                    gid_mgr=gid_mgr,
                    gid_order=gid_order,
                    gid_err=gid_err,
                    gid_bev_err=gid_bev_err,
                    loc_sheet=loc_sheet,
                    loc_range=loc_range,
                    update_range=update_range,
                    tz=tz,
                    tfmt=tfmt,
                    output_ttl=output_ttl,
                    failed_input_ttl=failed_input_ttl,
                    user_ttl=user_ttl,
                    use_rollover=use_rollover,
                    start_dow=start_dow,
                    end_dow=end_dow,
                    soft_cases_enabled=soft_cases_enabled,
                    soft_cases_threshold=soft_cases_threshold,
                    BEV_MAPPING_LINK=BEV_MAPPING_LINK,
                    edited_rows=edited_rows,
                )

                # Ensure we have a user token first
                creds = load_valid_token(cfg.SCOPES)
                if not creds:
                    st.error("Not authenticated. Please complete Google sign‑in first (top of page).")
                else:

                    # Drive service
                    _sheets, drive, _gmail = services(creds, cfg.HTTP_TIMEOUT_SECONDS)

                    # What we persist
                    drive_defaults = cfg.to_drive_defaults()

                    DEV_ENVIRONMENT = st.secrets.get("DEV_ENVIRONMENT", False)
                    DEV_CONFIG_FILE_ID = (st.secrets.get("DEV_CONFIG_FILE_ID", "") or "").strip()
                    CONFIG_FILE_ID = (st.secrets.get("CONFIG_FILE_ID", "") or "").strip()

                    # Decide where to SAVE
                    if DEV_ENVIRONMENT:
                        save_target_id = DEV_CONFIG_FILE_ID or None
                    else:
                        save_target_id = CONFIG_FILE_ID or None

                    new_id = save_config_to_drive(
                        drive,
                        drive_defaults,
                        file_id=save_target_id
                    )

                    # DEV auto-bootstrap case
                    if DEV_ENVIRONMENT and not DEV_CONFIG_FILE_ID:
                        st.success("✅ Created new DEV config file.")
                        st.info(
                            "Add this to Streamlit secrets as DEV_CONFIG_FILE_ID:\n\n"
                            f"`{new_id}`"
                        )
                    else:
                        st.success(f"✅ Defaults saved (file id: {new_id})")

            except Exception as e:
                st.error(f"Failed to save defaults to Drive: {e}")

        if submitted:
            # Apply per-run config
            rk_map = apply_per_run_config(
                cfg=cfg,
                to=to,
                cc=cc,
                error_recipients=error_recipients,
                use_all=use_all,
                report_keys=report_keys,
                include_full=include_full,
                send_full=send_full,
                email_mgr=email_mgr,
                calc_id=calc_id,
                incoming_id=incoming_id,
                mgr_folder=mgr_folder,
                order_folder=order_folder,
                error_folder=error_folder,
                user_folder=user_folder,
                redirect_port=redirect_port,
                gid_mgr=gid_mgr,
                gid_order=gid_order,
                gid_err=gid_err,
                gid_bev_err=gid_bev_err,
                loc_sheet=loc_sheet,
                loc_range=loc_range,
                update_range=update_range,
                tz=tz,
                tfmt=tfmt,
                output_ttl=output_ttl,
                failed_input_ttl=failed_input_ttl,
                user_ttl=user_ttl,
                use_rollover=use_rollover,
                start_dow=start_dow,
                end_dow=end_dow,
                soft_cases_enabled=soft_cases_enabled,
                soft_cases_threshold=soft_cases_threshold,
                BEV_MAPPING_LINK=BEV_MAPPING_LINK,
                edited_rows=edited_rows,
            )

            # --- ADD: warnings before kicking off the run ---
            if not cfg.USE_ALL_REPORT_KEYS and not cfg.REPORT_KEY_RUN_LIST:
                st.session_state.upload_error = (
                    "No report keys selected. Either enable 'Use all keys' "
                    "or provide explicit report keys."
                )
                st.session_state.ui_phase = UI_UPLOAD_ERROR
                _rerun()

            if not cfg.TO_RECIPIENTS and not cfg.DEFAULT_TO_RECIPIENTS and not rk_map:
                st.session_state.run_error = (
                    "No email recipients defined. At least one recipient is required."
                )
                st.session_state.ui_phase = UI_RESULT_ERROR
                _rerun()

            if rk_issues:
                st.session_state.run_error = (
                    "Invalid per‑report‑key recipient configuration:\n\n"
                    + "\n".join(rk_issues)
                )
                st.session_state.ui_phase = UI_RESULT_ERROR
                _rerun()

            
            # All validation must already be done
            st.session_state._run_start_time = None
            st.session_state.ui_phase = UI_RUNNING

            # Start pipeline in background (one time)
            if not st.session_state.pipeline_thread_started:
                start_run()

                t = threading.Thread(
                    target=run_pipeline_controller,
                    args=(cfg, st.session_state.pipe_run_id),
                    daemon=True
                )
                st.session_state.pipeline_thread = t
                t.start()

            _rerun()
            # --- END ADD ---

def run_pipeline_controller(cfg, run_id):

    logger = StatusLogger(
        print_to_console=True,
        file_path="last_run.log",
        overwrite=True,
    )

    try:
        result = run_pipeline(cfg, logger=logger)
        get_pipeline_queue().put(
            (run_id, PIPE_STATUS_DONE, result)
        )
    except Exception as e:
        get_pipeline_queue().put(
            (
                run_id,
                PIPE_STATUS_ERROR,
                {
                    "type": type(e).__name__,
                    "user_message": str(e),
                    "traceback": traceback.format_exc(),
                },
            )
        )

    finally:
        logger.close()


def render_running_status(cfg):
    import time
    import os
    import queue
    import streamlit as st
    from streamlit_autorefresh import st_autorefresh

    # ------------------------------------------------------------
    # Poll queue FIRST (edge-triggered)
    # ------------------------------------------------------------
    
    q = get_pipeline_queue()
    
    while True:
        try:
            run_id, status, payload = q.get_nowait()
        except queue.Empty:
            break

        if run_id != st.session_state.get("pipe_run_id"):
            continue

        if status == PIPE_STATUS_DONE:
            st.session_state.pipe_result = payload
            st.session_state.pipe_status = PIPE_STATUS_DONE
            st.session_state.pipe_finished = True

        elif status == PIPE_STATUS_ERROR:
            st.session_state.pipe_error = payload
            st.session_state.run_error = payload
            st.session_state.pipe_status = PIPE_STATUS_ERROR
            st.session_state.pipe_finished = True
    
    if st.session_state.pipe_finished:
        if st.session_state.pipe_status == PIPE_STATUS_DONE:
            st.session_state.ui_phase = UI_RESULT
        elif st.session_state.pipe_status == PIPE_STATUS_ERROR:
            st.session_state.ui_phase = UI_RESULT_ERROR
        _rerun()

    # ------------------------------------------------------------
    # ALWAYS render something
    # ------------------------------------------------------------
    with st.status("Running pipeline…", expanded=True):

        # ----- Bulletproof timer -----
        start_time = st.session_state.get("_run_start_time")
        if not isinstance(start_time, (int, float)):
            start_time = time.perf_counter()
            st.session_state._run_start_time = start_time

        elapsed = int(time.perf_counter() - start_time)
        h, m, s = elapsed // 3600, (elapsed % 3600) // 60, elapsed % 60
        st.markdown(f"**Elapsed:** `{h:02d}:{m:02d}:{s:02d}`")

        # ----- Log tail -----
        if os.path.exists("last_run.log"):
            try:
                with open("last_run.log", "r", encoding="utf-8") as f:
                    st.code("".join(f.readlines()[-8:]), language="text")
            except Exception:
                st.markdown("*Waiting for logs…*")
        else:
            st.markdown("*Waiting for logs…*")

    # ------------------------------------------------------------
    # LIVE status check (DO NOT CACHE THIS)
    # ------------------------------------------------------------
    status = st.session_state.get("pipe_status")

    # ✅ This is what keeps the UI alive
    
    if (
        st.session_state.pipe_status == PIPE_STATUS_RUNNING
        and not st.session_state.pipe_finished
    ):
        st_autorefresh(
            interval=1000,
            key=f"pipeline_tick_{st.session_state.pipe_run_id}",
        )



def render_results(cfg):
    result = st.session_state.get("pipe_result")

    
    if not _validate_pipeline_result(result):
        st.error(
            "Run completed, but the pipeline did not return a valid result object."
        )
        if os.path.exists("last_run.log"):
            with open("last_run.log", "rb") as f:
                st.download_button(
                    "⬇️ Download log",
                    f.read(),
                    file_name="last_run.log",
                    mime="text/plain",
                )


    with st.container(border=True):
        st.subheader("✅ Run Complete")

        st.write("### Outputs")
        col1, col2, col3 = st.columns(3)

        col1.metric("Location", result.location)
        col2.metric("Timestamp", result.timestamp)
        col3.metric(
            "Elapsed",
            f"{result.elapsed_seconds//3600:02d}:"
            f"{(result.elapsed_seconds%3600)//60:02d}:"
            f"{result.elapsed_seconds%60:02d}"
        )

        if getattr(result, "manager_pdf_link", None):
            st.success(f"Manager PDF: {result.manager_pdf_link}")
        if getattr(result, "full_order_link", None):
            st.success(f"Full Order Sheet: {result.full_order_link}")

    
        if os.path.exists("last_run.log"):
            with open("last_run.log", "rb") as f:
                st.session_state["last_run_log"] = f.read()
                st.session_state["last_run_timestamp"] = result.timestamp


        if "last_run_log" in st.session_state:
            st.download_button(
                "⬇️ Download full log (last_run.log)",
                st.session_state["last_run_log"],
                file_name=f"last_run_{st.session_state['last_run_timestamp']}.log",
                mime="text/plain",
                width='stretch'
                )


def render_sidebar(cfg):
    with st.sidebar:
        st.header("Utilities")

        # --- Existing buttons ---
        if st.button("Google Sign Out", type="secondary", width='stretch'):
            clear_token()
            for key in ["auth_required", "oauth_flow", "oauth_url", "auth_checked"]:
                if key in st.session_state:
                    del st.session_state[key]
            _rerun()

        st.link_button("Add Users to App", "https://console.cloud.google.com/auth/audience?project=favtripdev", width='stretch')
        st.link_button("Open Google Drive", "https://drive.google.com/drive/u/6/folders/1fhzbq0r8iugIJb9t-EQOdHGNvlr9gLT5", width='stretch')
        st.link_button("Open Modisoft", "https://insights.modisoft.com/account/logon", width='stretch')
        st.link_button("Open Bev Mapping File", cfg.BEV_MAPPING_LINK, width='stretch')

        if False:
            st.checkbox(
                "Offer full log download",
                key="offer_log_download",
                help="If enabled, a 'Download last_run.log' button appears when a run finishes."
            )


        # =============================================================
        # DEV-ONLY: Push DEV Defaults → PROD Defaults
        # =============================================================
        DEV_ENVIRONMENT = bool(st.secrets.get("DEV_ENVIRONMENT", False))

        if DEV_ENVIRONMENT:
            st.divider()
            st.subheader("DEV Tools")

            if st.button(
                "🚀 Push Dev Defaults to Prod",
                type="primary",
                width="stretch",
                help="Overwrite the PROD defaults JSON with the current DEV defaults",
            ):
                st.session_state["confirm_push_dev_to_prod"] = True
            
            if st.button(
                    "🚀 Push Code Changes to Prod",
                    type="primary",
                    width="stretch",
                    help="Merge the dev branch directly into main via GitHub",
                ):
                    st.session_state["confirm_merge_dev_to_main"] = True



        @st.dialog("⚠️ Confirm Default Push to Production")
        def confirm_push_dev_to_prod():
            st.markdown(
                """
                **You are about to overwrite the PROD defaults configuration.**

                - ✅ PROD file ID will remain unchanged  
                - ✅ DEV defaults will completely replace PROD defaults  
                - ❌ This action **cannot be undone**

                Please confirm you want to continue.
                """
            )

            col_confirm, col_cancel = st.columns(2)

            with col_confirm:
                if st.button("✅ Yes — Push to PROD", type="primary", width="stretch"):
                    try:
                        DEV_CONFIG_FILE_ID = (st.secrets.get("DEV_CONFIG_FILE_ID", "") or "").strip()
                        PROD_CONFIG_FILE_ID = (st.secrets.get("CONFIG_FILE_ID", "") or "").strip()

                        if not DEV_CONFIG_FILE_ID or not PROD_CONFIG_FILE_ID:
                            st.error("Missing DEV_CONFIG_FILE_ID or CONFIG_FILE_ID in secrets.")
                        else:
                            creds = load_valid_token(cfg.SCOPES)
                            if not creds:
                                st.error("Google authentication required.")
                            else:
                                _, drive, _ = services(creds, cfg.HTTP_TIMEOUT_SECONDS)

                                # Load DEV defaults
                                dev_blob = drive.files().get_media(
                                    fileId=DEV_CONFIG_FILE_ID
                                ).execute()
                                dev_defaults = json.loads(dev_blob.decode("utf-8"))

                                # Overwrite PROD defaults (same file ID)
                                save_config_to_drive(
                                    drive,
                                    dev_defaults,
                                    file_id=PROD_CONFIG_FILE_ID
                                )

                                st.success("✅ DEV defaults successfully pushed to PROD.")

                    except Exception as e:
                        st.error(f"Push failed: {e}")

                    finally:
                        st.session_state.pop("confirm_push_dev_to_prod", None)
                        _rerun()

            with col_cancel:
                if st.button("❌ Cancel", width="stretch"):
                    st.session_state.pop("confirm_push_dev_to_prod", None)
                    _rerun()

        @st.dialog("⚠️ Confirm Code Push to Production")
        def confirm_merge_dev_to_main():
            base = 'dev'
            target = 'main'
            st.markdown(
                f"""
                **You are about to merge `{base}` into `{target}`.**

                - ✅ GitHub history will be preserved  
                - ✅ Branch protections still apply  
                - ❌ This action **may deploy to production**
                - ❌ This action **cannot be undone**

                Please confirm you want to continue.
                """
            )

            col_confirm, col_cancel = st.columns(2)

            with col_confirm:
                if st.button(
                    f"✅ Yes — Merge {base} → {target}",
                    type="primary",
                    width="stretch",
                ):
                    try:
                        with st.spinner("Merging branches…"):
                            success = github_merge(base, target)

                        if success:
                            st.success(f"✅ {base} successfully merged into {target}.")
                        else:
                            st.error("❌ Merge did not complete.")

                    finally:
                        st.session_state.pop("confirm_merge_dev_to_main", None)
                        st.rerun()

            with col_cancel:
                if st.button("❌ Cancel", width="stretch"):
                    st.session_state.pop("confirm_merge_dev_to_main", None)
                    st.rerun()

        # Trigger dialog
        if st.session_state.get("confirm_push_dev_to_prod"):
            confirm_push_dev_to_prod()
        
        if st.session_state.get("confirm_merge_dev_to_main"):
            confirm_merge_dev_to_main()

        

def render_upload_different_button(cfg):
    if st.button("🔁 Upload different files", width="stretch"):
        reset_to_upload()
        reset_pipeline_state()

        st.session_state.ui_phase = UI_UPLOAD
        _rerun()



def render_result_error(cfg):
    payload = st.session_state.get("pipe_error")

    with st.container(border=True):
        st.subheader("❌ Run Failed")

        if isinstance(payload, dict):
            # ✅ Friendly message (wrapped, readable)
            st.error(f"{payload['type']}: {payload['user_message']}")

            # ✅ Technical details hidden by default
            with st.expander("Technical details"):
                st.text(payload["traceback"])

        else:
            st.error("Unknown error occurred.")
        
        if st.button("🔁 Upload different files", type="primary"):
            st.session_state.pop("run_error", None)
            reset_to_upload()
            reset_pipeline_state()
            _rerun()


    if os.path.exists("last_run.log") and "last_run_log" not in st.session_state:
        with open("last_run.log", "rb") as f:
            st.session_state["last_run_log"] = f.read()
        st.session_state.setdefault("last_run_timestamp", "error")
    
    if "last_run_log" in st.session_state:
        st.download_button(
            "⬇️ Download full log (last_run.log)",
            st.session_state["last_run_log"],
            file_name=f"last_run_{st.session_state['last_run_timestamp']}.log",
            mime="text/plain",
            width='stretch'
            )



def render_upload_error(cfg):
    with st.container(border=True):
        st.subheader("❌ Invalid Upload")

        st.error(
            "Your uploaded file is invalid.\n\n"
            "Please upload **1 or 2 full weeks of data only**."
        )

        st.warning(st.session_state.get("upload_error", ""))

        if st.button("🔁 Upload different files", type="primary"):
            st.session_state.pop("run_error", None)
            reset_to_upload()
            reset_pipeline_state()
            _rerun()


def render_app(cfg):
    phase = st.session_state.ui_phase

    if phase == UI_UPLOAD:
        render_sidebar(cfg)
        render_upload_card(cfg)

    elif phase == UI_READY:
        render_sidebar(cfg)
        render_run_options(cfg)
        render_upload_different_button(cfg)

    elif phase == UI_RUNNING:
        render_running_status(cfg)

    elif phase == UI_RESULT:
        render_sidebar(cfg)
        render_results(cfg)
        render_upload_different_button(cfg)
    
    elif phase == UI_RESULT_ERROR:
        render_sidebar(cfg)
        render_result_error(cfg)

    elif phase == UI_UPLOAD_ERROR:
        render_sidebar(cfg)
        render_upload_error(cfg)

        


# =========================
# App Entrypoint
# =========================

#st.title("🧾 FavTrip Reporting Pipeline")


st.set_page_config(
    page_title="FT Reporting",
    page_icon="🧾",          # emoji or path/URL to an image
    layout="wide",           # "centered" or "wide"
    initial_sidebar_state="collapsed",  # "auto", "expanded", "collapsed"
    menu_items={
        "Get Help": "mailto:ryan-morrow@uiowa.edu",
        "Report a bug": "https://github.com/ryan-j-morrow/favtrip_reporting/issues",
        "About": "FavTrip Reporting Pipeline",
    },
)

defaults = {
    "sales_selected_name": None,
    "vendor_selected_name": None,
    "sales_uploaded_ok": False,
    "vendor_uploaded_ok": False,
    "offer_log_download": False,
    "uploader_version": 0,
    "ui_phase": UI_UPLOAD,
    "auth_required": True,
    "running_ui_initialized": False,    
    "upload_epoch": 0,                 # increments on “Upload different files”
    "sales_selected_epoch": None,       # epoch when sales file was selected
    "vendor_selected_epoch": None,    
    "reset_generation": 0,
    "sales_selection_generation": None,
    "vendor_selection_generation": None,
    "sidebar_hint_seen": True
}

    
for key, default in defaults.items():
    if key not in st.session_state:
        st.session_state[key] = default

cfg = Config.load()

creds = load_valid_token(cfg.SCOPES)
st.session_state.auth_required = creds is None


# --- STATE INIT ---
init_thread_state()
init_pipeline_state()

# --- Finish OAuth inline when redirect comes back (this is in the NEW TAB) ---
params = st.query_params
if "code" in params and "state" in params:
    try:
        finish_web_oauth(params["code"], params["state"], cfg.SCOPES)
        # Token is saved locally in this new tab's app process
        st.success("✅ Google authentication complete.")
        
        has_token = (load_valid_token(cfg.SCOPES) is not None)
        st.session_state.auth_required = not has_token

        # Remove code/state from URL
        st.query_params.clear()

        # No messaging back to opener and NO window.close().
        # This tab becomes the main app; just rerun to flip UI.
        st.toast("Signed in. Loading the app…")
        _rerun()
    except Exception as e:
        st.error(f"OAuth error: {e}")

if (not st.session_state.auth_required) and ("sidebar_hint_seen" not in st.session_state):
    col_msg, col_btn = st.columns([6, 1], vertical_alignment="center")

    with col_msg:
        st.info(
            "⬅️ **Open the sidebar** for Utilities, Google auth, and DEV tools.",
            icon="👈",
        )

    with col_btn:
        if st.button("Got it", type="secondary"):
            st.session_state["sidebar_hint_seen"] = True
            _rerun()

# Auth gate
if st.session_state.auth_required:
    # ----------------------------
    # Authentication panel (shown only if auth required)
    # ----------------------------
    if st.session_state.auth_required:
        with st.expander("Google Authentication", expanded=True):
            st.caption(
                "Authentication is required before running. "
                "Click **Sign in with Google** to open the consent screen (it will open in a new tab)."
            )

            sign_in_ph = st.empty()
            clicked = sign_in_ph.button("Sign in with Google", type="primary", width='stretch')

            if clicked:
                try:
                    auth_url = start_web_oauth(cfg.SCOPES)
                    sign_in_ph.empty()

                    # Friendly message in this (original) tab
                    st.markdown(
                        """
                        <div style="
                            display:flex;align-items:center;justify-content:center;
                            height:55vh;text-align:center;
                            font-family: system-ui, Segoe UI, Roboto, Helvetica, Arial, sans-serif;">
                        <div>
                            <h2 style="margin-bottom:0.5rem;">You're being signed in…</h2>
                            <p style="font-size:1.05rem;opacity:.9;">
                            A new browser tab was opened for Google sign‑in.<br/>
                            <strong>After it completes, continue in that tab.</strong>
                            </p>
                        </div>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )

                    # Optional: refresh this tab when user returns (not required)
                    html(
                        """
                        <script>
                        document.addEventListener("visibilitychange", function() {
                            if (!document.hidden) { location.reload(); }
                        });
                        </script>
                        """,
                        height=0,
                    )

                    # Open Google auth in a NEW tab (this will ultimately become the main app)
                    html(
                        f"""
                        <script>
                        window.open({json.dumps(auth_url)}, "_blank", "noopener");
                        </script>
                        """,
                        height=0,
                    )

                    st.stop()
                except Exception as e:
                    st.error(f"Failed to start OAuth: {e}")

            with st.expander("Having trouble?", expanded=False):
                st.write(
                    "- The Google authorization page opens in a **new browser tab**.\n"
                    "- After completing consent, the **new tab** will load the app.\n"
                    "- If you renamed your Streamlit app or URL, ensure the Google OAuth "
                    "Authorized redirect URI matches exactly (including trailing slash)."
                )
            st.stop()

render_app(cfg)

```

---
### file: cli.py

```python
import argparse
from favtrip.config import Config
from favtrip.logger import StatusLogger
from favtrip.pipeline import run_pipeline


def main():
    parser = argparse.ArgumentParser(description="FavTrip Reporting Pipeline")
    parser.add_argument("--env", help="Path to .env file", default=None)

    # Per-run overrides (subset)
    parser.add_argument("--to", help="Comma-separated recipients", default=None)
    parser.add_argument("--cc", help="Comma-separated cc", default=None)
    parser.add_argument("--use-all-keys", action="store_true")
    parser.add_argument("--report-keys", help="Comma-separated report keys to run", default=None)
    parser.add_argument("--force-reauth", action="store_true")

    args = parser.parse_args()
    cfg = Config.load(args.env)

    if args.to:
        cfg.TO_RECIPIENTS = [s.strip() for s in args.to.split(',') if s.strip()]
    if args.cc:
        cfg.CC_RECIPIENTS = [s.strip() for s in args.cc.split(',') if s.strip()]
    if args.use_all_keys:
        cfg.USE_ALL_REPORT_KEYS = True
    if args.report_keys:
        cfg.REPORT_KEY_RUN_LIST = [s.strip().upper() for s in args.report_keys.split(',') if s.strip()]
    if args.force_reauth:
        cfg.FORCE_REAUTH = True

    logger = StatusLogger()
    result = run_pipeline(cfg, logger=logger)

    print("===== SUMMARY =====")
    print(logger.as_text())
    print("===================")


if __name__ == "__main__":
    main()

```

---
### file: core_functional_modules/__init__.py

```python
__all__ = [
    "config",
    "google_client",
    "sheets_utils",
    "drive_utils",
    "gmail_utils",
    "pipeline",
    "logger",
]

```

---
### file: core_functional_modules/config.py

```python
"""
config
======================================

Configuration loader and serializer for FavTrip reporting apps.

This module centralizes all runtime configuration for both local development
and cloud deployments (e.g., Streamlit Community Cloud). It provides a single,
typed `Config` dataclass plus helper functions that safely read from multiple
sources, coerce values to the expected Python types, and (optionally) overlay
a remote, Google Drive–hosted JSON configuration at runtime.

The loader is designed to be:
- **Layered**: Values are pulled from three tiers, in this order:
  1) Streamlit `st.secrets` (preferred in cloud; values may already be typed)
  2) Process environment and/or a local `.env` file (string-based; coerced) #Sandbox use only
  3) A Google Drive JSON override, applied last if credentials and
     a config file are available
- **Safe**: Missing keys never raise; reasonable defaults are used instead.
- **Type-aware**: Bools, lists, and dicts are parsed/coerced consistently so the
  same code works with typed TOML (in `st.secrets`) and string-based `.env`.

-------------------------------------------------------------------------------
Core API
-------------------------------------------------------------------------------

- `_get_secret(key: str, default: Any = None) -> Any`
  Attempts to read `key` from `streamlit.secrets` (if Streamlit is present and
  has `secrets`), else falls back to `os.getenv(key, default)`. Never raises for
  missing keys; always returns a value (possibly `default`). Streamlit import is
  lazy to avoid a hard dependency for non-Streamlit contexts.

- `_coerce_bool(v: Any, default: bool = False) -> bool`
  Accepts `bool | str | int | None` and returns a Python `bool`.
  Truthy strings (case-insensitive, trimmed) include:
  `{"1", "true", "yes", "on", "y", "t"}`. Non-parseable inputs fall back to
  `default`.

- `_coerce_csv(v: Any) -> List[str]`
  Accepts a list/tuple (already structured) or a comma-separated string and
  yields a list of **trimmed** strings. `None`/empty returns `[]`.

- `_coerce_json(v: Any) -> Dict[str, Any]`
  Accepts a `dict` or a JSON string. Returns a `dict`; parse failures yield `{}`.

- `@dataclass class Config`
  A top-level dataclass holding all tunable settings for the application:
  * **Drive/Sheets IDs**:
    - `CALC_SPREADSHEET_ID`, `INCOMING_FOLDER_ID`, `MANAGER_REPORT_FOLDER_ID`,
      `ORDER_REPORT_FOLDER_ID`, `USER_FOLDER_ID`
  * **GIDs, sheet metadata, timestamps**:
    - `GID_MANAGER_PDF`, `GID_ORDER_CSV`, `LOCATION_SHEET_TITLE`,
      `LOCATION_NAMED_RANGE`, `TEMPLATE_UPDATE_RANGE`
    - `TIMESTAMP_TZ` (e.g., "America/Chicago")
    - `TIMESTAMP_FMT` (default "%Y-%m-%d-%I-%M-%p")
  * **Email & distribution**:
    - `TO_RECIPIENTS`, `CC_RECIPIENTS`, `USE_ALL_REPORT_KEYS`,
      `REPORT_KEY_RUN_LIST`, `REPORT_KEY_RECIPIENTS`,
      `DEFAULT_ORDER_RECIPIENTS`
    - `INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL`,
      `SEND_SEPARATE_FULL_ORDER_EMAIL`, `EMAIL_MANAGER_REPORT`
  * **Google API**:
    - `SCOPES` (Drive/Sheets/Gmail), `FORCE_REAUTH`,
      `REDIRECT_PORT`, `HTTP_TIMEOUT_SECONDS`
  * **Advanced intake**:
    - `USE_AUTO_ROLLOVER_IF_ONE_WEEK`,
      `START_DAY_OF_WEEK`, `END_DAY_OF_WEEK`
      (Accepted values include: Sunday, Monday, Tuesday, Wednesday, Thursday,
      Friday, Saturday, Any)
  * **Cleanup (days)**:
    - `OUTPUT_TIME_TO_LIFE`, `FAILED_INPUT_TIME_TO_LIFE`, `USER_TIME_TO_LIFE`

  Defaults are provided for all fields. When loading from secrets or `.env`,
  values are coerced into the correct types; lists and dicts are parsed as
  necessary. `REPORT_KEY_RUN_LIST` values are uppercased to reduce downstream
  casing issues.

- `Config.load(env_path: Optional[pathlib.Path] = None) -> Config`
  Loads the final, effective configuration via a layered merge:
  1) Loads a local `.env` file from `env_path` (default: `cwd/.env`) using
     `python-dotenv` with `override=False` (so existing process env vars win).
  2) Reads settings from `st.secrets` if available; otherwise from environment.
     Values are passed through the coercers defined above.
  3) Attempts to overlay a Google Drive–hosted JSON config:
     - Uses `core_functional_modules.google_client.load_valid_token` and `services` to obtain a
       Drive client (respecting `HTTP_TIMEOUT_SECONDS`).
     - Reads a JSON dict via `core_functional_modules.config_store.load_config_from_drive`,
       optionally using `CONFIG_FILE_ID` from `st.secrets` if present.
     - Keys in the override dict that match `Config` attributes replace
       previously loaded values.
     - On any failure (no token, network error, file missing, etc.), the loader
       **fails open** and returns the base config without raising (best-effort).

- `Config.to_env() -> str`
  Serializes the current configuration to a string in `.env` format. Collections
  are flattened—lists are joined with commas, and dicts are JSON-encoded—so the
  output can be written to disk and re-read later in a purely string-based env.

- `Config.save(env_path: Optional[pathlib.Path] = None) -> None`
  Convenience wrapper around `to_env()` that writes the serialized configuration
  to `env_path` (default: `cwd/.env`, UTF-8).

-------------------------------------------------------------------------------
Environment / Secrets Reference (all optional; sensible defaults apply)
-------------------------------------------------------------------------------

Drive / Sheets IDs:
- `CALC_SPREADSHEET_ID`, `INCOMING_FOLDER_ID`, `MANAGER_REPORT_FOLDER_ID`,
  `ORDER_REPORT_FOLDER_ID`, `USER_FOLDER_ID`

Sheet metadata & timestamps:
- `GID_MANAGER_PDF`, `GID_ORDER_CSV`, `LOCATION_SHEET_TITLE`,
  `LOCATION_NAMED_RANGE`, `TEMPLATE_UPDATE_RANGE`
- `TIMESTAMP_TZ`, `TIMESTAMP_FMT`

Email & distribution:
- `TO_RECIPIENTS` (CSV), `CC_RECIPIENTS` (CSV)
- `USE_ALL_REPORT_KEYS` (bool)
- `REPORT_KEY_RUN_LIST` (CSV; uppercased during load)
- `REPORT_KEY_RECIPIENTS` (JSON dict)
- `DEFAULT_ORDER_RECIPIENTS` (CSV)
- `INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL` (bool)
- `SEND_SEPARATE_FULL_ORDER_EMAIL` (bool)
- `EMAIL_MANAGER_REPORT` (bool)

Google API:
- `SCOPES` (CSV; typical: Drive/Sheets/Gmail send)
- `FORCE_REAUTH` (bool)
- `REDIRECT_PORT` (int)
- `HTTP_TIMEOUT_SECONDS` (int)

Advanced intake / rollover:
- `USE_AUTO_ROLLOVER_IF_ONE_WEEK` (bool)
- `START_DAY_OF_WEEK`, `END_DAY_OF_WEEK`
  (Sunday|Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Any)

Cleanup (days):
- `OUTPUT_TIME_TO_LIFE`, `FAILED_INPUT_TIME_TO_LIFE`, `USER_TIME_TO_LIFE`

Drive override:
- `CONFIG_FILE_ID` (usually provided via `st.secrets`, if using a specific file)

-------------------------------------------------------------------------------
Operational Notes
-------------------------------------------------------------------------------

- **Lazy imports**: `streamlit` and Google client utilities are imported inside
  the loader so the module remains usable in non-Streamlit or headless contexts.
- **Fail-open Drive overrides**: If Drive credentials are unavailable or an
  override file cannot be retrieved/parsed, the loader returns the base config
  without raising (best-effort behavior).
- **Deterministic parsing**: Coercers are idempotent for already-typed values.
  For example, booleans in TOML remain booleans; CSV strings are split and
  trimmed; JSON strings are parsed into dicts.
- **Case normalization**: `REPORT_KEY_RUN_LIST` is uppercased at load time to
  minimize case-related mismatches elsewhere in the app.

Import this module early in your app to construct a single, consistent
`Config` instance and pass it through to components that require configuration.

"""


from __future__ import annotations

import os
import json
from dataclasses import dataclass, asdict, field
from typing import Any, Dict, List, Optional
from pathlib import Path
from dotenv import load_dotenv

# -----------------------------------------------------------------------------
# Helpers: read from Streamlit secrets (typed) or .env (strings) and coerce
# -----------------------------------------------------------------------------

def _get_secret(key: str, default: Any = None) -> Any:
    """
    Read from Streamlit secrets if present, else env var, else default.
    Does not raise if key missing; returns `default`.
    """
    try:
        import streamlit as st  # imported lazily to avoid hard dependency
        if hasattr(st, "secrets") and key in st.secrets:
            return st.secrets.get(key, default)
    except Exception:
        pass
    return os.getenv(key, default)

_TRUE = {"1", "true", "yes", "on", "y", "t"}

def _coerce_bool(v: Any, default: bool = False) -> bool:
    """
    Accept bool | str | int | None and return a Python bool.
    Works for typed TOML (bool) and .env strings.
    """
    if isinstance(v, bool):
        return v
    if v is None:
        return default
    try:
        return str(v).strip().lower() in _TRUE
    except Exception:
        return default

def _coerce_csv(v: Any) -> List[str]:
    """
    Accept list/tuple (already structured) or a comma-separated string.
    Returns a list of trimmed strings.
    """
    if v is None or v == "":
        return []
    if isinstance(v, (list, tuple)):
        return [str(x).strip() for x in v if str(x).strip()]
    return [p.strip() for p in str(v).split(",") if p.strip()]

def _coerce_json(v: Any) -> Dict[str, Any]:
    """
    Accept dict (already structured) or a JSON string.
    Returns a dict; falls back to {} on parse issues.
    """
    if v is None or v == "":
        return {}
    if isinstance(v, dict):
        return v
    try:
        return json.loads(v)
    except Exception:
        return {}

# -----------------------------------------------------------------------------
# Config dataclass (TOP-LEVEL — must start at column 0)
# -----------------------------------------------------------------------------

@dataclass
class Config:
    NON_PERSISTED_FIELDS = set()  
  
    # IDs and basic settings
    CALC_SPREADSHEET_ID: str = "1ibkGkQ2khYMJydeenJkTzC4KoLQAyBZW_esQrbjSHXs"
    INCOMING_FOLDER_ID: str = "1jJE3r9DOHXwBdd94E6ZhxBBH9xvSjI-b"
    MANAGER_REPORT_FOLDER_ID: str = "17Nqwo6HYe30JP0wnZYoLRG0F1s-X-IVZ"
    ORDER_REPORT_FOLDER_ID: str = "171dqzMim-IdpB_kzjYQnzoSbW89uJTfP"
    ERROR_REPORT_FOLDER_ID: str = "1T-rnyXmPD1eFcxi-s8i4b1EP6-pW5ETW"
    USER_FOLDER_ID: str = "1JBHBcnS6397ka2ITW6Wbuu2aKjbgCCHj"

    # GIDs, sheet metadata, timestamp settings
    GID_MANAGER_PDF: str = "1921812573"
    GID_ORDER_CSV: str = "1875928148"
    GID_ERROR_REPORT: str = "1581903111"
    GID_BEV_ERRORS: str = "72711538"
    LOCATION_SHEET_TITLE: str = "REFR: Values"
    LOCATION_NAMED_RANGE: str = "_locations"
    TIMESTAMP_TZ: str = "America/Chicago"
    TIMESTAMP_FMT: str = "%Y-%m-%d-%I-%M-%p"
    TEMPLATE_UPDATE_RANGE: str = "_update"

    # Email config
    TO_RECIPIENTS: List[str] = field(default_factory=lambda: ["FavtripReporting@gmail.com"])
    CC_RECIPIENTS: List[str] = None
    ERROR_RECIPIENTS: List[str] = field(default_factory=lambda: ["FavtripReporting@gmail.com"])
    USE_ALL_REPORT_KEYS: bool = False
    REPORT_KEY_RUN_LIST: List[str] = field(default_factory=lambda: ["COFFEE"])
    REPORT_KEY_RECIPIENTS: Dict[str, List[str]] = None
    DEFAULT_ORDER_RECIPIENTS: List[str] = None
    INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL: bool = False
    SEND_SEPARATE_FULL_ORDER_EMAIL: bool = False
    EMAIL_MANAGER_REPORT: bool = True

    # Google API
    SCOPES: List[str] = None
    FORCE_REAUTH: bool = False
    REDIRECT_PORT: int = 58285
    HTTP_TIMEOUT_SECONDS: int = 300

    # Advanced intake settings
    USE_AUTO_ROLLOVER_IF_ONE_WEEK: bool = True
    START_DAY_OF_WEEK: str = "Sunday"    # Sunday, Monday, Tuesday, Wednesday, Thursday, Friday, Saturday, Any
    END_DAY_OF_WEEK: str = "Saturday"    # Sunday, Monday, Tuesday, Wednesday, Thursday, Friday, Saturday, Any
    
    SOFT_CASES_ALERT_ENABLED: bool = True
    SOFT_CASES_ALERT_THRESHOLD: int = 10


    # Cleanup
    OUTPUT_TIME_TO_LIFE: int = 30
    FAILED_INPUT_TIME_TO_LIFE: int = 1
    USER_TIME_TO_LIFE: int = 90

    #Other
    BEV_MAPPING_LINK: str = "https://docs.google.com/spreadsheets/d/1O6MtF-GM0VayqMr_v3oJC5PnRK5yv6biiDtA_qw-Z3g/"

    @staticmethod
    def load(env_path: Optional[Path] = None) -> "Config":
        """
        Load config from Streamlit secrets (preferred on cloud) or from env/.env (local dev),
        then overlay any values found in a Drive-backed JSON config (optional).
        Secrets may be typed (bool/list/dict), so we coerce safely.
        """
        if env_path is None:
            env_path = Path.cwd() / ".env"
        load_dotenv(dotenv_path=env_path, override=False)

        cfg = Config(
            CALC_SPREADSHEET_ID=str(_get_secret("CALC_SPREADSHEET_ID", "")),
            INCOMING_FOLDER_ID=str(_get_secret("INCOMING_FOLDER_ID", "")),
            MANAGER_REPORT_FOLDER_ID=str(_get_secret("MANAGER_REPORT_FOLDER_ID", "")),
            ORDER_REPORT_FOLDER_ID=str(_get_secret("ORDER_REPORT_FOLDER_ID", "")),
            ERROR_REPORT_FOLDER_ID=str(_get_secret("ERROR_REPORT_FOLDER_ID", "")),
            USER_FOLDER_ID=str(_get_secret("USER_FOLDER_ID", "")),

            GID_MANAGER_PDF=str(_get_secret("GID_MANAGER_PDF", "1921812573")),
            GID_ORDER_CSV=str(_get_secret("GID_ORDER_CSV", "1875928148")),
            GID_ERROR_REPORT=str(_get_secret("GID_ERROR_REPORT", "1581903111")),
            GID_BEV_ERRORS=str(_get_secret("GID_BEV_ERRORS", "72711538")),
            LOCATION_SHEET_TITLE=str(_get_secret("LOCATION_SHEET_TITLE", "REFR: Values")),
            LOCATION_NAMED_RANGE=str(_get_secret("LOCATION_NAMED_RANGE", "_locations")),
            TEMPLATE_UPDATE_RANGE=str(_get_secret("TEMPLATE_UPDATE_RANGE", "_update")),
            TIMESTAMP_TZ=str(_get_secret("TIMESTAMP_TZ", "America/Chicago")),
            TIMESTAMP_FMT=str(_get_secret("TIMESTAMP_FMT", "%Y-%m-%d-%I-%M-%p")),

            OUTPUT_TIME_TO_LIFE=int(_get_secret("OUTPUT_TIME_TO_LIFE", 30)),
            FAILED_INPUT_TIME_TO_LIFE=int(_get_secret("FAILED_INPUT_TIME_TO_LIFE", 1)),
            USER_TIME_TO_LIFE=int(_get_secret("USER_TIME_TO_LIFE", 1)),

            TO_RECIPIENTS=_coerce_csv(_get_secret("TO_RECIPIENTS", "")),
            CC_RECIPIENTS=_coerce_csv(_get_secret("CC_RECIPIENTS", "")),
            ERROR_RECIPIENTS=_coerce_csv(_get_secret("ERROR_RECIPIENTS", "")),
            USE_ALL_REPORT_KEYS=_coerce_bool(_get_secret("USE_ALL_REPORT_KEYS", "false")),
            REPORT_KEY_RUN_LIST=[s.upper() for s in _coerce_csv(_get_secret("REPORT_KEY_RUN_LIST", ""))],
            REPORT_KEY_RECIPIENTS=_coerce_json(_get_secret("REPORT_KEY_RECIPIENTS", "{}")),
            DEFAULT_ORDER_RECIPIENTS=_coerce_csv(_get_secret("DEFAULT_ORDER_RECIPIENTS", "")),
            INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL=_coerce_bool(
                _get_secret("INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL", "false")
            ),
            SEND_SEPARATE_FULL_ORDER_EMAIL=_coerce_bool(
                _get_secret("SEND_SEPARATE_FULL_ORDER_EMAIL", "true")
            ),
            EMAIL_MANAGER_REPORT=_coerce_bool(_get_secret("EMAIL_MANAGER_REPORT", "true")),

            SCOPES=_coerce_csv(
                _get_secret(
                    "SCOPES",
                    "https://www.googleapis.com/auth/drive,"
                    "https://www.googleapis.com/auth/spreadsheets,"
                    "https://www.googleapis.com/auth/gmail.send",
                )
            ),
            FORCE_REAUTH=_coerce_bool(_get_secret("FORCE_REAUTH", "false")),
            REDIRECT_PORT=int(str(_get_secret("REDIRECT_PORT", "58285")) or "58285"),
            HTTP_TIMEOUT_SECONDS=int(str(_get_secret("HTTP_TIMEOUT_SECONDS", "300")) or "300"),

            USE_AUTO_ROLLOVER_IF_ONE_WEEK=_coerce_bool(_get_secret("USE_AUTO_ROLLOVER_IF_ONE_WEEK", "true")),
            START_DAY_OF_WEEK=str(_get_secret("START_DAY_OF_WEEK", "Sunday")),
            END_DAY_OF_WEEK=str(_get_secret("END_DAY_OF_WEEK", "Saturday")),

            
            SOFT_CASES_ALERT_ENABLED=_coerce_bool(_get_secret("SOFT_CASES_ALERT_ENABLED", "true")),
            SOFT_CASES_ALERT_THRESHOLD=int(_get_secret("SOFT_CASES_ALERT_THRESHOLD", 10)),

            BEV_MAPPING_LINK=str(_get_secret("BEV_MAPPING_LINK", "https://docs.google.com/spreadsheets/d/1O6MtF-GM0VayqMr_v3oJC5PnRK5yv6biiDtA_qw-Z3g/")),
        )

        normalized = {}
        for k, v in cfg.REPORT_KEY_RECIPIENTS.items():
            if isinstance(k, (list, tuple)):
                if len(k) == 2:
                    normalized[(k[0], k[1], None)] = v
                elif len(k) == 3:
                    normalized[tuple(k)] = v
            else:
                # defensive fallback
                normalized[(None, k, None)] = v

        cfg.REPORT_KEY_RECIPIENTS = normalized


        # Optional overlay from Drive JSON config (if creds + file present)
        # ---------------- Drive-backed config overlay ----------------
        try:
            import streamlit as st
            from core_functional_modules.google_client import load_valid_token, services
            from core_functional_modules.config_store import load_config_from_drive

            DEV_ENVIRONMENT = _coerce_bool(_get_secret("DEV_ENVIRONMENT", False))
            DEV_CONFIG_FILE_ID = str(_get_secret("DEV_CONFIG_FILE_ID", "") or "").strip()
            CONFIG_FILE_ID = str(_get_secret("CONFIG_FILE_ID", "") or "").strip()

            # Select which config file ID to READ from
            if DEV_ENVIRONMENT and DEV_CONFIG_FILE_ID:
                active_config_file_id = DEV_CONFIG_FILE_ID
            else:
                active_config_file_id = CONFIG_FILE_ID or None

            creds = load_valid_token(cfg.SCOPES)
            if creds:
                _sheets, drive, _gmail = services(creds, cfg.HTTP_TIMEOUT_SECONDS)

                overrides = {}
                
                # 1️⃣ Try DEV config first (if enabled)
                if DEV_ENVIRONMENT and DEV_CONFIG_FILE_ID:
                    overrides = load_config_from_drive(drive, DEV_CONFIG_FILE_ID)

                # 2️⃣ Fallback to PROD config if DEV missing/empty
                if not overrides and CONFIG_FILE_ID:
                    overrides = load_config_from_drive(drive, CONFIG_FILE_ID)

                # 3️⃣ Apply overrides if any
                if isinstance(overrides, dict):
                    for k, v in overrides.items():
                        if hasattr(cfg, k):
                            setattr(cfg, k, v)

        except Exception:
            # Fail-open by design
            
            import traceback
            print("[Config] Drive overlay failed:", e)
            traceback.print_exc()

            #pass
        
        return cfg


    # -------------------------------------------------------------------------
    # .env serialization (optional helper)
    # -------------------------------------------------------------------------
    def to_env(self) -> str:
        """Serialize to .env format (simple, string-based)."""
        data = asdict(self)
        as_env = {
            **data,
            "TO_RECIPIENTS": ",".join(self.TO_RECIPIENTS or []),
            "CC_RECIPIENTS": ",".join(self.CC_RECIPIENTS or []),
            "REPORT_KEY_RUN_LIST": ",".join(self.REPORT_KEY_RUN_LIST or []),
            "REPORT_KEY_RECIPIENTS": json.dumps(self.REPORT_KEY_RECIPIENTS or {}),
            "DEFAULT_ORDER_RECIPIENTS": ",".join(self.DEFAULT_ORDER_RECIPIENTS or []),
            "SCOPES": ",".join(self.SCOPES or []),
        }
        lines = [f"{k}={v}" for k, v in as_env.items()]
        return "\n".join(lines) + "\n"

    def save(self, env_path: Optional[Path] = None):
        if env_path is None:
            env_path = Path.cwd() / ".env"
        env_path.write_text(self.to_env(), encoding="utf-8")
    

    
    def to_drive_defaults(self) -> dict:
        return {
            k: v
            for k, v in vars(self).items()
            if not k.startswith("_")
            and k not in self.NON_PERSISTED_FIELDS
        }


```

---
### file: core_functional_modules/config_store.py

```python
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

```

---
### file: core_functional_modules/drive_utils.py

```python
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

def get_or_create_subfolder(drive_svc, parent_folder_id: str, name: str):
    """
    Return a Drive folder with the given name under parent_folder_id.
    Create it if it does not already exist.
    """
    q = (
        f"mimeType='application/vnd.google-apps.folder' "
        f"and name='{name}' "
        f"and '{parent_folder_id}' in parents "
        f"and trashed=false"
    )

    res = drive_svc.files().list(
        q=q,
        fields="files(id, name, webViewLink)",
        pageSize=1
    ).execute()

    files = res.get("files", [])
    if files:
        return files[0]

    metadata = {
        "name": name,
        "mimeType": "application/vnd.google-apps.folder",
        "parents": [parent_folder_id],
    }

    return drive_svc.files().create(
        body=metadata,
        fields="id, name, webViewLink"
    ).execute()


```

---
### file: core_functional_modules/gmail_utils.py

```python
""" 
gmail_utils
======================================
Email utilities for sending Gmail messages with PDF attachments via the Gmail API.

This module provides small, focused helpers for sending two types of emails through
the Gmail API using a pre-authorized `gmail_svc` client (e.g., returned by
`googleapiclient.discovery.build("gmail", "v1", ...)`). It includes:

- `send_email(...)`: Low-level helper that accepts an `email.message.EmailMessage`,
  base64-url encodes it as required by Gmail, and dispatches it via
  `users.messages.send`.

- `email_manager_report(...)`: Composes and sends a standardized "Manager Report"
  email with a primary PDF attachment and a backup link. Supports optional CC.

- `email_order_report(...)`: Composes and sends an "Order Report" email for a
  given vendor or category key, including a primary PDF attachment and an optional
  "full order" PDF. Also includes links to a backing Google Sheet and supports CC.

The functions here intentionally perform minimal validation and assume that callers
supply valid addresses, attachments, and links. Authentication, token refresh, and
error handling policy (e.g., retries, backoff, alerting) should be implemented by
the caller.

---
Key Behaviors
-------------
- **MIME construction**: Uses Python's stdlib `email.message.EmailMessage` to build
  multipart emails with both plain-text and HTML alternatives, and PDF attachments.
- **Gmail API compliance**: Serializes the email to bytes and encodes it with
  URL-safe Base64 as required by Gmail's `users.messages.send` endpoint.
- **Idempotency**: Sending is not idempotent; calling functions repeatedly may
  result in duplicate emails. Callers should implement their own guardrails if
  needed (e.g., deduplication keys, sent-flagging).
- **Internationalization**: The functions do not localize content; callers can adapt
  the text if i18n is required.
- **HTML content**: Simple HTML bodies are included via `add_alternative(..., subtype="html")`.
  The HTML snippets intentionally avoid external assets for reliable delivery.

---
Functions
---------
send_email(gmail_svc, user, msg)
    Low-level send helper. Encodes the `EmailMessage` and dispatches via the Gmail API.

email_manager_report(gmail_svc, sender, to_list, cc_list, pdf_name, pdf_bytes, pdf_link, ts, location)
    Sends a standardized "Manager Report" email with a PDF attachment and a backup link.

email_order_report(
    gmail_svc,
    sender,
    to_list,
    cc_list,
    key,
    tag,
    ts,
    location,
    pdf_name,
    pdf_bytes,
    sheet_link,
    include_full_order=False,
    full_pdf_bytes=None,
    full_pdf_name=None,
)
    Sends an "Order Report" email targeted to a `{key}` team with a primary PDF,
    optional full-order PDF, and a link to the backing Google Sheet.

---
Parameters (Shared Concepts)
----------------------------
gmail_svc : Any
    An authenticated Gmail API service client (e.g., from `googleapiclient.discovery.build`).

sender : str
    The "From" email address to display in the message header. The authenticated
    Gmail account must be authorized to send from this address.

to_list : Iterable[str]
    Recipient email addresses for the `To` field. Must contain at least one valid address.

cc_list : Optional[Iterable[str]]
    Optional CC recipient addresses. If empty or `None`, the `Cc` header is omitted.

pdf_name : str
    Filename for the attached PDF (e.g., `"report_2026-03-21.pdf"`).

pdf_bytes : bytes
    Raw bytes of the primary PDF attachment.

ts : str
    A timestamp string suitable for inclusion in the subject (e.g., `"2026-03-21"` or
    `"2026-03-21 18:25"`).

location : str
    A human-readable location name included in the subject/body (e.g., store or site).

pdf_link : str
    (Manager Report) A backup URL users can access if attachments are blocked.

key : str
    (Order Report) An identifier for the receiving team or vendor (e.g., `"Dairy"`, `"VendorX"`).

tag : str
    (Order Report) A secondary descriptor (e.g., `"Weekly"`, `"Overstock"`, `"Emergency"`).

sheet_link : str
    (Order Report) URL to the backing Google Sheet with order details.

include_full_order : bool
    (Order Report) Whether to attach an additional "full order" PDF.

full_pdf_bytes : Optional[bytes]
    (Order Report) Raw bytes of the full order PDF (required when `include_full_order=True`).

full_pdf_name : Optional[str]
    (Order Report) Filename for the full order PDF (required when `include_full_order=True`).

user : str
    (send_email) Gmail user identifier for the API call. Typically `"me"` to refer
    to the authenticated account.

msg : EmailMessage
    (send_email) A fully-constructed email message to be sent.

---
Returns
-------
dict
    The Gmail API response payload from `users.messages.send()` (e.g., includes `id`, `threadId`).

---
Raises
------
googleapiclient.errors.HttpError
    If the Gmail API call fails (e.g., quota exceeded, invalid permissions, bad request).
ValueError / TypeError
    If provided inputs (addresses, bytes, filenames) are invalid (may be raised by stdlib or caller validations).

---
"""


from __future__ import annotations
import base64
from email.message import EmailMessage


def send_email(gmail_svc, user: str, msg: EmailMessage):
    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
    return gmail_svc.users().messages().send(userId=user, body={"raw": raw}).execute()


def email_manager_report(gmail_svc, sender: str, to_list, cc_list, pdf_name, pdf_bytes, pdf_link, ts, location):
    msg = EmailMessage()
    msg["Subject"] = f"Manager Report – {location} – {ts}"
    msg["From"] = sender
    msg["To"] = ", ".join(to_list)
    if cc_list:
        msg["Cc"] = ", ".join(cc_list)
    msg.set_content(f"Hi team,\nAttached is the Manager Report ({location}).\nBackup link: {pdf_link}\n—Sent from an automated reporting pipeline")

    msg.add_alternative(
        f"""
        <p>Hi team,</p>
        <p>Your manager report for store <b>{location}</b> is ready.</p>
        <p><a href='{pdf_link}'>Backup Link</a></p>
        <p>Attached: {pdf_name}</p>
        <p>—Sent from an automated reporting pipeline</p>
        """,
        subtype="html",
    )

    msg.add_attachment(pdf_bytes, maintype="application", subtype="pdf", filename=pdf_name)
    return send_email(gmail_svc, sender, msg)


def email_order_report(
    gmail_svc,
    sender: str,
    to_list,
    cc_list,
    key: str,
    tag: str,
    ts: str,
    location: str,
    pdf_name: str,
    pdf_bytes: bytes,
    sheet_link: str,
    include_full_order: bool = False,
    full_pdf_bytes: bytes | None = None,
    full_pdf_name: str | None = None,
):
    msg = EmailMessage()

    msg["Subject"] = f"Order Report – {location} – {tag} – {ts}"
    msg["From"] = sender
    msg["To"] = ", ".join(to_list)

    if cc_list:
        msg["Cc"] = ", ".join(cc_list)

    msg.set_content(
        f"Hi {tag} team,\n"
        f"Your order report for {location} - {tag} is ready.\n"
        f"Google Sheet: {sheet_link}\n"
        f"Attached: {pdf_name}\n"
        "—Sent from an automated reporting pipeline"
    )

    msg.add_alternative(
        f"""
        <p>Hi {tag} team,</p>
        <p>Your order report for store <b>{location}</b> is ready.</p>
        <p><a href="{sheet_link}">Open Google Sheet</a></p>
        <p>Attached: {pdf_name}</p>
        <p>—Sent from an automated reporting pipeline</p>
        """,
        subtype="html",
    )

    msg.add_attachment(
        pdf_bytes,
        maintype="application",
        subtype="pdf",
        filename=pdf_name,
    )

    if include_full_order and full_pdf_bytes:
        msg.add_attachment(
            full_pdf_bytes,
            maintype="application",
            subtype="pdf",
            filename=full_pdf_name,
        )

    return send_email(gmail_svc, sender, msg)


def email_error_report(
    gmail_svc,
    sender: str,
    to_list,
    cc_list,
    ts: str,
    pdf_name: str,
    pdf_bytes: bytes,
    sheet_link: str
    ):
    msg = EmailMessage()

    msg["Subject"] = f"Error Report – {ts}"
    msg["From"] = sender
    msg["To"] = ", ".join(to_list)

    if cc_list:
        msg["Cc"] = ", ".join(cc_list)

    msg.set_content(
        f"Hi Technical Support Team,\n"
        f"A user tried using the reporting pipeline, however some of the items that were uploaded are not listed on the Vendor Price Book.\n"
        f"The default recipient of the pipeline run is CC'd on this email for visibility and communication purposes.\n"
        f"Please reply to this email once the Vendor Price Book is updated so that the user knows they can rerun the pipeline.\n\n"
        f"Google Sheet: {sheet_link}\n"
        f"Attached: {pdf_name}\n"
        "—Sent from an automated reporting pipeline"
    )

    msg.add_alternative(
        f"""
        <p>Hi Technical Support Team,</p>
        <p>A user tried using the reporting pipeline, however some of the items that were uploaded are not listed on the Vendor Price Book.</p>
        <p>The default recipient of the pipeline run is CC'd on this email for visibility and communication purposes.</p>
        <p>Please reply to this email once the Vendor Price Book is updated so that the user knows they can rerun the pipeline.</p>
        <p></p>
        <p><a href="{sheet_link}">Open Error Report in Google Sheets</a></p>
        <p>Attached: {pdf_name}</p>
        <p>—Sent from an automated reporting pipeline</p>
        """,
        subtype="html",
    )

    msg.add_attachment(
        pdf_bytes,
        maintype="application",
        subtype="pdf",
        filename=pdf_name,
    )

    return send_email(gmail_svc, sender, msg)

def email_bev_error_report(
    gmail_svc,
    sender: str,
    to_list,
    cc_list,
    ts: str,
    pdf_name: str,
    pdf_bytes: bytes,
    sheet_link: str,
    mapping_link: str
    ):
    msg = EmailMessage()

    msg["Subject"] = f"Soft Alert – Unassigned Beverages Report – {ts}"
    msg["From"] = sender
    msg["To"] = ", ".join(to_list)

    if cc_list:
        msg["Cc"] = ", ".join(cc_list)

    msg.set_content(
        f"Hi Technical Support Team,\n"
        f"A user just ran the ordering pipeline successfully, however some beverages were not mapped to a sub-report-key / vendor.\n"
        f"The BEV orders were still sent, however these unassigned beverages were included on their own order under BEV - UNASSIGNED.\n\n"
        f"The default recipient of the pipeline run is CC'd on this email for visibility and communication purposes.\n"
        f"To fix this error for future runs look at the Unassigned Beverages Report and add those Scan Codes to the Mapping File.\n"
        f"Please reply to this email once the Mapping File is updated so that the user knows they can rerun the pipeline if needed.\n\n"
        f"Beverage Mapping File: {mapping_link}\n"
        f"Unassigned Beverages Google Sheet: {sheet_link}\n"
        f"Attached: {pdf_name}\n"
        "—Sent from an automated reporting pipeline"
    )

    msg.add_alternative(
        f"""
        <p>Hi Technical Support Team,</p>
        <p>A user just ran the ordering pipeline successfully, however some beverages were not mapped to a sub-report-key / vendor.</p>
        <p>The BEV orders were still sent; however, these unassigned beverages were included on their own order under <strong>BEV - UNASSIGNED</strong>.</p><p></p>
        <p>The default recipient of the pipeline run is CC'd on this email for visibility and communication purposes.</p>
        <p>To fix this error for future runs, please review the Unassigned Beverages Report and add those Scan Codes to the Mapping File.</p>
        <p>Please reply to this email once the Mapping File is updated so that the user knows they can rerun the pipeline if needed.</p><p></p>
        <p><a href="{mapping_link}">Open Beverage Mapping File</a></p>
        <p><a href="{sheet_link}">Open Unassigned Beverages Google Sheet</a></p>
        <p>Attached: {pdf_name}</p>
        <p>—Sent from an automated reporting pipeline</p>
        """,
        subtype="html",
    )

    msg.add_attachment(
        pdf_bytes,
        maintype="application",
        subtype="pdf",
        filename=pdf_name,
    )

    return send_email(gmail_svc, sender, msg)

def email_large_case_alert_report(
    gmail_svc,
    sender: str,
    to_list,
    cc_list,
    ts: str,
    location: str,
    threshold: int,
    pdf_name: str,
    pdf_bytes: bytes,
    sheet_link: str,
):
    """
    Send a soft-alert email when FULL order lines exceed a case threshold.

    This is a NON-blocking informational alert intended for
    technical review, not end users.
    """

    msg = EmailMessage()

    msg["Subject"] = (
        f"Soft Alert – High Case Quantities – {location} – {ts}"
    )
    msg["From"] = sender
    msg["To"] = ", ".join(to_list)

    if cc_list:
        msg["Cc"] = ", ".join(cc_list)

    msg.set_content(
        f"""Hi Technical Support Team,

        This is a soft alert generated by the ordering pipeline.

        One or more items exceeded the configured cases-to-order
        threshold of {threshold}. This usually is a result of Units Per Case being set to 1 in the Vendor Price Book by mistake.

        This alert does NOT block the pipeline and the order reports were still sent.

        Google Sheet:
        {sheet_link}

        Attached: {pdf_name}

        — Sent from an automated reporting pipeline
        """
    )

    msg.add_alternative(
        f"""
        <p><strong>Soft Alert – High Case Quantities</strong></p>
        <p>One or more items exceeded the configured cases-to-order threshold of {threshold}. This usually is a result of Units Per Case being set to 1 in the Vendor Price Book by mistake.</p>
        <p>This alert does NOT block the pipeline and the order reports were still sent.</p>
        <p><a href="{sheet_link}">Open Alert Sheet in Google Sheets</a></p>
        <p>Attached: {pdf_name}</p>
        <p>— Sent from an automated reporting pipeline</p>
        """,
        subtype="html",
    )

    msg.add_attachment(
        pdf_bytes,
        maintype="application",
        subtype="pdf",
        filename=pdf_name,
    )

    return send_email(gmail_svc, sender, msg)
```

---
### file: core_functional_modules/google_client.py

```python
"""
google_client
======================================
This module centralizes Google OAuth 2.0 sign‑in for local Python applications,
supporting both:

1) **Classic CLI flow** – prints an authorization URL to the console and accepts
   a pasted redirect URL or auth code (for headless shells, remote SSH, or when
   opening a browser is impractical).

2) **Streamlit-/Desktop-friendly local-server flow** – opens the user's default
   browser and spins up a temporary HTTP listener on ``127.0.0.1`` to complete
   the OAuth redirect without any console copy/paste.

It also provides small helpers to manage a persistent token cache
(``token.json``), refresh expired tokens when possible, and construct Google API
service clients (Sheets, Drive, Gmail) using the official ``google-api-python-client``.

-------------------------------------------------------------------------------
Key Features
-------------------------------------------------------------------------------
- **Token cache**: Reads/writes ``token.json`` to persist credentials between runs.
  Includes best-effort refresh of expired tokens when a refresh token exists.
- **Two auth paths**:
  - *CLI/manual path*: URL is printed; user pastes back the full redirect URL or
    just the ``code`` parameter.
  - *Local server/one-click path*: Automatically opens browser and listens on a
    local port (tries an OS-chosen free port first, then a configured fallback).
- **Graceful fallbacks**: If automated browser auth fails, raises a descriptive
  error suggesting the manual method.
- **Service builders**: Convenience helpers to create Sheets v4, Drive v3, and
  Gmail v1 service clients with the provided credentials.

-------------------------------------------------------------------------------
Files Used
-------------------------------------------------------------------------------
- ``credentials.json`` (required):
  The OAuth 2.0 client secrets file downloaded from Google Cloud Console.

- ``token.json`` (optional, auto-created):
  The persisted user credentials (access/refresh tokens). If present and valid,
  it is reused to avoid re-authentication. If expired but refreshable, it is
  refreshed automatically and re-written.

-------------------------------------------------------------------------------
Function Overview
-------------------------------------------------------------------------------
- ``clear_token()``:
    Deletes ``token.json`` if present (best-effort). Useful to force a
    re-authentication scenario.

- ``load_valid_token(scopes) -> Optional[Credentials]``:
    Loads credentials from ``token.json`` for the given scopes. If expired but
    refreshable, refreshes and persists the updated token. Returns a valid
    ``Credentials`` or ``None``.

- ``get_credentials(scopes, redirect_port, force_reauth=False) -> Credentials``:
    **CLI-friendly** method. If no valid token exists, prints an auth URL and
    prompts for a pasted redirect URL or code. Persists the resulting token to
    ``token.json``.

- ``login_via_local_server(scopes, redirect_port) -> Credentials``:
    **Streamlit-/desktop-friendly** one-click OAuth that opens a browser and
    listens on ``127.0.0.1``. Tries an OS-chosen free port first (``port=0``),
    then the provided ``redirect_port``. Uses a 120s timeout for safety.

- ``start_oauth(scopes, redirect_port) -> (InstalledAppFlow, auth_url)``:
    Starts the manual flow by creating an ``InstalledAppFlow`` with a configured
    redirect URI and returns the authorization URL to display in your own UI.

- ``finish_oauth(flow, pasted) -> Credentials``:
    Completes the manual flow using the pasted redirect URL (or raw ``code``),
    fetches tokens, writes ``token.json``, and returns ``Credentials``.

- ``_service(api, version, creds)``:
    Internal helper to construct a Google API service for the given
    ``api``/``version`` using the supplied ``Credentials``.

- ``services(creds, _http_timeout_seconds)``:
    Convenience function returning a tuple of ready-to-use clients:
    ``(sheets, drive, gmail)``. The ``_http_timeout_seconds`` parameter is
    currently reserved for future use.

-------------------------------------------------------------------------------
Error Handling & Edge Cases
-------------------------------------------------------------------------------
- If ``credentials.json`` is missing, a ``FileNotFoundError`` is raised early.
- Token refresh failures fall back to a fresh login.
- The local-server path uses a 120-second timeout to avoid hanging the process.
- If both automatic local-server attempts fail, a ``RuntimeError`` is raised
  advising the manual copy/paste method with detailed error messages from both
  attempts.
-------------------------------------------------------------------------------
Maintainer Tips
-------------------------------------------------------------------------------
- If you add new Google APIs, extend ``services(...)`` or call ``_service(...)``
  directly with the desired API name/version.
- Consider surfacing the timeout and host/port as user configuration if your app
  needs more control in diverse environments.

"""

from __future__ import annotations
import os
from urllib.parse import urlparse, parse_qs

from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build


# ---------- Token helpers ----------

def clear_token():
    """Delete token.json if present."""
    try:
        if os.path.exists("token.json"):
            os.remove("token.json")
    except Exception:
        pass


def load_valid_token(scopes):
    """
    Try to load token.json. If expired but refreshable, refresh it and persist.
    Returns valid Credentials or None.
    """
    if not os.path.exists("token.json"):
        return None
    try:
        creds = Credentials.from_authorized_user_file("token.json", scopes)
    except Exception:
        return None

    if creds and creds.valid:
        return creds

    if creds and creds.expired and creds.refresh_token:
        try:
            creds.refresh(Request())
            with open("token.json", "w") as f:
                f.write(creds.to_json())
            return creds
        except Exception:
            return None

    return None


# ---------- Classic CLI path (kept for completeness) ----------

def get_credentials(scopes, redirect_port: int, force_reauth: bool = False) -> Credentials:
    """
    CLI-friendly: prints URL and waits for input() if token is missing/invalid.
    The Streamlit UI uses the in-UI functions below instead.
    """
    if force_reauth:
        clear_token()

    creds = load_valid_token(scopes)
    if creds:
        return creds

    if not os.path.exists("credentials.json"):
        raise FileNotFoundError("Missing credentials.json in working directory")

    flow = InstalledAppFlow.from_client_secrets_file("credentials.json", scopes)
    flow.redirect_uri = f"http://127.0.0.1:{redirect_port}/"
    auth_url, _ = flow.authorization_url(
        access_type="offline",
        include_granted_scopes="true",
        prompt="consent",
    )
    print("Open this URL and complete the login:\n", auth_url)
    pasted = input("Paste full redirect URL or auth code here: ").strip()
    code = pasted
    if pasted.startswith("http"):
        qs = parse_qs(urlparse(pasted).query)
        if "code" in qs:
            code = qs["code"][0]
    flow.fetch_token(code=code)
    creds = flow.credentials
    with open("token.json", "w") as f:
        f.write(creds.to_json())
    return creds


# ---------- Streamlit-friendly OAuth (no console) ----------

# favtrip/google_client.py

def login_via_local_server(scopes, redirect_port: int) -> Credentials:
    """
    One-click OAuth: open browser and listen on 127.0.0.1.
    Tries OS-chosen port first, then the configured port.
    Uses a timeout to avoid hanging indefinitely.
    NOTE: No optional text parameters are passed, for compatibility with older google-auth-oauthlib.
    """
    if not os.path.exists("credentials.json"):
        raise FileNotFoundError("Missing credentials.json in working directory")

    # Attempt 1: OS-chosen free port (port=0)
    flow = InstalledAppFlow.from_client_secrets_file("credentials.json", scopes)
    try:
        creds = flow.run_local_server(
            host="127.0.0.1",
            port=0,                 # let OS choose a free port
            open_browser=True,
            timeout_seconds=120,    # bail out after 2 minutes
        )
        with open("token.json", "w") as f:
            f.write(creds.to_json())
        return creds
    except Exception as first_err:
        # Attempt 2: user-configured port (from .env)
        flow = InstalledAppFlow.from_client_secrets_file("credentials.json", scopes)
        try:
            creds = flow.run_local_server(
                host="127.0.0.1",
                port=int(redirect_port),
                open_browser=True,
                timeout_seconds=120,
            )
            with open("token.json", "w") as f:
                f.write(creds.to_json())
            return creds
        except Exception as second_err:
            raise RuntimeError(
                "Automatic browser auth failed both on a random port and on your configured REDIRECT_PORT. "
                "Please use the manual method (copy/paste URL). "
                f"Details: first={first_err}; second={second_err}"
            )


def start_oauth(scopes, redirect_port: int):
    """
    Manual fallback: returns (flow, auth_url) for paste-based completion.
    """
    if not os.path.exists("credentials.json"):
        raise FileNotFoundError("Missing credentials.json in working directory")
    flow = InstalledAppFlow.from_client_secrets_file("credentials.json", scopes)
    flow.redirect_uri = f"http://127.0.0.1:{redirect_port}/"
    auth_url, _ = flow.authorization_url(
        access_type="offline",
        include_granted_scopes="true",
        prompt="consent",
    )
    return flow, auth_url


def finish_oauth(flow: InstalledAppFlow, pasted: str) -> Credentials:
    """
    Manual fallback: accepts the pasted redirect URL or the code; returns Credentials and writes token.json.
    """
    code = pasted.strip()
    if pasted.startswith("http"):
        qs = parse_qs(urlparse(pasted).query)
        if "code" in qs:
            code = qs["code"][0]
    flow.fetch_token(code=code)
    creds = flow.credentials
    with open("token.json", "w") as f:
        f.write(creds.to_json())
    return creds


# ---------- Google services ----------

def _service(api: str, version: str, creds: Credentials):
    # Pass credentials directly (no google_auth_httplib2 dependency)
    return build(api, version, credentials=creds, cache_discovery=False)


def services(creds: Credentials, _http_timeout_seconds: int):
    sheets = _service("sheets", "v4", creds)
    drive = _service("drive", "v3", creds)
    gmail = _service("gmail", "v1", creds)
    return sheets, drive, gmail

```

---
### file: core_functional_modules/logger.py

```python
"""
logger
======================================
This module provides two dataclasses—`LogEvent` and `StatusLogger`—to record simple,
human-readable status messages during a process or script run. It is designed to be:

- **Simple**: minimal API (`info`, `warn`, `error`) and a small in-memory log.
- **Immediate**: console prints occur synchronously; file writes are line-buffered and flushed.
- **Fail-open**: if a log file cannot be opened or written, logging proceeds to console and memory.
- **Portable**: standard library only (dataclasses, datetime, typing).

-------------------------------------------------------------------------------
Data Model
-------------------------------------------------------------------------------
- LogEvent
    - ts (datetime.datetime): Timestamp captured via `datetime.now()` when the event is recorded.
      Note: this is a **naive** datetime in local time.
    - level (str): Log level label (e.g., "INFO", "WARN", "ERROR").
    - message (str): The event text.

- StatusLogger
    - events (list[LogEvent]): In-memory event history in append order.
    - print_to_console (bool): If True (default), each log line is printed to stdout.
    - file_path (str | None): If set, lines are also written to this file. If `None`, file logging
      is disabled. Default is "last_run.log".
    - overwrite (bool): If True (default), the log file is opened in write mode on instantiation;
      otherwise it is appended to.

-------------------------------------------------------------------------------
Output Format
-------------------------------------------------------------------------------
- Console/file lines: `[YYYY-MM-DD HH:MM:SS] LEVEL: message`
- `as_text()`:         `[HH:MM:SS] LEVEL: message` per line (no date, suitable for compact display)
- `last_line()`:       Returns the most recent line in `as_text()` format, or `"Starting…"` if empty.

-------------------------------------------------------------------------------
Behavior & Guarantees
-------------------------------------------------------------------------------
- **File handling**: On initialization, if `file_path` is provided, the file is opened once in
  line-buffered text mode (`buffering=1`) and UTF-8 encoding. If opening fails, the logger
  continues without a file handle.
- **Atomicity**: Each `_emit` call attempts to write a single line and then flush. Any file write
  errors are swallowed; console output and in-memory storage are unaffected.
- **Timestamps**: Timestamps are captured at call time (`datetime.now()`), local time, naive datetimes.
- **Memory growth**: All events are retained in `events`; for long-running processes, consider
  pruning or exporting periodically.
- **Thread-safety**: Not thread-safe. If you need concurrent logging, protect calls with a lock or
  adapt the implementation for multi-thread/process usage.
- **No rotation**: No file rotation or size limiting. Use external tools or extend as needed.
"""

from dataclasses import dataclass, field
from datetime import datetime
from typing import List, Optional

@dataclass
class LogEvent:
    ts: datetime
    level: str
    message: str

@dataclass
class StatusLogger:
    events: List[LogEvent] = field(default_factory=list)
    print_to_console: bool = True
    file_path: Optional[str] = "last_run.log"
    overwrite: bool = True

    def __post_init__(self):
        # Prepare the file on first use
        self._fh = None
        if self.file_path:
            mode = "w" if self.overwrite else "a"
            try:
                self._fh = open(self.file_path, mode, encoding="utf-8", buffering=1)  # line-buffered
            except Exception:
                # If we cannot open a file, we keep running without file logging
                self._fh = None

    def _emit(self, line: str):
        if self.print_to_console:
            print(line)
        if self._fh:
            try:
                self._fh.write(line + "\n")
                self._fh.flush()  # ensure immediate persistence
            except Exception:
                pass

    def _log(self, level: str, message: str):
        evt = LogEvent(datetime.now(), level, message)
        self.events.append(evt)
        self._emit(f"[{evt.ts:%Y-%m-%d %H:%M:%S}] {level}: {message}")

    def info(self, message: str):
        self._log("INFO", message)

    def warn(self, message: str):
        self._log("WARN", message)

    def error(self, message: str):
        self._log("ERROR", message)

    def as_text(self) -> str:
        return "\n".join(f"[{e.ts:%H:%M:%S}] {e.level}: {e.message}" for e in self.events)

    def last_line(self) -> str:
        if not self.events:
            return "Starting…"
        e = self.events[-1]
        return f"[{e.ts:%H:%M:%S}] {e.level}: {e.message}"

    def close(self):
        try:
            if self._fh:
                self._fh.close()
        except Exception:
            pass

```

---
### file: core_functional_modules/pipeline.py

```python
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
    add_or_replace_sheet, put_values_2d, _force_column_as_text, delete_row_indices, delete_rows_range, copy_sheet_to_another_spreadsheet, autoresize_columns, export_sheet
)
from .drive_utils import find_latest_sheet, upload_to_drive, _rfc3339, trash_file, cleanup_folder_by_age, find_sheet_by_name, copy_file_to_folder, rename_file, get_or_create_subfolder
from .gmail_utils import send_email, email_manager_report, email_order_report, email_error_report, email_bev_error_report, email_large_case_alert_report

CSV_MIME = "text/csv"


def clean_tag(s: str | None) -> str | None:
    if s is None:
        return None
    s = str(s).strip()
    if not s:
        return None
    return re.sub(r"[^A-Za-z0-9._-]+", "-", s).strip("-")



import requests
from io import BytesIO
from openpyxl import Workbook


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


def filter_master_csv_to_ran(master_csv_bytes, cfg):
    text = master_csv_bytes.decode("utf-8-sig", errors="replace")
    reader = csv.reader(io.StringIO(text))
    rows = list(reader)

    if not rows:
        return master_csv_bytes  # nothing to do

    headers = [h.strip() for h in rows[0]]
    lower_idx = {h.lower(): i for i, h in enumerate(headers)}

    report_idx = lower_idx.get("report_key")
    sub_idx = lower_idx.get("sub_report_key")

    if report_idx is None:
        raise RuntimeError("Master CSV missing Report_Key column")

    output = io.StringIO()
    writer = csv.writer(output, lineterminator="\n")
    writer.writerow(headers)

    for r in rows[1:]:
        key = (r[report_idx] if report_idx < len(r) else "").strip().upper()
        sub = None
        if sub_idx is not None:
            sub = (r[sub_idx] if sub_idx < len(r) else "").strip().upper() or None

        if should_run(cfg, key, sub):
            writer.writerow(r)

    return output.getvalue().encode("utf-8-sig")

def sort_master_csv(csv_bytes: bytes) -> bytes:
    df = pandas.read_csv(io.BytesIO(csv_bytes))

    # Normalize column names
    cols = {c.lower(): c for c in df.columns}

    report_col = cols.get("report_key")
    sub_col = cols.get("sub_report_key")
    cases_col = cols.get("cases to order")

    if not report_col:
        raise RuntimeError("Report_Key column not found for sorting")

    # Fill missing sub-keys so they sort last
    if sub_col:
        df[sub_col] = df[sub_col].fillna("ZZZ")

    # Ensure numeric sorting for cases
    if cases_col:
        df[cases_col] = pandas.to_numeric(df[cases_col], errors="coerce").fillna(0)

    sort_cols = [report_col]
    sort_ascending = [True]

    if sub_col:
        sort_cols.append(sub_col)
        sort_ascending.append(True)

    if cases_col:
        sort_cols.append(cases_col)
        sort_ascending.append(False)  # DESCENDING

    df = df.sort_values(
        by=sort_cols,
        ascending=sort_ascending,
        kind="mergesort"  # stable sort
    )

    # Restore blanks if we filled them
    if sub_col:
        df[sub_col] = df[sub_col].replace("ZZZ", "")

    out = io.StringIO()
    df.to_csv(out, index=False)
    return out.getvalue().encode("utf-8-sig")


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

    
    ran_report_keys = set()
    ran_sub_report_keys = set()
    ran_reports = {}


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

    # Step 4: Manager Report PDF
    if logger:
        logger.info("Exporting Manager Report (PDF)…")
    pdf_bytes = export_sheet(creds, calc_ss_id, cfg.GID_MANAGER_PDF, "pdf", True)
    pdf_name = f"Manager_Report_{ts}_{location}.pdf"
    uploaded_pdf = upload_to_drive(drive_svc, pdf_bytes, pdf_name, "application/pdf", cfg.MANAGER_REPORT_FOLDER_ID, to_sheet=False)
    manager_link = uploaded_pdf.get("webViewLink")
    if logger:
        logger.info(f"Uploaded Manager PDF: {manager_link}")

    # Step 5: Master Order CSV
    if logger:
        logger.info("Exporting Master Order (CSV)…")
    master_csv_bytes = export_sheet(creds, calc_ss_id, cfg.GID_ORDER_CSV, "csv")
    master_csv_bytes = filter_master_csv_to_ran(master_csv_bytes, cfg)
    master_csv_bytes = sort_master_csv(master_csv_bytes)

    # Step 6: Error Report CSV, Upload, Export PDF
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
        
        autoresize_columns(sheets_svc, err_file_id, err_gid)

        err_pdf = export_sheet(creds, err_file_id, err_gid, "pdf", False)
        err_pdf_name = f"Error_Report_{ts}.pdf"

        if logger:
            logger.info(f"Uploaded filtered Error Sheet: {err_link}")

        # Step 6.1: Send Error Report if Needed

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




    # Step 7: Full order upload (CSV) and export (PDF)
    full_csv_name = f"Order_Report_FULL_{location}_{ts}.csv"
    full_created = upload_to_drive(drive_svc, master_csv_bytes, full_csv_name, CSV_MIME, cfg.ORDER_REPORT_FOLDER_ID, to_sheet=True)
    full_file_id = full_created["id"]
    full_link = full_created.get('webViewLink')
    full_gid = first_gid(sheets_svc, full_file_id)
    autoresize_columns(sheets_svc, full_file_id, full_gid)
    full_pdf = export_sheet(creds, full_file_id, full_gid, "pdf", False)
    full_pdf_name = f"Order_Report_FULL_{location}_{ts}.pdf"
    if logger:
        logger.info(f"Uploaded FULL sheet: {full_created.get('webViewLink')}")


    #Step 7.1: Large Case Alert
    try:
        if cfg.SOFT_CASES_ALERT_ENABLED:
            if logger:
                logger.info(
                    f"Checking FULL order for case quantities > "
                    f"{cfg.SOFT_CASES_ALERT_THRESHOLD}"
                )

            # Load FULL order CSV into DataFrame
            df = pandas.read_csv(io.BytesIO(master_csv_bytes))

            # Normalize column lookup (case-insensitive)
            lower_cols = {c.lower(): c for c in df.columns}
            cases_col = lower_cols.get("cases to order")

            if not cases_col:
                if logger:
                    logger.warn(
                        "Large case alert skipped — 'Cases to Order' "
                        "column not found in FULL order CSV"
                    )
            else:
                # Ensure numeric comparison
                df[cases_col] = pandas.to_numeric(
                    df[cases_col], errors="coerce"
                ).fillna(0)

                flagged = df[
                    df[cases_col] > cfg.SOFT_CASES_ALERT_THRESHOLD
                ]

                if flagged.empty:
                    if logger:
                        logger.info(
                            "No FULL order rows exceed case threshold"
                        )
                else:
                    if logger:
                        logger.warn(
                            f"Soft alert triggered: {len(flagged)} "
                            f"rows exceed case threshold"
                        )

                    # --------------------------------------------------
                    # Create filtered CSV (only flagged rows)
                    # --------------------------------------------------
                    buf = io.StringIO()
                    flagged.to_csv(buf, index=False)
                    alert_csv_bytes = buf.getvalue().encode("utf-8-sig")

                    alert_csv_name = (
                        f"Large_Case_Alert_{location}_{ts}.csv"
                    )

                    # Upload alert CSV → Google Sheet
                    created = upload_to_drive(
                        drive_svc,
                        alert_csv_bytes,
                        alert_csv_name,
                        CSV_MIME,
                        cfg.ERROR_REPORT_FOLDER_ID,
                        to_sheet=True,
                    )

                    alert_sheet_id = created["id"]
                    alert_sheet_link = created.get("webViewLink")

                    alert_gid = first_gid(sheets_svc, alert_sheet_id)
                    autoresize_columns(
                        sheets_svc,
                        alert_sheet_id,
                        alert_gid,
                    )

                    # Export alert PDF
                    alert_pdf_bytes = export_sheet(
                        creds,
                        alert_sheet_id,
                        alert_gid,
                        "pdf",
                        False,
                    )
                    alert_pdf_name = (
                        f"Large_Case_Alert_{location}_{ts}.pdf"
                    )

                    # --------------------------------------------------
                    # Resolve recipients (technical first)
                    # --------------------------------------------------
                    to_list = _fallback_recipients(
                        "LARGE CASE ALERT",
                        cfg.ERROR_RECIPIENTS,
                        cfg.TO_RECIPIENTS,
                        cfg.DEFAULT_ORDER_RECIPIENTS,
                    )

                    cc_list = [
                        e for e in _clean_emails(cfg.CC_RECIPIENTS)
                        if e not in to_list
                    ]

                    # --------------------------------------------------
                    # Send email (SOFT ALERT)
                    # --------------------------------------------------
                    email_large_case_alert_report(
                        gmail_svc=gmail_svc,
                        sender="me",
                        to_list=to_list,
                        cc_list=cc_list,
                        ts=ts,
                        location=location,
                        threshold=cfg.SOFT_CASES_ALERT_THRESHOLD,
                        pdf_name=alert_pdf_name,
                        pdf_bytes=alert_pdf_bytes,
                        sheet_link=alert_sheet_link,
                    )

                    if logger:
                        logger.info(
                            "Large case quantity alert email sent"
                        )

    except Exception as e:
        # 🔐 Soft failure only — never block pipeline
        if logger:
            logger.warn(
                f"Large case quantity alert failed (soft): {e}"
            )

    # Step 8: Create per-report-key outputs (CSV) and email

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
    

    bev_order_sent = False
    
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
        autoresize_columns(sheets_svc, file_id, gid)
        pdf = export_sheet(creds, file_id, gid, "pdf", False)
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

        
        email_tag_parts = [tag]
        if sub_tag:
            email_tag_parts.append(sub_tag)

        email_tag = " - ".join(email_tag_parts)



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

        
        if key.upper() == "BEV":
            bev_order_sent = True

    
        if logger:
            logger.info(f"Emailed {store} - {email_tag} to {recipients}")
    
    # Step 9: Unassigned Beverages Report (Soft Error)
    try:
        # Must have successfully sent a BEV order
        if not bev_order_sent:
            if logger:
                logger.info("Skipping Unassigned Beverages Report — no BEV order was sent.")
        else:
            if logger:
                logger.info("Exporting Unassigned Beverages Report (CSV)…")

            unassigned_csv_bytes = export_sheet(
                creds,
                calc_ss_id,
                cfg.GID_BEV_ERRORS,
                "csv"
            )

            # Inspect CSV to ensure it actually has data
            text = unassigned_csv_bytes.decode("utf-8-sig", errors="replace")
            reader = csv.reader(io.StringIO(text))
            rows = list(reader)

            if not rows or len(rows) <= 1:
                if logger:
                    logger.info("No unassigned beverages found — report will not be sent.")
            else:
                headers = [h.strip() for h in rows[0]]
                lower_idx = {h.lower(): i for i, h in enumerate(headers)}

                if "report_key" not in lower_idx:
                    raise RuntimeError("Unassigned BEV report missing Report_Key column")

                if "sub_report_key" not in lower_idx:
                    raise RuntimeError("Unassigned BEV report missing Sub_Report_Key column")

                report_idx = lower_idx["report_key"]
                sub_idx = lower_idx["sub_report_key"]

                # Filter to BEV + BEV_UNASSIGNED only
                unassigned_rows = [
                    r for r in rows[1:]
                    if r[report_idx].strip().upper() == "BEV"
                    and r[sub_idx].strip().upper() == "UNASSIGNED"
                ]

                if not unassigned_rows:
                    if logger:
                        logger.info("No BEV_UNASSIGNED rows found — skipping email.")
                else:
                    # Upload filtered sheet
                    output = io.StringIO()
                    writer = csv.writer(output)
                    writer.writerow(rows[0])
                    writer.writerows(unassigned_rows)

                    filtered_bytes = output.getvalue().encode("utf-8-sig")

                    csv_name = f"Unassigned_Beverages_Report_{ts}.csv"

                    created = upload_to_drive(
                        drive_svc,
                        filtered_bytes,
                        csv_name,
                        CSV_MIME,
                        cfg.ERROR_REPORT_FOLDER_ID,
                        to_sheet=True
                    )

                    sheet_id = created["id"]
                    sheet_link = created.get("webViewLink")
                    gid = first_gid(sheets_svc, sheet_id)

                    autoresize_columns(sheets_svc, sheet_id, gid)
                    pdf_bytes = export_sheet(creds, sheet_id, gid, "pdf", False)
                    pdf_name = f"Unassigned_Beverages_Report_{ts}.pdf"

                    # Resolve recipients
                    to_list = _fallback_recipients(
                        "UNASSIGNED BEVERAGES REPORT",
                        cfg.ERROR_RECIPIENTS,
                        cfg.TO_RECIPIENTS,
                        cfg.DEFAULT_ORDER_RECIPIENTS,
                    )

                    cc_list = list(dict.fromkeys(
                        set(_clean_emails(cfg.TO_RECIPIENTS))
                        | set(_clean_emails(cfg.CC_RECIPIENTS))
                        - set(to_list)
                    ))

                    email_bev_error_report(
                        gmail_svc=gmail_svc,
                        sender="me",
                        to_list=to_list,
                        cc_list=cc_list,
                        ts=ts,
                        pdf_name=pdf_name,
                        pdf_bytes=pdf_bytes,
                        sheet_link=sheet_link,
                        mapping_link=cfg.BEV_MAPPING_LINK,
                    )

                    if logger:
                        logger.info("Unassigned Beverages Report email sent")

    except Exception as e:
        # Soft error — log and continue
        if logger:
            logger.warn(f"Unassigned Beverages Report failed (soft): {e}")
        
    # Step 10: Send Manager Report (guarded by cfg.EMAIL_MANAGER_REPORT)
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

    

    # Step 11: Send Full Order if needed
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

    # Step 12: File Cleanup

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

```

---
### file: core_functional_modules/pipeline_bus.py

```python
# pipeline_bus.py
import queue

_PIPELINE_QUEUE = None

def get_pipeline_queue():
    global _PIPELINE_QUEUE
    if _PIPELINE_QUEUE is None:
        _PIPELINE_QUEUE = queue.Queue()
    return _PIPELINE_QUEUE
```

---
### file: core_functional_modules/sheets_utils.py

```python
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
import requests


def list_sheets(svc, spreadsheet_id: str) -> List[Dict[str, Any]]:
    return svc.spreadsheets().get(spreadsheetId=spreadsheet_id).execute().get("sheets", [])


def get_sheet(sheets, title: str):
    for s in sheets:
        if s["properties"]["title"] == title:
            return s["properties"]
    return None


def export_sheet(creds, spreadsheet_id: str, gid: str | int, fmt: str, portrait: bool = True,) -> bytes:
    params = {
        "format": fmt,
        "gid": gid,
    }

    # PDF-only layout options
    if fmt.lower() == "pdf":
        params.update({
            "portrait": "true" if portrait else "false",
            "fitw": "true",   # fit to width
        })

    # Build query string
    query = "&".join(f"{k}={v}" for k, v in params.items())

    url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export?{query}"

    headers = {"Authorization": f"Bearer {creds.token}"}
    r = requests.get(url, headers=headers)
    r.raise_for_status()
    return r.content



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

def autoresize_columns(sheets_svc, spreadsheet_id, sheet_id):
    sheets_svc.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={
            "requests": [{
                "autoResizeDimensions": {
                    "dimensions": {
                        "sheetId": sheet_id,
                        "dimension": "COLUMNS",
                        "startIndex": 0,
                        "endIndex": 26  # A–Z (adjust if needed)
                    }
                }
            }]
        }
    ).execute()

```

---
### file: documentation/PROJECT_SNAPSHOT_CODEBUNDLE.md

```markdown
# Project Snapshot (CodeBundle)

                This single Markdown file contains a **self-contained snapshot** of your project so another AI/engineer can review or modify it without needing the original folder.

                **How to use this file with an AI**
                1. Upload or paste this file as a single attachment.
                2. Ask for changes; the AI can reference specific `file:` sections below.
                3. Copy updated blocks back into the corresponding files in your project.

                > Notes: secrets like `token.json` are intentionally excluded. Virtual envs and build artifacts are omitted to keep this readable.

                ## Directory tree (filtered)
                .env [skipped: secret]
.env.example
.gitignore
FavTripPipeline.spec
FavTripPipelineUI.spec
_user_interface_.py
cli.py
credentials.json [skipped: secret]
last_run.log
launcher_streamlit.py
requirements.txt
setup_py2app.py
token.json [skipped: secret]
web_url_credentials.json [skipped: secret]
  core_functional_modules/
    __init__.py
    config.py
    config_store.py
    drive_utils.py
    gmail_utils.py
    google_client.py
    logger.py
    pipeline.py
    pipeline_bus.py
    sheets_utils.py
  documentation/
    PROJECT_SNAPSHOT_CODEBUNDLE.md  [skipped: too large]
    README.md
    generate_code_bundle.py
    git_workflow.txt
    requirements.txt
  __dev_input_sales_files/
    Advanced Testing Sales Report - 1 Week.xlsx
    Advanced Testing Sales Report - 2 Weeks 2 Stores.xlsx  [skipped: too large]
    Advanced Testing Sales Report - 2 Weeks.xlsx  [skipped: too large]
    Advanced Testing Sales Report - Bad End.xlsx  [skipped: too large]
    Advanced Testing Sales Report - Bad Start.xlsx  [skipped: too large]
    Advanced Testing Sales Report - No Errors.xlsx  [skipped: too large]
    Testing Sales Report - Week 2.xlsx  [skipped: too large]
    Testing Sales Report - Week 3.xlsx  [skipped: too large]
    Testing Sales Report - Week 4.xlsx  [skipped: too large]
    Testing Sales Report - Week 5.xlsx  [skipped: too large]
  __dev_input_vendor_file/
    Vendors Price Book.xlsx
  __executable/
    run_windows.bat
---
### file: .env.example

```
# --- Required IDs ---
CALC_SPREADSHEET_ID=
INCOMING_FOLDER_ID=
MANAGER_REPORT_FOLDER_ID=
ORDER_REPORT_FOLDER_ID=

# --- Optional IDs / settings ---
GID_MANAGER_PDF=1921812573
GID_ORDER_CSV=1875928148
LOCATION_SHEET_TITLE=REFR: Values
LOCATION_NAMED_RANGE=_locations

TIMESTAMP_TZ=America/Chicago
TIMESTAMP_FMT=%Y-%m-%d-%I-%M-%p

# Recipients
TO_RECIPIENTS=
CC_RECIPIENTS=
DEFAULT_ORDER_RECIPIENTS=

# Report keys
USE_ALL_REPORT_KEYS=false
REPORT_KEY_RUN_LIST=GROCERY,COFFEE
# JSON mapping: {"GROCERY":["a@b.com","c@d.com"],"COFFEE":["x@y.com"]}
REPORT_KEY_RECIPIENTS={}

INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL=false
SEND_SEPARATE_FULL_ORDER_EMAIL=true

# Google API scopes (normally leave as-is)
SCOPES=https://www.googleapis.com/auth/drive,https://www.googleapis.com/auth/spreadsheets,https://www.googleapis.com/auth/gmail.send

FORCE_REAUTH=false
REDIRECT_PORT=58285
HTTP_TIMEOUT_SECONDS=300

```

---
### file: .gitignore

```
*.env
*credentials.json
token.json
web_url_credentials.json

```

---
### file: FavTripPipeline.spec

```
# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['cli.py'],
    pathex=[],
    binaries=[],
    datas=[('.env', '.'), ('credentials.json', '.')],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='FavTripPipeline',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

```

---
### file: FavTripPipelineUI.spec

```
# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['launcher_streamlit.py'],
    pathex=[],
    binaries=[],
    datas=[('.env', '.'), ('credentials.json', '.'), ('ui_streamlit.py', '.'), ('favtrip', 'favtrip')],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='FavTripPipelineUI',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

```

---
### file: __executable/run_windows.bat

```bat
@echo off
setlocal
REM ---------------------------------------------------------------------------
REM Run Streamlit UI without a persistent console window.
REM Location: __executable\run_web_windows_silent.bat
REM Behavior: brief flash at launch, then only the browser tab remains.
REM ---------------------------------------------------------------------------

REM Move into the folder of this .bat
pushd "%~dp0"

REM Go to the project root (one level up from __executable)
cd ..

REM Choose Python: prefer venv's interpreter if present
set "PY_VENV=.\.venv\Scripts\python.exe"
set "PY="
if exist "%PY_VENV%" (
  set "PY=%PY_VENV%"
) else (
  for %%P in (python.exe py.exe) do (
    where %%P >nul 2>&1 && (set "PY=%%P" & goto :gotpy)
  )
)
:gotpy
if not defined PY (
  echo [Launcher] Python was not found. Install Python or create .\.venv and try again.
  popd
  exit /b 1
)

REM Streamlit prefs: ensure it opens the browser and stays local
set "STREAMLIT_SERVER_HEADLESS=false"
set "STREAMLIT_BROWSER_GATHER_USAGE_STATS=false"

REM Start Streamlit hidden and detach from this console (which then closes)
REM - We invoke PowerShell only to spawn the hidden child process.
start "" /MIN powershell -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -Command ^
  "Start-Process -FilePath '%PY%' -ArgumentList '-m','streamlit','run','ui_streamlit.py' -WindowStyle Hidden"

popd
exit /b 0
```

---
### file: _user_interface_.py

```python
"""
FavTrip Reporting Streamlit UI
================================

Overview
--------
This Streamlit app is the front-end for the FavTrip Reporting pipeline. It lets an authenticated
Google user upload a Modisoft "Live Items Report", configure who receives emails for each report
(or fallback recipients), tune advanced IDs/GIDs/time settings, and then orchestrate the
`core_functional_modules.pipeline.run_pipeline` execution while streaming status updates and a timer.

Design goals
------------
* **No local secrets in code**: OAuth client JSON is read from `st.secrets["GOOGLE_CREDENTIALS"]` and
  the app base URL from `st.secrets["APP_BASE_URL"]`. Optional `CONFIG_FILE_ID` pins a Drive JSON that
  stores editable defaults.
* **Constrained, clear UX**: A two-step flow (Upload ➜ Run). The Run button is enabled only after a
  successful upload, reducing accidental runs on stale input.
* **Robust OAuth (PKCE)**: Uses an explicit code verifier/challenge and encodes the verifier inside
  the `state` payload to remain stateless across redirects.
* **Operational safety**: Detects common mistakes (e.g., invalid email inputs, duplicate keys,
  wrong-week uploads) and surfaces warnings or blocks execution accordingly.

Key concepts
------------
* **Incoming file**: The uploaded Modisoft report (CSV/XLSX). It is pushed to a configured Google
  Drive folder and optionally converted to a Google Sheet for the downstream pipeline.
* **Report Keys**: Categories/tags used by the pipeline to partition output and email recipients.
  You can either process *all* keys present in the data or restrict to a comma-separated subset.
* **Per-Report-Key Recipients**: Optional overrides that map `(Store, Report Key)` pairs to recipient
  lists. Fallback recipients apply where no specific mapping exists.
* **Drive-backed defaults**: The app can persist your current UI settings to a JSON in Drive. Supplying
  `CONFIG_FILE_ID` in Streamlit secrets will cause subsequent sessions to update that exact file.

Security model
--------------
* OAuth scopes are supplied by `Config` and used to mint a user token saved locally as `token.json`.
* The app opens the Google consent screen in a **new tab**, and that tab becomes the main app after
  redirect. Tokens are not sent back to the opener page; they are stored only in the process serving
  the tab that completed OAuth.

Operational notes
-----------------
* If a run fails with the message "Please only upload 1 or 2 full weeks of data", the UI locks to
  prevent immediate re-runs. Use **Retry** to clear the lock and upload a correct file.
* Set **Offer full log download** (sidebar) to expose a download button for `last_run.log` after a run.
* The **green Run button** appears once a fresh upload succeeds, indicating the pipeline is ready to run.

Dependencies & integration points
---------------------------------
* `core_functional_modules.google_client`: token loading/clearing and service factories (Drive, Sheets, Gmail)
* `core_functional_modules.config`: the central configuration object. `Config.load()` merges defaults, secrets, and any
  Drive-stored overrides.
* `core_functional_modules.drive_utils.upload_to_drive`: uploads the incoming report and (optionally) converts to Sheet.
* `core_functional_modules.pipeline.run_pipeline`: the orchestrated processing step; returns an object with links and
  timing information used to render the result panel.

This file intentionally includes **documentation-only** additions (module docstring and inline comments)
without modifying the executable logic.
"""

# ------------------------------
# Quick-start for maintainers
# ------------------------------
# 1) Configure Streamlit secrets:
#    - APP_BASE_URL: The exact external base URL of your deployed app (with trailing slash normalization).
#    - GOOGLE_CREDENTIALS: A JSON string containing your OAuth client configuration.
#    - CONFIG_FILE_ID (optional): The Drive file ID for persisted UI defaults.
# 2) Grant your Google Cloud OAuth Client access to the app origin and redirect URI.
# 3) Run: `streamlit run ui_streamlit.py` (ensure the backend `favtrip` package is importable).
# 4) Upload a Modisoft report, adjust recipients and options, then click **Run Pipeline**.


import os
import time
import threading
import json
import base64
import hashlib
import secrets
import re
import queue
import uuid
import traceback

import streamlit as st
from streamlit.components.v1 import html
from streamlit.components.v1 import html as _html_listener
from streamlit_autorefresh import st_autorefresh

from google_auth_oauthlib.flow import Flow
from google.oauth2.credentials import Credentials

from core_functional_modules.google_client import load_valid_token, services, clear_token
from core_functional_modules.config_store import save_config_to_drive
from core_functional_modules.config import Config
from core_functional_modules.logger import StatusLogger
from core_functional_modules.pipeline import run_pipeline
from core_functional_modules.drive_utils import upload_to_drive, get_or_create_subfolder
from core_functional_modules.pipeline_bus import get_pipeline_queue


# =========================
# Constants & Simple Helpers
# =========================

EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")

err = None

UI_UPLOAD  = "UPLOAD"
UI_READY   = "READY"
UI_RUNNING = "RUNNING"
UI_RESULT = "RESULT"
UI_RESULT_ERROR = "RESULT_ERROR"     # run_pipeline failed
UI_UPLOAD_ERROR = "UPLOAD_ERROR"     # invalid input (1–2 weeks)


PIPE_STATUS_IDLE = "idle"
PIPE_STATUS_RUNNING = "running"
PIPE_STATUS_DONE = "done"
PIPE_STATUS_ERROR = "error"


class UIError(Exception):
    """
    Base class for errors that should show a friendly message
    plus optional technical details.
    """
    user_message: str
    title: str = "Error"

    def __init__(self, user_message: str, *, title: str | None = None):
        super().__init__(user_message)
        self.user_message = user_message
        if title:
            self.title = title


def _split_emails(csv_str: str):
    return [e.strip() for e in (csv_str or "").split(",") if e.strip()]


def _parse_emails(csv_str: str):
    return _split_emails(csv_str)


def _invalid_emails(csv_str: str):
    return [e for e in _parse_emails(csv_str) if not EMAIL_RE.match(e)]


def _analyze_rk_rows(rows):
    """
    Validate the 'Per-Report-Key Recipients' editor rows.
    Returns (issues: List[str], preview_lines: List[str], rk_map: Dict[str, List[str]])
    """
    issues, preview, rk_map = [], [], {}
    seen, dupes = set(), set()

    for idx, r in enumerate(rows or [], start=1):
        raw_key = (r.get("REPORT KEY (ALL CAPS)") or "").strip()
        emails_csv = r.get("Emails (comma)") or ""
        if not raw_key and not emails_csv:
            # allow a blank template row
            continue

        # uppercase flag
        if raw_key != raw_key.upper():
            issues.append(f"Row {idx}: key '{raw_key}' is not ALL CAPS.")

        # duplicate detection
        if raw_key:
            if raw_key in seen:
                dupes.add(raw_key)
            else:
                seen.add(raw_key)

        # email validation
        bads = _invalid_emails(emails_csv)
        if bads:
            issues.append(f"Row {idx}: invalid emails → {', '.join(bads)}")

        # mapping + preview
        if raw_key:
            emails = _parse_emails(emails_csv)
            if emails:
                rk_map[raw_key] = emails
            preview.append(f"{raw_key} → {', '.join(emails) if emails else emails_csv}")

    if dupes:
        issues.append(f"Duplicate keys detected: {', '.join(sorted(dupes))}")
    return issues, preview, rk_map


def _b64url(b: bytes) -> str:
    return base64.urlsafe_b64encode(b).rstrip(b"=").decode("ascii")


def _pkce_pair() -> tuple[str, str]:
    """
    Generate a high-entropy PKCE code_verifier and its S256 code_challenge.
    RFC 7636 requires 43–128 chars; this approach yields a URL-safe value.
    """
    verifier = _b64url(secrets.token_bytes(64))        # ~86 chars, URL-safe, no padding
    challenge = _b64url(hashlib.sha256(verifier.encode("ascii")).digest())
    return verifier, challenge


def _redirect_base() -> str:
    """
    Always return a non-empty redirect base that exactly matches your OAuth client's
    Authorized redirect URI. Prefer Secrets; normalize to one trailing slash.
    """
    base = (st.secrets.get("APP_BASE_URL", "") or "").strip()
    if not base:
        # Fallback to request (often available), still normalized
        try:
            base = (st.request.url_root or "").strip()
        except Exception:
            base = ""
    if not base:
        st.error("OAuth redirect base is not set. Define APP_BASE_URL in Secrets.")
        st.stop()
    return base.rstrip("/") + "/"


def _parse_state(state_b64: str) -> dict:
    # Add padding back for base64 decoding if needed
    padding = "=" * ((4 - len(state_b64) % 4) % 4)
    raw = base64.urlsafe_b64decode(state_b64 + padding)
    return json.loads(raw.decode("utf-8"))


def _infer_media_mime(name: str) -> str:
    n = (name or "").lower()
    if n.endswith(".csv"):
        return "text/csv"
    if n.endswith(".xlsx"):
        return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    return "application/octet-stream"


def _get_drive_service_or_raise(cfg):
    creds = load_valid_token(cfg.SCOPES)
    if not creds:
        raise RuntimeError("Google authorization required. Please sign in first.")
    _sheets, drive, _gmail = services(creds, cfg.HTTP_TIMEOUT_SECONDS)
    return drive


def _rerun():
    try:
        st.rerun()
    except AttributeError:
        st.experimental_rerun()

def reset_to_upload():
    st.session_state.sales_uploaded_ok = False
    st.session_state.vendor_uploaded_ok = False

    st.session_state.sales_selected_name = None
    st.session_state.vendor_selected_name = None

    st.session_state.reset_generation += 1

    st.session_state.sales_selection_generation = None
    st.session_state.vendor_selection_generation = None


    # 🔑 Increment the upload epoch
    st.session_state.upload_epoch += 1
    st.session_state.sales_selected_epoch = None
    st.session_state.vendor_selected_epoch = None

    st.session_state.running_ui_initialized = False
    st.session_state.uploader_version += 1
    st.session_state.ui_phase = UI_UPLOAD

    # Fully clear uploader widget state
    for k in list(st.session_state.keys()):
        if k.startswith("sales_upload_") or k.startswith("vendor_upload_"):
            st.session_state.pop(k, None)



def init_thread_state():
    if "pipeline_thread_started" not in st.session_state:
        st.session_state.pipeline_thread_started = False
    if "pipeline_done" not in st.session_state:
        st.session_state.pipeline_done = False
    if "pipeline_error" not in st.session_state:
        st.session_state.pipeline_error = None
    if "pipeline_thread" not in st.session_state:
        st.session_state.pipeline_thread = None

def init_pipeline_state():
    st.session_state.setdefault("pipe_status", PIPE_STATUS_IDLE)
    st.session_state.setdefault("pipe_result", None)
    st.session_state.setdefault("pipe_finished", False)
    st.session_state.setdefault("pipe_error", None)
    st.session_state.setdefault("pipe_run_id", None)

def reset_pipeline_state():
    # Thread control
    st.session_state.pipeline_thread_started = False
    st.session_state.pipeline_done = False
    st.session_state.pipeline_error = None
    st.session_state.pipeline_thread = None

    # Pipeline result & lifecycle
    st.session_state.pipe_status = PIPE_STATUS_IDLE
    st.session_state.pipe_finished = False
    st.session_state.pipe_result = None
    st.session_state.pipe_error = None
    st.session_state.pipe_run_id = None

    # Timer
    st.session_state._run_start_time = None


def start_run():
    st.session_state.pipe_run_id = str(uuid.uuid4())

    st.session_state.pipe_status = PIPE_STATUS_RUNNING
    st.session_state.pipeline_thread_started = True
    st.session_state.pipeline_done = False
    st.session_state.pipeline_error = None


    st.session_state.pipeline_refresh_key = f"pipeline_refresh_{time.time()}"



def _both_uploads_ok():
    epoch = st.session_state.upload_epoch

    return (
        st.session_state.sales_selected_epoch == epoch
        and st.session_state.vendor_selected_epoch == epoch
    )



def _validate_pipeline_result(result):
    required_attrs = ("location", "timestamp", "elapsed_seconds")
    return (
        result is not None
        and all(hasattr(result, attr) for attr in required_attrs)
    )



# =========================
# OAuth (Web / PKCE)
# =========================

def start_web_oauth(scopes):
    """
    Build an authorization URL that:
      - uses a stable redirect_uri (from Secrets)
      - uses explicit PKCE (S256)
      - embeds the code_verifier inside the state (base64url(JSON))
    """
    cfg = json.loads(st.secrets["GOOGLE_CREDENTIALS"])
    redirect = _redirect_base()

    # Explicit PKCE (stateless across redirect)
    code_verifier, code_challenge = _pkce_pair()

    # CSRF token + verifier encoded into state that Google will return unchanged.
    state_obj = {
        "csrf": _b64url(secrets.token_bytes(16)),
        "v": code_verifier,
        "r": redirect,
    }
    state_b64 = _b64url(json.dumps(state_obj).encode("utf-8"))

    flow = Flow.from_client_config(cfg, scopes=scopes, redirect_uri=redirect)
    auth_url, _ = flow.authorization_url(
        access_type="offline",
        include_granted_scopes="true",
        prompt="consent",
        state=state_b64,
        code_challenge=code_challenge,
        code_challenge_method="S256",
    )

    # Keep only minimal context; state carries the verifier.
    st.session_state["_oauth_redirect"] = redirect
    return auth_url


def finish_web_oauth(code: str, state_b64: str, scopes):
    """
    Recreate a Flow with the same redirect_uri and exchange code + code_verifier for tokens.
    (No UI side effects here.)
    """
    cfg = json.loads(st.secrets["GOOGLE_CREDENTIALS"])
    state_obj = _parse_state(state_b64)
    code_verifier = state_obj.get("v")
    redirect = state_obj.get("r") or st.session_state.get("_oauth_redirect") or _redirect_base()

    if not code_verifier:
        st.error("OAuth state did not include a PKCE code_verifier.")
        st.stop()

    flow = Flow.from_client_config(cfg, scopes=scopes, redirect_uri=redirect)
    flow.fetch_token(code=code, code_verifier=code_verifier)

    creds = flow.credentials
    with open("token.json", "w") as f:
        f.write(creds.to_json())
    return creds

def clean_tag(s: str) -> str:
    import re
    s = (s or "").strip()
    s = re.sub(r"[^A-Za-z0-9._-]+", "-", s)
    return s.strip("-") or "UNKNOWN"

# --- OAuth Redirect Handler (Web/PKCE only) ---


# =========================
# UI Sections
# =========================

def render_upload_card(cfg):
    with st.container(border=True):
        st.subheader("Upload Required Input Files")
        st.caption("Both Sales Data and Vendor Price Data are required. Sales Data must be 1 or 2 complete weeks.")

        up_col, _, upbtn_col = st.columns([4, 1, 1])

        current_gen = st.session_state.reset_generation

        with up_col:
            sales_key = f"sales_upload_v{st.session_state.uploader_version}"
            sales_file = st.file_uploader(
                "Upload Sales Data",
                type=["xlsx", "csv"],
                key=sales_key,
                help="Go to Modisoft -> Sales -> Live Items, Select Stores & Dates, Download as Excel"
            )
            
            if sales_file:
                if st.session_state.sales_selection_generation != current_gen:
                    st.session_state.sales_selected_name = sales_file.name
                    st.session_state.sales_selected_epoch = st.session_state.upload_epoch
                    st.session_state.sales_selection_generation = current_gen


            vendor_key = f"vendor_upload_v{st.session_state.uploader_version}"
            vendor_file = st.file_uploader(
                "Upload Vendor Price Data",
                type=["xlsx", "csv"],
                key=vendor_key,
                help="Go to Modisoft -> Products -> Price Book , Download as Excel"
            )
            
            if vendor_file:
                if st.session_state.vendor_selection_generation != current_gen:
                    st.session_state.vendor_selected_name = vendor_file.name
                    st.session_state.vendor_selected_epoch = st.session_state.upload_epoch
                    st.session_state.vendor_selection_generation = current_gen

        
        with upbtn_col:
            st.markdown('<div class="ft-right-btn">', unsafe_allow_html=True)

            upload_clicked = st.button(
                "⬆️ Upload Now",
                width="stretch",
                type="primary",
                disabled= (not _both_uploads_ok()),
                key="upload_submit",
            )

            st.markdown('</div>', unsafe_allow_html=True)


        # --- Handle the upload action immediately ---
        if upload_clicked:
            if not cfg.INCOMING_FOLDER_ID:
                st.session_state.upload_error = "Incoming Folder ID is empty."
                st.session_state.ui_phase = UI_UPLOAD_ERROR
                _rerun()

            if sales_file is None or vendor_file is None:
                st.session_state.upload_error = (
                    "Both Sales Data and Vendor Price Data are required."
                )
                st.session_state.ui_phase = UI_UPLOAD_ERROR
                _rerun()

            try:
                st.warning("Uploading files to google drive...")
                drive = _get_drive_service_or_raise(cfg)
                

                # --- Resolve user ---
                me = drive.about().get(
                    fields="user(emailAddress,permissionId,displayName)"
                ).execute().get("user", {})

                user_email = (me or {}).get("emailAddress") or "UNKNOWN_USER"

                # --- Per-user folder ---
                user_folder = get_or_create_subfolder(
                    drive,
                    cfg.INCOMING_FOLDER_ID,
                    user_email,
                )

                sales_folder = get_or_create_subfolder(
                    drive,
                    user_folder["id"],
                    "01 Sales Data Inputs",
                )

                vendor_folder = get_or_create_subfolder(
                    drive,
                    user_folder["id"],
                    "02 Vendor Price Data Inputs",
                )

                # --- Upload SALES ---
                sales_created = upload_to_drive(
                    drive,
                    data=sales_file.getvalue(),
                    name=f"{os.path.splitext(sales_file.name)[0]} (Sales Data via UI)",
                    mime=_infer_media_mime(sales_file.name),
                    folder_id=sales_folder["id"],
                    to_sheet=True,
                )

                # --- Upload VENDOR ---
                vendor_created = upload_to_drive(
                    drive,
                    data=vendor_file.getvalue(),
                    name=f"{os.path.splitext(vendor_file.name)[0]} (Vendor Price Data via UI)",
                    mime=_infer_media_mime(vendor_file.name),
                    folder_id=vendor_folder["id"],
                    to_sheet=True,
                )

                st.session_state.ui_phase = UI_READY
                _rerun()

            except Exception as e:
                st.session_state.upload_error = f"Upload failed: {e}"
                st.session_state.ui_phase = UI_UPLOAD_ERROR
                _rerun()
    
def render_run_options(cfg):
    run_form_wrapper_classes = "ft-card ft-row"

    # A file is "dirty" if the user has selected something not yet uploaded
    files_dirty = (
        st.session_state.sales_selected_name is not None
        or st.session_state.vendor_selected_name is not None
    )

    # Have we successfully uploaded both files?

    # OPEN the wrapper with real HTML (no entities)
    st.markdown(f'<div class="{run_form_wrapper_classes}">', unsafe_allow_html=True)



    with st.form("run_form"):
        # Header row uses the same columns to align the Run button with Upload button above
        tl, _, col_run = st.columns([4, 1, 1])
        with tl:
            st.subheader("Run Options")
            st.caption("Configure email behavior and report keys. Use **Advanced** for IDs/GIDs/timezone.")

        # --- Unified gating logic ---
        # A) If a file is currently selected but not uploaded -> disable Run
        # B) If no file selected and we have a prior successful upload -> enable Run
        # C) Otherwise (no prior upload or ambiguous state) -> disable Run

        with col_run:
            # Right-align and full-width, matching Upload Now
            st.markdown('<div class="ft-right-btn">', unsafe_allow_html=True)
            submitted = st.form_submit_button(
                "▶️ Run Pipeline",
                width='stretch',
                disabled=False,
                type="primary",
                key="run_submit"
            )
            st.markdown('</div>', unsafe_allow_html=True)

        # ----- Main options -----

        # Recipients
        st.markdown("##### Recipients")
        col1, col2 = st.columns([1, 1])
        with col1:
            to = st.text_input(
                "To (comma)", value=",".join(cfg.TO_RECIPIENTS or []),
                help="Fallback recipients for Manager & Order emails."
            )
        with col2:
            cc = st.text_input(
                "CC (comma)", value=",".join(cfg.CC_RECIPIENTS or []),
                help="Optional CC added to all emails."
            )

        # Report Keys
        st.markdown("##### Report Keys")
        colk1, colk2, colk3 = st.columns([1, 1, 2])
        with colk1:
            use_all = st.toggle(
                "Use all keys from CSV",
                value=cfg.USE_ALL_REPORT_KEYS,
                help="ON: process every key found. OFF: only the keys you list."
            )

        with colk2:
            pass

        with colk3:
            report_keys = st.text_input(
                "Keys to run (comma)",
                value=",".join(cfg.REPORT_KEY_RUN_LIST or []),
                help="Used when 'Use all keys' is OFF. For Sub_Report_Keys use Report_Key-Sub_Report_Key. Example: COFFEE,GROCERY,BEV-7UP"
            )

        # General Behavior
        st.markdown("##### General Behavior")
        cole1, cole2, cole3, cole4 = st.columns([1, 1, 1, 1])
        with cole1:
            include_full = st.toggle(
                "Attach FULL order in each email",
                value=cfg.INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL
            )
        with cole2:
            send_full = st.toggle(
                "Send separate FULL order email",
                value=cfg.SEND_SEPARATE_FULL_ORDER_EMAIL
            )
        with cole3:
            email_mgr = st.toggle(
                "Email Manager Report",
                value=getattr(cfg, "EMAIL_MANAGER_REPORT", True),
                help="When ON, the Manager Report email is sent. When OFF, it is skipped."
            )
        with cole4:
            use_rollover = st.toggle(
                    'Use auto-rollover for single week uploads',
                    value=cfg.USE_AUTO_ROLLOVER_IF_ONE_WEEK,
                    help='If this is on, when only 1 week is uploaded, the most recent previously uploaded data will become the "Last Week" data; If this is off then the "Last Week" data will be left blank'
                )

        # Per-Report-Key Recipients
        with st.expander("Per-Report-Key Recipients (optional)", expanded=False):

            st.caption("""
              Map **Store, Report Key → Emails (comma)**.
              
              **Email Delivery Priority:**  
              - `(Store, Key)` → First priority set of emails  
              - `(, Key)` → Second priority set of emails  
              - `(Store,)` → Third priority set of emails
              - If not defined, it will use the default set of emails in `To (comma)` field above
              """)
        
            rows = []
        
            if cfg.REPORT_KEY_RECIPIENTS:
                for (store, key, sub_key), emails in cfg.REPORT_KEY_RECIPIENTS.items():
                    rows.append({
                        "Store (optional)": store or "",
                        "Report Key (optional)": key or "",
                        "Sub-Report Key (optional)": sub_key or "",
                        "Emails (comma)": ",".join(emails or [])
                    })
            else:
                rows = [{
                    "Store (optional)": "",
                    "Report Key (optional)": "",
                    "Sub-Report Key (optional)": "",
                    "Emails (comma)": ""
                }]
        
            edited_rows = st.data_editor(
                rows,
                num_rows="dynamic",
                width='stretch',
                key="rk_editor",
            )
        
            rk_map = {}
            rk_preview = []
            rk_issues = []
        
            for i, r in enumerate(edited_rows):
        
                store = (r.get("Store (optional)") or "").strip().upper()
                key = (r.get("Report Key (optional)") or "").strip().upper()
                sub_key = (r.get("Sub-Report Key (optional)") or "").strip().upper() or None
                emails_raw = (r.get("Emails (comma)") or "").strip()
        
                emails = [e.strip() for e in emails_raw.split(",") if e.strip()]
        
                store_val = store if store else None
                key_val = key if key else None
                sub_val = sub_key if sub_key else None
        
                if emails and not (store_val or key_val or sub_key):
                    rk_issues.append(f"Row {i+1}: Must include Store, Key, or both.")
                    continue
        
                if (store_val or key_val or sub_key) and not emails:
                    rk_issues.append(f"Row {i+1}: Missing email(s).")
                    continue

                store_tag = clean_tag(store_val)
                key_tag = clean_tag(key_val)
                sub_tag = clean_tag(sub_val)
        
                rk_map[(store_tag, key_tag, sub_tag)] = emails
        
                rk_preview.append(f"{(store_val, key_val, sub_val)} -> {emails}")
        
            #if rk_preview:
            #    with st.expander("Recipient mapping preview"):
            #        st.code("\n".join(rk_preview), language="text")
        
            #if rk_issues:
            #    st.warning("Recipient configuration issues:\n\n- " + "\n- ".join(rk_issues))

        # Advanced
        with st.expander("Advanced", expanded=False):
            tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(["Folders", "Files & Links", "Ranges", "Timing", "Lifecycle", "Technical"])
            with tab1:
                ga1, ga2 = st.columns([1, 1])
                with ga1:
                    st.markdown("###### Input Folders")
                    user_folder = st.text_input("User Calculations Folder ID", value=cfg.USER_FOLDER_ID,
                        help="The file ID of the google drive folder that user workhorse files should be stored.")
                    incoming_id = st.text_input("Incoming Folder ID", value=cfg.INCOMING_FOLDER_ID,
                        help="The file ID of the google drive folder that user input folders & files should be stored.")
                    
                with ga2:
                    st.markdown("###### Output Folders")
                    order_folder = st.text_input("Order Report Folder ID", value=cfg.ORDER_REPORT_FOLDER_ID,
                        help="The file ID of the google drive folder that order report csv output files should be stored.")
                    mgr_folder = st.text_input("Manager Report Folder ID", value=cfg.MANAGER_REPORT_FOLDER_ID,
                        help="The file ID of the google drive folder that manager report pdf output files should be stored.")
                    error_folder = st.text_input("Error Report Folder ID", value=cfg.ERROR_REPORT_FOLDER_ID,
                        help="The file ID of the google drive folder that error report csv output files should be stored.")

            with tab2:
                gb1, gb2 = st.columns([1, 1])
                with gb1:
                    calc_id = st.text_input("Master Calculations Spreadsheet ID", value=cfg.CALC_SPREADSHEET_ID,
                        help="The file ID of the Master Calculations google sheets file that user workhorse files should be based off of.")
                with gb2:
                    BEV_MAPPING_LINK = st.text_input("BEV Mapping Link", value=cfg.BEV_MAPPING_LINK,
                        help="The url to the live, editable BEV Sub-Key Mapping google sheets file.")
            
            with tab3:
                gc1, gc2 = st.columns([1, 1])
                with gc1:
                    st.markdown("###### GIDs")
                    gid_mgr = st.text_input("Manager Report gid", value=str(cfg.GID_MANAGER_PDF),
                        help="The GID of the Manager Report Tab within the Master Calculations Sheet that should be used for outputs.")
                    gid_err = st.text_input("Error Report gid", value=str(cfg.GID_ERROR_REPORT),
                        help="The GID of the Error Report Tab within the Master Calculations Sheet that should be used for outputs.")
                    gid_order = st.text_input("Order Report gid", value=str(cfg.GID_ORDER_CSV),
                        help="The GID of the Order Report Tab within the Master Calculations Sheet that should be used for outputs.")
                    gid_bev_err = st.text_input("Unassigned Beverages Report gid", value=str(cfg.GID_ORDER_CSV),
                        help="The GID of the UB Report Tab within the Master Calculations Sheet that should be used for outputs.")
                with gc2:
                    st.markdown("###### Titles")
                    loc_sheet = st.text_input("Named Range Sheet Title", value=cfg.LOCATION_SHEET_TITLE,
                        help="The Sheet Title of the tab within the Master Calculations Sheet where the below named ranges exist.")
                    loc_range = st.text_input("Location Named Range", value=cfg.LOCATION_NAMED_RANGE,
                        help="The named range within the Master Calculations Sheet that refrences the cleaned location(s) name.")
                    update_range = st.text_input("Update Timestamp Range", value=cfg.TEMPLATE_UPDATE_RANGE,
                        help="The named range within the Master Calculations Sheet that refrences the last time the template was updated.")

            with tab4:
                gd1, gd2 = st.columns([1, 1])
                with gd1:
                    st.markdown("###### Data Integrity Controls")
                    _days = ["Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Any"]
                    start_dow = st.selectbox(
                        "Start day of week", _days, index=_days.index(cfg.START_DAY_OF_WEEK),
                        help="The day of week that the uploaded data should start at, any other day will raise an error."
                    )
                    end_dow = st.selectbox(
                        "End day of week", _days, index=_days.index(cfg.END_DAY_OF_WEEK),
                        help="The day of week that the uploaded data should end at, any other day will raise an error."
                    )
                with gd2:
                    st.markdown("###### Formatting")                                
                    tz = st.text_input("Timestamp Timezone", value=cfg.TIMESTAMP_TZ,
                        help="The timezone that should be used in all timestamps.")
                    tfmt = st.text_input("Timestamp Format", value=cfg.TIMESTAMP_FMT,
                        help="The format that should be used in all timestamps.")
            with tab5:
                ge1, ge2 = st.columns([1, 1])
                with ge1:
                    st.markdown("###### One-Time Use Files")
                    failed_input_ttl = st.number_input(
                        "Failed Input Time-To-Life (days)",
                        min_value=0,
                        max_value=3650,
                        value=int(cfg.FAILED_INPUT_TIME_TO_LIFE),
                        help="Delete old unused incoming files older than this many days after a successful run."
                        )
                    output_ttl = st.number_input(
                        "Output Time-To-Life (days)",
                        min_value=0,
                        max_value=3650,
                        value=int(cfg.OUTPUT_TIME_TO_LIFE),
                        help="Delete output files older than this many days after a successful run."
                        )
                with ge2:
                    st.markdown("###### Recurring Use Files")
                    user_ttl = st.number_input(
                        "User Calculations Time-To-Life (days)",
                        min_value=0,
                        max_value=3650,
                        value=int(cfg.USER_TIME_TO_LIFE),
                        help="Delete old unused user calculations files older than this many days after a successful run."
                        )
            
            with tab6:
                gf1, gf2 = st.columns([1, 1])
                with gf1:
                    raw_redirect_port = int(cfg.REDIRECT_PORT) if str(cfg.REDIRECT_PORT).isdigit() else 0
                    redirect_port = st.number_input(
                        "Redirect Port (0 = auto)",
                        min_value=0, max_value=65535,
                        value=raw_redirect_port if raw_redirect_port in (0, *range(1024, 65536)) else 0,
                        help="Use 0 to auto-pick a free port. Otherwise choose 1024–65535."
                    )
                with gf2:
                    error_recipients = st.text_input(
                        "Technical Support (comma)",
                        value=",".join(cfg.ERROR_RECIPIENTS or []),
                        help="If errors arise such as missing items in the Vendor Price Book, the error report will be sent here."
                    )


        save_drive_defaults = st.checkbox("Update defaults", value=False)

        # ----- Submission handling -----

        if submitted:
            # Apply per-run config
            cfg.TO_RECIPIENTS = _split_emails(to)
            cfg.CC_RECIPIENTS = _split_emails(cc)
            cfg.ERROR_RECIPIENTS = _split_emails(error_recipients)
            cfg.USE_ALL_REPORT_KEYS = use_all
            cfg.REPORT_KEY_RUN_LIST = [s.strip().upper() for s in (report_keys or "").split(",") if s.strip()]

            cfg.INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL = include_full
            cfg.SEND_SEPARATE_FULL_ORDER_EMAIL = send_full
            cfg.EMAIL_MANAGER_REPORT = bool(email_mgr)

            cfg.CALC_SPREADSHEET_ID = calc_id
            cfg.INCOMING_FOLDER_ID = incoming_id
            cfg.MANAGER_REPORT_FOLDER_ID = mgr_folder
            cfg.ORDER_REPORT_FOLDER_ID = order_folder
            cfg.ERROR_REPORT_FOLDER_ID = error_folder
            cfg.USER_FOLDER_ID = user_folder
            cfg.REDIRECT_PORT = int(redirect_port)

            cfg.GID_MANAGER_PDF = gid_mgr
            cfg.GID_ORDER_CSV = gid_order
            cfg.GID_ERROR_REPORT = gid_err
            cfg.GID_BEV_ERRORS = gid_bev_err
            cfg.LOCATION_SHEET_TITLE = loc_sheet
            cfg.LOCATION_NAMED_RANGE = loc_range
            cfg.TEMPLATE_UPDATE_RANGE = update_range
            cfg.TIMESTAMP_TZ = tz
            cfg.TIMESTAMP_FMT = tfmt
            
            cfg.OUTPUT_TIME_TO_LIFE = int(output_ttl)
            cfg.FAILED_INPUT_TIME_TO_LIFE = int(failed_input_ttl)
            cfg.USER_TIME_TO_LIFE = int(user_ttl)

            cfg.USE_AUTO_ROLLOVER_IF_ONE_WEEK = bool(use_rollover)
            cfg.START_DAY_OF_WEEK = start_dow
            cfg.END_DAY_OF_WEEK = end_dow

            cfg.BEV_MAPPING_LINK = BEV_MAPPING_LINK


            # Per-key recipients from editor
            
            rk_map = {}
            rk_preview = []  # optional: for a quick visual confirmation in the UI
            
            for r in edited_rows:
                store = (r.get("Store (optional)") or "").strip().upper() or None
                key   = (r.get("Report Key (optional)") or "").strip().upper() or None
                sub_key   = (r.get("Sub-Report Key (optional)") or "").strip().upper() or None
            
                emails_raw = (r.get("Emails (comma)") or "").strip()
                emails = [e.strip() for e in emails_raw.split(",") if e.strip()]
            
                store_tag = clean_tag(store) if store else None
                key_tag = clean_tag(key) if key else None
                sub_tag = clean_tag(sub_key) if sub_key else None
            
                if (store_tag or key_tag or sub_tag) and emails:
                    # >>> THIS is the part that actually records the mapping
                    rk_map[(store_tag, key_tag, sub_tag)] = emails
            
            cfg.REPORT_KEY_RECIPIENTS = rk_map

            # --- ADD: warnings before kicking off the run ---
            if not cfg.USE_ALL_REPORT_KEYS and not cfg.REPORT_KEY_RUN_LIST:
                st.session_state.upload_error = (
                    "No report keys selected. Either enable 'Use all keys' "
                    "or provide explicit report keys."
                )
                st.session_state.ui_phase = UI_UPLOAD_ERROR
                _rerun()

            if not cfg.TO_RECIPIENTS and not cfg.DEFAULT_TO_RECIPIENTS and not rk_map:
                st.session_state.run_error = (
                    "No email recipients defined. At least one recipient is required."
                )
                st.session_state.ui_phase = UI_RESULT_ERROR
                _rerun()

            if rk_issues:
                st.session_state.run_error = (
                    "Invalid per‑report‑key recipient configuration:\n\n"
                    + "\n".join(rk_issues)
                )
                st.session_state.ui_phase = UI_RESULT_ERROR
                _rerun()

            # All validation must already be done
            st.session_state._run_start_time = None
            st.session_state.ui_phase = UI_RUNNING

            # Start pipeline in background (one time)
            if not st.session_state.pipeline_thread_started:
                start_run()

                t = threading.Thread(
                    target=run_pipeline_controller,
                    args=(cfg, st.session_state.pipe_run_id),
                    daemon=True
                )
                st.session_state.pipeline_thread = t
                t.start()

            _rerun()
            # --- END ADD ---


            # --- Save edited defaults to Drive JSON (optional) ---
            if save_drive_defaults:
                try:
                    # Ensure we have a user token first
                    creds = load_valid_token(cfg.SCOPES)
                    if not creds:
                        st.error("Not authenticated. Please complete Google sign‑in first (top of page).")
                    else:
                        # Drive service
                        _sheets, drive, _gmail = services(creds, cfg.HTTP_TIMEOUT_SECONDS)

                        # What we persist
                        drive_defaults = {
                            "CALC_SPREADSHEET_ID": cfg.CALC_SPREADSHEET_ID,
                            "INCOMING_FOLDER_ID": cfg.INCOMING_FOLDER_ID,
                            "MANAGER_REPORT_FOLDER_ID": cfg.MANAGER_REPORT_FOLDER_ID,
                            "ORDER_REPORT_FOLDER_ID": cfg.ORDER_REPORT_FOLDER_ID,
                            "ERROR_REPORT_FOLDER_ID": cfg.ERROR_REPORT_FOLDER_ID,
                            "USER_FOLDER_ID": cfg.USER_FOLDER_ID,

                            "GID_MANAGER_PDF": cfg.GID_MANAGER_PDF,
                            "GID_ORDER_CSV": cfg.GID_ORDER_CSV,
                            "GID_ERROR_REPORT": cfg.GID_ERROR_REPORT,
                            "GID_BEV_ERRORS": cfg.GID_BEV_ERRORS,

                            "LOCATION_SHEET_TITLE": cfg.LOCATION_SHEET_TITLE,
                            "LOCATION_NAMED_RANGE": cfg.LOCATION_NAMED_RANGE,
                            "TEMPLATE_UPDATE_RANGE": cfg.TEMPLATE_UPDATE_RANGE,

                            "TIMESTAMP_TZ": cfg.TIMESTAMP_TZ,
                            "TIMESTAMP_FMT": cfg.TIMESTAMP_FMT,

                            "OUTPUT_TIME_TO_LIFE": cfg.OUTPUT_TIME_TO_LIFE,
                            "FAILED_INPUT_TIME_TO_LIFE": cfg.FAILED_INPUT_TIME_TO_LIFE,
                            "USER_TIME_TO_LIFE": cfg.USER_TIME_TO_LIFE,

                            "TO_RECIPIENTS": cfg.TO_RECIPIENTS,   # lists are fine; JSON keeps types
                            "CC_RECIPIENTS": cfg.CC_RECIPIENTS,
                            "ERROR_RECIPIENTS": cfg.ERROR_RECIPIENTS,

                            "USE_ALL_REPORT_KEYS": cfg.USE_ALL_REPORT_KEYS,
                            "REPORT_KEY_RUN_LIST": cfg.REPORT_KEY_RUN_LIST,

                            "REPORT_KEY_RECIPIENTS": cfg.REPORT_KEY_RECIPIENTS,

                            "DEFAULT_ORDER_RECIPIENTS": cfg.DEFAULT_ORDER_RECIPIENTS,

                            "INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL": cfg.INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL,
                            "SEND_SEPARATE_FULL_ORDER_EMAIL": cfg.SEND_SEPARATE_FULL_ORDER_EMAIL,
                            "EMAIL_MANAGER_REPORT": cfg.EMAIL_MANAGER_REPORT,

                            "USE_AUTO_ROLLOVER_IF_ONE_WEEK" : cfg.USE_AUTO_ROLLOVER_IF_ONE_WEEK,
                            "START_DAY_OF_WEEK" : cfg.START_DAY_OF_WEEK,
                            "END_DAY_OF_WEEK" : cfg.END_DAY_OF_WEEK,

                            "BEV_MAPPING_LINK" : cfg.BEV_MAPPING_LINK
                        }

                        DEV_ENVIRONMENT = st.secrets.get("DEV_ENVIRONMENT", False)
                        DEV_CONFIG_FILE_ID = (st.secrets.get("DEV_CONFIG_FILE_ID", "") or "").strip()
                        CONFIG_FILE_ID = (st.secrets.get("CONFIG_FILE_ID", "") or "").strip()

                        # Decide where to SAVE
                        if DEV_ENVIRONMENT:
                            save_target_id = DEV_CONFIG_FILE_ID or None
                        else:
                            save_target_id = CONFIG_FILE_ID or None

                        new_id = save_config_to_drive(
                            drive,
                            drive_defaults,
                            file_id=save_target_id
                        )

                        # DEV auto-bootstrap case
                        if DEV_ENVIRONMENT and not DEV_CONFIG_FILE_ID:
                            st.success("✅ Created new DEV config file.")
                            st.info(
                                "Add this to Streamlit secrets as DEV_CONFIG_FILE_ID:\n\n"
                                f"`{new_id}`"
                            )
                        else:
                            st.success(f"✅ Defaults saved (file id: {new_id})")

                except Exception as e:
                    st.error(f"Failed to save defaults to Drive: {e}")


            # If user checked "Force Google re-auth for this run", kick them into auth gating first.
            """
            if cfg.FORCE_REAUTH:
                clear_token()
                try:
                    flow, url = start_web_oauth(cfg.SCOPES, cfg.REDIRECT_PORT)
                    st.session_state.oauth_flow = flow
                    st.session_state.oauth_url = url
                    st.session_state.auth_required = True
                    st.info("Re-auth required for this run. Open the URL shown in the Authentication panel.")
                    _rerun()
                except Exception as e:
                    st.error(f"Failed to start OAuth: {e}")
            else:
                st.session_state.ui_phase = UI_RUNNING
                _rerun()
            """

def run_pipeline_controller(cfg, run_id):

    logger = StatusLogger(
        print_to_console=True,
        file_path="last_run.log",
        overwrite=True,
    )

    try:
        result = run_pipeline(cfg, logger=logger)
        get_pipeline_queue().put(
            (run_id, PIPE_STATUS_DONE, result)
        )
    except Exception as e:
        get_pipeline_queue().put(
            (
                run_id,
                PIPE_STATUS_ERROR,
                {
                    "type": type(e).__name__,
                    "user_message": str(e),
                    "traceback": traceback.format_exc(),
                },
            )
        )

    finally:
        logger.close()


def render_running_status(cfg):
    import time
    import os
    import queue
    import streamlit as st
    from streamlit_autorefresh import st_autorefresh

    # ------------------------------------------------------------
    # Poll queue FIRST (edge-triggered)
    # ------------------------------------------------------------
    
    q = get_pipeline_queue()
    
    while True:
        try:
            run_id, status, payload = q.get_nowait()
        except queue.Empty:
            break

        if run_id != st.session_state.get("pipe_run_id"):
            continue

        if status == PIPE_STATUS_DONE:
            st.session_state.pipe_result = payload
            st.session_state.pipe_status = PIPE_STATUS_DONE
            st.session_state.pipe_finished = True

        elif status == PIPE_STATUS_ERROR:
            st.session_state.pipe_error = payload
            st.session_state.run_error = payload
            st.session_state.pipe_status = PIPE_STATUS_ERROR
            st.session_state.pipe_finished = True
    
    if st.session_state.pipe_finished:
        if st.session_state.pipe_status == PIPE_STATUS_DONE:
            st.session_state.ui_phase = UI_RESULT
        elif st.session_state.pipe_status == PIPE_STATUS_ERROR:
            st.session_state.ui_phase = UI_RESULT_ERROR
        _rerun()

    # ------------------------------------------------------------
    # ALWAYS render something
    # ------------------------------------------------------------
    with st.status("Running pipeline…", expanded=True):

        # ----- Bulletproof timer -----
        start_time = st.session_state.get("_run_start_time")
        if not isinstance(start_time, (int, float)):
            start_time = time.perf_counter()
            st.session_state._run_start_time = start_time

        elapsed = int(time.perf_counter() - start_time)
        h, m, s = elapsed // 3600, (elapsed % 3600) // 60, elapsed % 60
        st.markdown(f"**Elapsed:** `{h:02d}:{m:02d}:{s:02d}`")

        # ----- Log tail -----
        if os.path.exists("last_run.log"):
            try:
                with open("last_run.log", "r", encoding="utf-8") as f:
                    st.code("".join(f.readlines()[-8:]), language="text")
            except Exception:
                st.markdown("*Waiting for logs…*")
        else:
            st.markdown("*Waiting for logs…*")

    # ------------------------------------------------------------
    # LIVE status check (DO NOT CACHE THIS)
    # ------------------------------------------------------------
    status = st.session_state.get("pipe_status")

    # ✅ This is what keeps the UI alive
    
    if (
        st.session_state.pipe_status == PIPE_STATUS_RUNNING
        and not st.session_state.pipe_finished
    ):
        st_autorefresh(
            interval=1000,
            key=f"pipeline_tick_{st.session_state.pipe_run_id}",
        )



def render_results(cfg):
    result = st.session_state.get("pipe_result")

    
    if not _validate_pipeline_result(result):
        st.error(
            "Run completed, but the pipeline did not return a valid result object."
        )
        if os.path.exists("last_run.log"):
            with open("last_run.log", "rb") as f:
                st.download_button(
                    "⬇️ Download log",
                    f.read(),
                    file_name="last_run.log",
                    mime="text/plain",
                )


    with st.container(border=True):
        st.subheader("✅ Run Complete")

        st.write("### Outputs")
        col1, col2, col3 = st.columns(3)

        col1.metric("Location", result.location)
        col2.metric("Timestamp", result.timestamp)
        col3.metric(
            "Elapsed",
            f"{result.elapsed_seconds//3600:02d}:"
            f"{(result.elapsed_seconds%3600)//60:02d}:"
            f"{result.elapsed_seconds%60:02d}"
        )

        if getattr(result, "manager_pdf_link", None):
            st.success(f"Manager PDF: {result.manager_pdf_link}")
        if getattr(result, "full_order_link", None):
            st.success(f"Full Order Sheet: {result.full_order_link}")

    
        if os.path.exists("last_run.log"):
            with open("last_run.log", "rb") as f:
                st.session_state["last_run_log"] = f.read()
                st.session_state["last_run_timestamp"] = result.timestamp


        if "last_run_log" in st.session_state:
            st.download_button(
                "⬇️ Download full log (last_run.log)",
                st.session_state["last_run_log"],
                file_name=f"last_run_{st.session_state['last_run_timestamp']}.log",
                mime="text/plain",
                width='stretch'
                )


def render_sidebar(cfg):
    with st.sidebar:
        st.header("Utilities")

        # --- Existing buttons ---
        if st.button("Google Sign Out", type="secondary", width='stretch'):
            clear_token()
            for key in ["auth_required", "oauth_flow", "oauth_url", "auth_checked"]:
                if key in st.session_state:
                    del st.session_state[key]
            _rerun()

        st.link_button("Add Users to App", "https://console.cloud.google.com/auth/audience?project=favtripdev", width='stretch')
        st.link_button("Open Google Drive", "https://drive.google.com/drive/u/0/folders/1Wpq1JBQDZSJsxBPi5q4rtZfSjD7ZkU4k", width='stretch')
        st.link_button("Open Modisoft", "https://insights.modisoft.com/account/logon", width='stretch')
        st.link_button("Open Bev Mapping File", cfg.BEV_MAPPING_LINK, width='stretch')

        if False:
            st.checkbox(
                "Offer full log download",
                key="offer_log_download",
                help="If enabled, a 'Download last_run.log' button appears when a run finishes."
            )


        # =============================================================
        # DEV-ONLY: Push DEV Defaults → PROD Defaults
        # =============================================================
        DEV_ENVIRONMENT = bool(st.secrets.get("DEV_ENVIRONMENT", False))

        if DEV_ENVIRONMENT:
            st.divider()
            st.subheader("DEV Tools")

            if st.button(
                "🚀 Push Dev Defaults to Prod",
                type="primary",
                width="stretch",
                help="Overwrite the PROD defaults JSON with the current DEV defaults",
            ):
                st.session_state["confirm_push_dev_to_prod"] = True


        @st.dialog("⚠️ Confirm Push to Production")
        def confirm_push_dev_to_prod():
            st.markdown(
                """
                **You are about to overwrite the PROD defaults configuration.**

                - ✅ PROD file ID will remain unchanged  
                - ✅ DEV defaults will completely replace PROD defaults  
                - ❌ This action **cannot be undone**

                Please confirm you want to continue.
                """
            )

            col_confirm, col_cancel = st.columns(2)

            with col_confirm:
                if st.button("✅ Yes — Push to PROD", type="primary", width="stretch"):
                    try:
                        DEV_CONFIG_FILE_ID = (st.secrets.get("DEV_CONFIG_FILE_ID", "") or "").strip()
                        PROD_CONFIG_FILE_ID = (st.secrets.get("CONFIG_FILE_ID", "") or "").strip()

                        if not DEV_CONFIG_FILE_ID or not PROD_CONFIG_FILE_ID:
                            st.error("Missing DEV_CONFIG_FILE_ID or CONFIG_FILE_ID in secrets.")
                        else:
                            creds = load_valid_token(cfg.SCOPES)
                            if not creds:
                                st.error("Google authentication required.")
                            else:
                                _, drive, _ = services(creds, cfg.HTTP_TIMEOUT_SECONDS)

                                # Load DEV defaults
                                dev_blob = drive.files().get_media(
                                    fileId=DEV_CONFIG_FILE_ID
                                ).execute()
                                dev_defaults = json.loads(dev_blob.decode("utf-8"))

                                # Overwrite PROD defaults (same file ID)
                                save_config_to_drive(
                                    drive,
                                    dev_defaults,
                                    file_id=PROD_CONFIG_FILE_ID
                                )

                                st.success("✅ DEV defaults successfully pushed to PROD.")

                    except Exception as e:
                        st.error(f"Push failed: {e}")

                    finally:
                        st.session_state.pop("confirm_push_dev_to_prod", None)
                        _rerun()

            with col_cancel:
                if st.button("❌ Cancel", width="stretch"):
                    st.session_state.pop("confirm_push_dev_to_prod", None)
                    _rerun()


        # Trigger dialog
        if st.session_state.get("confirm_push_dev_to_prod"):
            confirm_push_dev_to_prod()

def render_upload_different_button(cfg):
    if st.button("🔁 Upload different files", width="stretch"):
        reset_to_upload()
        reset_pipeline_state()

        st.session_state.ui_phase = UI_UPLOAD
        _rerun()



def render_result_error(cfg):
    payload = st.session_state.get("pipe_error")

    with st.container(border=True):
        st.subheader("❌ Run Failed")

        if isinstance(payload, dict):
            # ✅ Friendly message (wrapped, readable)
            st.error(f"{payload['type']}: {payload['user_message']}")

            # ✅ Technical details hidden by default
            with st.expander("Technical details"):
                st.text(payload["traceback"])

        else:
            st.error("Unknown error occurred.")
        
        if st.button("🔁 Upload different files", type="primary"):
            st.session_state.pop("run_error", None)
            reset_to_upload()
            reset_pipeline_state()
            _rerun()


    if os.path.exists("last_run.log") and "last_run_log" not in st.session_state:
        with open("last_run.log", "rb") as f:
            st.session_state["last_run_log"] = f.read()
        st.session_state.setdefault("last_run_timestamp", "error")
    
    if "last_run_log" in st.session_state:
        st.download_button(
            "⬇️ Download full log (last_run.log)",
            st.session_state["last_run_log"],
            file_name=f"last_run_{st.session_state['last_run_timestamp']}.log",
            mime="text/plain",
            width='stretch'
            )



def render_upload_error(cfg):
    with st.container(border=True):
        st.subheader("❌ Invalid Upload")

        st.error(
            "Your uploaded file is invalid.\n\n"
            "Please upload **1 or 2 full weeks of data only**."
        )

        st.warning(st.session_state.get("upload_error", ""))

        if st.button("🔁 Upload different files", type="primary"):
            st.session_state.pop("run_error", None)
            reset_to_upload()
            reset_pipeline_state()
            _rerun()


def render_app(cfg):
    phase = st.session_state.ui_phase

    if phase == UI_UPLOAD:
        render_sidebar(cfg)
        render_upload_card(cfg)

    elif phase == UI_READY:
        render_sidebar(cfg)
        render_run_options(cfg)
        render_upload_different_button(cfg)

    elif phase == UI_RUNNING:
        render_running_status(cfg)

    elif phase == UI_RESULT:
        render_sidebar(cfg)
        render_results(cfg)
        render_upload_different_button(cfg)
    
    elif phase == UI_RESULT_ERROR:
        render_sidebar(cfg)
        render_result_error(cfg)

    elif phase == UI_UPLOAD_ERROR:
        render_sidebar(cfg)
        render_upload_error(cfg)

        


# =========================
# App Entrypoint
# =========================

#st.title("🧾 FavTrip Reporting Pipeline")


st.set_page_config(
    page_title="FT Reporting",
    page_icon="🧾",          # emoji or path/URL to an image
    layout="wide",           # "centered" or "wide"
    initial_sidebar_state="collapsed",  # "auto", "expanded", "collapsed"
    menu_items={
        "Get Help": "mailto:ryan-morrow@uiowa.edu",
        "Report a bug": "https://github.com/ryan-j-morrow/favtrip_reporting/issues",
        "About": "FavTrip Reporting Pipeline",
    },
)

defaults = {
    "sales_selected_name": None,
    "vendor_selected_name": None,
    "sales_uploaded_ok": False,
    "vendor_uploaded_ok": False,
    "offer_log_download": False,
    "uploader_version": 0,
    "ui_phase": UI_UPLOAD,
    "auth_required": True,
    "running_ui_initialized": False,    
    "upload_epoch": 0,                 # increments on “Upload different files”
    "sales_selected_epoch": None,       # epoch when sales file was selected
    "vendor_selected_epoch": None,    
    "reset_generation": 0,
    "sales_selection_generation": None,
    "vendor_selection_generation": None
}

    
for key, default in defaults.items():
    if key not in st.session_state:
        st.session_state[key] = default

cfg = Config.load()

creds = load_valid_token(cfg.SCOPES)
st.session_state.auth_required = creds is None


# --- STATE INIT ---
init_thread_state()
init_pipeline_state()

# --- Finish OAuth inline when redirect comes back (this is in the NEW TAB) ---
params = st.query_params
if "code" in params and "state" in params:
    try:
        finish_web_oauth(params["code"], params["state"], cfg.SCOPES)
        # Token is saved locally in this new tab's app process
        st.success("✅ Google authentication complete.")
        
        has_token = (load_valid_token(cfg.SCOPES) is not None)
        st.session_state.auth_required = not has_token

        # Remove code/state from URL
        st.query_params.clear()

        # No messaging back to opener and NO window.close().
        # This tab becomes the main app; just rerun to flip UI.
        st.toast("Signed in. Loading the app…")
        _rerun()
    except Exception as e:
        st.error(f"OAuth error: {e}")

if (not st.session_state.auth_required) and ("sidebar_hint_seen" not in st.session_state):
    col_msg, col_btn = st.columns([6, 1], vertical_alignment="center")

    with col_msg:
        st.info(
            "⬅️ **Open the sidebar** for Utilities, Google auth, and DEV tools.",
            icon="👈",
        )

    with col_btn:
        if st.button("Got it", type="secondary"):
            st.session_state["sidebar_hint_seen"] = True
            _rerun()

# Auth gate
if st.session_state.auth_required:
    # ----------------------------
    # Authentication panel (shown only if auth required)
    # ----------------------------
    if st.session_state.auth_required:
        with st.expander("Google Authentication", expanded=True):
            st.caption(
                "Authentication is required before running. "
                "Click **Sign in with Google** to open the consent screen (it will open in a new tab)."
            )

            sign_in_ph = st.empty()
            clicked = sign_in_ph.button("Sign in with Google", type="primary", width='stretch')

            if clicked:
                try:
                    auth_url = start_web_oauth(cfg.SCOPES)
                    sign_in_ph.empty()

                    # Friendly message in this (original) tab
                    st.markdown(
                        """
                        <div style="
                            display:flex;align-items:center;justify-content:center;
                            height:55vh;text-align:center;
                            font-family: system-ui, Segoe UI, Roboto, Helvetica, Arial, sans-serif;">
                        <div>
                            <h2 style="margin-bottom:0.5rem;">You're being signed in…</h2>
                            <p style="font-size:1.05rem;opacity:.9;">
                            A new browser tab was opened for Google sign‑in.<br/>
                            <strong>After it completes, continue in that tab.</strong>
                            </p>
                        </div>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )

                    # Optional: refresh this tab when user returns (not required)
                    html(
                        """
                        <script>
                        document.addEventListener("visibilitychange", function() {
                            if (!document.hidden) { location.reload(); }
                        });
                        </script>
                        """,
                        height=0,
                    )

                    # Open Google auth in a NEW tab (this will ultimately become the main app)
                    html(
                        f"""
                        <script>
                        window.open({json.dumps(auth_url)}, "_blank", "noopener");
                        </script>
                        """,
                        height=0,
                    )

                    st.stop()
                except Exception as e:
                    st.error(f"Failed to start OAuth: {e}")

            with st.expander("Having trouble?", expanded=False):
                st.write(
                    "- The Google authorization page opens in a **new browser tab**.\n"
                    "- After completing consent, the **new tab** will load the app.\n"
                    "- If you renamed your Streamlit app or URL, ensure the Google OAuth "
                    "Authorized redirect URI matches exactly (including trailing slash)."
                )
            st.stop()

render_app(cfg)

```

---
### file: cli.py

```python
import argparse
from favtrip.config import Config
from favtrip.logger import StatusLogger
from favtrip.pipeline import run_pipeline


def main():
    parser = argparse.ArgumentParser(description="FavTrip Reporting Pipeline")
    parser.add_argument("--env", help="Path to .env file", default=None)

    # Per-run overrides (subset)
    parser.add_argument("--to", help="Comma-separated recipients", default=None)
    parser.add_argument("--cc", help="Comma-separated cc", default=None)
    parser.add_argument("--use-all-keys", action="store_true")
    parser.add_argument("--report-keys", help="Comma-separated report keys to run", default=None)
    parser.add_argument("--force-reauth", action="store_true")

    args = parser.parse_args()
    cfg = Config.load(args.env)

    if args.to:
        cfg.TO_RECIPIENTS = [s.strip() for s in args.to.split(',') if s.strip()]
    if args.cc:
        cfg.CC_RECIPIENTS = [s.strip() for s in args.cc.split(',') if s.strip()]
    if args.use_all_keys:
        cfg.USE_ALL_REPORT_KEYS = True
    if args.report_keys:
        cfg.REPORT_KEY_RUN_LIST = [s.strip().upper() for s in args.report_keys.split(',') if s.strip()]
    if args.force_reauth:
        cfg.FORCE_REAUTH = True

    logger = StatusLogger()
    result = run_pipeline(cfg, logger=logger)

    print("===== SUMMARY =====")
    print(logger.as_text())
    print("===================")


if __name__ == "__main__":
    main()

```

---
### file: core_functional_modules/__init__.py

```python
__all__ = [
    "config",
    "google_client",
    "sheets_utils",
    "drive_utils",
    "gmail_utils",
    "pipeline",
    "logger",
]

```

---
### file: core_functional_modules/config.py

```python
"""
config
======================================

Configuration loader and serializer for FavTrip reporting apps.

This module centralizes all runtime configuration for both local development
and cloud deployments (e.g., Streamlit Community Cloud). It provides a single,
typed `Config` dataclass plus helper functions that safely read from multiple
sources, coerce values to the expected Python types, and (optionally) overlay
a remote, Google Drive–hosted JSON configuration at runtime.

The loader is designed to be:
- **Layered**: Values are pulled from three tiers, in this order:
  1) Streamlit `st.secrets` (preferred in cloud; values may already be typed)
  2) Process environment and/or a local `.env` file (string-based; coerced) #Sandbox use only
  3) A Google Drive JSON override, applied last if credentials and
     a config file are available
- **Safe**: Missing keys never raise; reasonable defaults are used instead.
- **Type-aware**: Bools, lists, and dicts are parsed/coerced consistently so the
  same code works with typed TOML (in `st.secrets`) and string-based `.env`.

-------------------------------------------------------------------------------
Core API
-------------------------------------------------------------------------------

- `_get_secret(key: str, default: Any = None) -> Any`
  Attempts to read `key` from `streamlit.secrets` (if Streamlit is present and
  has `secrets`), else falls back to `os.getenv(key, default)`. Never raises for
  missing keys; always returns a value (possibly `default`). Streamlit import is
  lazy to avoid a hard dependency for non-Streamlit contexts.

- `_coerce_bool(v: Any, default: bool = False) -> bool`
  Accepts `bool | str | int | None` and returns a Python `bool`.
  Truthy strings (case-insensitive, trimmed) include:
  `{"1", "true", "yes", "on", "y", "t"}`. Non-parseable inputs fall back to
  `default`.

- `_coerce_csv(v: Any) -> List[str]`
  Accepts a list/tuple (already structured) or a comma-separated string and
  yields a list of **trimmed** strings. `None`/empty returns `[]`.

- `_coerce_json(v: Any) -> Dict[str, Any]`
  Accepts a `dict` or a JSON string. Returns a `dict`; parse failures yield `{}`.

- `@dataclass class Config`
  A top-level dataclass holding all tunable settings for the application:
  * **Drive/Sheets IDs**:
    - `CALC_SPREADSHEET_ID`, `INCOMING_FOLDER_ID`, `MANAGER_REPORT_FOLDER_ID`,
      `ORDER_REPORT_FOLDER_ID`, `USER_FOLDER_ID`
  * **GIDs, sheet metadata, timestamps**:
    - `GID_MANAGER_PDF`, `GID_ORDER_CSV`, `LOCATION_SHEET_TITLE`,
      `LOCATION_NAMED_RANGE`, `TEMPLATE_UPDATE_RANGE`
    - `TIMESTAMP_TZ` (e.g., "America/Chicago")
    - `TIMESTAMP_FMT` (default "%Y-%m-%d-%I-%M-%p")
  * **Email & distribution**:
    - `TO_RECIPIENTS`, `CC_RECIPIENTS`, `USE_ALL_REPORT_KEYS`,
      `REPORT_KEY_RUN_LIST`, `REPORT_KEY_RECIPIENTS`,
      `DEFAULT_ORDER_RECIPIENTS`
    - `INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL`,
      `SEND_SEPARATE_FULL_ORDER_EMAIL`, `EMAIL_MANAGER_REPORT`
  * **Google API**:
    - `SCOPES` (Drive/Sheets/Gmail), `FORCE_REAUTH`,
      `REDIRECT_PORT`, `HTTP_TIMEOUT_SECONDS`
  * **Advanced intake**:
    - `USE_AUTO_ROLLOVER_IF_ONE_WEEK`,
      `START_DAY_OF_WEEK`, `END_DAY_OF_WEEK`
      (Accepted values include: Sunday, Monday, Tuesday, Wednesday, Thursday,
      Friday, Saturday, Any)
  * **Cleanup (days)**:
    - `OUTPUT_TIME_TO_LIFE`, `FAILED_INPUT_TIME_TO_LIFE`, `USER_TIME_TO_LIFE`

  Defaults are provided for all fields. When loading from secrets or `.env`,
  values are coerced into the correct types; lists and dicts are parsed as
  necessary. `REPORT_KEY_RUN_LIST` values are uppercased to reduce downstream
  casing issues.

- `Config.load(env_path: Optional[pathlib.Path] = None) -> Config`
  Loads the final, effective configuration via a layered merge:
  1) Loads a local `.env` file from `env_path` (default: `cwd/.env`) using
     `python-dotenv` with `override=False` (so existing process env vars win).
  2) Reads settings from `st.secrets` if available; otherwise from environment.
     Values are passed through the coercers defined above.
  3) Attempts to overlay a Google Drive–hosted JSON config:
     - Uses `core_functional_modules.google_client.load_valid_token` and `services` to obtain a
       Drive client (respecting `HTTP_TIMEOUT_SECONDS`).
     - Reads a JSON dict via `core_functional_modules.config_store.load_config_from_drive`,
       optionally using `CONFIG_FILE_ID` from `st.secrets` if present.
     - Keys in the override dict that match `Config` attributes replace
       previously loaded values.
     - On any failure (no token, network error, file missing, etc.), the loader
       **fails open** and returns the base config without raising (best-effort).

- `Config.to_env() -> str`
  Serializes the current configuration to a string in `.env` format. Collections
  are flattened—lists are joined with commas, and dicts are JSON-encoded—so the
  output can be written to disk and re-read later in a purely string-based env.

- `Config.save(env_path: Optional[pathlib.Path] = None) -> None`
  Convenience wrapper around `to_env()` that writes the serialized configuration
  to `env_path` (default: `cwd/.env`, UTF-8).

-------------------------------------------------------------------------------
Environment / Secrets Reference (all optional; sensible defaults apply)
-------------------------------------------------------------------------------

Drive / Sheets IDs:
- `CALC_SPREADSHEET_ID`, `INCOMING_FOLDER_ID`, `MANAGER_REPORT_FOLDER_ID`,
  `ORDER_REPORT_FOLDER_ID`, `USER_FOLDER_ID`

Sheet metadata & timestamps:
- `GID_MANAGER_PDF`, `GID_ORDER_CSV`, `LOCATION_SHEET_TITLE`,
  `LOCATION_NAMED_RANGE`, `TEMPLATE_UPDATE_RANGE`
- `TIMESTAMP_TZ`, `TIMESTAMP_FMT`

Email & distribution:
- `TO_RECIPIENTS` (CSV), `CC_RECIPIENTS` (CSV)
- `USE_ALL_REPORT_KEYS` (bool)
- `REPORT_KEY_RUN_LIST` (CSV; uppercased during load)
- `REPORT_KEY_RECIPIENTS` (JSON dict)
- `DEFAULT_ORDER_RECIPIENTS` (CSV)
- `INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL` (bool)
- `SEND_SEPARATE_FULL_ORDER_EMAIL` (bool)
- `EMAIL_MANAGER_REPORT` (bool)

Google API:
- `SCOPES` (CSV; typical: Drive/Sheets/Gmail send)
- `FORCE_REAUTH` (bool)
- `REDIRECT_PORT` (int)
- `HTTP_TIMEOUT_SECONDS` (int)

Advanced intake / rollover:
- `USE_AUTO_ROLLOVER_IF_ONE_WEEK` (bool)
- `START_DAY_OF_WEEK`, `END_DAY_OF_WEEK`
  (Sunday|Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Any)

Cleanup (days):
- `OUTPUT_TIME_TO_LIFE`, `FAILED_INPUT_TIME_TO_LIFE`, `USER_TIME_TO_LIFE`

Drive override:
- `CONFIG_FILE_ID` (usually provided via `st.secrets`, if using a specific file)

-------------------------------------------------------------------------------
Operational Notes
-------------------------------------------------------------------------------

- **Lazy imports**: `streamlit` and Google client utilities are imported inside
  the loader so the module remains usable in non-Streamlit or headless contexts.
- **Fail-open Drive overrides**: If Drive credentials are unavailable or an
  override file cannot be retrieved/parsed, the loader returns the base config
  without raising (best-effort behavior).
- **Deterministic parsing**: Coercers are idempotent for already-typed values.
  For example, booleans in TOML remain booleans; CSV strings are split and
  trimmed; JSON strings are parsed into dicts.
- **Case normalization**: `REPORT_KEY_RUN_LIST` is uppercased at load time to
  minimize case-related mismatches elsewhere in the app.

Import this module early in your app to construct a single, consistent
`Config` instance and pass it through to components that require configuration.

"""


from __future__ import annotations

import os
import json
from dataclasses import dataclass, asdict, field
from typing import Any, Dict, List, Optional
from pathlib import Path
from dotenv import load_dotenv

# -----------------------------------------------------------------------------
# Helpers: read from Streamlit secrets (typed) or .env (strings) and coerce
# -----------------------------------------------------------------------------

def _get_secret(key: str, default: Any = None) -> Any:
    """
    Read from Streamlit secrets if present, else env var, else default.
    Does not raise if key missing; returns `default`.
    """
    try:
        import streamlit as st  # imported lazily to avoid hard dependency
        if hasattr(st, "secrets") and key in st.secrets:
            return st.secrets.get(key, default)
    except Exception:
        pass
    return os.getenv(key, default)

_TRUE = {"1", "true", "yes", "on", "y", "t"}

def _coerce_bool(v: Any, default: bool = False) -> bool:
    """
    Accept bool | str | int | None and return a Python bool.
    Works for typed TOML (bool) and .env strings.
    """
    if isinstance(v, bool):
        return v
    if v is None:
        return default
    try:
        return str(v).strip().lower() in _TRUE
    except Exception:
        return default

def _coerce_csv(v: Any) -> List[str]:
    """
    Accept list/tuple (already structured) or a comma-separated string.
    Returns a list of trimmed strings.
    """
    if v is None or v == "":
        return []
    if isinstance(v, (list, tuple)):
        return [str(x).strip() for x in v if str(x).strip()]
    return [p.strip() for p in str(v).split(",") if p.strip()]

def _coerce_json(v: Any) -> Dict[str, Any]:
    """
    Accept dict (already structured) or a JSON string.
    Returns a dict; falls back to {} on parse issues.
    """
    if v is None or v == "":
        return {}
    if isinstance(v, dict):
        return v
    try:
        return json.loads(v)
    except Exception:
        return {}

# -----------------------------------------------------------------------------
# Config dataclass (TOP-LEVEL — must start at column 0)
# -----------------------------------------------------------------------------

@dataclass
class Config:
    # IDs and basic settings
    CALC_SPREADSHEET_ID: str = "1ibkGkQ2khYMJydeenJkTzC4KoLQAyBZW_esQrbjSHXs"
    INCOMING_FOLDER_ID: str = "1jJE3r9DOHXwBdd94E6ZhxBBH9xvSjI-b"
    MANAGER_REPORT_FOLDER_ID: str = "17Nqwo6HYe30JP0wnZYoLRG0F1s-X-IVZ"
    ORDER_REPORT_FOLDER_ID: str = "171dqzMim-IdpB_kzjYQnzoSbW89uJTfP"
    ERROR_REPORT_FOLDER_ID: str = "1T-rnyXmPD1eFcxi-s8i4b1EP6-pW5ETW"
    USER_FOLDER_ID: str = "1JBHBcnS6397ka2ITW6Wbuu2aKjbgCCHj"

    # GIDs, sheet metadata, timestamp settings
    GID_MANAGER_PDF: str = "1921812573"
    GID_ORDER_CSV: str = "1875928148"
    GID_ERROR_REPORT: str = "1581903111"
    GID_BEV_ERRORS: str = "72711538"
    LOCATION_SHEET_TITLE: str = "REFR: Values"
    LOCATION_NAMED_RANGE: str = "_locations"
    TIMESTAMP_TZ: str = "America/Chicago"
    TIMESTAMP_FMT: str = "%Y-%m-%d-%I-%M-%p"
    TEMPLATE_UPDATE_RANGE: str = "_update"

    # Email config
    TO_RECIPIENTS: List[str] = field(default_factory=lambda: ["FavtripReporting@gmail.com"])
    CC_RECIPIENTS: List[str] = None
    ERROR_RECIPIENTS: List[str] = field(default_factory=lambda: ["FavtripReporting@gmail.com"])
    USE_ALL_REPORT_KEYS: bool = False
    REPORT_KEY_RUN_LIST: List[str] = field(default_factory=lambda: ["COFFEE"])
    REPORT_KEY_RECIPIENTS: Dict[str, List[str]] = None
    DEFAULT_ORDER_RECIPIENTS: List[str] = None
    INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL: bool = False
    SEND_SEPARATE_FULL_ORDER_EMAIL: bool = False
    EMAIL_MANAGER_REPORT: bool = True

    # Google API
    SCOPES: List[str] = None
    FORCE_REAUTH: bool = False
    REDIRECT_PORT: int = 58285
    HTTP_TIMEOUT_SECONDS: int = 300

    # Advanced intake settings
    USE_AUTO_ROLLOVER_IF_ONE_WEEK: bool = True
    START_DAY_OF_WEEK: str = "Sunday"    # Sunday, Monday, Tuesday, Wednesday, Thursday, Friday, Saturday, Any
    END_DAY_OF_WEEK: str = "Saturday"    # Sunday, Monday, Tuesday, Wednesday, Thursday, Friday, Saturday, Any

    # Cleanup
    OUTPUT_TIME_TO_LIFE: int = 30
    FAILED_INPUT_TIME_TO_LIFE: int = 1
    USER_TIME_TO_LIFE: int = 90

    #Other
    BEV_MAPPING_LINK: str = "https://docs.google.com/spreadsheets/d/1O6MtF-GM0VayqMr_v3oJC5PnRK5yv6biiDtA_qw-Z3g/"

    @staticmethod
    def load(env_path: Optional[Path] = None) -> "Config":
        """
        Load config from Streamlit secrets (preferred on cloud) or from env/.env (local dev),
        then overlay any values found in a Drive-backed JSON config (optional).
        Secrets may be typed (bool/list/dict), so we coerce safely.
        """
        if env_path is None:
            env_path = Path.cwd() / ".env"
        load_dotenv(dotenv_path=env_path, override=False)

        cfg = Config(
            CALC_SPREADSHEET_ID=str(_get_secret("CALC_SPREADSHEET_ID", "")),
            INCOMING_FOLDER_ID=str(_get_secret("INCOMING_FOLDER_ID", "")),
            MANAGER_REPORT_FOLDER_ID=str(_get_secret("MANAGER_REPORT_FOLDER_ID", "")),
            ORDER_REPORT_FOLDER_ID=str(_get_secret("ORDER_REPORT_FOLDER_ID", "")),
            ERROR_REPORT_FOLDER_ID=str(_get_secret("ERROR_REPORT_FOLDER_ID", "")),
            USER_FOLDER_ID=str(_get_secret("USER_FOLDER_ID", "")),

            GID_MANAGER_PDF=str(_get_secret("GID_MANAGER_PDF", "1921812573")),
            GID_ORDER_CSV=str(_get_secret("GID_ORDER_CSV", "1875928148")),
            GID_ERROR_REPORT=str(_get_secret("GID_ERROR_REPORT", "1581903111")),
            GID_BEV_ERRORS=str(_get_secret("GID_BEV_ERRORS", "72711538")),
            LOCATION_SHEET_TITLE=str(_get_secret("LOCATION_SHEET_TITLE", "REFR: Values")),
            LOCATION_NAMED_RANGE=str(_get_secret("LOCATION_NAMED_RANGE", "_locations")),
            TEMPLATE_UPDATE_RANGE=str(_get_secret("TEMPLATE_UPDATE_RANGE", "_update")),
            TIMESTAMP_TZ=str(_get_secret("TIMESTAMP_TZ", "America/Chicago")),
            TIMESTAMP_FMT=str(_get_secret("TIMESTAMP_FMT", "%Y-%m-%d-%I-%M-%p")),

            OUTPUT_TIME_TO_LIFE=int(_get_secret("OUTPUT_TIME_TO_LIFE", 30)),
            FAILED_INPUT_TIME_TO_LIFE=int(_get_secret("FAILED_INPUT_TIME_TO_LIFE", 1)),
            USER_TIME_TO_LIFE=int(_get_secret("USER_TIME_TO_LIFE", 1)),

            TO_RECIPIENTS=_coerce_csv(_get_secret("TO_RECIPIENTS", "")),
            CC_RECIPIENTS=_coerce_csv(_get_secret("CC_RECIPIENTS", "")),
            ERROR_RECIPIENTS=_coerce_csv(_get_secret("ERROR_RECIPIENTS", "")),
            USE_ALL_REPORT_KEYS=_coerce_bool(_get_secret("USE_ALL_REPORT_KEYS", "false")),
            REPORT_KEY_RUN_LIST=[s.upper() for s in _coerce_csv(_get_secret("REPORT_KEY_RUN_LIST", ""))],
            REPORT_KEY_RECIPIENTS=_coerce_json(_get_secret("REPORT_KEY_RECIPIENTS", "{}")),
            DEFAULT_ORDER_RECIPIENTS=_coerce_csv(_get_secret("DEFAULT_ORDER_RECIPIENTS", "")),
            INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL=_coerce_bool(
                _get_secret("INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL", "false")
            ),
            SEND_SEPARATE_FULL_ORDER_EMAIL=_coerce_bool(
                _get_secret("SEND_SEPARATE_FULL_ORDER_EMAIL", "true")
            ),
            EMAIL_MANAGER_REPORT=_coerce_bool(_get_secret("EMAIL_MANAGER_REPORT", "true")),

            SCOPES=_coerce_csv(
                _get_secret(
                    "SCOPES",
                    "https://www.googleapis.com/auth/drive,"
                    "https://www.googleapis.com/auth/spreadsheets,"
                    "https://www.googleapis.com/auth/gmail.send",
                )
            ),
            FORCE_REAUTH=_coerce_bool(_get_secret("FORCE_REAUTH", "false")),
            REDIRECT_PORT=int(str(_get_secret("REDIRECT_PORT", "58285")) or "58285"),
            HTTP_TIMEOUT_SECONDS=int(str(_get_secret("HTTP_TIMEOUT_SECONDS", "300")) or "300"),

            USE_AUTO_ROLLOVER_IF_ONE_WEEK=_coerce_bool(_get_secret("USE_AUTO_ROLLOVER_IF_ONE_WEEK", "true")),
            START_DAY_OF_WEEK=str(_get_secret("START_DAY_OF_WEEK", "Sunday")),
            END_DAY_OF_WEEK=str(_get_secret("END_DAY_OF_WEEK", "Saturday")),

            BEV_MAPPING_LINK=str(_get_secret("BEV_MAPPING_LINK", "https://docs.google.com/spreadsheets/d/1O6MtF-GM0VayqMr_v3oJC5PnRK5yv6biiDtA_qw-Z3g/")),
        )

        normalized = {}
        for k, v in cfg.REPORT_KEY_RECIPIENTS.items():
            if isinstance(k, (list, tuple)):
                if len(k) == 2:
                    normalized[(k[0], k[1], None)] = v
                elif len(k) == 3:
                    normalized[tuple(k)] = v
            else:
                # defensive fallback
                normalized[(None, k, None)] = v

        cfg.REPORT_KEY_RECIPIENTS = normalized


        # Optional overlay from Drive JSON config (if creds + file present)
        # ---------------- Drive-backed config overlay ----------------
        try:
            import streamlit as st
            from core_functional_modules.google_client import load_valid_token, services
            from core_functional_modules.config_store import load_config_from_drive

            DEV_ENVIRONMENT = _coerce_bool(_get_secret("DEV_ENVIRONMENT", False))
            DEV_CONFIG_FILE_ID = str(_get_secret("DEV_CONFIG_FILE_ID", "") or "").strip()
            CONFIG_FILE_ID = str(_get_secret("CONFIG_FILE_ID", "") or "").strip()

            # Select which config file ID to READ from
            if DEV_ENVIRONMENT and DEV_CONFIG_FILE_ID:
                active_config_file_id = DEV_CONFIG_FILE_ID
            else:
                active_config_file_id = CONFIG_FILE_ID or None

            creds = load_valid_token(cfg.SCOPES)
            if creds:
                _sheets, drive, _gmail = services(creds, cfg.HTTP_TIMEOUT_SECONDS)

                overrides = {}
                
                # 1️⃣ Try DEV config first (if enabled)
                if DEV_ENVIRONMENT and DEV_CONFIG_FILE_ID:
                    overrides = load_config_from_drive(drive, DEV_CONFIG_FILE_ID)

                # 2️⃣ Fallback to PROD config if DEV missing/empty
                if not overrides and CONFIG_FILE_ID:
                    overrides = load_config_from_drive(drive, CONFIG_FILE_ID)

                # 3️⃣ Apply overrides if any
                if isinstance(overrides, dict):
                    for k, v in overrides.items():
                        if hasattr(cfg, k):
                            setattr(cfg, k, v)

        except Exception:
            # Fail-open by design
            
            import traceback
            print("[Config] Drive overlay failed:", e)
            traceback.print_exc()

            #pass
        
        return cfg


    # -------------------------------------------------------------------------
    # .env serialization (optional helper)
    # -------------------------------------------------------------------------
    def to_env(self) -> str:
        """Serialize to .env format (simple, string-based)."""
        data = asdict(self)
        as_env = {
            **data,
            "TO_RECIPIENTS": ",".join(self.TO_RECIPIENTS or []),
            "CC_RECIPIENTS": ",".join(self.CC_RECIPIENTS or []),
            "REPORT_KEY_RUN_LIST": ",".join(self.REPORT_KEY_RUN_LIST or []),
            "REPORT_KEY_RECIPIENTS": json.dumps(self.REPORT_KEY_RECIPIENTS or {}),
            "DEFAULT_ORDER_RECIPIENTS": ",".join(self.DEFAULT_ORDER_RECIPIENTS or []),
            "SCOPES": ",".join(self.SCOPES or []),
        }
        lines = [f"{k}={v}" for k, v in as_env.items()]
        return "\n".join(lines) + "\n"

    def save(self, env_path: Optional[Path] = None):
        if env_path is None:
            env_path = Path.cwd() / ".env"
        env_path.write_text(self.to_env(), encoding="utf-8")

```

---
### file: core_functional_modules/config_store.py

```python
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

```

---
### file: core_functional_modules/drive_utils.py

```python
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

def get_or_create_subfolder(drive_svc, parent_folder_id: str, name: str):
    """
    Return a Drive folder with the given name under parent_folder_id.
    Create it if it does not already exist.
    """
    q = (
        f"mimeType='application/vnd.google-apps.folder' "
        f"and name='{name}' "
        f"and '{parent_folder_id}' in parents "
        f"and trashed=false"
    )

    res = drive_svc.files().list(
        q=q,
        fields="files(id, name, webViewLink)",
        pageSize=1
    ).execute()

    files = res.get("files", [])
    if files:
        return files[0]

    metadata = {
        "name": name,
        "mimeType": "application/vnd.google-apps.folder",
        "parents": [parent_folder_id],
    }

    return drive_svc.files().create(
        body=metadata,
        fields="id, name, webViewLink"
    ).execute()


```

---
### file: core_functional_modules/gmail_utils.py

```python
""" 
gmail_utils
======================================
Email utilities for sending Gmail messages with PDF attachments via the Gmail API.

This module provides small, focused helpers for sending two types of emails through
the Gmail API using a pre-authorized `gmail_svc` client (e.g., returned by
`googleapiclient.discovery.build("gmail", "v1", ...)`). It includes:

- `send_email(...)`: Low-level helper that accepts an `email.message.EmailMessage`,
  base64-url encodes it as required by Gmail, and dispatches it via
  `users.messages.send`.

- `email_manager_report(...)`: Composes and sends a standardized "Manager Report"
  email with a primary PDF attachment and a backup link. Supports optional CC.

- `email_order_report(...)`: Composes and sends an "Order Report" email for a
  given vendor or category key, including a primary PDF attachment and an optional
  "full order" PDF. Also includes links to a backing Google Sheet and supports CC.

The functions here intentionally perform minimal validation and assume that callers
supply valid addresses, attachments, and links. Authentication, token refresh, and
error handling policy (e.g., retries, backoff, alerting) should be implemented by
the caller.

---
Key Behaviors
-------------
- **MIME construction**: Uses Python's stdlib `email.message.EmailMessage` to build
  multipart emails with both plain-text and HTML alternatives, and PDF attachments.
- **Gmail API compliance**: Serializes the email to bytes and encodes it with
  URL-safe Base64 as required by Gmail's `users.messages.send` endpoint.
- **Idempotency**: Sending is not idempotent; calling functions repeatedly may
  result in duplicate emails. Callers should implement their own guardrails if
  needed (e.g., deduplication keys, sent-flagging).
- **Internationalization**: The functions do not localize content; callers can adapt
  the text if i18n is required.
- **HTML content**: Simple HTML bodies are included via `add_alternative(..., subtype="html")`.
  The HTML snippets intentionally avoid external assets for reliable delivery.

---
Functions
---------
send_email(gmail_svc, user, msg)
    Low-level send helper. Encodes the `EmailMessage` and dispatches via the Gmail API.

email_manager_report(gmail_svc, sender, to_list, cc_list, pdf_name, pdf_bytes, pdf_link, ts, location)
    Sends a standardized "Manager Report" email with a PDF attachment and a backup link.

email_order_report(
    gmail_svc,
    sender,
    to_list,
    cc_list,
    key,
    tag,
    ts,
    location,
    pdf_name,
    pdf_bytes,
    sheet_link,
    include_full_order=False,
    full_pdf_bytes=None,
    full_pdf_name=None,
)
    Sends an "Order Report" email targeted to a `{key}` team with a primary PDF,
    optional full-order PDF, and a link to the backing Google Sheet.

---
Parameters (Shared Concepts)
----------------------------
gmail_svc : Any
    An authenticated Gmail API service client (e.g., from `googleapiclient.discovery.build`).

sender : str
    The "From" email address to display in the message header. The authenticated
    Gmail account must be authorized to send from this address.

to_list : Iterable[str]
    Recipient email addresses for the `To` field. Must contain at least one valid address.

cc_list : Optional[Iterable[str]]
    Optional CC recipient addresses. If empty or `None`, the `Cc` header is omitted.

pdf_name : str
    Filename for the attached PDF (e.g., `"report_2026-03-21.pdf"`).

pdf_bytes : bytes
    Raw bytes of the primary PDF attachment.

ts : str
    A timestamp string suitable for inclusion in the subject (e.g., `"2026-03-21"` or
    `"2026-03-21 18:25"`).

location : str
    A human-readable location name included in the subject/body (e.g., store or site).

pdf_link : str
    (Manager Report) A backup URL users can access if attachments are blocked.

key : str
    (Order Report) An identifier for the receiving team or vendor (e.g., `"Dairy"`, `"VendorX"`).

tag : str
    (Order Report) A secondary descriptor (e.g., `"Weekly"`, `"Overstock"`, `"Emergency"`).

sheet_link : str
    (Order Report) URL to the backing Google Sheet with order details.

include_full_order : bool
    (Order Report) Whether to attach an additional "full order" PDF.

full_pdf_bytes : Optional[bytes]
    (Order Report) Raw bytes of the full order PDF (required when `include_full_order=True`).

full_pdf_name : Optional[str]
    (Order Report) Filename for the full order PDF (required when `include_full_order=True`).

user : str
    (send_email) Gmail user identifier for the API call. Typically `"me"` to refer
    to the authenticated account.

msg : EmailMessage
    (send_email) A fully-constructed email message to be sent.

---
Returns
-------
dict
    The Gmail API response payload from `users.messages.send()` (e.g., includes `id`, `threadId`).

---
Raises
------
googleapiclient.errors.HttpError
    If the Gmail API call fails (e.g., quota exceeded, invalid permissions, bad request).
ValueError / TypeError
    If provided inputs (addresses, bytes, filenames) are invalid (may be raised by stdlib or caller validations).

---
"""


from __future__ import annotations
import base64
from email.message import EmailMessage


def send_email(gmail_svc, user: str, msg: EmailMessage):
    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
    return gmail_svc.users().messages().send(userId=user, body={"raw": raw}).execute()


def email_manager_report(gmail_svc, sender: str, to_list, cc_list, pdf_name, pdf_bytes, pdf_link, ts, location):
    msg = EmailMessage()
    msg["Subject"] = f"Manager Report – {location} – {ts}"
    msg["From"] = sender
    msg["To"] = ", ".join(to_list)
    if cc_list:
        msg["Cc"] = ", ".join(cc_list)
    msg.set_content(f"Hi team,\nAttached is the Manager Report ({location}).\nBackup link: {pdf_link}\n—Sent from an automated reporting pipeline")

    msg.add_alternative(
        f"""
        <p>Hi team,</p>
        <p>Your manager report for store <b>{location}</b> is ready.</p>
        <p><a href='{pdf_link}'>Backup Link</a></p>
        <p>Attached: {pdf_name}</p>
        <p>—Sent from an automated reporting pipeline</p>
        """,
        subtype="html",
    )

    msg.add_attachment(pdf_bytes, maintype="application", subtype="pdf", filename=pdf_name)
    return send_email(gmail_svc, sender, msg)


def email_order_report(
    gmail_svc,
    sender: str,
    to_list,
    cc_list,
    key: str,
    tag: str,
    ts: str,
    location: str,
    pdf_name: str,
    pdf_bytes: bytes,
    sheet_link: str,
    include_full_order: bool = False,
    full_pdf_bytes: bytes | None = None,
    full_pdf_name: str | None = None,
):
    msg = EmailMessage()

    msg["Subject"] = f"Order Report – {location} – {tag} – {ts}"
    msg["From"] = sender
    msg["To"] = ", ".join(to_list)

    if cc_list:
        msg["Cc"] = ", ".join(cc_list)

    msg.set_content(
        f"Hi {tag} team,\n"
        f"Your order report for {location} - {tag} is ready.\n"
        f"Google Sheet: {sheet_link}\n"
        f"Attached: {pdf_name}\n"
        "—Sent from an automated reporting pipeline"
    )

    msg.add_alternative(
        f"""
        <p>Hi {tag} team,</p>
        <p>Your order report for store <b>{location}</b> is ready.</p>
        <p><a href="{sheet_link}">Open Google Sheet</a></p>
        <p>Attached: {pdf_name}</p>
        <p>—Sent from an automated reporting pipeline</p>
        """,
        subtype="html",
    )

    msg.add_attachment(
        pdf_bytes,
        maintype="application",
        subtype="pdf",
        filename=pdf_name,
    )

    if include_full_order and full_pdf_bytes:
        msg.add_attachment(
            full_pdf_bytes,
            maintype="application",
            subtype="pdf",
            filename=full_pdf_name,
        )

    return send_email(gmail_svc, sender, msg)


def email_error_report(
    gmail_svc,
    sender: str,
    to_list,
    cc_list,
    ts: str,
    pdf_name: str,
    pdf_bytes: bytes,
    sheet_link: str
    ):
    msg = EmailMessage()

    msg["Subject"] = f"Error Report – {ts}"
    msg["From"] = sender
    msg["To"] = ", ".join(to_list)

    if cc_list:
        msg["Cc"] = ", ".join(cc_list)

    msg.set_content(
        f"Hi Technical Support Team,\n"
        f"A user tried using the reporting pipeline, however some of the items that were uploaded are not listed on the Vendor Price Book.\n"
        f"The default recipient of the pipeline run is CC'd on this email for visibility and communication purposes.\n"
        f"Please reply to this email once the Vendor Price Book is updated so that the user knows they can rerun the pipeline.\n\n"
        f"Google Sheet: {sheet_link}\n"
        f"Attached: {pdf_name}\n"
        "—Sent from an automated reporting pipeline"
    )

    msg.add_alternative(
        f"""
        <p>Hi Technical Support Team,</p>
        <p>A user tried using the reporting pipeline, however some of the items that were uploaded are not listed on the Vendor Price Book.</p>
        <p>The default recipient of the pipeline run is CC'd on this email for visibility and communication purposes.</p>
        <p>Please reply to this email once the Vendor Price Book is updated so that the user knows they can rerun the pipeline.</p>
        <p></p>
        <p><a href="{sheet_link}">Open Error Report in Google Sheets</a></p>
        <p>Attached: {pdf_name}</p>
        <p>—Sent from an automated reporting pipeline</p>
        """,
        subtype="html",
    )

    msg.add_attachment(
        pdf_bytes,
        maintype="application",
        subtype="pdf",
        filename=pdf_name,
    )

    return send_email(gmail_svc, sender, msg)

def email_bev_error_report(
    gmail_svc,
    sender: str,
    to_list,
    cc_list,
    ts: str,
    pdf_name: str,
    pdf_bytes: bytes,
    sheet_link: str,
    mapping_link: str
    ):
    msg = EmailMessage()

    msg["Subject"] = f"Unassigned Beverages Report – {ts}"
    msg["From"] = sender
    msg["To"] = ", ".join(to_list)

    if cc_list:
        msg["Cc"] = ", ".join(cc_list)

    msg.set_content(
        f"Hi Technical Support Team,\n"
        f"A user just ran the ordering pipeline successfully, however some beverages were not mapped to a sub-report-key / vendor.\n"
        f"The BEV orders were still sent, however these unassigned beverages were included on their own order under BEV - UNASSIGNED.\n\n"
        f"The default recipient of the pipeline run is CC'd on this email for visibility and communication purposes.\n"
        f"To fix this error for future runs look at the Unassigned Beverages Report and add those Scan Codes to the Mapping File.\n"
        f"Please reply to this email once the Mapping File is updated so that the user knows they can rerun the pipeline if needed.\n\n"
        f"Beverage Mapping File: {mapping_link}\n"
        f"Unassigned Beverages Google Sheet: {sheet_link}\n"
        f"Attached: {pdf_name}\n"
        "—Sent from an automated reporting pipeline"
    )

    msg.add_alternative(
        f"""
        <p>Hi Technical Support Team,</p>
        <p>A user just ran the ordering pipeline successfully, however some beverages were not mapped to a sub-report-key / vendor.</p>
        <p>The BEV orders were still sent; however, these unassigned beverages were included on their own order under <strong>BEV - UNASSIGNED</strong>.</p><p></p>
        <p>The default recipient of the pipeline run is CC'd on this email for visibility and communication purposes.</p>
        <p>To fix this error for future runs, please review the Unassigned Beverages Report and add those Scan Codes to the Mapping File.</p>
        <p>Please reply to this email once the Mapping File is updated so that the user knows they can rerun the pipeline if needed.</p><p></p>
        <p><a href="{mapping_link}">Open Beverage Mapping File</a></p>
        <p><a href="{sheet_link}">Open Unassigned Beverages Google Sheet</a></p>
        <p>Attached: {pdf_name}</p>
        <p>—Sent from an automated reporting pipeline</p>
        """,
        subtype="html",
    )

    msg.add_attachment(
        pdf_bytes,
        maintype="application",
        subtype="pdf",
        filename=pdf_name,
    )

    return send_email(gmail_svc, sender, msg)
```

---
### file: core_functional_modules/google_client.py

```python
"""
google_client
======================================
This module centralizes Google OAuth 2.0 sign‑in for local Python applications,
supporting both:

1) **Classic CLI flow** – prints an authorization URL to the console and accepts
   a pasted redirect URL or auth code (for headless shells, remote SSH, or when
   opening a browser is impractical).

2) **Streamlit-/Desktop-friendly local-server flow** – opens the user's default
   browser and spins up a temporary HTTP listener on ``127.0.0.1`` to complete
   the OAuth redirect without any console copy/paste.

It also provides small helpers to manage a persistent token cache
(``token.json``), refresh expired tokens when possible, and construct Google API
service clients (Sheets, Drive, Gmail) using the official ``google-api-python-client``.

-------------------------------------------------------------------------------
Key Features
-------------------------------------------------------------------------------
- **Token cache**: Reads/writes ``token.json`` to persist credentials between runs.
  Includes best-effort refresh of expired tokens when a refresh token exists.
- **Two auth paths**:
  - *CLI/manual path*: URL is printed; user pastes back the full redirect URL or
    just the ``code`` parameter.
  - *Local server/one-click path*: Automatically opens browser and listens on a
    local port (tries an OS-chosen free port first, then a configured fallback).
- **Graceful fallbacks**: If automated browser auth fails, raises a descriptive
  error suggesting the manual method.
- **Service builders**: Convenience helpers to create Sheets v4, Drive v3, and
  Gmail v1 service clients with the provided credentials.

-------------------------------------------------------------------------------
Files Used
-------------------------------------------------------------------------------
- ``credentials.json`` (required):
  The OAuth 2.0 client secrets file downloaded from Google Cloud Console.

- ``token.json`` (optional, auto-created):
  The persisted user credentials (access/refresh tokens). If present and valid,
  it is reused to avoid re-authentication. If expired but refreshable, it is
  refreshed automatically and re-written.

-------------------------------------------------------------------------------
Function Overview
-------------------------------------------------------------------------------
- ``clear_token()``:
    Deletes ``token.json`` if present (best-effort). Useful to force a
    re-authentication scenario.

- ``load_valid_token(scopes) -> Optional[Credentials]``:
    Loads credentials from ``token.json`` for the given scopes. If expired but
    refreshable, refreshes and persists the updated token. Returns a valid
    ``Credentials`` or ``None``.

- ``get_credentials(scopes, redirect_port, force_reauth=False) -> Credentials``:
    **CLI-friendly** method. If no valid token exists, prints an auth URL and
    prompts for a pasted redirect URL or code. Persists the resulting token to
    ``token.json``.

- ``login_via_local_server(scopes, redirect_port) -> Credentials``:
    **Streamlit-/desktop-friendly** one-click OAuth that opens a browser and
    listens on ``127.0.0.1``. Tries an OS-chosen free port first (``port=0``),
    then the provided ``redirect_port``. Uses a 120s timeout for safety.

- ``start_oauth(scopes, redirect_port) -> (InstalledAppFlow, auth_url)``:
    Starts the manual flow by creating an ``InstalledAppFlow`` with a configured
    redirect URI and returns the authorization URL to display in your own UI.

- ``finish_oauth(flow, pasted) -> Credentials``:
    Completes the manual flow using the pasted redirect URL (or raw ``code``),
    fetches tokens, writes ``token.json``, and returns ``Credentials``.

- ``_service(api, version, creds)``:
    Internal helper to construct a Google API service for the given
    ``api``/``version`` using the supplied ``Credentials``.

- ``services(creds, _http_timeout_seconds)``:
    Convenience function returning a tuple of ready-to-use clients:
    ``(sheets, drive, gmail)``. The ``_http_timeout_seconds`` parameter is
    currently reserved for future use.

-------------------------------------------------------------------------------
Error Handling & Edge Cases
-------------------------------------------------------------------------------
- If ``credentials.json`` is missing, a ``FileNotFoundError`` is raised early.
- Token refresh failures fall back to a fresh login.
- The local-server path uses a 120-second timeout to avoid hanging the process.
- If both automatic local-server attempts fail, a ``RuntimeError`` is raised
  advising the manual copy/paste method with detailed error messages from both
  attempts.
-------------------------------------------------------------------------------
Maintainer Tips
-------------------------------------------------------------------------------
- If you add new Google APIs, extend ``services(...)`` or call ``_service(...)``
  directly with the desired API name/version.
- Consider surfacing the timeout and host/port as user configuration if your app
  needs more control in diverse environments.

"""

from __future__ import annotations
import os
from urllib.parse import urlparse, parse_qs

from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build


# ---------- Token helpers ----------

def clear_token():
    """Delete token.json if present."""
    try:
        if os.path.exists("token.json"):
            os.remove("token.json")
    except Exception:
        pass


def load_valid_token(scopes):
    """
    Try to load token.json. If expired but refreshable, refresh it and persist.
    Returns valid Credentials or None.
    """
    if not os.path.exists("token.json"):
        return None
    try:
        creds = Credentials.from_authorized_user_file("token.json", scopes)
    except Exception:
        return None

    if creds and creds.valid:
        return creds

    if creds and creds.expired and creds.refresh_token:
        try:
            creds.refresh(Request())
            with open("token.json", "w") as f:
                f.write(creds.to_json())
            return creds
        except Exception:
            return None

    return None


# ---------- Classic CLI path (kept for completeness) ----------

def get_credentials(scopes, redirect_port: int, force_reauth: bool = False) -> Credentials:
    """
    CLI-friendly: prints URL and waits for input() if token is missing/invalid.
    The Streamlit UI uses the in-UI functions below instead.
    """
    if force_reauth:
        clear_token()

    creds = load_valid_token(scopes)
    if creds:
        return creds

    if not os.path.exists("credentials.json"):
        raise FileNotFoundError("Missing credentials.json in working directory")

    flow = InstalledAppFlow.from_client_secrets_file("credentials.json", scopes)
    flow.redirect_uri = f"http://127.0.0.1:{redirect_port}/"
    auth_url, _ = flow.authorization_url(
        access_type="offline",
        include_granted_scopes="true",
        prompt="consent",
    )
    print("Open this URL and complete the login:\n", auth_url)
    pasted = input("Paste full redirect URL or auth code here: ").strip()
    code = pasted
    if pasted.startswith("http"):
        qs = parse_qs(urlparse(pasted).query)
        if "code" in qs:
            code = qs["code"][0]
    flow.fetch_token(code=code)
    creds = flow.credentials
    with open("token.json", "w") as f:
        f.write(creds.to_json())
    return creds


# ---------- Streamlit-friendly OAuth (no console) ----------

# favtrip/google_client.py

def login_via_local_server(scopes, redirect_port: int) -> Credentials:
    """
    One-click OAuth: open browser and listen on 127.0.0.1.
    Tries OS-chosen port first, then the configured port.
    Uses a timeout to avoid hanging indefinitely.
    NOTE: No optional text parameters are passed, for compatibility with older google-auth-oauthlib.
    """
    if not os.path.exists("credentials.json"):
        raise FileNotFoundError("Missing credentials.json in working directory")

    # Attempt 1: OS-chosen free port (port=0)
    flow = InstalledAppFlow.from_client_secrets_file("credentials.json", scopes)
    try:
        creds = flow.run_local_server(
            host="127.0.0.1",
            port=0,                 # let OS choose a free port
            open_browser=True,
            timeout_seconds=120,    # bail out after 2 minutes
        )
        with open("token.json", "w") as f:
            f.write(creds.to_json())
        return creds
    except Exception as first_err:
        # Attempt 2: user-configured port (from .env)
        flow = InstalledAppFlow.from_client_secrets_file("credentials.json", scopes)
        try:
            creds = flow.run_local_server(
                host="127.0.0.1",
                port=int(redirect_port),
                open_browser=True,
                timeout_seconds=120,
            )
            with open("token.json", "w") as f:
                f.write(creds.to_json())
            return creds
        except Exception as second_err:
            raise RuntimeError(
                "Automatic browser auth failed both on a random port and on your configured REDIRECT_PORT. "
                "Please use the manual method (copy/paste URL). "
                f"Details: first={first_err}; second={second_err}"
            )


def start_oauth(scopes, redirect_port: int):
    """
    Manual fallback: returns (flow, auth_url) for paste-based completion.
    """
    if not os.path.exists("credentials.json"):
        raise FileNotFoundError("Missing credentials.json in working directory")
    flow = InstalledAppFlow.from_client_secrets_file("credentials.json", scopes)
    flow.redirect_uri = f"http://127.0.0.1:{redirect_port}/"
    auth_url, _ = flow.authorization_url(
        access_type="offline",
        include_granted_scopes="true",
        prompt="consent",
    )
    return flow, auth_url


def finish_oauth(flow: InstalledAppFlow, pasted: str) -> Credentials:
    """
    Manual fallback: accepts the pasted redirect URL or the code; returns Credentials and writes token.json.
    """
    code = pasted.strip()
    if pasted.startswith("http"):
        qs = parse_qs(urlparse(pasted).query)
        if "code" in qs:
            code = qs["code"][0]
    flow.fetch_token(code=code)
    creds = flow.credentials
    with open("token.json", "w") as f:
        f.write(creds.to_json())
    return creds


# ---------- Google services ----------

def _service(api: str, version: str, creds: Credentials):
    # Pass credentials directly (no google_auth_httplib2 dependency)
    return build(api, version, credentials=creds, cache_discovery=False)


def services(creds: Credentials, _http_timeout_seconds: int):
    sheets = _service("sheets", "v4", creds)
    drive = _service("drive", "v3", creds)
    gmail = _service("gmail", "v1", creds)
    return sheets, drive, gmail

```

---
### file: core_functional_modules/logger.py

```python
"""
logger
======================================
This module provides two dataclasses—`LogEvent` and `StatusLogger`—to record simple,
human-readable status messages during a process or script run. It is designed to be:

- **Simple**: minimal API (`info`, `warn`, `error`) and a small in-memory log.
- **Immediate**: console prints occur synchronously; file writes are line-buffered and flushed.
- **Fail-open**: if a log file cannot be opened or written, logging proceeds to console and memory.
- **Portable**: standard library only (dataclasses, datetime, typing).

-------------------------------------------------------------------------------
Data Model
-------------------------------------------------------------------------------
- LogEvent
    - ts (datetime.datetime): Timestamp captured via `datetime.now()` when the event is recorded.
      Note: this is a **naive** datetime in local time.
    - level (str): Log level label (e.g., "INFO", "WARN", "ERROR").
    - message (str): The event text.

- StatusLogger
    - events (list[LogEvent]): In-memory event history in append order.
    - print_to_console (bool): If True (default), each log line is printed to stdout.
    - file_path (str | None): If set, lines are also written to this file. If `None`, file logging
      is disabled. Default is "last_run.log".
    - overwrite (bool): If True (default), the log file is opened in write mode on instantiation;
      otherwise it is appended to.

-------------------------------------------------------------------------------
Output Format
-------------------------------------------------------------------------------
- Console/file lines: `[YYYY-MM-DD HH:MM:SS] LEVEL: message`
- `as_text()`:         `[HH:MM:SS] LEVEL: message` per line (no date, suitable for compact display)
- `last_line()`:       Returns the most recent line in `as_text()` format, or `"Starting…"` if empty.

-------------------------------------------------------------------------------
Behavior & Guarantees
-------------------------------------------------------------------------------
- **File handling**: On initialization, if `file_path` is provided, the file is opened once in
  line-buffered text mode (`buffering=1`) and UTF-8 encoding. If opening fails, the logger
  continues without a file handle.
- **Atomicity**: Each `_emit` call attempts to write a single line and then flush. Any file write
  errors are swallowed; console output and in-memory storage are unaffected.
- **Timestamps**: Timestamps are captured at call time (`datetime.now()`), local time, naive datetimes.
- **Memory growth**: All events are retained in `events`; for long-running processes, consider
  pruning or exporting periodically.
- **Thread-safety**: Not thread-safe. If you need concurrent logging, protect calls with a lock or
  adapt the implementation for multi-thread/process usage.
- **No rotation**: No file rotation or size limiting. Use external tools or extend as needed.
"""

from dataclasses import dataclass, field
from datetime import datetime
from typing import List, Optional

@dataclass
class LogEvent:
    ts: datetime
    level: str
    message: str

@dataclass
class StatusLogger:
    events: List[LogEvent] = field(default_factory=list)
    print_to_console: bool = True
    file_path: Optional[str] = "last_run.log"
    overwrite: bool = True

    def __post_init__(self):
        # Prepare the file on first use
        self._fh = None
        if self.file_path:
            mode = "w" if self.overwrite else "a"
            try:
                self._fh = open(self.file_path, mode, encoding="utf-8", buffering=1)  # line-buffered
            except Exception:
                # If we cannot open a file, we keep running without file logging
                self._fh = None

    def _emit(self, line: str):
        if self.print_to_console:
            print(line)
        if self._fh:
            try:
                self._fh.write(line + "\n")
                self._fh.flush()  # ensure immediate persistence
            except Exception:
                pass

    def _log(self, level: str, message: str):
        evt = LogEvent(datetime.now(), level, message)
        self.events.append(evt)
        self._emit(f"[{evt.ts:%Y-%m-%d %H:%M:%S}] {level}: {message}")

    def info(self, message: str):
        self._log("INFO", message)

    def warn(self, message: str):
        self._log("WARN", message)

    def error(self, message: str):
        self._log("ERROR", message)

    def as_text(self) -> str:
        return "\n".join(f"[{e.ts:%H:%M:%S}] {e.level}: {e.message}" for e in self.events)

    def last_line(self) -> str:
        if not self.events:
            return "Starting…"
        e = self.events[-1]
        return f"[{e.ts:%H:%M:%S}] {e.level}: {e.message}"

    def close(self):
        try:
            if self._fh:
                self._fh.close()
        except Exception:
            pass

```

---
### file: core_functional_modules/pipeline.py

```python
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
    add_or_replace_sheet, put_values_2d, _force_column_as_text, delete_row_indices, delete_rows_range, copy_sheet_to_another_spreadsheet, autoresize_columns, export_sheet
)
from .drive_utils import find_latest_sheet, upload_to_drive, _rfc3339, trash_file, cleanup_folder_by_age, find_sheet_by_name, copy_file_to_folder, rename_file, get_or_create_subfolder
from .gmail_utils import send_email, email_manager_report, email_order_report, email_error_report, email_bev_error_report

CSV_MIME = "text/csv"


def clean_tag(s: str) -> str:
    return re.sub(r"[^A-Za-z0-9._-]+", "-", s.strip()).strip("-") or "UNKNOWN"


import requests
from io import BytesIO
from openpyxl import Workbook


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


def filter_master_csv_to_ran(master_csv_bytes, cfg):
    text = master_csv_bytes.decode("utf-8-sig", errors="replace")
    reader = csv.reader(io.StringIO(text))
    rows = list(reader)

    if not rows:
        return master_csv_bytes  # nothing to do

    headers = [h.strip() for h in rows[0]]
    lower_idx = {h.lower(): i for i, h in enumerate(headers)}

    report_idx = lower_idx.get("report_key")
    sub_idx = lower_idx.get("sub_report_key")

    if report_idx is None:
        raise RuntimeError("Master CSV missing Report_Key column")

    output = io.StringIO()
    writer = csv.writer(output, lineterminator="\n")
    writer.writerow(headers)

    for r in rows[1:]:
        key = (r[report_idx] if report_idx < len(r) else "").strip().upper()
        sub = None
        if sub_idx is not None:
            sub = (r[sub_idx] if sub_idx < len(r) else "").strip().upper() or None

        if should_run(cfg, key, sub):
            writer.writerow(r)

    return output.getvalue().encode("utf-8-sig")

def sort_master_csv(csv_bytes: bytes) -> bytes:
    df = pandas.read_csv(io.BytesIO(csv_bytes))

    # Normalize column names
    cols = {c.lower(): c for c in df.columns}

    report_col = cols.get("report_key")
    sub_col = cols.get("sub_report_key")
    cases_col = cols.get("cases to order")

    if not report_col:
        raise RuntimeError("Report_Key column not found for sorting")

    # Fill missing sub-keys so they sort last
    if sub_col:
        df[sub_col] = df[sub_col].fillna("ZZZ")

    # Ensure numeric sorting for cases
    if cases_col:
        df[cases_col] = pandas.to_numeric(df[cases_col], errors="coerce").fillna(0)

    sort_cols = [report_col]
    sort_ascending = [True]

    if sub_col:
        sort_cols.append(sub_col)
        sort_ascending.append(True)

    if cases_col:
        sort_cols.append(cases_col)
        sort_ascending.append(False)  # DESCENDING

    df = df.sort_values(
        by=sort_cols,
        ascending=sort_ascending,
        kind="mergesort"  # stable sort
    )

    # Restore blanks if we filled them
    if sub_col:
        df[sub_col] = df[sub_col].replace("ZZZ", "")

    out = io.StringIO()
    df.to_csv(out, index=False)
    return out.getvalue().encode("utf-8-sig")


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

    
    ran_report_keys = set()
    ran_sub_report_keys = set()
    ran_reports = {}


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

    # Step 4: Manager Report PDF
    if logger:
        logger.info("Exporting Manager Report (PDF)…")
    pdf_bytes = export_sheet(creds, calc_ss_id, cfg.GID_MANAGER_PDF, "pdf", True)
    pdf_name = f"Manager_Report_{ts}_{location}.pdf"
    uploaded_pdf = upload_to_drive(drive_svc, pdf_bytes, pdf_name, "application/pdf", cfg.MANAGER_REPORT_FOLDER_ID, to_sheet=False)
    manager_link = uploaded_pdf.get("webViewLink")
    if logger:
        logger.info(f"Uploaded Manager PDF: {manager_link}")

    # Step 5: Master Order CSV
    if logger:
        logger.info("Exporting Master Order (CSV)…")
    master_csv_bytes = export_sheet(creds, calc_ss_id, cfg.GID_ORDER_CSV, "csv")
    master_csv_bytes = filter_master_csv_to_ran(master_csv_bytes, cfg)
    master_csv_bytes = sort_master_csv(master_csv_bytes)

    # Step 6: Error Report CSV, Upload, Export PDF
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
        
        autoresize_columns(sheets_svc, err_file_id, err_gid)

        err_pdf = export_sheet(creds, err_file_id, err_gid, "pdf", False)
        err_pdf_name = f"Error_Report_{ts}.pdf"

        if logger:
            logger.info(f"Uploaded filtered Error Sheet: {err_link}")

        # Step 6.1: Send Error Report if Needed

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




    # Step 7: Full order upload (CSV) and export (PDF)
    full_csv_name = f"Order_Report_FULL_{location}_{ts}.csv"
    full_created = upload_to_drive(drive_svc, master_csv_bytes, full_csv_name, CSV_MIME, cfg.ORDER_REPORT_FOLDER_ID, to_sheet=True)
    full_file_id = full_created["id"]
    full_link = full_created.get('webViewLink')
    full_gid = first_gid(sheets_svc, full_file_id)
    autoresize_columns(sheets_svc, full_file_id, full_gid)
    full_pdf = export_sheet(creds, full_file_id, full_gid, "pdf", False)
    full_pdf_name = f"Order_Report_FULL_{location}_{ts}.pdf"
    if logger:
        logger.info(f"Uploaded FULL sheet: {full_created.get('webViewLink')}")

    # Step 8: Create per-report-key outputs (CSV) and email

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
    

    bev_order_sent = False
    
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
        autoresize_columns(sheets_svc, file_id, gid)
        pdf = export_sheet(creds, file_id, gid, "pdf", False)
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

        
        if key.upper() == "BEV":
            bev_order_sent = True

    
        if logger:
            logger.info(f"Emailed {store} - {email_tag} to {recipients}")
    
    # Step 9: Unassigned Beverages Report (Soft Error)
    try:
        # Must have successfully sent a BEV order
        if not bev_order_sent:
            if logger:
                logger.info("Skipping Unassigned Beverages Report — no BEV order was sent.")
        else:
            if logger:
                logger.info("Exporting Unassigned Beverages Report (CSV)…")

            unassigned_csv_bytes = export_sheet(
                creds,
                calc_ss_id,
                cfg.GID_BEV_ERRORS,
                "csv"
            )

            # Inspect CSV to ensure it actually has data
            text = unassigned_csv_bytes.decode("utf-8-sig", errors="replace")
            reader = csv.reader(io.StringIO(text))
            rows = list(reader)

            if not rows or len(rows) <= 1:
                if logger:
                    logger.info("No unassigned beverages found — report will not be sent.")
            else:
                headers = [h.strip() for h in rows[0]]
                lower_idx = {h.lower(): i for i, h in enumerate(headers)}

                if "report_key" not in lower_idx:
                    raise RuntimeError("Unassigned BEV report missing Report_Key column")

                if "sub_report_key" not in lower_idx:
                    raise RuntimeError("Unassigned BEV report missing Sub_Report_Key column")

                report_idx = lower_idx["report_key"]
                sub_idx = lower_idx["sub_report_key"]

                # Filter to BEV + BEV_UNASSIGNED only
                unassigned_rows = [
                    r for r in rows[1:]
                    if r[report_idx].strip().upper() == "BEV"
                    and r[sub_idx].strip().upper() == "UNASSIGNED"
                ]

                if not unassigned_rows:
                    if logger:
                        logger.info("No BEV_UNASSIGNED rows found — skipping email.")
                else:
                    # Upload filtered sheet
                    output = io.StringIO()
                    writer = csv.writer(output)
                    writer.writerow(rows[0])
                    writer.writerows(unassigned_rows)

                    filtered_bytes = output.getvalue().encode("utf-8-sig")

                    csv_name = f"Unassigned_Beverages_Report_{ts}.csv"

                    created = upload_to_drive(
                        drive_svc,
                        filtered_bytes,
                        csv_name,
                        CSV_MIME,
                        cfg.ERROR_REPORT_FOLDER_ID,
                        to_sheet=True
                    )

                    sheet_id = created["id"]
                    sheet_link = created.get("webViewLink")
                    gid = first_gid(sheets_svc, sheet_id)

                    autoresize_columns(sheets_svc, sheet_id, gid)
                    pdf_bytes = export_sheet(creds, sheet_id, gid, "pdf", False)
                    pdf_name = f"Unassigned_Beverages_Report_{ts}.pdf"

                    # Resolve recipients
                    to_list = _fallback_recipients(
                        "UNASSIGNED BEVERAGES REPORT",
                        cfg.ERROR_RECIPIENTS,
                        cfg.TO_RECIPIENTS,
                        cfg.DEFAULT_ORDER_RECIPIENTS,
                    )

                    cc_list = list(dict.fromkeys(
                        set(_clean_emails(cfg.TO_RECIPIENTS))
                        | set(_clean_emails(cfg.CC_RECIPIENTS))
                        - set(to_list)
                    ))

                    email_bev_error_report(
                        gmail_svc=gmail_svc,
                        sender="me",
                        to_list=to_list,
                        cc_list=cc_list,
                        ts=ts,
                        pdf_name=pdf_name,
                        pdf_bytes=pdf_bytes,
                        sheet_link=sheet_link,
                        mapping_link=cfg.BEV_MAPPING_LINK,
                    )

                    if logger:
                        logger.info("Unassigned Beverages Report email sent")

    except Exception as e:
        # Soft error — log and continue
        if logger:
            logger.warn(f"Unassigned Beverages Report failed (soft): {e}")
        
    # Step 10: Send Manager Report (guarded by cfg.EMAIL_MANAGER_REPORT)
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

    

    # Step 11: Send Full Order if needed
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

    # Step 12: File Cleanup

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

```

---
### file: core_functional_modules/pipeline_bus.py

```python
# pipeline_bus.py
import queue

_PIPELINE_QUEUE = None

def get_pipeline_queue():
    global _PIPELINE_QUEUE
    if _PIPELINE_QUEUE is None:
        _PIPELINE_QUEUE = queue.Queue()
    return _PIPELINE_QUEUE
```

---
### file: core_functional_modules/sheets_utils.py

```python
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
import requests


def list_sheets(svc, spreadsheet_id: str) -> List[Dict[str, Any]]:
    return svc.spreadsheets().get(spreadsheetId=spreadsheet_id).execute().get("sheets", [])


def get_sheet(sheets, title: str):
    for s in sheets:
        if s["properties"]["title"] == title:
            return s["properties"]
    return None


def export_sheet(creds, spreadsheet_id: str, gid: str | int, fmt: str, portrait: bool = True,) -> bytes:
    params = {
        "format": fmt,
        "gid": gid,
    }

    # PDF-only layout options
    if fmt.lower() == "pdf":
        params.update({
            "portrait": "true" if portrait else "false",
            "fitw": "true",   # fit to width
            "scale": "4",     # normal scaling
        })

    # Build query string
    query = "&".join(f"{k}={v}" for k, v in params.items())

    url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export?{query}"

    headers = {"Authorization": f"Bearer {creds.token}"}
    r = requests.get(url, headers=headers)
    r.raise_for_status()
    return r.content



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

def autoresize_columns(sheets_svc, spreadsheet_id, sheet_id):
    sheets_svc.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={
            "requests": [{
                "autoResizeDimensions": {
                    "dimensions": {
                        "sheetId": sheet_id,
                        "dimension": "COLUMNS",
                        "startIndex": 0,
                        "endIndex": 26  # A–Z (adjust if needed)
                    }
                }
            }]
        }
    ).execute()

```

---
### file: documentation/README.md

```markdown
# FavTrip Reporting Pipeline — Codebase Overview

## Purpose & High-Level Architecture

The FavTrip Reporting Pipeline is a **Google Workspace–integrated reporting system** designed to automate weekly store reporting based on Modisoft sales data. It accepts a “Live Items Report” input file, validates and processes one or two weeks of sales data, updates calculation workbooks, produces PDFs and Sheets, and distributes reports via Gmail.

The codebase is structured into **three clear layers**:

1. **User Interface Layer** – Handles user interaction, authentication, configuration, and orchestration.
2. **Core Functional Modules** – Implements all business logic and Google API interactions.
3. **Supporting Assets & Documentation** – Reference materials, examples, and developer tooling.

This README intentionally focuses only on the **UI entrypoint** and the **core functional modules**.  
For detailed behavior, contracts, and edge cases, refer to the **module-level docstrings** in each file.

---

## User Interface Layer

### `_user_interface_.py`

This file implements the **primary Streamlit web application** and is the main operational entrypoint for most users.

At a high level, the UI is responsible for:

- Handling **Google OAuth authentication** (PKCE-based, stateless across redirects).
- Accepting Modisoft report uploads (CSV/XLSX) and storing them in Google Drive.
- Exposing **runtime configuration controls**:
  - Email recipients
  - Report keys
  - Feature toggles
  - Advanced IDs, GIDs, and date validation rules
- Validating inputs early to prevent unsafe or invalid pipeline runs.
- Executing the backend pipeline and streaming **live status updates and timing**.
- Rendering outputs (Drive links, timestamps) and handling failure recovery.

Key architectural characteristics:

- The UI **does not contain business logic**.
- All processing is delegated to `core_functional_modules.pipeline.run_pipeline`.
- Configuration is assembled via a shared `Config` object and passed downstream.
- The UI can evolve independently of the pipeline without risking logic drift.

For details about OAuth flow, state management, upload gating, and UI locking behavior, see the module docstring in this file.

---

## Core Functional Modules

All core logic lives under `core_functional_modules/`. These modules are designed to be:

- UI-agnostic
- Composable and reusable
- Safe to invoke from both the Streamlit UI and the CLI

Only responsibilities are summarized here; implementation details live in docstrings.

---

### `config.py` — Central Configuration Model

- Defines the canonical `Config` dataclass used everywhere in the system.
- Loads configuration via a **layered merge**:
  1. Streamlit secrets (typed, preferred in cloud)
  2. Environment variables / `.env`
  3. Optional Google Drive–hosted JSON overrides
- Normalizes types (booleans, lists, dicts) for consistent behavior.
- Acts as the single source of truth for IDs, flags, recipients, and cleanup policies.

This module underpins all other core components.

---

### `config_store.py` — Drive-Backed Config Persistence

- Reads and writes JSON configuration files stored in Google Drive.
- Enables the UI’s “Update defaults” feature.
- Uses resilient, fail-open behavior so missing or malformed configs never break execution.

---

### `google_client.py` — OAuth & Google Service Bootstrapping

- Manages Google OAuth sign-in and token lifecycle (`token.json`).
- Supports both browser-assisted and manual (CLI-style) flows.
- Produces authenticated service clients for:
  - Google Drive
  - Google Sheets
  - Gmail

All Google API usage throughout the system flows through this module.

---

### `drive_utils.py` — Google Drive Operations

- Uploads files (raw or converted to Google Sheets).
- Finds the latest files in folders.
- Copies, renames, and trashes Drive resources.
- Performs **time-based cleanup** of old files.
- Handles Drive query escaping and RFC‑3339 timestamps.

---

### `sheets_utils.py` — Google Sheets Manipulation

- Implements all Sheet-level operations:
  - Copying, deleting, and renaming sheets
  - Writing 2D values
  - Refreshing formulas
  - Forcing text coercion on specific columns
- Provides chunked refresh utilities for large sheets.

This module isolates low-level Sheets API complexity from the pipeline.

---

### `gmail_utils.py` — Email Composition & Sending

- Builds MIME-compliant emails with:
  - PDF attachments
  - Optional full-order attachments
  - Plain-text and HTML bodies
- Sends messages through the Gmail API as the authenticated user.
- Assumes recipient selection has already been resolved upstream.

---

### `logger.py` — Run-Time Status Logging

- Lightweight status logger with:
  - In-memory event tracking
  - Optional file-backed logging (`last_run.log`)
- Designed for real-time UI updates and post-run log downloads.
- Avoids heavy logging frameworks for predictability and portability.

---

### `pipeline.py` — Orchestrated Processing Engine

This is the **core execution engine** of the system.

At a high level, it:

1. Authenticates and initializes Google services.
2. Locates the latest incoming report.
3. Validates date coverage and week boundaries.
4. Prepares or rolls forward calculation workbooks.
5. Refreshes reference sheets.
6. Generates PDFs and Sheets for:
   - Manager reports
   - Full orders
   - Per-report-key outputs
7. Routes emails using a prioritized fallback model.
8. Performs Drive cleanup and retention enforcement.

The pipeline is intentionally **UI-agnostic** and is safe to invoke from:
- The Streamlit UI
- The CLI
- Future automation or scheduling contexts

---

## Supporting Folders

### `documentation/`

This directory contains **developer-facing documentation and tooling**, including:

- Architectural and usage documentation
- Setup, packaging, and deployment notes
- Git workflow guidance
- Dependency lists
- Tooling used to generate the project snapshot markdown bundle

Nothing in this folder is required for runtime execution. It exists to support onboarding, maintenance, and knowledge transfer.

---

### `__dev_input_sales_files/`

This directory contains **example Modisoft sales reports** used for testing and validation.

The files intentionally cover multiple scenarios, including:

- One-week vs. two-week uploads
- Multiple stores
- Invalid or misaligned week boundaries
- Bad start or end dates

These files are **not consumed automatically** by the application and exist purely for manual testing and development.

---

## Files and Folders Not Covered Here

Any files or folders **not explicitly mentioned in this README** (for example: logs, generated outputs, cached tokens, per-user folders, or packaging artifacts) are:

- **Auto-created**
- **Auto-maintained**
- **Managed by the application at runtime**

⚠️ **Do not edit, delete, or manually modify these files unless you fully understand their lifecycle and side effects.**  
Incorrect changes may break authentication, corrupt Drive state, or cause data loss.

---

## How to Navigate the Codebase

- Start with `_user_interface_.py` to understand user flow and orchestration.
- Follow execution into `pipeline.run_pipeline`.
- Use module-level docstrings in `core_functional_modules/` for precise logic and guarantees.
- Treat `Config` as the shared contract that ties the system together.

```

---
### file: documentation/generate_code_bundle.py

```python
#!/usr/bin/env python3


from __future__ import annotations
import argparse
import fnmatch
import os
from pathlib import Path
from typing import Iterable, List, Set, Tuple

# ----------------------------
# Defaults (sane and safe)
# ----------------------------

DEFAULT_IGNORE_DIRS: Set[str] = {
    ".git", ".hg", ".svn",
    ".venv", "venv", "env",
    "build", "dist", "__pycache__",
    ".pytest_cache", ".mypy_cache", ".idea", ".vscode",
}


DEFAULT_IGNORE_NAMES: Set[str] = {
    # OAuth / secrets
    "token.json",
    "credentials.json",
    "web_url_credentials.json",
    ".env",

    # Common secret patterns
    "client_secrets.json",
}


DEFAULT_IGNORE_EXTS: Set[str] = {
    ".pyc", ".pyo", ".pyd",
    ".so", ".dll", ".dylib",
    ".zip", ".tar", ".gz", ".7z",
    ".exe", ".bin",
}


DEFAULT_IGNORE_GLOBS = [
    "**/*credentials*.json",
    "**/*.env",
    "**/*.env.*",
    "**/*secret*.json",
]


# Basic language hints for code fences
LANG_BY_EXT = {
    ".py": "python",
    ".md": "markdown",
    ".txt": "text",
    ".bat": "bat",
    ".cmd": "bat",
    ".sh": "bash",
    ".ps1": "powershell",
    ".json": "json",
    ".ini": "ini",
    ".env": "ini",
    ".yml": "yaml",
    ".yaml": "yaml",
    ".csv": "csv",
    ".ts": "ts",
    ".tsx": "tsx",
    ".js": "javascript",
    ".jsx": "jsx",
    ".html": "html",
    ".css": "css",
    ".toml": "toml",
}


def match_any(path: Path, patterns: Iterable[str]) -> bool:
    """Return True if path matches ANY of the glob patterns (POSIX-style)."""
    s = path.as_posix()
    for pat in patterns:
        if fnmatch.fnmatch(s, pat):
            return True
    return False


def build_tree(root: Path,
               ignore_dirs: Set[str],
               ignore_exts: Set[str],
               ignore_names: Set[str],
               include_globs: List[str] | None,
               exclude_globs: List[str] | None,
               max_bytes: int) -> Tuple[str, List[Path]]:
    """
    Return a (tree_text, files_list) tuple.
    files_list contains all files to embed in the bundle.
    """
    lines_tree: List[str] = []
    files_out: List[Path] = []

    for dirpath, dirnames, filenames in os.walk(root):
        # prune ignored directories
        dirnames[:] = [
            d for d in dirnames
            if d not in ignore_dirs and not d.startswith(".")
        ]
        cur = Path(dirpath)
        rel_dir = cur.relative_to(root)
        depth = 0 if rel_dir == Path(".") else len(rel_dir.parts)
        indent = "  " * depth
        if rel_dir != Path("."):
            lines_tree.append(f"{indent}{rel_dir.as_posix()}/")

        for fname in sorted(filenames):
            p = cur / fname

            # Ignore by name/extension
            if fname in ignore_names:
                entry = fname if rel_dir == Path(".") else f"{indent} {fname}"
                lines_tree.append(entry + " [skipped: secret]")
                continue
            if p.suffix.lower() in ignore_exts:
                continue

            
            # Ignore by glob (secrets)
            if match_any(p.relative_to(root), DEFAULT_IGNORE_GLOBS):
                entry = fname if rel_dir == Path(".") else f"{indent} {fname}"
                lines_tree.append(entry + " [skipped: secret]")
                continue


            # Exclude hidden top-level noise by pattern
            if exclude_globs and match_any(p.relative_to(root), exclude_globs):
                continue
            # If include globs were given, only take matches
            if include_globs and not match_any(p.relative_to(root), include_globs):
                continue

            # Size check
            try:
                if p.stat().st_size > max_bytes:
                    # Do not list it in files_out, but show in tree (optional)
                    entry = fname if rel_dir == Path(".") else f"{indent}  {fname}"
                    lines_tree.append(entry + "  [skipped: too large]")
                    continue
            except Exception:
                # If can't stat, skip silently
                continue

            # Accept file
            entry = fname if rel_dir == Path(".") else f"{indent}  {fname}"
            lines_tree.append(entry)
            files_out.append(p)

    return "\n".join(lines_tree), files_out


def read_text_safe(p: Path) -> str | None:
    try:
        return p.read_text(encoding="utf-8")
    except UnicodeDecodeError:
        # Binary or non-UTF8; skip
        return None
    except Exception:
        return None


def make_bundle_markdown(root: Path,
                         out: Path,
                         include_globs: List[str] | None,
                         exclude_globs: List[str] | None,
                         max_bytes: int,
                         ignore_dirs: Set[str],
                         ignore_exts: Set[str],
                         ignore_names: Set[str]) -> Path:
    tree_text, files_to_embed = build_tree(
        root, ignore_dirs, ignore_exts, ignore_names,
        include_globs, exclude_globs, max_bytes
    )

    header = f"""# Project Snapshot (CodeBundle)

                This single Markdown file contains a **self-contained snapshot** of your project so another AI/engineer can review or modify it without needing the original folder.

                **How to use this file with an AI**
                1. Upload or paste this file as a single attachment.
                2. Ask for changes; the AI can reference specific `file:` sections below.
                3. Copy updated blocks back into the corresponding files in your project.

                > Notes: secrets like `token.json` are intentionally excluded. Virtual envs and build artifacts are omitted to keep this readable.

                ## Directory tree (filtered)
                {tree_text}"""

    sections: List[str] = [header]

    for p in sorted(files_to_embed, key=lambda x: x.relative_to(root).as_posix()):
        rel = p.relative_to(root)
        lang = LANG_BY_EXT.get(p.suffix.lower(), "")
        content = read_text_safe(p)
        if content is None:
            # Skip non-text files silently
            continue
        fence = "```"
        sections.append(
            f"\n---\n### file: {rel.as_posix()}\n\n{fence}{lang}\n{content}\n{fence}\n"
        )

    out.write_text("".join(sections), encoding="utf-8")
    return out


def parse_args() -> argparse.Namespace:
    ap = argparse.ArgumentParser(
        description="Generate a single-file Markdown 'CodeBundle' of a project."
    )
    # Defaults are computed in main()
    ap.add_argument(
        "--root",
        default=None,
        help="Project root folder (default: parent of this script)",
    )
    ap.add_argument(
        "--out",
        default=None,
        help="Output Markdown file path (default: PROJECT_SNAPSHOT_CODEBUNDLE.md in the script's directory)",
    )
    ap.add_argument("--include", nargs="*", default=None,
                    help='Optional list of glob patterns to include (e.g. "**/*.py" "**/*.md").')
    ap.add_argument("--exclude", nargs="*", default=None,
                    help='Optional list of glob patterns to exclude (e.g. ".venv/**" "dist/**").')
    ap.add_argument("--max-bytes", type=int, default=1_000_000,
                    help="Max file size to embed (bytes). Oversized files are listed but skipped. Default: 1,000,000.")
    ap.add_argument("--no-default-ignores", action="store_true",
                    help="Disable default ignore sets for dirs/exts/names.")
    return ap.parse_args()


def main():
    print("CodeBundle writing starting...")
    args = parse_args()

    # Locate the script and its parent (project root)
    script_dir = Path(__file__).resolve().parent          # .../favtrip_reporting/documentation
    project_root = script_dir.parent                      # .../favtrip_reporting

    # ROOT: default to the parent of this script unless overridden
    if args.root is None:
        root = project_root
    else:
        root_arg = Path(args.root)
        root = root_arg.resolve() if root_arg.is_absolute() else (Path.cwd() / root_arg).resolve()

    if not root.exists() or not root.is_dir():
        raise SystemExit(f"Root does not exist or is not a directory: {root}")

    # OUT: default to PROJECT_SNAPSHOT_CODEBUNDLE.md in the script's directory (the child)
    if args.out is None:
        out = (script_dir / "PROJECT_SNAPSHOT_CODEBUNDLE.md").resolve()
    else:
        out_arg = Path(args.out)
        # If user gave a relative path, resolve it under the chosen root; absolute paths are respected
        out = out_arg.resolve() if out_arg.is_absolute() else (root / out_arg).resolve()

    # ... keep the rest of your function unchanged ...
    if args.no_default_ignores:
        ignore_dirs = set()
        ignore_exts = set()
        ignore_names = set()
    else:
        ignore_dirs = set(DEFAULT_IGNORE_DIRS)
        ignore_exts = set(DEFAULT_IGNORE_EXTS)
        ignore_names = set(DEFAULT_IGNORE_NAMES)

    out.parent.mkdir(parents=True, exist_ok=True)
    result_path = make_bundle_markdown(
        root=root,
        out=out,
        include_globs=args.include,
        exclude_globs=args.exclude,
        max_bytes=args.max_bytes,
        ignore_dirs=ignore_dirs,
        ignore_exts=ignore_exts,
        ignore_names=ignore_names,
    )
    print(f"CodeBundle written to: {result_path}")


if __name__ == "__main__":
    main()

#cd "C:\Users\rjrul\OneDrive - University of Iowa\000 Current Semester\004 BAIS 4150 BAIS Capstone\favtrip_reporting"
#python documentation/generate_code_bundle.py
```

---
### file: documentation/git_workflow.txt

```text

---
1. Create a dev branch

# Switch to dev branch
    git checkout dev

# Push dev branch to GitHub (the -u flag sets upstream tracking)
    git push -u origin dev

---
2. Do all development on dev
Make changes normally, then stage, commit, and push:

    git add .
    git commit -m "your message here"
    git push

All pushes go to dev (not main).

---
3. Push dev → main when ready (production release)
Use a Pull Request on GitHub:
1. Go to the repo on GitHub.
2. You will see a banner offering “Compare & Pull Request” (dev → main).
3. Open the Pull Request.
4. Review and click “Merge Pull Request”.

After merging, update your local main branch:

    git checkout main
    git pull

---
4. Keep dev updated with the latest main
After merging into main, update dev:

    git checkout dev
    git pull origin main

Or:

    git merge main

This prevents dev from drifting behind main.

---
5. Summary
- main = production, always stable
- dev = development, experimental work
- Merge dev → main only after testing

```

---
### file: documentation/requirements.txt

```text
google-api-python-client
google-auth
google-auth-oauthlib
google-auth-httplib2
httplib2
requests
python-dotenv
streamlit
streamlit-autorefresh
openpyxl
pandas
uuid
```

---
### file: last_run.log

```
[2026-03-03 23:16:33] INFO: Authorizing with Google APIs…
[2026-03-03 23:16:33] INFO: Google services ready
[2026-03-03 23:16:33] INFO: Finding latest incoming spreadsheet…
[2026-03-03 23:16:34] INFO: Latest incoming: Testing Sales Report - Week 5 (1kNjmEbljdUIqJUjwh8e2-plfLWRZSGHxMHxL50-2ce4)
[2026-03-03 23:16:34] INFO: Preparing calculations workbook…
[2026-03-03 23:16:42] INFO: Copied old 'Current Week' to 'Last Week'
[2026-03-03 23:16:50] INFO: Inserted new 'Current Week' from latest incoming report
[2026-03-03 23:16:50] INFO: Refreshing reference sheets (prefix 'REFR: ')…
[2026-03-03 23:17:17] INFO: [1/3] Recalc OK: REFR: Charts
[2026-03-03 23:17:22] INFO: [2/3] Recalc OK: REFR: Values
[2026-03-03 23:18:15] INFO: [3/3] Recalc OK: REFR: Order Calcs
[2026-03-03 23:18:17] INFO: Location: Favtrip_Independence; Timestamp: 2026-03-03-11-18-PM
[2026-03-03 23:18:17] INFO: Exporting Manager Report (PDF)…
[2026-03-03 23:18:21] INFO: Uploaded Manager PDF: https://drive.google.com/file/d/1UILYk_poamIwe6JrFQUGZIz9az9iBsdU/view?usp=drivesdk
[2026-03-03 23:18:21] INFO: Exporting Master Order (CSV)…
[2026-03-03 23:18:34] INFO: Uploaded FULL sheet: https://docs.google.com/spreadsheets/d/1bOcle5eMseFc43TZMnXLqgHR7_DcPs13ntb6PSBians/edit?usp=drivesdk
[2026-03-03 23:18:42] INFO: Emailed COFFEE
[2026-03-03 23:18:43] INFO: Manager email sent
[2026-03-03 23:18:43] INFO: Separate full order email disabled
[2026-03-03 23:18:43] INFO: Run completed in 00:02:09

```

---
### file: launcher_streamlit.py

```python
import os, sys, subprocess

def main():
    # Ensure current working directory is bundle dir
    base = getattr(sys, '_MEIPASS', os.path.abspath(os.path.dirname(__file__)))
    os.chdir(base)
    subprocess.call([sys.executable, "-m", "streamlit", "run", "ui_streamlit.py"])

if __name__ == "__main__":
    main()

```

---
### file: requirements.txt

```text
-r documentation/requirements.txt
```

---
### file: setup_py2app.py

```python
from setuptools import setup

APP = ['launcher_streamlit.py']
OPTIONS = {
    'argv_emulation': True,
    'plist': {
        'CFBundleName': 'FavTripPipelineUI',
    },
    'packages': ['googleapiclient', 'google', 'httplib2', 'google_auth_oauthlib', 'google_auth_httplib2', 'dotenv', 'requests', 'streamlit'],
}

data_files = ['ui_streamlit.py', 'cli.py', 'requirements.txt', '.env', 'credentials.json']

setup(
    app=APP,
    options={'py2app': OPTIONS},
    data_files=data_files,
)

```

```

---
### file: documentation/README.md

```markdown
# FavTrip Reporting Pipeline — Codebase Overview

## Purpose & High-Level Architecture

The FavTrip Reporting Pipeline is a **Google Workspace–integrated reporting system** designed to automate weekly store reporting based on Modisoft sales data. It accepts a “Live Items Report” input file, validates and processes one or two weeks of sales data, updates calculation workbooks, produces PDFs and Sheets, and distributes reports via Gmail.

The codebase is structured into **three clear layers**:

1. **User Interface Layer** – Handles user interaction, authentication, configuration, and orchestration.
2. **Core Functional Modules** – Implements all business logic and Google API interactions.
3. **Supporting Assets & Documentation** – Reference materials, examples, and developer tooling.

This README intentionally focuses only on the **UI entrypoint** and the **core functional modules**.  
For detailed behavior, contracts, and edge cases, refer to the **module-level docstrings** in each file.

---

## User Interface Layer

### `_user_interface_.py`

This file implements the **primary Streamlit web application** and is the main operational entrypoint for most users.

At a high level, the UI is responsible for:

- Handling **Google OAuth authentication** (PKCE-based, stateless across redirects).
- Accepting Modisoft report uploads (CSV/XLSX) and storing them in Google Drive.
- Exposing **runtime configuration controls**:
  - Email recipients
  - Report keys
  - Feature toggles
  - Advanced IDs, GIDs, and date validation rules
- Validating inputs early to prevent unsafe or invalid pipeline runs.
- Executing the backend pipeline and streaming **live status updates and timing**.
- Rendering outputs (Drive links, timestamps) and handling failure recovery.

Key architectural characteristics:

- The UI **does not contain business logic**.
- All processing is delegated to `core_functional_modules.pipeline.run_pipeline`.
- Configuration is assembled via a shared `Config` object and passed downstream.
- The UI can evolve independently of the pipeline without risking logic drift.

For details about OAuth flow, state management, upload gating, and UI locking behavior, see the module docstring in this file.

---

## Core Functional Modules

All core logic lives under `core_functional_modules/`. These modules are designed to be:

- UI-agnostic
- Composable and reusable
- Safe to invoke from both the Streamlit UI and the CLI

Only responsibilities are summarized here; implementation details live in docstrings.

---

### `config.py` — Central Configuration Model

- Defines the canonical `Config` dataclass used everywhere in the system.
- Loads configuration via a **layered merge**:
  1. Streamlit secrets (typed, preferred in cloud)
  2. Environment variables / `.env`
  3. Optional Google Drive–hosted JSON overrides
- Normalizes types (booleans, lists, dicts) for consistent behavior.
- Acts as the single source of truth for IDs, flags, recipients, and cleanup policies.

This module underpins all other core components.

---

### `config_store.py` — Drive-Backed Config Persistence

- Reads and writes JSON configuration files stored in Google Drive.
- Enables the UI’s “Update defaults” feature.
- Uses resilient, fail-open behavior so missing or malformed configs never break execution.

---

### `google_client.py` — OAuth & Google Service Bootstrapping

- Manages Google OAuth sign-in and token lifecycle (`token.json`).
- Supports both browser-assisted and manual (CLI-style) flows.
- Produces authenticated service clients for:
  - Google Drive
  - Google Sheets
  - Gmail

All Google API usage throughout the system flows through this module.

---

### `drive_utils.py` — Google Drive Operations

- Uploads files (raw or converted to Google Sheets).
- Finds the latest files in folders.
- Copies, renames, and trashes Drive resources.
- Performs **time-based cleanup** of old files.
- Handles Drive query escaping and RFC‑3339 timestamps.

---

### `sheets_utils.py` — Google Sheets Manipulation

- Implements all Sheet-level operations:
  - Copying, deleting, and renaming sheets
  - Writing 2D values
  - Refreshing formulas
  - Forcing text coercion on specific columns
- Provides chunked refresh utilities for large sheets.

This module isolates low-level Sheets API complexity from the pipeline.

---

### `gmail_utils.py` — Email Composition & Sending

- Builds MIME-compliant emails with:
  - PDF attachments
  - Optional full-order attachments
  - Plain-text and HTML bodies
- Sends messages through the Gmail API as the authenticated user.
- Assumes recipient selection has already been resolved upstream.

---

### `logger.py` — Run-Time Status Logging

- Lightweight status logger with:
  - In-memory event tracking
  - Optional file-backed logging (`last_run.log`)
- Designed for real-time UI updates and post-run log downloads.
- Avoids heavy logging frameworks for predictability and portability.

---

### `pipeline.py` — Orchestrated Processing Engine

This is the **core execution engine** of the system.

At a high level, it:

1. Authenticates and initializes Google services.
2. Locates the latest incoming report.
3. Validates date coverage and week boundaries.
4. Prepares or rolls forward calculation workbooks.
5. Refreshes reference sheets.
6. Generates PDFs and Sheets for:
   - Manager reports
   - Full orders
   - Per-report-key outputs
7. Routes emails using a prioritized fallback model.
8. Performs Drive cleanup and retention enforcement.

The pipeline is intentionally **UI-agnostic** and is safe to invoke from:
- The Streamlit UI
- The CLI
- Future automation or scheduling contexts

---

## Supporting Folders

### `documentation/`

This directory contains **developer-facing documentation and tooling**, including:

- Architectural and usage documentation
- Setup, packaging, and deployment notes
- Git workflow guidance
- Dependency lists
- Tooling used to generate the project snapshot markdown bundle

Nothing in this folder is required for runtime execution. It exists to support onboarding, maintenance, and knowledge transfer.

---

### `__dev_input_sales_files/`

This directory contains **example Modisoft sales reports** used for testing and validation.

The files intentionally cover multiple scenarios, including:

- One-week vs. two-week uploads
- Multiple stores
- Invalid or misaligned week boundaries
- Bad start or end dates

These files are **not consumed automatically** by the application and exist purely for manual testing and development.

---

## Files and Folders Not Covered Here

Any files or folders **not explicitly mentioned in this README** (for example: logs, generated outputs, cached tokens, per-user folders, or packaging artifacts) are:

- **Auto-created**
- **Auto-maintained**
- **Managed by the application at runtime**

⚠️ **Do not edit, delete, or manually modify these files unless you fully understand their lifecycle and side effects.**  
Incorrect changes may break authentication, corrupt Drive state, or cause data loss.

---

## How to Navigate the Codebase

- Start with `_user_interface_.py` to understand user flow and orchestration.
- Follow execution into `pipeline.run_pipeline`.
- Use module-level docstrings in `core_functional_modules/` for precise logic and guarantees.
- Treat `Config` as the shared contract that ties the system together.

```

---
### file: documentation/generate_code_bundle.py

```python
#!/usr/bin/env python3


from __future__ import annotations
import argparse
import fnmatch
import os
from pathlib import Path
from typing import Iterable, List, Set, Tuple

# ----------------------------
# Defaults (sane and safe)
# ----------------------------

DEFAULT_IGNORE_DIRS: Set[str] = {
    ".git", ".hg", ".svn",
    ".venv", "venv", "env",
    "build", "dist", "__pycache__",
    ".pytest_cache", ".mypy_cache", ".idea", ".vscode",
}


DEFAULT_IGNORE_NAMES: Set[str] = {
    # OAuth / secrets
    "token.json",
    "credentials.json",
    "web_url_credentials.json",
    ".env",

    # Common secret patterns
    "client_secrets.json",
}


DEFAULT_IGNORE_EXTS: Set[str] = {
    ".pyc", ".pyo", ".pyd",
    ".so", ".dll", ".dylib",
    ".zip", ".tar", ".gz", ".7z",
    ".exe", ".bin",
}


DEFAULT_IGNORE_GLOBS = [
    "**/*credentials*.json",
    "**/*.env",
    "**/*.env.*",
    "**/*secret*.json",
]


# Basic language hints for code fences
LANG_BY_EXT = {
    ".py": "python",
    ".md": "markdown",
    ".txt": "text",
    ".bat": "bat",
    ".cmd": "bat",
    ".sh": "bash",
    ".ps1": "powershell",
    ".json": "json",
    ".ini": "ini",
    ".env": "ini",
    ".yml": "yaml",
    ".yaml": "yaml",
    ".csv": "csv",
    ".ts": "ts",
    ".tsx": "tsx",
    ".js": "javascript",
    ".jsx": "jsx",
    ".html": "html",
    ".css": "css",
    ".toml": "toml",
}


def match_any(path: Path, patterns: Iterable[str]) -> bool:
    """Return True if path matches ANY of the glob patterns (POSIX-style)."""
    s = path.as_posix()
    for pat in patterns:
        if fnmatch.fnmatch(s, pat):
            return True
    return False


def build_tree(root: Path,
               ignore_dirs: Set[str],
               ignore_exts: Set[str],
               ignore_names: Set[str],
               include_globs: List[str] | None,
               exclude_globs: List[str] | None,
               max_bytes: int) -> Tuple[str, List[Path]]:
    """
    Return a (tree_text, files_list) tuple.
    files_list contains all files to embed in the bundle.
    """
    lines_tree: List[str] = []
    files_out: List[Path] = []

    for dirpath, dirnames, filenames in os.walk(root):
        # prune ignored directories
        dirnames[:] = [
            d for d in dirnames
            if d not in ignore_dirs and not d.startswith(".")
        ]
        cur = Path(dirpath)
        rel_dir = cur.relative_to(root)
        depth = 0 if rel_dir == Path(".") else len(rel_dir.parts)
        indent = "  " * depth
        if rel_dir != Path("."):
            lines_tree.append(f"{indent}{rel_dir.as_posix()}/")

        for fname in sorted(filenames):
            p = cur / fname

            # Ignore by name/extension
            if fname in ignore_names:
                entry = fname if rel_dir == Path(".") else f"{indent} {fname}"
                lines_tree.append(entry + " [skipped: secret]")
                continue
            if p.suffix.lower() in ignore_exts:
                continue

            
            # Ignore by glob (secrets)
            if match_any(p.relative_to(root), DEFAULT_IGNORE_GLOBS):
                entry = fname if rel_dir == Path(".") else f"{indent} {fname}"
                lines_tree.append(entry + " [skipped: secret]")
                continue


            # Exclude hidden top-level noise by pattern
            if exclude_globs and match_any(p.relative_to(root), exclude_globs):
                continue
            # If include globs were given, only take matches
            if include_globs and not match_any(p.relative_to(root), include_globs):
                continue

            # Size check
            try:
                if p.stat().st_size > max_bytes:
                    # Do not list it in files_out, but show in tree (optional)
                    entry = fname if rel_dir == Path(".") else f"{indent}  {fname}"
                    lines_tree.append(entry + "  [skipped: too large]")
                    continue
            except Exception:
                # If can't stat, skip silently
                continue

            # Accept file
            entry = fname if rel_dir == Path(".") else f"{indent}  {fname}"
            lines_tree.append(entry)
            files_out.append(p)

    return "\n".join(lines_tree), files_out


def read_text_safe(p: Path) -> str | None:
    try:
        return p.read_text(encoding="utf-8")
    except UnicodeDecodeError:
        # Binary or non-UTF8; skip
        return None
    except Exception:
        return None


def make_bundle_markdown(root: Path,
                         out: Path,
                         include_globs: List[str] | None,
                         exclude_globs: List[str] | None,
                         max_bytes: int,
                         ignore_dirs: Set[str],
                         ignore_exts: Set[str],
                         ignore_names: Set[str]) -> Path:
    tree_text, files_to_embed = build_tree(
        root, ignore_dirs, ignore_exts, ignore_names,
        include_globs, exclude_globs, max_bytes
    )

    header = f"""# Project Snapshot (CodeBundle)

                This single Markdown file contains a **self-contained snapshot** of your project so another AI/engineer can review or modify it without needing the original folder.

                **How to use this file with an AI**
                1. Upload or paste this file as a single attachment.
                2. Ask for changes; the AI can reference specific `file:` sections below.
                3. Copy updated blocks back into the corresponding files in your project.

                > Notes: secrets like `token.json` are intentionally excluded. Virtual envs and build artifacts are omitted to keep this readable.

                ## Directory tree (filtered)
                {tree_text}"""

    sections: List[str] = [header]

    for p in sorted(files_to_embed, key=lambda x: x.relative_to(root).as_posix()):
        rel = p.relative_to(root)
        lang = LANG_BY_EXT.get(p.suffix.lower(), "")
        content = read_text_safe(p)
        if content is None:
            # Skip non-text files silently
            continue
        fence = "```"
        sections.append(
            f"\n---\n### file: {rel.as_posix()}\n\n{fence}{lang}\n{content}\n{fence}\n"
        )

    out.write_text("".join(sections), encoding="utf-8")
    return out


def parse_args() -> argparse.Namespace:
    ap = argparse.ArgumentParser(
        description="Generate a single-file Markdown 'CodeBundle' of a project."
    )
    # Defaults are computed in main()
    ap.add_argument(
        "--root",
        default=None,
        help="Project root folder (default: parent of this script)",
    )
    ap.add_argument(
        "--out",
        default=None,
        help="Output Markdown file path (default: PROJECT_SNAPSHOT_CODEBUNDLE.md in the script's directory)",
    )
    ap.add_argument("--include", nargs="*", default=None,
                    help='Optional list of glob patterns to include (e.g. "**/*.py" "**/*.md").')
    ap.add_argument("--exclude", nargs="*", default=None,
                    help='Optional list of glob patterns to exclude (e.g. ".venv/**" "dist/**").')
    ap.add_argument("--max-bytes", type=int, default=1_000_000,
                    help="Max file size to embed (bytes). Oversized files are listed but skipped. Default: 1,000,000.")
    ap.add_argument("--no-default-ignores", action="store_true",
                    help="Disable default ignore sets for dirs/exts/names.")
    return ap.parse_args()


def main():
    print("CodeBundle writing starting...")
    args = parse_args()

    # Locate the script and its parent (project root)
    script_dir = Path(__file__).resolve().parent          # .../favtrip_reporting/documentation
    project_root = script_dir.parent                      # .../favtrip_reporting

    # ROOT: default to the parent of this script unless overridden
    if args.root is None:
        root = project_root
    else:
        root_arg = Path(args.root)
        root = root_arg.resolve() if root_arg.is_absolute() else (Path.cwd() / root_arg).resolve()

    if not root.exists() or not root.is_dir():
        raise SystemExit(f"Root does not exist or is not a directory: {root}")

    # OUT: default to PROJECT_SNAPSHOT_CODEBUNDLE.md in the script's directory (the child)
    if args.out is None:
        out = (script_dir / "PROJECT_SNAPSHOT_CODEBUNDLE.md").resolve()
    else:
        out_arg = Path(args.out)
        # If user gave a relative path, resolve it under the chosen root; absolute paths are respected
        out = out_arg.resolve() if out_arg.is_absolute() else (root / out_arg).resolve()

    # ... keep the rest of your function unchanged ...
    if args.no_default_ignores:
        ignore_dirs = set()
        ignore_exts = set()
        ignore_names = set()
    else:
        ignore_dirs = set(DEFAULT_IGNORE_DIRS)
        ignore_exts = set(DEFAULT_IGNORE_EXTS)
        ignore_names = set(DEFAULT_IGNORE_NAMES)

    out.parent.mkdir(parents=True, exist_ok=True)
    result_path = make_bundle_markdown(
        root=root,
        out=out,
        include_globs=args.include,
        exclude_globs=args.exclude,
        max_bytes=args.max_bytes,
        ignore_dirs=ignore_dirs,
        ignore_exts=ignore_exts,
        ignore_names=ignore_names,
    )
    print(f"CodeBundle written to: {result_path}")


if __name__ == "__main__":
    main()

#cd "C:\Users\rjrul\OneDrive - University of Iowa\000 Current Semester\004 BAIS 4150 BAIS Capstone\favtrip_reporting"
#python documentation/generate_code_bundle.py
```

---
### file: documentation/git_workflow.txt

```text

---
1. Create a dev branch

# Switch to dev branch
    git checkout dev

# Push dev branch to GitHub (the -u flag sets upstream tracking)
    git push -u origin dev

---
2. Do all development on dev
Make changes normally, then stage, commit, and push:

    git add .
    git commit -m "your message here"
    git push

All pushes go to dev (not main).

---
3. Push dev → main when ready (production release)
Use a Pull Request on GitHub:
1. Go to the repo on GitHub.
2. You will see a banner offering “Compare & Pull Request” (dev → main).
3. Open the Pull Request.
4. Review and click “Merge Pull Request”.

After merging, update your local main branch:

    git checkout main
    git pull

---
4. Keep dev updated with the latest main
After merging into main, update dev:

    git checkout dev
    git pull origin main

Or:

    git merge main

This prevents dev from drifting behind main.

---
5. Summary
- main = production, always stable
- dev = development, experimental work
- Merge dev → main only after testing

```

---
### file: documentation/requirements.txt

```text
google-api-python-client
google-auth
google-auth-oauthlib
google-auth-httplib2
httplib2
requests
python-dotenv
streamlit
streamlit-autorefresh
openpyxl
pandas
uuid
```

---
### file: last_run.log

```
[2026-03-03 23:16:33] INFO: Authorizing with Google APIs…
[2026-03-03 23:16:33] INFO: Google services ready
[2026-03-03 23:16:33] INFO: Finding latest incoming spreadsheet…
[2026-03-03 23:16:34] INFO: Latest incoming: Testing Sales Report - Week 5 (1kNjmEbljdUIqJUjwh8e2-plfLWRZSGHxMHxL50-2ce4)
[2026-03-03 23:16:34] INFO: Preparing calculations workbook…
[2026-03-03 23:16:42] INFO: Copied old 'Current Week' to 'Last Week'
[2026-03-03 23:16:50] INFO: Inserted new 'Current Week' from latest incoming report
[2026-03-03 23:16:50] INFO: Refreshing reference sheets (prefix 'REFR: ')…
[2026-03-03 23:17:17] INFO: [1/3] Recalc OK: REFR: Charts
[2026-03-03 23:17:22] INFO: [2/3] Recalc OK: REFR: Values
[2026-03-03 23:18:15] INFO: [3/3] Recalc OK: REFR: Order Calcs
[2026-03-03 23:18:17] INFO: Location: Favtrip_Independence; Timestamp: 2026-03-03-11-18-PM
[2026-03-03 23:18:17] INFO: Exporting Manager Report (PDF)…
[2026-03-03 23:18:21] INFO: Uploaded Manager PDF: https://drive.google.com/file/d/1UILYk_poamIwe6JrFQUGZIz9az9iBsdU/view?usp=drivesdk
[2026-03-03 23:18:21] INFO: Exporting Master Order (CSV)…
[2026-03-03 23:18:34] INFO: Uploaded FULL sheet: https://docs.google.com/spreadsheets/d/1bOcle5eMseFc43TZMnXLqgHR7_DcPs13ntb6PSBians/edit?usp=drivesdk
[2026-03-03 23:18:42] INFO: Emailed COFFEE
[2026-03-03 23:18:43] INFO: Manager email sent
[2026-03-03 23:18:43] INFO: Separate full order email disabled
[2026-03-03 23:18:43] INFO: Run completed in 00:02:09

```

---
### file: launcher_streamlit.py

```python
import os, sys, subprocess

def main():
    # Ensure current working directory is bundle dir
    base = getattr(sys, '_MEIPASS', os.path.abspath(os.path.dirname(__file__)))
    os.chdir(base)
    subprocess.call([sys.executable, "-m", "streamlit", "run", "ui_streamlit.py"])

if __name__ == "__main__":
    main()

```

---
### file: requirements.txt

```text
-r documentation/requirements.txt
```

---
### file: setup_py2app.py

```python
from setuptools import setup

APP = ['launcher_streamlit.py']
OPTIONS = {
    'argv_emulation': True,
    'plist': {
        'CFBundleName': 'FavTripPipelineUI',
    },
    'packages': ['googleapiclient', 'google', 'httplib2', 'google_auth_oauthlib', 'google_auth_httplib2', 'dotenv', 'requests', 'streamlit'],
}

data_files = ['ui_streamlit.py', 'cli.py', 'requirements.txt', '.env', 'credentials.json']

setup(
    app=APP,
    options={'py2app': OPTIONS},
    data_files=data_files,
)

```

```

---
### file: documentation/README.md

```markdown
# FavTrip Reporting Pipeline — Codebase Overview

## Purpose & High-Level Architecture

The FavTrip Reporting Pipeline is a **Google Workspace–integrated reporting system** designed to automate weekly store reporting based on Modisoft sales data. It accepts a “Live Items Report” input file, validates and processes one or two weeks of sales data, updates calculation workbooks, produces PDFs and Sheets, and distributes reports via Gmail.

The codebase is structured into **three clear layers**:

1. **User Interface Layer** – Handles user interaction, authentication, configuration, and orchestration.
2. **Core Functional Modules** – Implements all business logic and Google API interactions.
3. **Supporting Assets & Documentation** – Reference materials, examples, and developer tooling.

This README intentionally focuses only on the **UI entrypoint** and the **core functional modules**.  
For detailed behavior, contracts, and edge cases, refer to the **module-level docstrings** in each file.

---

## User Interface Layer

### `_user_interface_.py`

This file implements the **primary Streamlit web application** and is the main operational entrypoint for most users.

At a high level, the UI is responsible for:

- Handling **Google OAuth authentication** (PKCE-based, stateless across redirects).
- Accepting Modisoft report uploads (CSV/XLSX) and storing them in Google Drive.
- Exposing **runtime configuration controls**:
  - Email recipients
  - Report keys
  - Feature toggles
  - Advanced IDs, GIDs, and date validation rules
- Validating inputs early to prevent unsafe or invalid pipeline runs.
- Executing the backend pipeline and streaming **live status updates and timing**.
- Rendering outputs (Drive links, timestamps) and handling failure recovery.

Key architectural characteristics:

- The UI **does not contain business logic**.
- All processing is delegated to `core_functional_modules.pipeline.run_pipeline`.
- Configuration is assembled via a shared `Config` object and passed downstream.
- The UI can evolve independently of the pipeline without risking logic drift.

For details about OAuth flow, state management, upload gating, and UI locking behavior, see the module docstring in this file.

---

## Core Functional Modules

All core logic lives under `core_functional_modules/`. These modules are designed to be:

- UI-agnostic
- Composable and reusable
- Safe to invoke from both the Streamlit UI and the CLI

Only responsibilities are summarized here; implementation details live in docstrings.

---

### `config.py` — Central Configuration Model

- Defines the canonical `Config` dataclass used everywhere in the system.
- Loads configuration via a **layered merge**:
  1. Streamlit secrets (typed, preferred in cloud)
  2. Environment variables / `.env`
  3. Optional Google Drive–hosted JSON overrides
- Normalizes types (booleans, lists, dicts) for consistent behavior.
- Acts as the single source of truth for IDs, flags, recipients, and cleanup policies.

This module underpins all other core components.

---

### `config_store.py` — Drive-Backed Config Persistence

- Reads and writes JSON configuration files stored in Google Drive.
- Enables the UI’s “Update defaults” feature.
- Uses resilient, fail-open behavior so missing or malformed configs never break execution.

---

### `google_client.py` — OAuth & Google Service Bootstrapping

- Manages Google OAuth sign-in and token lifecycle (`token.json`).
- Supports both browser-assisted and manual (CLI-style) flows.
- Produces authenticated service clients for:
  - Google Drive
  - Google Sheets
  - Gmail

All Google API usage throughout the system flows through this module.

---

### `drive_utils.py` — Google Drive Operations

- Uploads files (raw or converted to Google Sheets).
- Finds the latest files in folders.
- Copies, renames, and trashes Drive resources.
- Performs **time-based cleanup** of old files.
- Handles Drive query escaping and RFC‑3339 timestamps.

---

### `sheets_utils.py` — Google Sheets Manipulation

- Implements all Sheet-level operations:
  - Copying, deleting, and renaming sheets
  - Writing 2D values
  - Refreshing formulas
  - Forcing text coercion on specific columns
- Provides chunked refresh utilities for large sheets.

This module isolates low-level Sheets API complexity from the pipeline.

---

### `gmail_utils.py` — Email Composition & Sending

- Builds MIME-compliant emails with:
  - PDF attachments
  - Optional full-order attachments
  - Plain-text and HTML bodies
- Sends messages through the Gmail API as the authenticated user.
- Assumes recipient selection has already been resolved upstream.

---

### `logger.py` — Run-Time Status Logging

- Lightweight status logger with:
  - In-memory event tracking
  - Optional file-backed logging (`last_run.log`)
- Designed for real-time UI updates and post-run log downloads.
- Avoids heavy logging frameworks for predictability and portability.

---

### `pipeline.py` — Orchestrated Processing Engine

This is the **core execution engine** of the system.

At a high level, it:

1. Authenticates and initializes Google services.
2. Locates the latest incoming report.
3. Validates date coverage and week boundaries.
4. Prepares or rolls forward calculation workbooks.
5. Refreshes reference sheets.
6. Generates PDFs and Sheets for:
   - Manager reports
   - Full orders
   - Per-report-key outputs
7. Routes emails using a prioritized fallback model.
8. Performs Drive cleanup and retention enforcement.

The pipeline is intentionally **UI-agnostic** and is safe to invoke from:
- The Streamlit UI
- The CLI
- Future automation or scheduling contexts

---

## Supporting Folders

### `documentation/`

This directory contains **developer-facing documentation and tooling**, including:

- Architectural and usage documentation
- Setup, packaging, and deployment notes
- Git workflow guidance
- Dependency lists
- Tooling used to generate the project snapshot markdown bundle

Nothing in this folder is required for runtime execution. It exists to support onboarding, maintenance, and knowledge transfer.

---

### `__dev_input_sales_files/`

This directory contains **example Modisoft sales reports** used for testing and validation.

The files intentionally cover multiple scenarios, including:

- One-week vs. two-week uploads
- Multiple stores
- Invalid or misaligned week boundaries
- Bad start or end dates

These files are **not consumed automatically** by the application and exist purely for manual testing and development.

---

## Files and Folders Not Covered Here

Any files or folders **not explicitly mentioned in this README** (for example: logs, generated outputs, cached tokens, per-user folders, or packaging artifacts) are:

- **Auto-created**
- **Auto-maintained**
- **Managed by the application at runtime**

⚠️ **Do not edit, delete, or manually modify these files unless you fully understand their lifecycle and side effects.**  
Incorrect changes may break authentication, corrupt Drive state, or cause data loss.

---

## How to Navigate the Codebase

- Start with `_user_interface_.py` to understand user flow and orchestration.
- Follow execution into `pipeline.run_pipeline`.
- Use module-level docstrings in `core_functional_modules/` for precise logic and guarantees.
- Treat `Config` as the shared contract that ties the system together.

```

---
### file: documentation/generate_code_bundle.py

```python
#!/usr/bin/env python3


from __future__ import annotations
import argparse
import fnmatch
import os
from pathlib import Path
from typing import Iterable, List, Set, Tuple

# ----------------------------
# Defaults (sane and safe)
# ----------------------------

DEFAULT_IGNORE_DIRS: Set[str] = {
    ".git", ".hg", ".svn",
    ".venv", "venv", "env",
    "build", "dist", "__pycache__",
    ".pytest_cache", ".mypy_cache", ".idea", ".vscode",
}


DEFAULT_IGNORE_NAMES: Set[str] = {
    # OAuth / secrets
    "token.json",
    "credentials.json",
    "web_url_credentials.json",
    ".env",

    # Common secret patterns
    "client_secrets.json",
}


DEFAULT_IGNORE_EXTS: Set[str] = {
    ".pyc", ".pyo", ".pyd",
    ".so", ".dll", ".dylib",
    ".zip", ".tar", ".gz", ".7z",
    ".exe", ".bin",
}


DEFAULT_IGNORE_GLOBS = [
    "**/*credentials*.json",
    "**/*.env",
    "**/*.env.*",
    "**/*secret*.json",
]


# Basic language hints for code fences
LANG_BY_EXT = {
    ".py": "python",
    ".md": "markdown",
    ".txt": "text",
    ".bat": "bat",
    ".cmd": "bat",
    ".sh": "bash",
    ".ps1": "powershell",
    ".json": "json",
    ".ini": "ini",
    ".env": "ini",
    ".yml": "yaml",
    ".yaml": "yaml",
    ".csv": "csv",
    ".ts": "ts",
    ".tsx": "tsx",
    ".js": "javascript",
    ".jsx": "jsx",
    ".html": "html",
    ".css": "css",
    ".toml": "toml",
}


def match_any(path: Path, patterns: Iterable[str]) -> bool:
    """Return True if path matches ANY of the glob patterns (POSIX-style)."""
    s = path.as_posix()
    for pat in patterns:
        if fnmatch.fnmatch(s, pat):
            return True
    return False


def build_tree(root: Path,
               ignore_dirs: Set[str],
               ignore_exts: Set[str],
               ignore_names: Set[str],
               include_globs: List[str] | None,
               exclude_globs: List[str] | None,
               max_bytes: int) -> Tuple[str, List[Path]]:
    """
    Return a (tree_text, files_list) tuple.
    files_list contains all files to embed in the bundle.
    """
    lines_tree: List[str] = []
    files_out: List[Path] = []

    for dirpath, dirnames, filenames in os.walk(root):
        # prune ignored directories
        dirnames[:] = [
            d for d in dirnames
            if d not in ignore_dirs and not d.startswith(".")
        ]
        cur = Path(dirpath)
        rel_dir = cur.relative_to(root)
        depth = 0 if rel_dir == Path(".") else len(rel_dir.parts)
        indent = "  " * depth
        if rel_dir != Path("."):
            lines_tree.append(f"{indent}{rel_dir.as_posix()}/")

        for fname in sorted(filenames):
            p = cur / fname

            # Ignore by name/extension
            if fname in ignore_names:
                entry = fname if rel_dir == Path(".") else f"{indent} {fname}"
                lines_tree.append(entry + " [skipped: secret]")
                continue
            if p.suffix.lower() in ignore_exts:
                continue

            
            # Ignore by glob (secrets)
            if match_any(p.relative_to(root), DEFAULT_IGNORE_GLOBS):
                entry = fname if rel_dir == Path(".") else f"{indent} {fname}"
                lines_tree.append(entry + " [skipped: secret]")
                continue


            # Exclude hidden top-level noise by pattern
            if exclude_globs and match_any(p.relative_to(root), exclude_globs):
                continue
            # If include globs were given, only take matches
            if include_globs and not match_any(p.relative_to(root), include_globs):
                continue

            # Size check
            try:
                if p.stat().st_size > max_bytes:
                    # Do not list it in files_out, but show in tree (optional)
                    entry = fname if rel_dir == Path(".") else f"{indent}  {fname}"
                    lines_tree.append(entry + "  [skipped: too large]")
                    continue
            except Exception:
                # If can't stat, skip silently
                continue

            # Accept file
            entry = fname if rel_dir == Path(".") else f"{indent}  {fname}"
            lines_tree.append(entry)
            files_out.append(p)

    return "\n".join(lines_tree), files_out


def read_text_safe(p: Path) -> str | None:
    try:
        return p.read_text(encoding="utf-8")
    except UnicodeDecodeError:
        # Binary or non-UTF8; skip
        return None
    except Exception:
        return None


def make_bundle_markdown(root: Path,
                         out: Path,
                         include_globs: List[str] | None,
                         exclude_globs: List[str] | None,
                         max_bytes: int,
                         ignore_dirs: Set[str],
                         ignore_exts: Set[str],
                         ignore_names: Set[str]) -> Path:
    tree_text, files_to_embed = build_tree(
        root, ignore_dirs, ignore_exts, ignore_names,
        include_globs, exclude_globs, max_bytes
    )

    header = f"""# Project Snapshot (CodeBundle)

                This single Markdown file contains a **self-contained snapshot** of your project so another AI/engineer can review or modify it without needing the original folder.

                **How to use this file with an AI**
                1. Upload or paste this file as a single attachment.
                2. Ask for changes; the AI can reference specific `file:` sections below.
                3. Copy updated blocks back into the corresponding files in your project.

                > Notes: secrets like `token.json` are intentionally excluded. Virtual envs and build artifacts are omitted to keep this readable.

                ## Directory tree (filtered)
                {tree_text}"""

    sections: List[str] = [header]

    for p in sorted(files_to_embed, key=lambda x: x.relative_to(root).as_posix()):
        rel = p.relative_to(root)
        lang = LANG_BY_EXT.get(p.suffix.lower(), "")
        content = read_text_safe(p)
        if content is None:
            # Skip non-text files silently
            continue
        fence = "```"
        sections.append(
            f"\n---\n### file: {rel.as_posix()}\n\n{fence}{lang}\n{content}\n{fence}\n"
        )

    out.write_text("".join(sections), encoding="utf-8")
    return out


def parse_args() -> argparse.Namespace:
    ap = argparse.ArgumentParser(
        description="Generate a single-file Markdown 'CodeBundle' of a project."
    )
    # Defaults are computed in main()
    ap.add_argument(
        "--root",
        default=None,
        help="Project root folder (default: parent of this script)",
    )
    ap.add_argument(
        "--out",
        default=None,
        help="Output Markdown file path (default: PROJECT_SNAPSHOT_CODEBUNDLE.md in the script's directory)",
    )
    ap.add_argument("--include", nargs="*", default=None,
                    help='Optional list of glob patterns to include (e.g. "**/*.py" "**/*.md").')
    ap.add_argument("--exclude", nargs="*", default=None,
                    help='Optional list of glob patterns to exclude (e.g. ".venv/**" "dist/**").')
    ap.add_argument("--max-bytes", type=int, default=1_000_000,
                    help="Max file size to embed (bytes). Oversized files are listed but skipped. Default: 1,000,000.")
    ap.add_argument("--no-default-ignores", action="store_true",
                    help="Disable default ignore sets for dirs/exts/names.")
    return ap.parse_args()


def main():
    print("CodeBundle writing starting...")
    args = parse_args()

    # Locate the script and its parent (project root)
    script_dir = Path(__file__).resolve().parent          # .../favtrip_reporting/documentation
    project_root = script_dir.parent                      # .../favtrip_reporting

    # ROOT: default to the parent of this script unless overridden
    if args.root is None:
        root = project_root
    else:
        root_arg = Path(args.root)
        root = root_arg.resolve() if root_arg.is_absolute() else (Path.cwd() / root_arg).resolve()

    if not root.exists() or not root.is_dir():
        raise SystemExit(f"Root does not exist or is not a directory: {root}")

    # OUT: default to PROJECT_SNAPSHOT_CODEBUNDLE.md in the script's directory (the child)
    if args.out is None:
        out = (script_dir / "PROJECT_SNAPSHOT_CODEBUNDLE.md").resolve()
    else:
        out_arg = Path(args.out)
        # If user gave a relative path, resolve it under the chosen root; absolute paths are respected
        out = out_arg.resolve() if out_arg.is_absolute() else (root / out_arg).resolve()

    # ... keep the rest of your function unchanged ...
    if args.no_default_ignores:
        ignore_dirs = set()
        ignore_exts = set()
        ignore_names = set()
    else:
        ignore_dirs = set(DEFAULT_IGNORE_DIRS)
        ignore_exts = set(DEFAULT_IGNORE_EXTS)
        ignore_names = set(DEFAULT_IGNORE_NAMES)

    out.parent.mkdir(parents=True, exist_ok=True)
    result_path = make_bundle_markdown(
        root=root,
        out=out,
        include_globs=args.include,
        exclude_globs=args.exclude,
        max_bytes=args.max_bytes,
        ignore_dirs=ignore_dirs,
        ignore_exts=ignore_exts,
        ignore_names=ignore_names,
    )
    print(f"CodeBundle written to: {result_path}")


if __name__ == "__main__":
    main()

#cd "C:\Users\rjrul\OneDrive - University of Iowa\000 Current Semester\004 BAIS 4150 BAIS Capstone\favtrip_reporting"
#python documentation/generate_code_bundle.py
```

---
### file: documentation/git_workflow.txt

```text

---
1. Create a dev branch

# Switch to dev branch
    git checkout dev

# Push dev branch to GitHub (the -u flag sets upstream tracking)
    git push -u origin dev

---
2. Do all development on dev
Make changes normally, then stage, commit, and push:

    git add .
    git commit -m "your message here"
    git push

All pushes go to dev (not main).

---
3. Push dev → main when ready (production release)
Use a Pull Request on GitHub:
1. Go to the repo on GitHub.
2. You will see a banner offering “Compare & Pull Request” (dev → main).
3. Open the Pull Request.
4. Review and click “Merge Pull Request”.

After merging, update your local main branch:

    git checkout main
    git pull

---
4. Keep dev updated with the latest main
After merging into main, update dev:

    git checkout dev
    git pull origin main

Or:

    git merge main

This prevents dev from drifting behind main.

---
5. Summary
- main = production, always stable
- dev = development, experimental work
- Merge dev → main only after testing

```

---
### file: documentation/requirements.txt

```text
google-api-python-client
google-auth
google-auth-oauthlib
google-auth-httplib2
httplib2
requests
python-dotenv
streamlit
streamlit-autorefresh
openpyxl
pandas
uuid
```

---
### file: last_run.log

```
[2026-03-03 23:16:33] INFO: Authorizing with Google APIs…
[2026-03-03 23:16:33] INFO: Google services ready
[2026-03-03 23:16:33] INFO: Finding latest incoming spreadsheet…
[2026-03-03 23:16:34] INFO: Latest incoming: Testing Sales Report - Week 5 (1kNjmEbljdUIqJUjwh8e2-plfLWRZSGHxMHxL50-2ce4)
[2026-03-03 23:16:34] INFO: Preparing calculations workbook…
[2026-03-03 23:16:42] INFO: Copied old 'Current Week' to 'Last Week'
[2026-03-03 23:16:50] INFO: Inserted new 'Current Week' from latest incoming report
[2026-03-03 23:16:50] INFO: Refreshing reference sheets (prefix 'REFR: ')…
[2026-03-03 23:17:17] INFO: [1/3] Recalc OK: REFR: Charts
[2026-03-03 23:17:22] INFO: [2/3] Recalc OK: REFR: Values
[2026-03-03 23:18:15] INFO: [3/3] Recalc OK: REFR: Order Calcs
[2026-03-03 23:18:17] INFO: Location: Favtrip_Independence; Timestamp: 2026-03-03-11-18-PM
[2026-03-03 23:18:17] INFO: Exporting Manager Report (PDF)…
[2026-03-03 23:18:21] INFO: Uploaded Manager PDF: https://drive.google.com/file/d/1UILYk_poamIwe6JrFQUGZIz9az9iBsdU/view?usp=drivesdk
[2026-03-03 23:18:21] INFO: Exporting Master Order (CSV)…
[2026-03-03 23:18:34] INFO: Uploaded FULL sheet: https://docs.google.com/spreadsheets/d/1bOcle5eMseFc43TZMnXLqgHR7_DcPs13ntb6PSBians/edit?usp=drivesdk
[2026-03-03 23:18:42] INFO: Emailed COFFEE
[2026-03-03 23:18:43] INFO: Manager email sent
[2026-03-03 23:18:43] INFO: Separate full order email disabled
[2026-03-03 23:18:43] INFO: Run completed in 00:02:09

```

---
### file: launcher_streamlit.py

```python
import os, sys, subprocess

def main():
    # Ensure current working directory is bundle dir
    base = getattr(sys, '_MEIPASS', os.path.abspath(os.path.dirname(__file__)))
    os.chdir(base)
    subprocess.call([sys.executable, "-m", "streamlit", "run", "ui_streamlit.py"])

if __name__ == "__main__":
    main()

```

---
### file: requirements.txt

```text
-r documentation/requirements.txt
```

---
### file: setup_py2app.py

```python
from setuptools import setup

APP = ['launcher_streamlit.py']
OPTIONS = {
    'argv_emulation': True,
    'plist': {
        'CFBundleName': 'FavTripPipelineUI',
    },
    'packages': ['googleapiclient', 'google', 'httplib2', 'google_auth_oauthlib', 'google_auth_httplib2', 'dotenv', 'requests', 'streamlit'],
}

data_files = ['ui_streamlit.py', 'cli.py', 'requirements.txt', '.env', 'credentials.json']

setup(
    app=APP,
    options={'py2app': OPTIONS},
    data_files=data_files,
)

```
