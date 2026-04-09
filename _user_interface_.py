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
from core_functional_modules.sheets_utils import force_named_range_timestamp
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


STORE_LIST = ['FAV TRIP GRANDVIEW LLC', 'FAV TRIP KCMO LLC', 'FAVTRIP INDEPENDENCE', 'TRUMAN STOP']
KEY_LIST = ['AUTO', 'BAKERY', 'BEV', 'CARTON', 'CBD', 'CHEW', 'CIGARS', 'COFFEE', 'DELI ITEM', 'DELIVERY', 'E CIGG', 'E-CIG', 'FOUNTAIN', 'GROCERY', 'HBA', 'ICE', 'KANBE', 'PK CIGG', 'PLU NOT FOUND', 'REFILL', 'SLUSH', 'SNACK/LO TAX GRO']
BEV_SUB_KEY_LIST = ['7UP', 'COKE', 'HILAND', 'PEPSI', 'REDBULL', 'WF', 'Unassigned']



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
            options = []
            
            for key in KEY_LIST:
                if key != "PLU NOT FOUND":
                    options.append(key)
            options.append("PLU NOT FOUND")
            
            for sub_key in BEV_SUB_KEY_LIST:
                if sub_key != "Unassigned":
                    options.append(f"BEV-{sub_key}")

            options.append(f"BEV-Unassigned")

            # Remove duplicates + sort
            options = sorted(set(options))

            if "PLU NOT FOUND" in KEY_LIST:
                options.append("PLU NOT FOUND")
            

            selected_keys = st.multiselect(
                "Keys to run",
                options=options,
                default=cfg.REPORT_KEY_RUN_LIST or [],
                help="Select one or more keys. Sub-keys shown as ReportKey-SubKey."
            )

            # Convert back to your config format
            report_keys = ",".join(selected_keys)

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
                column_config={
                    "Store (optional)": st.column_config.SelectboxColumn(
                        "Store (optional)",
                        options=STORE_LIST
                    ),
                    "Report Key (optional)": st.column_config.SelectboxColumn(
                        "Report Key (optional)",
                        options=KEY_LIST
                    ),
                    "Sub-Report Key (optional)": st.column_config.SelectboxColumn(
                        "Sub-Report Key (optional)",
                        options=BEV_SUB_KEY_LIST
                    ),
                }
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

                    
            if st.button(
                "⏱️ Force File Renews",
                type="secondary",
                width="stretch",
                help="Forces all user calculation files to be renewed the next time the user runs the pipeline",
            ):
                st.session_state["confirm_force_sheet_timestamp"] = True

            
            
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
                        _rerun()

            with col_cancel:
                if st.button("❌ Cancel", width="stretch"):
                    st.session_state.pop("confirm_merge_dev_to_main", None)
                    _rerun()
        
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
                    _rerun()
        
        @st.dialog("⚠️ Confirm Force Sheet Timestamp")
        def confirm_force_sheet_timestamp():
            st.markdown(
                """
                **You are about to force all users to renew their calculation files.**

                Proceed only if you know why you are doing this.
                """
            )

            col_confirm, col_cancel = st.columns(2)

            with col_confirm:
                if st.button("✅ Yes — Force File Renews", type="primary", width="stretch"):
                    try:
                        cfg.CALC_SPREADSHEET_ID

                        if not cfg.CALC_SPREADSHEET_ID:
                            st.error("Missing TARGET_SPREADSHEET_ID in secrets.")
                        else:
                            creds = load_valid_token(cfg.SCOPES)
                            if not creds:
                                st.error("Google authentication required.")
                            else:
                                sheets_svc, _, _ = services(creds, cfg.HTTP_TIMEOUT_SECONDS)

                                force_named_range_timestamp(
                                    sheets_svc,
                                    spreadsheet_id=cfg.CALC_SPREADSHEET_ID,
                                    named_range="_update",
                                )

                                st.success("✅ Timestamp successfully refreshed.")

                    except Exception as e:
                        st.error(f"Timestamp refresh failed: {e}")

                    finally:
                        st.session_state.pop("confirm_force_sheet_timestamp", None)
                        _rerun()

            with col_cancel:
                if st.button("❌ Cancel", width="stretch"):
                    st.session_state.pop("confirm_force_sheet_timestamp", None)
                    _rerun()



                    


        # Trigger dialog
        if st.session_state.get("confirm_push_dev_to_prod"):
            confirm_push_dev_to_prod()
        
        if st.session_state.get("confirm_merge_dev_to_main"):
            confirm_merge_dev_to_main()
        
        if st.session_state.get("confirm_rebuild_workspace"):
            confirm_rebuild_workspace()
        
        if st.session_state.get("confirm_force_sheet_timestamp"):
            confirm_force_sheet_timestamp()

        

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
