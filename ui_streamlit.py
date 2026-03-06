import os
import time
import threading
import json
import base64
import hashlib
import secrets
import re

import streamlit as st
from streamlit.components.v1 import html
from streamlit.components.v1 import html as _html_listener

from google_auth_oauthlib.flow import Flow
from google.oauth2.credentials import Credentials

from favtrip.google_client import load_valid_token, services, clear_token
from favtrip.config_store import save_config_to_drive
from favtrip.config import Config
from favtrip.logger import StatusLogger
from favtrip.pipeline import run_pipeline
from favtrip.drive_utils import upload_to_drive


# =========================
# Constants & Simple Helpers
# =========================

EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")


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
    """
    cfg = json.loads(st.secrets["GOOGLE_CREDENTIALS"])
    state_obj = _parse_state(state_b64)
    code_verifier = state_obj.get("v")
    redirect = state_obj.get("r") or st.session_state.get("_oauth_redirect") or _redirect_base()

    if not code_verifier:
        st.error("OAuth state did not include a PKCE code_verifier.")
        st.stop()

    flow = Flow.from_client_config(cfg, scopes=scopes, redirect_uri=redirect)
    # IMPORTANT: include code_verifier here to satisfy PKCE.
    flow.fetch_token(code=code, code_verifier=code_verifier)

    creds = flow.credentials
    with open("token.json", "w") as f:
        f.write(creds.to_json())
    return creds

# --- OAuth Redirect Handler (Web/PKCE only) ---
def handle_oauth_redirect_if_any(cfg):
    # Parse query params safely
    query_params = st.query_params if hasattr(st, "query_params") else st.experimental_get_query_params()
    code = None
    state = None

    # Accept both styles of query param access
    if isinstance(query_params, dict):
        # streamlit>=1.31: st.query_params is Mapping[str, str]
        code = query_params.get("code")
        state = query_params.get("state")
        # Older API returns list values
        if isinstance(code, list): code = code[0] if code else None
        if isinstance(state, list): state = state[0] if state else None

    if not code or not state:
        return  # no redirect present

    try:
        creds = finish_web_oauth(code=code, state_b64=state, scopes=cfg.SCOPES)

        st.success("✅ Google sign‑in complete.")

        # Clean URL
        if hasattr(st, "query_params"):
            st.query_params.clear()
        else:
            st.experimental_set_query_params()

        # 🆕 Signal opener and close this tab
        html("""
        <script>
        try {
            // Notify opener (original tab) that OAuth finished
            if (window.opener && !window.opener.closed) {
            window.opener.postMessage({ type: "favtrip_oauth_done" }, "*");
            }
        } catch (e) {}
        // Close this OAuth tab
        window.close();
        </script>
        """, height=0)

        # Fallback UI in case close() is blocked
        st.success("✅ Google sign‑in complete.")
        st.toast("Signed in to Google.")
        st.caption("You can close this tab if it didn't close automatically.")
        st.rerun()
    except Exception as e:
        st.error(f"OAuth finish failed: {e}")


# =========================
# UI Sections
# =========================

def render_run_form(cfg):
    # --- STATE INIT ---
    if "incoming_uploader_version" not in st.session_state:
        st.session_state.incoming_uploader_version = 0
    if "incoming_selected_name" not in st.session_state:
        st.session_state.incoming_selected_name = None
    if "incoming_uploaded_ok" not in st.session_state:
        st.session_state.incoming_uploaded_ok = False

    uploader_key = f"incoming_upload_v{st.session_state.incoming_uploader_version}"

    # =========================
    # UPLOAD CARD (outside form)
    # =========================
    
    with st.container(border=True):
        st.subheader("Upload Current Week Sales Report")
        # Use the SAME column grid as the run form header: [4, 1, 1]
        up_col, _, upbtn_col = st.columns([4, 1, 1])

        with up_col:
            incoming_file = st.file_uploader(
                "Upload Current Week Sales Report",
                type=["xlsx", "csv"],
                key=uploader_key,
                help="Please upload the current week's 'Live Items Report' from Modisoft as an XLSX or CSV file.",
                label_visibility="collapsed",
                accept_multiple_files=False,
            )

        # Track selection to manage gating (must click Upload Now successfully before Run)
        file_selected = incoming_file is not None
        if file_selected and st.session_state.get("incoming_selected_name") != incoming_file.name:
            st.session_state.incoming_selected_name = incoming_file.name
            st.session_state.incoming_uploaded_ok = False

        with upbtn_col:
            st.markdown('<div class="ft-right-btn">', unsafe_allow_html=True)
            upload_clicked = st.button(
                "⬆️ Upload Now",
                use_container_width=True,
                disabled=(not file_selected),
                type="secondary",
                key="upload_submit",
            )
            st.markdown('</div>', unsafe_allow_html=True)

        # --- Handle the upload action immediately ---
        if upload_clicked:
            if not cfg.INCOMING_FOLDER_ID:
                st.error("Incoming Folder ID is empty. Set it under **Advanced → Incoming Folder ID**.")
            elif incoming_file is None:
                st.warning("Choose a .xlsx or .csv file first.")
            else:
                try:
                    drive = _get_drive_service_or_raise(cfg)
                    media_mime = _infer_media_mime(incoming_file.name)
                    base_name = os.path.splitext(incoming_file.name)[0]
                    nice_name = f"{base_name} (uploaded via UI)"
                    created = upload_to_drive(
                        drive,
                        data=incoming_file.getvalue(),
                        name=nice_name,
                        mime=media_mime,
                        folder_id=cfg.INCOMING_FOLDER_ID,
                        to_sheet=True,
                    )
                    link = created.get("webViewLink", "")

                    st.session_state.incoming_uploaded_ok = True
                    st.session_state.incoming_uploader_version += 1
                    st.session_state.incoming_selected_name = None

                    st.success("✅ Uploaded to Incoming as a Google Sheet.")
                    if link:
                        st.link_button("Open uploaded Sheet", link, use_container_width=True)
                    st.caption("This will be treated as the latest incoming report on the next run.")
                    _rerun()
                except Exception as e:
                    st.error(f"Upload failed: {e}")


    # =========================
    # RUN FORM CARD
    # =========================
    # When upload is OK, make the Run button green by adding .ft-run-green to the page segment
    run_form_wrapper_classes = "ft-card ft-row"
    if st.session_state.get("incoming_uploaded_ok", False):
        run_form_wrapper_classes += " ft-run-green"

    
    with st.form("run_form"):
        # Header row uses the same columns to align the Run button with Upload button above
        tl, _, col_run = st.columns([4, 1, 1])
        with tl:
            st.subheader("Run Options")
            st.caption("Configure email behavior and report keys. Use **Advanced** for IDs/GIDs/timezone.")

        # Gate: require successful upload only if a new file is currently selected but not uploaded
        if not file_selected:
            run_disabled = False
        elif file_selected and not st.session_state.get("incoming_uploaded_ok", False):
            run_disabled = True
        else:
            run_disabled = False

        with col_run:
            # Right-align and full-width, matching Upload Now
            st.markdown('<div class="ft-right-btn">', unsafe_allow_html=True)
            submitted = st.form_submit_button(
                "▶️ Run Pipeline",
                use_container_width=True,
                disabled=run_disabled,
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
        colk1, colk2 = st.columns([0.45, 1.55])
        with colk1:
            use_all = st.toggle(
                "Use all keys from CSV",
                value=cfg.USE_ALL_REPORT_KEYS,
                help="ON: process every key found. OFF: only the keys you list."
            )
        with colk2:
            report_keys = st.text_input(
                "Keys to run (comma)",
                value=",".join(cfg.REPORT_KEY_RUN_LIST or []),
                help="Used when 'Use all keys' is OFF. Example: COFFEE,GROCERY"
            )

        # Email Behavior
        st.markdown("##### Email Behavior")
        cole1, cole2, cole3 = st.columns([1, 1, 1])
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

        # Per‑Report‑Key Recipients
        with st.expander("Per‑Report‑Key Recipients (optional)", expanded=False):
            st.caption("Map **REPORT KEY (ALL CAPS)** → **Emails (comma)**.")
            rows = []
            if cfg.REPORT_KEY_RECIPIENTS:
                for k, v in cfg.REPORT_KEY_RECIPIENTS.items():
                    rows.append({"REPORT KEY (ALL CAPS)": k, "Emails (comma)": ",".join(v or [])})
            else:
                rows = [{"REPORT KEY (ALL CAPS)": "", "Emails (comma)": ""}]
            edited_rows = st.data_editor(
                rows,
                num_rows="dynamic",
                use_container_width=True,
                key="rk_editor",
            )

            rk_issues, rk_preview, rk_map_preview = _analyze_rk_rows(edited_rows)
            if rk_preview:
                with st.expander("Row template preview"):
                    st.code("\n".join(rk_preview), language="text")
            if rk_issues:
                st.warning("Per‑report‑key recipient issues:\n\n- " + "\n- ".join(rk_issues))

        # Advanced
        with st.expander("Advanced (IDs, GIDs, Timezone, Redirect Port)", expanded=False):
            st.markdown("##### Google Drive / Sheets IDs")
            ga1, ga2 = st.columns([1, 1])
            with ga1:
                calc_id = st.text_input("Calculations Spreadsheet ID", value=cfg.CALC_SPREADSHEET_ID)
                mgr_folder = st.text_input("Manager Report Folder ID", value=cfg.MANAGER_REPORT_FOLDER_ID)
            with ga2:
                incoming_id = st.text_input("Incoming Folder ID", value=cfg.INCOMING_FOLDER_ID)
                order_folder = st.text_input("Order Report Folder ID", value=cfg.ORDER_REPORT_FOLDER_ID)

            st.markdown("##### GIDs & Named Ranges")
            gb1, gb2 = st.columns([1, 1])
            with gb1:
                gid_mgr = st.text_input("Manager Report gid", value=str(cfg.GID_MANAGER_PDF))
                loc_sheet = st.text_input("Location Sheet Title", value=cfg.LOCATION_SHEET_TITLE)
            with gb2:
                gid_order = st.text_input("Order CSV gid", value=str(cfg.GID_ORDER_CSV))
                loc_range = st.text_input("Location Named Range", value=cfg.LOCATION_NAMED_RANGE)

            st.markdown("##### Time & OAuth")
            gc1, gc2 = st.columns([1, 1])
            with gc1:
                tz = st.text_input("Timestamp Timezone", value=cfg.TIMESTAMP_TZ)
                tfmt = st.text_input("Timestamp Format", value=cfg.TIMESTAMP_FMT)
            with gc2:
                raw_redirect_port = int(cfg.REDIRECT_PORT) if str(cfg.REDIRECT_PORT).isdigit() else 0
                redirect_port = st.number_input(
                    "Redirect Port (0 = auto)",
                    min_value=0, max_value=65535,
                    value=raw_redirect_port if raw_redirect_port in (0, *range(1024, 65536)) else 0,
                    help="Use 0 to auto-pick a free port. Otherwise choose 1024–65535."
                )

        save_drive_defaults = st.checkbox("Update defaults", value=False)

        # ----- Submission handling -----
        if submitted:
            # Apply per-run config
            cfg.TO_RECIPIENTS = _split_emails(to)
            cfg.CC_RECIPIENTS = _split_emails(cc)
            cfg.USE_ALL_REPORT_KEYS = use_all
            cfg.REPORT_KEY_RUN_LIST = [s.strip().upper() for s in (report_keys or "").split(",") if s.strip()]

            cfg.INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL = include_full
            cfg.SEND_SEPARATE_FULL_ORDER_EMAIL = send_full
            cfg.EMAIL_MANAGER_REPORT = bool(email_mgr)

            cfg.CALC_SPREADSHEET_ID = calc_id
            cfg.INCOMING_FOLDER_ID = incoming_id
            cfg.MANAGER_REPORT_FOLDER_ID = mgr_folder
            cfg.ORDER_REPORT_FOLDER_ID = order_folder
            cfg.REDIRECT_PORT = int(redirect_port)

            cfg.GID_MANAGER_PDF = gid_mgr
            cfg.GID_ORDER_CSV = gid_order
            cfg.LOCATION_SHEET_TITLE = loc_sheet
            cfg.LOCATION_NAMED_RANGE = loc_range
            cfg.TIMESTAMP_TZ = tz
            cfg.TIMESTAMP_FMT = tfmt

            # Per-key recipients
            rk_map = {}
            for r in edited_rows:
                key = (r.get("REPORT KEY (ALL CAPS)") or "").strip()
                emails_csv = r.get("Emails (comma)") or ""
                emails = _parse_emails(emails_csv)
                if key and emails:
                    rk_map[key] = emails
            cfg.REPORT_KEY_RECIPIENTS = rk_map

            # Warnings
            requested_keys = [s.strip().upper() for s in (report_keys or "").split(",") if s.strip()]
            if not use_all and not requested_keys:
                st.warning(
                    "You left **Use all Report Keys** OFF but provided **no keys** to run. "
                    "No per‑key outputs will be generated unless you add keys.",
                    icon="⚠️"
                )

            any_to = bool(_parse_emails(to) or (cfg.TO_RECIPIENTS or []))
            any_default = bool(cfg.DEFAULT_ORDER_RECIPIENTS or [])
            any_per_key = bool(cfg.REPORT_KEY_RECIPIENTS)
            if not (any_to or any_default or any_per_key):
                st.warning(
                    "No recipients are defined: **TO**, **DEFAULT_ORDER_RECIPIENTS**, and **Per‑Report‑Key** are all empty. "
                    "Emails will not be sent.",
                    icon="⚠️"
                )

            if 'rk_issues' in locals() and rk_issues:
                st.warning(
                    "Per‑report‑key recipient issues detected above. These may prevent emails from sending correctly:\n\n- "
                    + "\n- ".join(rk_issues),
                    icon="⚠️"
                )

            # Save edited defaults to Drive JSON (optional)
            if save_drive_defaults:
                try:
                    creds = load_valid_token(cfg.SCOPES)
                    if not creds:
                        st.error("Not authenticated. Please complete Google sign‑in first (top of page).")
                    else:
                        _sheets, drive, _gmail = services(creds, cfg.HTTP_TIMEOUT_SECONDS)
                        drive_defaults = {
                            "CALC_SPREADSHEET_ID": cfg.CALC_SPREADSHEET_ID,
                            "INCOMING_FOLDER_ID": cfg.INCOMING_FOLDER_ID,
                            "MANAGER_REPORT_FOLDER_ID": cfg.MANAGER_REPORT_FOLDER_ID,
                            "ORDER_REPORT_FOLDER_ID": cfg.ORDER_REPORT_FOLDER_ID,
                            "GID_MANAGER_PDF": cfg.GID_MANAGER_PDF,
                            "GID_ORDER_CSV": cfg.GID_ORDER_CSV,
                            "LOCATION_SHEET_TITLE": cfg.LOCATION_SHEET_TITLE,
                            "LOCATION_NAMED_RANGE": cfg.LOCATION_NAMED_RANGE,
                            "TIMESTAMP_TZ": cfg.TIMESTAMP_TZ,
                            "TIMESTAMP_FMT": cfg.TIMESTAMP_FMT,
                            "TO_RECIPIENTS": cfg.TO_RECIPIENTS,
                            "CC_RECIPIENTS": cfg.CC_RECIPIENTS,
                            "USE_ALL_REPORT_KEYS": cfg.USE_ALL_REPORT_KEYS,
                            "REPORT_KEY_RUN_LIST": cfg.REPORT_KEY_RUN_LIST,
                            "REPORT_KEY_RECIPIENTS": cfg.REPORT_KEY_RECIPIENTS,
                            "DEFAULT_ORDER_RECIPIENTS": cfg.DEFAULT_ORDER_RECIPIENTS,
                            "INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL": cfg.INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL,
                            "SEND_SEPARATE_FULL_ORDER_EMAIL": cfg.SEND_SEPARATE_FULL_ORDER_EMAIL,
                            "EMAIL_MANAGER_REPORT": cfg.EMAIL_MANAGER_REPORT,
                        }
                        CONFIG_FILE_ID = (st.secrets.get("CONFIG_FILE_ID", "") or "").strip()
                        new_id = save_config_to_drive(
                            drive,
                            drive_defaults,
                            file_id=CONFIG_FILE_ID or None,
                        )
                        st.success(f"Saved defaults to Drive config (file id: {new_id}).")
                        if not CONFIG_FILE_ID:
                            st.info(
                                "Tip: add this ID to Streamlit Secrets as `CONFIG_FILE_ID` to pin the same file for all runs:\n"
                                f"`{new_id}`"
                            )
                except Exception as e:
                    st.error(f"Failed to save defaults to Drive: {e}")

            # If forced reauth is requested via config, just clear token and return to auth gate
            if getattr(cfg, "FORCE_REAUTH", False):
                clear_token()
                st.session_state.auth_required = True
                st.info("Re-auth required for this run. Please sign in again.")
                _rerun()
                return

            # Run
            logger = StatusLogger(print_to_console=True, file_path="last_run.log", overwrite=True)
            run_pipeline(cfg, logger)
            
    # 🔚 CLOSE the run card wrapper (ALWAYS close after the form block)
    

def render_auth_panel(cfg):
    st.markdown("### Google Sign‑in")
    st.caption("Sign in with your Google account to gain access to the app.")

    if st.button("🔐 Sign in with Google", type="primary", use_container_width=True):
        try:
            auth_url = start_web_oauth(cfg.SCOPES)

            # 🆕 Open Google OAuth in a NEW TAB (popup-style), like your old version.
            html(f"""
            <script>
              // Open in a new tab; 'noopener' avoids giving the new tab access to window.opener
              window.open({json.dumps(auth_url)}, "_blank", "noopener");
            </script>
            """, height=0)

            # Optional: friendly message in the current tab
            st.info("A new browser tab was opened for Google sign‑in. After it completes, return to this tab.")
            st.stop()
        except Exception as e:
            st.error(f"Failed to start OAuth: {e}")

    st.info("If you just completed sign‑in and returned here, the page will process your login automatically.")


def render_sidebar():
    st.sidebar.header("Utilities")
    if st.session_state.get("auth_required", True):
        if st.sidebar.button("Sign in"):
            try:
                auth_url = start_web_oauth(Config.load().SCOPES)

                html(f"""
                <script>
                  window.open({json.dumps(auth_url)}, "_blank", "noopener");
                </script>
                """, height=0)
                st.sidebar.info("A new tab was opened for Google sign‑in.")
                st.stop()
            except Exception as e:
                st.sidebar.error(f"Failed to start OAuth: {e}")
    else:
        st.sidebar.success("Signed in")
        if st.sidebar.button("Sign out"):
            clear_token()
            st.session_state.auth_required = True
            st.rerun()

# =========================
# App Entrypoint
# =========================

#st.title("🧾 FavTrip Reporting Pipeline")


st.set_page_config(
    page_title="FT Reporting",
    page_icon="🧾",          # emoji or path/URL to an image
    layout="wide",           # "centered" or "wide"
    initial_sidebar_state="expanded",  # "auto", "expanded", "collapsed"
    menu_items={
        "Get Help": "mailto:ryan-morrow@uiowa.edu",
        "Report a bug": "https://github.com/ryan-j-morrow/favtrip_reporting/issues",
        "About": "FavTrip Reporting Pipeline",
    },
)


_html_listener("""
<script>
  // Listen for 'oauth done' message and reload the page to flip the gate
  window.addEventListener("message", function(e) {
    try {
      if (e && e.data && e.data.type === "favtrip_oauth_done") {
        window.location.reload();
      }
    } catch (_) {}
  }, false);
</script>
""", height=0)

st.title("🧾 FavTrip Reporting Pipeline")

# Global card/layout styles (once)
st.markdown("""
<style>
.ft-card {
  border: 1px solid rgba(0,0,0,0.08);
  border-radius: 10px;
  padding: 1rem;
  background: #fff;
  margin-bottom: 1rem;
}
.ft-card h4 { margin: 0 0 .5rem 0; }
.ft-row { display: block; }
.ft-right-btn > div { display: flex; justify-content: flex-end; }
.ft-right-btn > div button, .ft-right-btn > div a { width: 100%; }

/* Makes the Run button green when upload succeeded (applied via wrapper class) */
.ft-run-green div[data-testid="stFormSubmitButton"] button {
  background-color: #188038 !important;
  color: #fff !important;
}
.ft-run-green div[data-testid="stFormSubmitButton"] button:hover {
  filter: brightness(0.95);
}
</style>
""", unsafe_allow_html=True)

cfg = Config.load()

# Handle OAuth redirect (if present in URL)
handle_oauth_redirect_if_any(cfg)

st.session_state.auth_required = (load_valid_token(cfg.SCOPES) is None)

# Session state defaults
if "incoming_selected_name" not in st.session_state:
    st.session_state.incoming_selected_name = None
if "incoming_uploaded_ok" not in st.session_state:
    st.session_state.incoming_uploaded_ok = False
if "offer_log_download" not in st.session_state:
    st.session_state.offer_log_download = False


# Sidebar (always visible)
with st.sidebar:
    render_sidebar()

# Auth gate
if st.session_state.auth_required:
    render_auth_panel(cfg)
else:
    # Ensure config reflects Drive-based overrides after auth
    cfg = Config.load()
    render_run_form(cfg)