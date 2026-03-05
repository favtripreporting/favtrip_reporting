import os
import time
import threading
import streamlit as st


import base64, hashlib, secrets, json
from google_auth_oauthlib.flow import Flow
from google.oauth2.credentials import Credentials


# --- add near the top of ui_streamlit.py (after imports) ---
import base64, hashlib, json, secrets
import streamlit as st
from google_auth_oauthlib.flow import Flow
from google.oauth2.credentials import Credentials


from favtrip.google_client import load_valid_token, services
from favtrip.config_store import load_config_from_drive


def _b64url(b: bytes) -> str:
    return base64.urlsafe_b64encode(b).rstrip(b"=").decode("ascii")

def _pkce_pair() -> tuple[str, str]:
    """
    Generate a high-entropy PKCE code_verifier and its S256 code_challenge.
    RFC 7636 requires 43–128 chars; this approach yields a URL-safe value.  [2](https://auth0.com/docs/get-started/authentication-and-authorization-flow/authorization-code-flow-with-pkce)
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
        # PKCE parameters:
        code_challenge=code_challenge,
        code_challenge_method="S256",
    )

    # Keep only minimal context; state carries the verifier.
    st.session_state["_oauth_redirect"] = redirect
    return auth_url

def _parse_state(state_b64: str) -> dict:
    # Add padding back for base64 decoding if needed
    padding = "=" * ((4 - len(state_b64) % 4) % 4)
    raw = base64.urlsafe_b64decode(state_b64 + padding)
    return json.loads(raw.decode("utf-8"))

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
    # IMPORTANT: include code_verifier here to satisfy PKCE.  [1](https://github.com/googleapis/google-auth-library-python-oauthlib/issues/46)
    flow.fetch_token(code=code, code_verifier=code_verifier)

    creds = flow.credentials
    with open("token.json", "w") as f:
        f.write(creds.to_json())
    return creds


from favtrip.config_store import save_config_to_drive
from favtrip.config import Config
from favtrip.logger import StatusLogger
from favtrip.pipeline import run_pipeline
from favtrip.google_client import (
    start_oauth,
    finish_oauth,
    load_valid_token,
    clear_token,
    services,
)

def _rerun():
    # Works on Streamlit >= 1.27 (st.rerun) and older (experimental_rerun)
    try:
        import streamlit as st
        st.rerun()
    except AttributeError:
        st.experimental_rerun()

st.set_page_config(page_title="FavTrip Reporting Pipeline", page_icon="🧾", layout="wide")

# --- Larger Run button ---
st.markdown(
    """
    <style>
      div.stButton > button:first-child {
        font-size: 1.1rem;
        padding: 0.6rem 1.2rem;
      }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("🧾 FavTrip Reporting Pipeline")

cfg = Config.load()

#Old Config Diagonistics
def hide():
    with st.expander("🔎 Config diagnostics", expanded=False):
        try:
            CONFIG_FILE_ID = (st.secrets.get("CONFIG_FILE_ID", "") or "").strip()
            st.write(f"CONFIG_FILE_ID (Secrets): `{CONFIG_FILE_ID or '(unset)'}`")

            token = load_valid_token(cfg.SCOPES)
            if not token:
                st.warning("No valid Google token yet. Drive overlay hasn't run.")
            else:
                _sheets, drive, _gmail = services(token, cfg.HTTP_TIMEOUT_SECONDS)
                overrides = load_config_from_drive(drive, CONFIG_FILE_ID or None)
                if overrides:
                    st.success(f"Loaded Drive overrides ({len(overrides)} keys).")
                    # show just the most relevant keys
                    show = {k: overrides.get(k) for k in [
                        "CALC_SPREADSHEET_ID", "INCOMING_FOLDER_ID",
                        "MANAGER_REPORT_FOLDER_ID", "ORDER_REPORT_FOLDER_ID",
                        "TO_RECIPIENTS", "REPORT_KEY_RUN_LIST"
                    ]}
                    st.json(show)
                else:
                    st.error("No Drive config found (or file empty).")
                    st.caption("If you just saved defaults to Drive, copy its file id into Secrets as CONFIG_FILE_ID and rerun.")
        except Exception as e:
            st.error(f"Diagnostics failed: {e}")

    st.caption("**Effective config (current page state):**")
    colA, colB = st.columns(2)
    with colA:
        st.write("**IDs**")
        st.code({
            "CALC_SPREADSHEET_ID": cfg.CALC_SPREADSHEET_ID,
            "INCOMING_FOLDER_ID": cfg.INCOMING_FOLDER_ID,
            "MANAGER_REPORT_FOLDER_ID": cfg.MANAGER_REPORT_FOLDER_ID,
            "ORDER_REPORT_FOLDER_ID": cfg.ORDER_REPORT_FOLDER_ID,
        }, language="json")
        st.write("**Recipients**")
        st.code({
            "TO_RECIPIENTS": cfg.TO_RECIPIENTS,
            "CC_RECIPIENTS": cfg.CC_RECIPIENTS,
            "DEFAULT_ORDER_RECIPIENTS": cfg.DEFAULT_ORDER_RECIPIENTS,
        }, language="json")
    with colB:
        st.write("**Flags & lists**")
        st.code({
            "USE_ALL_REPORT_KEYS": cfg.USE_ALL_REPORT_KEYS,
            "REPORT_KEY_RUN_LIST": cfg.REPORT_KEY_RUN_LIST,
            "REPORT_KEY_RECIPIENTS": cfg.REPORT_KEY_RECIPIENTS,
        }, language="json")
        st.write("**GIDs / TZ**")
        st.code({
            "GID_MANAGER_PDF": cfg.GID_MANAGER_PDF,
            "GID_ORDER_CSV": cfg.GID_ORDER_CSV,
            "LOCATION_SHEET_TITLE": cfg.LOCATION_SHEET_TITLE,
            "LOCATION_NAMED_RANGE": cfg.LOCATION_NAMED_RANGE,
            "TIMESTAMP_TZ": cfg.TIMESTAMP_TZ,
            "TIMESTAMP_FMT": cfg.TIMESTAMP_FMT,
        }, language="json")
    
params = st.query_params  # Streamlit >=1.31; for older use st.experimental_get_query_params()
if "code" in params and "state" in params:
    try:
        finish_web_oauth(params["code"], params["state"], cfg.SCOPES)
        st.success("✅ Google authentication complete.")
        st.query_params.clear()  # avoid reusing on refresh
        st.rerun()
    except Exception as e:
        st.error(f"OAuth error: {e}")
if load_valid_token(cfg.SCOPES):
    cfg = Config.load()

# ----------------------------
# Session state: auth gating
# ----------------------------
if "auth_checked" not in st.session_state:
    # On first load: if token is missing/invalid/unrefreshable -> require auth
    st.session_state.auth_required = (load_valid_token(cfg.SCOPES) is None)
    st.session_state.oauth_flow = None
    st.session_state.oauth_url = None
    st.session_state.auth_checked = True

# Sidebar controls (always visible)
with st.sidebar:
    st.header("Defaults (.env)")

    if st.button("Force Google Re-Auth", type="secondary", use_container_width=True):
        clear_token()
        try:
            # Prefer auto flow
            from favtrip.google_client import login_via_local_server
            with st.status("Re-auth in progress (browser will open)…", expanded=True):
                creds = login_via_local_server(cfg.SCOPES, cfg.REDIRECT_PORT)
                st.success("✅ Re-auth complete.")
            st.session_state.auth_required = False
            st.rerun()
        except Exception as e:
            # Fallback to manual method
            try:
                flow, url = start_oauth(cfg.SCOPES, cfg.REDIRECT_PORT)
                st.session_state.oauth_flow = flow
                st.session_state.oauth_url = url
                st.session_state.auth_required = True  # show the auth panel
                st.info("Auto re-auth failed; showing manual method. Open the URL shown in the Authentication panel.")
                st.rerun()
            except Exception as e2:
                st.error(f"Failed to start re-auth: {e2}")

    # Developer mode (suppresses console prints when unchecked)
    dev_mode = st.checkbox("Run in developer mode", value=False)

    st.markdown("Edit values for this *one run*. Optionally tick **Update .env** to persist.")

# ----------------------------
# Authentication panel (shown only if auth required)
# ----------------------------

if st.session_state.auth_required:
    with st.expander("Google Authentication", expanded=True):
        st.caption(
            "Authentication is required before running. "
            "Click **Sign in with Google** to open the consent screen. "
            "You will be redirected back here automatically."
        )

        # New: web-redirect OAuth flow suitable for Streamlit Cloud
        if st.button("Sign in with Google"):
            try:
                auth_url = start_web_oauth(cfg.SCOPES)
                st.link_button("Open Google Authorization", auth_url, use_container_width=True)
                st.info("After allowing access, you'll be redirected back here automatically.")
            except Exception as e:
                st.error(f"Failed to start OAuth: {e}")

        # (Optional) Helpful hints / debug info
        with st.expander("Having trouble?", expanded=False):
            st.write(
                "- Make sure the Google OAuth **Authorized redirect URI** in Cloud Console "
                "matches your Streamlit app URL exactly (e.g., `https://your-app.streamlit.app`).\n"
                "- If you change the app URL or rename the app, update the redirect URI in Google Cloud."
            )

# ----------------------------
# Only show Run Options if NOT requiring auth
# ----------------------------
if not st.session_state.auth_required:

    # ---- Run Form (Run button top-right) ----
    with st.form("run_form"):
        tl, tr = st.columns([4, 1])
        with tl:
            st.subheader("Run Options")
            st.caption("Configure email behavior and report keys. Use **Advanced** for IDs/GIDs/timezone.")
        with tr:
            submitted = st.form_submit_button("▶️ Run Pipeline", use_container_width=True)

        # --- Main options ---
        colA, colB, colC = st.columns(3)
        with colA:
            to = st.text_input("To Recipients (comma)", value=",".join(cfg.TO_RECIPIENTS or []))
            cc = st.text_input("CC Recipients (comma)", value=",".join(cfg.CC_RECIPIENTS or []))
        with colB:
            use_all = st.checkbox("Use all Report Keys in CSV", value=cfg.USE_ALL_REPORT_KEYS)
            report_keys = st.text_input("Report Keys to run (comma)", value=",".join(cfg.REPORT_KEY_RUN_LIST or []))
        with colC:
            include_full = st.checkbox("Attach FULL order in each email", value=cfg.INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL)
            send_full = st.checkbox("Send separate FULL order email", value=cfg.SEND_SEPARATE_FULL_ORDER_EMAIL)
            #force_reauth = st.checkbox("Force Google re-auth for this run", value=cfg.FORCE_REAUTH)

        # --- Per-report-key recipients editor (above Advanced) ---
        with st.expander("Per-Report-Key Recipients (optional)"):
            st.caption('Map **REPORT KEY (ALL CAPS)** → **Emails (comma)** (friendly editor for `REPORT_KEY_RECIPIENTS`).')
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

        # --- Advanced (IDs, GIDs, Timezone, Redirect Port) ---
        with st.expander("Advanced (IDs, GIDs, Timezone, Redirect Port)"):
            col1, col2 = st.columns(2)
            with col1:
                calc_id = st.text_input("Calculations Spreadsheet ID", value=cfg.CALC_SPREADSHEET_ID)
                incoming_id = st.text_input("Incoming Folder ID", value=cfg.INCOMING_FOLDER_ID)
                mgr_folder = st.text_input("Manager Report Folder ID", value=cfg.MANAGER_REPORT_FOLDER_ID)
                order_folder = st.text_input("Order Report Folder ID", value=cfg.ORDER_REPORT_FOLDER_ID)
                
                raw_redirect_port = int(cfg.REDIRECT_PORT) if str(cfg.REDIRECT_PORT).isdigit() else 0
                redirect_port = st.number_input(
                    "Redirect Port (0 = auto)",
                    min_value=0,             # allow 0 explicitly
                    max_value=65535,
                    value=raw_redirect_port if raw_redirect_port in (0, *range(1024, 65536)) else 0,
                    help="Use 0 to auto-pick a free port. Otherwise choose 1024–65535.",
                )

            with col2:
                gid_mgr = st.text_input("Manager Report gid", value=str(cfg.GID_MANAGER_PDF))
                gid_order = st.text_input("Order CSV gid", value=str(cfg.GID_ORDER_CSV))
                loc_sheet = st.text_input("Location Sheet Title", value=cfg.LOCATION_SHEET_TITLE)
                loc_range = st.text_input("Location Named Range", value=cfg.LOCATION_NAMED_RANGE)
                tz = st.text_input("Timestamp Timezone", value=cfg.TIMESTAMP_TZ)
                tfmt = st.text_input("Timestamp Format", value=cfg.TIMESTAMP_FMT)

        save_drive_defaults = st.checkbox("Update defaults", value=False)

    # ----------------------------
    # Submission handling
    # ----------------------------
    def _split_emails(csv_str: str):
        return [e.strip() for e in (csv_str or "").split(",") if e.strip()]

    if submitted:
        # Apply per-run config
        cfg.TO_RECIPIENTS = _split_emails(to)
        cfg.CC_RECIPIENTS = _split_emails(cc)
        cfg.USE_ALL_REPORT_KEYS = use_all
        cfg.REPORT_KEY_RUN_LIST = [s.strip().upper() for s in (report_keys or "").split(",") if s.strip()]

        cfg.INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL = include_full
        cfg.SEND_SEPARATE_FULL_ORDER_EMAIL = send_full

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

        # Per-key recipients from editor
        rk_map = {}
        for r in edited_rows:
            key = (r.get("REPORT KEY (ALL CAPS)") or "").strip().upper()
            emails_csv = r.get("Emails (comma)") or ""
            emails = _split_emails(emails_csv)
            if key and emails:
                rk_map[key] = emails
        cfg.REPORT_KEY_RECIPIENTS = rk_map

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

                    # What we persist (the fields you asked to move out of Secrets)
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

                        "TO_RECIPIENTS": cfg.TO_RECIPIENTS,   # lists are fine; JSON keeps types
                        "CC_RECIPIENTS": cfg.CC_RECIPIENTS,

                        "USE_ALL_REPORT_KEYS": cfg.USE_ALL_REPORT_KEYS,
                        "REPORT_KEY_RUN_LIST": cfg.REPORT_KEY_RUN_LIST,

                        "REPORT_KEY_RECIPIENTS": cfg.REPORT_KEY_RECIPIENTS,

                        "DEFAULT_ORDER_RECIPIENTS": cfg.DEFAULT_ORDER_RECIPIENTS,

                        "INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL": cfg.INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL,
                        "SEND_SEPARATE_FULL_ORDER_EMAIL": cfg.SEND_SEPARATE_FULL_ORDER_EMAIL,
                    }

                    # If you have CONFIG_FILE_ID in Secrets, we update that exact file.
                    # Otherwise we'll upsert a file named 'favtrip_config.json' and return its id.
                    CONFIG_FILE_ID = (st.secrets.get("CONFIG_FILE_ID", "") or "").strip()
                    new_id = save_config_to_drive(
                        drive,
                        drive_defaults,
                        file_id=CONFIG_FILE_ID or None,   # update/pin if set, else upsert by name
                        # parent_folder_id=None,          # optional: set a folder id to create under
                    )

                    st.success(f"Saved defaults to Drive config (file id: {new_id}).")
                    if not CONFIG_FILE_ID:
                        st.info(
                            "Tip: add this ID to Streamlit Secrets as `CONFIG_FILE_ID` to pin the same file for all runs:\n"
                            f"`{new_id}`"
                        )
            except Exception as e:
                st.error(f"Failed to save defaults to Drive: {e}")

        # If user checked "Force Google re-auth for this run", kick them into auth gating first.
        if cfg.FORCE_REAUTH:
            clear_token()
            try:
                flow, url = start_oauth(cfg.SCOPES, cfg.REDIRECT_PORT)
                st.session_state.oauth_flow = flow
                st.session_state.oauth_url = url
                st.session_state.auth_required = True
                st.info("Re-auth required for this run. Open the URL shown in the Authentication panel.")
                _rerun()
            except Exception as e:
                st.error(f"Failed to start OAuth: {e}")
        else:
            # --- Live run with timer + last log (no full log after completion) ---
            # Write all logs to last_run.log; overwrite on each run
            logger = StatusLogger(print_to_console=dev_mode, file_path="last_run.log", overwrite=True)
            result_holder = {"value": None, "error": None}

            def _runner():
                try:
                    result_holder["value"] = run_pipeline(cfg, logger=logger)
                # Catch BaseException so SystemExit / KeyboardInterrupt are captured too
                except BaseException as e:
                    result_holder["error"] = e

            t0 = time.perf_counter()
            th = threading.Thread(target=_runner, daemon=True)
            th.start()

            with st.status("Running pipeline…", expanded=True) as status:
                timer_ph = st.empty()
                lastlog_ph = st.empty()
                while th.is_alive():
                    elapsed = int(time.perf_counter() - t0)
                    timer_ph.markdown(f"**Elapsed:** `{elapsed//3600:02d}:{(elapsed%3600)//60:02d}:{elapsed%60:02d}`")
                    lastlog_ph.markdown(f"**Last:** {logger.last_line()}")
                    time.sleep(0.5)

                th.join()
                elapsed = int(time.perf_counter() - t0)
                timer_ph.markdown(f"**Elapsed:** `{elapsed//3600:02d}:{(elapsed%3600)//60:02d}:{elapsed%60:02d}`")
                lastlog_ph.markdown(f"**Last:** {logger.last_line()}")

                if result_holder["error"]:
                    st.error(f"Run failed: {result_holder['error']}")
                    # Optional during debugging: show stack trace (remove later for a cleaner UI)
                    try:
                        st.exception(result_holder["error"])
                    except Exception:
                        pass
                    status.update(label="❌ Failed", state="error")
                else:
                    result = result_holder["value"]
                    if result is None:
                        st.error("Run finished without returning a result. Check logs and inputs (IDs, Drive access).")
                        status.update(label="⚠️ No result", state="error")
                    else:
                        st.write("### Outputs")
                        col1, col2, col3 = st.columns(3)
                        col1.metric("Location", result.location)
                        col2.metric("Timestamp", result.timestamp)
                        mm = result.elapsed_seconds
                        col3.metric("Elapsed", f"{mm//3600:02d}:{(mm%3600)//60:02d}:{mm%60:02d}")

                        if getattr(result, "manager_pdf_link", None):
                            st.success(f"Manager PDF: {result.manager_pdf_link}")
                        if getattr(result, "full_order_link", None):
                            st.success(f"Full Order Sheet: {result.full_order_link}")

                        status.update(label="✅ Completed", state="complete")