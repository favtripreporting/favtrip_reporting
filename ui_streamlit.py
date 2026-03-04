import os
import time
import threading
import streamlit as st


import json
from google_auth_oauthlib.flow import Flow
from google.oauth2.credentials import Credentials



def app_base_url():
    # Works on Streamlit Cloud; returns https://<your-app>.streamlit.app
    try:
        import streamlit as st
        # newer Streamlit:
        base = st.request.url_root.rstrip("/")
    except Exception:
        # fallback; you can hardcode if needed
        base = ""
    return base

def start_web_oauth(scopes):
    import streamlit as st
    cfg = json.loads(st.secrets["GOOGLE_CREDENTIALS"])  # the full JSON pasted in Secrets
    flow = Flow.from_client_config(cfg, scopes=scopes, redirect_uri=app_base_url())
    auth_url, state = flow.authorization_url(
        access_type="offline",
        include_granted_scopes="true",
        prompt="consent",
    )
    st.session_state["_oauth_state"] = state
    st.session_state["_oauth_cfg"] = cfg
    return auth_url

def finish_web_oauth(code, state, scopes):
    import streamlit as st
    cfg = st.session_state.get("_oauth_cfg") or json.loads(st.secrets["GOOGLE_CREDENTIALS"])
    flow = Flow.from_client_config(cfg, scopes=scopes, redirect_uri=app_base_url())
    flow.fetch_token(code=code)
    creds = flow.credentials
    with open("token.json", "w") as f:
        f.write(creds.to_json())
    return creds


from favtrip.config import Config
from favtrip.logger import StatusLogger
from favtrip.pipeline import run_pipeline
from favtrip.google_client import (
    start_oauth,
    finish_oauth,
    load_valid_token,
    clear_token,
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

params = st.query_params  # Streamlit >=1.31; for older use st.experimental_get_query_params()
if "code" in params and "state" in params:
    try:
        finish_web_oauth(params["code"], params["state"], cfg.SCOPES)
        st.success("✅ Google authentication complete.")
        st.query_params.clear()  # avoid reusing on refresh
        st.rerun()
    except Exception as e:
        st.error(f"OAuth error: {e}")


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

        save_env = st.checkbox("Update defaults in .env with the edited fields (optional)")

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

        if save_env:
            cfg.save()
            st.success("Saved updated defaults to .env")

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
                except Exception as e:
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
                    status.update(label="❌ Failed", state="error")
                else:
                    result = result_holder["value"]
                    st.write("### Outputs")
                    col1, col2, col3 = st.columns(3)
                    col1.metric("Location", result.location)
                    col2.metric("Timestamp", result.timestamp)
                    mm = result.elapsed_seconds
                    col3.metric("Elapsed", f"{mm//3600:02d}:{(mm%3600)//60:02d}:{mm%60:02d}")
                    if result.manager_pdf_link:
                        st.success(f"Manager PDF: {result.manager_pdf_link}")
                    if result.full_order_link:
                        st.success(f"Full Order Sheet: {result.full_order_link}")
                    # Note: intentionally NOT showing the full live log post-run
                    status.update(label="✅ Completed", state="complete")