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

from streamlit.components.v1 import html
from streamlit.components.v1 import html as _html_msg



import re

EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")

def _parse_emails(csv_str: str):
    return [e.strip() for e in (csv_str or "").split(",") if e.strip()]

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
        st.rerun()
    except AttributeError:
        st.experimental_rerun()

st.set_page_config(page_title="FavTrip Reporting Pipeline", page_icon="🧾", layout="wide")


# --- Larger Run button ---
# --- REPLACE existing CSS block with this compact style ---
# --- ADD / MERGE: Big, right-aligned "Run Pipeline" button styles ---
st.markdown(
    """
    <style>
      /* Ensure the submit container in our header aligns to the right */
      .ft-runwrap {
        display: flex;
        justify-content: flex-end;
        align-items: center;
      }

      /* Style ONLY the form submit button (safe selector Streamlit uses for submit) */
      div[data-testid="stFormSubmitButton"] button {
        font-size: 1.15rem;      /* bigger text */
        font-weight: 600;         /* bolder text */
        padding: 0.85rem 1.35rem; /* bigger hit area */
        border-radius: 10px;      /* pill-ish corners */
        box-shadow: 0 1px 2px rgba(0,0,0,0.08);
      }

      /* Optional: make the button more prominent with a stronger accent color */
      div[data-testid="stFormSubmitButton"] button[kind="primary"] {
        background: var(--primary-color);
        color: white;
      }

      /* On wider screens, keep the submit container tight to the right edge */
      @media (min-width: 900px) {
        div[data-testid="stFormSubmitButton"] {
          width: 100%;
          display: flex;
          justify-content: flex-end;
        }
      }

      /* Optional hover/active states */
      div[data-testid="stFormSubmitButton"] button:hover {
        filter: brightness(0.96);
      }
      div[data-testid="stFormSubmitButton"] button:active {
        transform: translateY(0.5px);
      }
    </style>
    """,
    unsafe_allow_html=True
)
# --- END ADD / MERGE ---

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
    
params = st.query_params
if "code" in params and "state" in params:
    try:
        finish_web_oauth(params["code"], params["state"], cfg.SCOPES)
        st.success("✅ Google authentication complete.")
        st.query_params.clear()
        # NEW: tell the opener to refresh and then close the popup
        html(
            """
            <script>
              try {
                if (window.opener && !window.opener.closed) {
                  window.opener.postMessage({type: "favtrip_oauth_done"}, "*");
                }
              } catch (e) {}
              // Close this popup
              window.close();
            </script>
            """,
            height=0,
        )
        # Fallback: in case the browser blocks close(), allow manual continue
        st.caption("You can close this window if it didn't close automatically.")
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


if "offer_log_download" not in st.session_state:
    st.session_state.offer_log_download = False


# Sidebar controls (always visible)
with st.sidebar:
    st.header("Utilities")

    if st.button("Google Sign Out", type="secondary", use_container_width=True):
        clear_token()
        for key in ["auth_required", "oauth_flow", "oauth_url", "auth_checked"]:
            if key in st.session_state:
                del st.session_state[key]
        _rerun()
        
    st.checkbox(
        "Offer full log download after completion",
        key="offer_log_download",
        help="If enabled, a 'Download last_run.log' button appears when a run finishes.")


# ----------------------------
# Authentication panel (shown only if auth required)
# ----------------------------

# --- Authentication expander (shown only if auth required) ---
if st.session_state.auth_required:
    with st.expander("Google Authentication", expanded=True):
        st.caption(
            "Authentication is required before running. "
            "Click **Sign in with Google** to open the consent screen (it will open in a new tab)."
        )

        # Use a placeholder so we can remove the button immediately
        sign_in_ph = st.empty()

        # Render the button inside the placeholder
        clicked = sign_in_ph.button("Sign in with Google", type="primary", use_container_width=True)

        if clicked:
            try:
                auth_url = start_web_oauth(cfg.SCOPES)

                # 1) Clear the placeholder to remove the button immediately
                sign_in_ph.empty()

                # 2) Show the message in the current tab
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
                          <strong>After it completes, you may close this tab.</strong>
                        </p>
                      </div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

                # (Optional) If user returns to this tab later, refresh to show signed-in state
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

                # 3) Open Google auth in a NEW tab (leave this tab on the message)
                html(
                    f"""
                    <script>
                      window.open({json.dumps(auth_url)}, "_blank", "noopener");
                    </script>
                    """,
                    height=0,
                )

                # 4) Stop further rendering so the message persists
                st.stop()

            except Exception as e:
                st.error(f"Failed to start OAuth: {e}")

        with st.expander("Having trouble?", expanded=False):
            st.write(
                "- The Google authorization page opens in a **new browser tab**.\n"
                "- After completing consent, **close this tab** and use the new tab.\n"
                "- If you renamed your Streamlit app or URL, ensure the Google OAuth "
                "Authorized redirect URI matches exactly (including trailing slash)."
            )

# ----------------------------
# Only show Run Options if NOT requiring auth
# ----------------------------
if not st.session_state.auth_required:

    # ---- Run Form (Run button top-right) ----
    with st.form("run_form"):
        
        st.markdown('<div class="ft-form-header">', unsafe_allow_html=True)
        st.markdown('<div class="ft-form-header-left">', unsafe_allow_html=True)
        

        
        hdr_left, hdr_right = st.columns([1, 0.36])

        with hdr_left:
            st.subheader("Run Options")
            st.caption("Configure email behavior and report keys. Use **Advanced** for IDs/GIDs/timezone.")

        with hdr_right:
            # A small wrapper so the button is visually anchored to the right edge
            st.markdown('<div class="ft-runwrap">', unsafe_allow_html=True)
            submitted = st.form_submit_button(
                "▶️ Run Pipeline",
                help="Start the pipeline with the options configured in this form."
            )
            st.markdown('</div>', unsafe_allow_html=True)


        # --- Main options ---
        # --- REPLACE your current recipients/keys/email-behavior sections with this ---

        # ===== Recipients (compact two columns) =====
        with st.container():
            st.markdown("##### Recipients")
            with st.container():
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

        # ===== Report Keys (toggle + input aligned) =====
        with st.container():
            st.markdown("##### Report Keys")
            with st.container():
                colk1, colk2 = st.columns([0.45, 1.55])
                with colk1:
                    # Toggle reads better than checkbox for 'all vs selected'
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

        # ===== Email Behavior (two toggles in one row) =====
        with st.container():
            st.markdown("##### Email Behavior")
            cole1, cole2 = st.columns([1, 1])
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

        # ===== Put optional editors/switches in cards for light framing =====
        st.markdown('<div class="ft-card">', unsafe_allow_html=True)
        with st.expander("Per‑Report‑Key Recipients (optional)", expanded=False):
            st.caption("Map **REPORT KEY (ALL CAPS)** → **Emails (comma)**.")
            # existing rows construction remains the same...
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

            # Preview + validation (from earlier step)
            rk_issues, rk_preview, rk_map_preview = _analyze_rk_rows(edited_rows)
            if rk_preview:
                with st.expander("Row template preview"):
                    st.code("\n".join(rk_preview), language="text")
            if rk_issues:
                st.warning("Per‑report‑key recipient issues:\n\n- " + "\n- ".join(rk_issues))

        st.markdown('</div>', unsafe_allow_html=True)

        st.markdown('<div class="ft-card">', unsafe_allow_html=True)
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
        st.markdown('</div>', unsafe_allow_html=True)

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
            key = (r.get("REPORT KEY (ALL CAPS)") or "").strip()
            emails_csv = r.get("Emails (comma)") or ""
            emails = _parse_emails(emails_csv)
            if key and emails:
                rk_map[key] = emails
        cfg.REPORT_KEY_RECIPIENTS = rk_map

        # --- ADD: warnings before kicking off the run ---
        # 1) Warn if use_all is OFF and no explicit keys provided
        requested_keys = [s.strip().upper() for s in (report_keys or "").split(",") if s.strip()]
        if not use_all and not requested_keys:
            st.warning(
                "You left **Use all Report Keys** OFF but provided **no keys** to run. "
                "No per‑key outputs will be generated unless you add keys.",
                icon="⚠️"
            )

        # 2) Warn if no recipients anywhere (TO, DEFAULT_ORDER, or per‑key map)
        any_to = bool(_parse_emails(to) or (cfg.TO_RECIPIENTS or []))
        any_default = bool(cfg.DEFAULT_ORDER_RECIPIENTS or [])
        any_per_key = bool(cfg.REPORT_KEY_RECIPIENTS)
        if not (any_to or any_default or any_per_key):
            st.warning(
                "No recipients are defined: **TO**, **DEFAULT_ORDER_RECIPIENTS**, and **Per‑Report‑Key** are all empty. "
                "Emails will not be sent.",
                icon="⚠️"
            )

        # 3) Surface per-key table issues detected earlier
        if rk_issues:
            st.warning(
                "Per‑report‑key recipient issues detected above. These may prevent emails from sending correctly:\n\n- "
                + "\n- ".join(rk_issues),
                icon="⚠️"
            )
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
            logger = StatusLogger(print_to_console=True, file_path="last_run.log", overwrite=True)
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

                            
                        if st.session_state.offer_log_download and os.path.exists("last_run.log"):
                            with open("last_run.log", "rb") as f:
                                st.download_button(
                                    "⬇️ Download full log (last_run.log)",
                                    f,
                                    file_name=f"last_run_{result.timestamp}.log",
                                    mime="text/plain",
                                    use_container_width=True
                                )
                            

                        status.update(label="✅ Completed", state="complete")