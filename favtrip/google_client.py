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