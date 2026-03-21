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
