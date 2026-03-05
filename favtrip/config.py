from __future__ import annotations

import os
import json
from dataclasses import dataclass, asdict
from typing import Dict, List, Optional
from pathlib import Path
from dotenv import load_dotenv
from typing import Any, Dict

# -----------------------------------------------------------------------------
# Helpers: read from Streamlit secrets (typed) or .env (strings) and coerce
# -----------------------------------------------------------------------------

def _get_secret(key: str, default: str = ""):
    """
    Prefer Streamlit secrets (typed TOML values), else fall back to environment.
    Returns the raw value from st.secrets (could be bool/list/dict/str) or a str from env.
    """
    try:
        import streamlit as st  # imported lazily so local CLI still works
        val = st.secrets.get(key, None)
        if val is None:
            return os.getenv(key, default)
        return val
    except Exception:
        return os.getenv(key, default)


_TRUE = {"1", "true", "yes", "on", "y", "t"}


def _coerce_bool(v, default: bool = False) -> bool:
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


def _coerce_csv(v) -> List[str]:
    """
    Accept list/tuple (already structured) or a comma-separated string.
    Returns a list of trimmed strings.
    """
    if v is None or v == "":
        return []
    if isinstance(v, (list, tuple)):
        return [str(x).strip() for x in v if str(x).strip()]
    return [p.strip() for p in str(v).split(",") if p.strip()]


def _coerce_json(v):
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
    CALC_SPREADSHEET_ID: str
    INCOMING_FOLDER_ID: str
    MANAGER_REPORT_FOLDER_ID: str
    ORDER_REPORT_FOLDER_ID: str

    # GIDs, sheet metadata, timestamp settings
    GID_MANAGER_PDF: str = "1921812573"
    GID_ORDER_CSV: str = "1875928148"
    LOCATION_SHEET_TITLE: str = "REFR: Values"
    LOCATION_NAMED_RANGE: str = "_locations"
    TIMESTAMP_TZ: str = "America/Chicago"
    TIMESTAMP_FMT: str = "%Y-%m-%d-%I-%M-%p"

    # Email config
    TO_RECIPIENTS: List[str] = None
    CC_RECIPIENTS: List[str] = None
    USE_ALL_REPORT_KEYS: bool = False
    REPORT_KEY_RUN_LIST: List[str] = None
    REPORT_KEY_RECIPIENTS: Dict[str, List[str]] = None
    DEFAULT_ORDER_RECIPIENTS: List[str] = None
    INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL: bool = False
    SEND_SEPARATE_FULL_ORDER_EMAIL: bool = True

    # Google API
    SCOPES: List[str] = None
    FORCE_REAUTH: bool = False
    REDIRECT_PORT: int = 58285
    HTTP_TIMEOUT_SECONDS: int = 300


import os
import json
from dataclasses import dataclass, asdict
from typing import Dict, List, Optional
from pathlib import Path
from dotenv import load_dotenv
from typing import Any, Dict

# -----------------------------------------------------------------------------
# Helpers: read from Streamlit secrets (typed) or .env (strings) and coerce
# -----------------------------------------------------------------------------

def _get_secret(key: str, default: str = ""):
    """
    Prefer Streamlit secrets (typed TOML values), else fall back to environment.
    Returns the raw value from st.secrets (could be bool/list/dict/str) or a str from env.
    """
    try:
        import streamlit as st  # imported lazily so local CLI still works
        val = st.secrets.get(key, None)
        if val is None:
            return os.getenv(key, default)
        return val
    except Exception:
        return os.getenv(key, default)


_TRUE = {"1", "true", "yes", "on", "y", "t"}


def _coerce_bool(v, default: bool = False) -> bool:
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


def _coerce_csv(v) -> List[str]:
    """
    Accept list/tuple (already structured) or a comma-separated string.
    Returns a list of trimmed strings.
    """
    if v is None or v == "":
        return []
    if isinstance(v, (list, tuple)):
        return [str(x).strip() for x in v if str(x).strip()]
    return [p.strip() for p in str(v).split(",") if p.strip()]


def _coerce_json(v):
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
    CALC_SPREADSHEET_ID: str
    INCOMING_FOLDER_ID: str
    MANAGER_REPORT_FOLDER_ID: str
    ORDER_REPORT_FOLDER_ID: str

    # GIDs, sheet metadata, timestamp settings
    GID_MANAGER_PDF: str = "1921812573"
    GID_ORDER_CSV: str = "1875928148"
    LOCATION_SHEET_TITLE: str = "REFR: Values"
    LOCATION_NAMED_RANGE: str = "_locations"
    TIMESTAMP_TZ: str = "America/Chicago"
    TIMESTAMP_FMT: str = "%Y-%m-%d-%I-%M-%p"

    # Email config
    TO_RECIPIENTS: List[str] = None
    CC_RECIPIENTS: List[str] = None
    USE_ALL_REPORT_KEYS: bool = False
    REPORT_KEY_RUN_LIST: List[str] = None
    REPORT_KEY_RECIPIENTS: Dict[str, List[str]] = None
    DEFAULT_ORDER_RECIPIENTS: List[str] = None
    INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL: bool = False
    SEND_SEPARATE_FULL_ORDER_EMAIL: bool = True

    # Google API
    SCOPES: List[str] = None
    FORCE_REAUTH: bool = False
    REDIRECT_PORT: int = 58285
    HTTP_TIMEOUT_SECONDS: int = 300

    @staticmethod
    def load(env_path: Optional[Path] = None) -> "Config":
        """
        Load config from Streamlit secrets (preferred on cloud) or from .env (local dev),
        then overlay any values found in a Drive-backed JSON config (optional).
        Secrets may be typed (bool/list/dict), so we coerce safely.
        """
        if env_path is None:
            env_path = Path.cwd() / ".env"
        load_dotenv(dotenv_path=env_path, override=False)

        # ---------- 1) Base config from Secrets / .env ----------
        cfg = Config(
            CALC_SPREADSHEET_ID=str(_get_secret("CALC_SPREADSHEET_ID", "")),
            INCOMING_FOLDER_ID=str(_get_secret("INCOMING_FOLDER_ID", "")),
            MANAGER_REPORT_FOLDER_ID=str(_get_secret("MANAGER_REPORT_FOLDER_ID", "")),
            ORDER_REPORT_FOLDER_ID=str(_get_secret("ORDER_REPORT_FOLDER_ID", "")),

            GID_MANAGER_PDF=str(_get_secret("GID_MANAGER_PDF", "1921812573")),
            GID_ORDER_CSV=str(_get_secret("GID_ORDER_CSV", "1875928148")),
            LOCATION_SHEET_TITLE=str(_get_secret("LOCATION_SHEET_TITLE", "REFR: Values")),
            LOCATION_NAMED_RANGE=str(_get_secret("LOCATION_NAMED_RANGE", "_locations")),
            TIMESTAMP_TZ=str(_get_secret("TIMESTAMP_TZ", "America/Chicago")),
            TIMESTAMP_FMT=str(_get_secret("TIMESTAMP_FMT", "%Y-%m-%d-%I-%M-%p")),

            TO_RECIPIENTS=_coerce_csv(_get_secret("TO_RECIPIENTS", "")),
            CC_RECIPIENTS=_coerce_csv(_get_secret("CC_RECIPIENTS", "")),
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
        )

        # ---------- 2) Overlay from Drive JSON config (optional) ----------
        # If present, read a JSON file on Drive and apply known keys over cfg.
        try:
            import streamlit as st
            from favtrip.google_client import load_valid_token, services
            from favtrip.config_store import load_config_from_drive

            CONFIG_FILE_ID = ""
            if hasattr(st, "secrets"):
                CONFIG_FILE_ID = (st.secrets.get("CONFIG_FILE_ID", "") or "").strip()

            creds = load_valid_token(cfg.SCOPES)
            if creds:
                _sheets, drive, _gmail = services(creds, cfg.HTTP_TIMEOUT_SECONDS)
                overrides = load_config_from_drive(drive, CONFIG_FILE_ID or None)  # {} if not found
                if isinstance(overrides, dict) and overrides:
                    for k, v in overrides.items():
                        if hasattr(cfg, k):
                            setattr(cfg, k, v)
        except Exception:
            # Fail-open: if Drive/token not ready yet, just return base cfg
            pass

        return cfg

    # -----------------------------------------------------------------------------
    # .env serialization (still useful if you let users "Update defaults in .env")
    # -----------------------------------------------------------------------------

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

    # -----------------------------------------------------------------------------
    # .env serialization (still useful if you let users "Update defaults in .env")
    # -----------------------------------------------------------------------------

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