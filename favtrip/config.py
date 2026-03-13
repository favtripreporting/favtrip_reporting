from __future__ import annotations

import os
import json
from dataclasses import dataclass, asdict
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
    USER_FOLDER_ID: str = "1JBHBcnS6397ka2ITW6Wbuu2aKjbgCCHj"

    # GIDs, sheet metadata, timestamp settings
    GID_MANAGER_PDF: str = "1921812573"
    GID_ORDER_CSV: str = "1875928148"
    LOCATION_SHEET_TITLE: str = "REFR: Values"
    LOCATION_NAMED_RANGE: str = "_locations"
    TIMESTAMP_TZ: str = "America/Chicago"
    TIMESTAMP_FMT: str = "%Y-%m-%d-%I-%M-%p"
    TEMPLATE_UPDATE_RANGE: str = "_update"

    # Email config
    TO_RECIPIENTS: List[str] = ["FavtripReporting@gmail.com"]
    CC_RECIPIENTS: List[str] = None
    USE_ALL_REPORT_KEYS: bool = False
    REPORT_KEY_RUN_LIST: List[str] = ["COFFEE"]
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
            USER_FOLDER_ID=str(_get_secret("USER_FOLDER_ID", "")),

            GID_MANAGER_PDF=str(_get_secret("GID_MANAGER_PDF", "1921812573")),
            GID_ORDER_CSV=str(_get_secret("GID_ORDER_CSV", "1875928148")),
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
        )

        # Optional overlay from Drive JSON config (if creds + file present)
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
