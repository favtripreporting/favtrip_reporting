from __future__ import annotations
import os
import json
from dataclasses import dataclass, asdict
from typing import Dict, List, Optional
from pathlib import Path
from dotenv import load_dotenv

_BOOL_TRUE = {"1", "true", "yes", "on", "y", "t"}


def _get_secret(key, default=""):
    try:
        import streamlit as st
        val = st.secrets.get(key, None)
        if val is None:
            import os
            return os.getenv(key, default)
        return val
    except Exception:
        import os
        return os.getenv(key, default)



def _b(s: str, default: bool = False) -> bool:
    if s is None:
        return default
    return str(s).strip().lower() in _BOOL_TRUE


def _csv(s: str) -> List[str]:
    if not s:
        return []
    return [p.strip() for p in s.split(",") if p.strip()]


def _json_or_empty(s: str):
    if not s:
        return {}
    try:
        return json.loads(s)
    except Exception:
        return {}


@dataclass
class Config:
    # IDs and basic settings
    CALC_SPREADSHEET_ID: str
    INCOMING_FOLDER_ID: str
    MANAGER_REPORT_FOLDER_ID: str
    ORDER_REPORT_FOLDER_ID: str

    GID_MANAGER_PDF: str = "1921812573"
    GID_ORDER_CSV: str = "1875928148"

    LOCATION_SHEET_TITLE: str = "REFR: Values"
    LOCATION_NAMED_RANGE: str = "_locations"

    TIMESTAMP_TZ: str = "America/Chicago"
    TIMESTAMP_FMT: str = "%Y-%m-%d-%I-%M-%p"

    TO_RECIPIENTS: List[str] = None
    CC_RECIPIENTS: List[str] = None

    USE_ALL_REPORT_KEYS: bool = False
    REPORT_KEY_RUN_LIST: List[str] = None
    REPORT_KEY_RECIPIENTS: Dict[str, List[str]] = None
    DEFAULT_ORDER_RECIPIENTS: List[str] = None

    INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL: bool = False
    SEND_SEPARATE_FULL_ORDER_EMAIL: bool = True

    SCOPES: List[str] = None
    FORCE_REAUTH: bool = False
    REDIRECT_PORT: int = 58285

    HTTP_TIMEOUT_SECONDS: int = 300

    @staticmethod
    def load(env_path: Optional[Path] = None) -> "Config":
        if env_path is None:
            env_path = Path.cwd() / ".env"
        load_dotenv(dotenv_path=env_path, override=False)
        
        return Config(
            CALC_SPREADSHEET_ID=_get_secret("CALC_SPREADSHEET_ID", ""),
            INCOMING_FOLDER_ID=_get_secret("INCOMING_FOLDER_ID", ""),
            MANAGER_REPORT_FOLDER_ID=_get_secret("MANAGER_REPORT_FOLDER_ID", ""),
            ORDER_REPORT_FOLDER_ID=_get_secret("ORDER_REPORT_FOLDER_ID", ""),
            GID_MANAGER_PDF=_get_secret("GID_MANAGER_PDF", "1921812573"),
            GID_ORDER_CSV=_get_secret("GID_ORDER_CSV", "1875928148"),
            LOCATION_SHEET_TITLE=_get_secret("LOCATION_SHEET_TITLE", "REFR: Values"),
            LOCATION_NAMED_RANGE=_get_secret("LOCATION_NAMED_RANGE", "_locations"),
            TIMESTAMP_TZ=_get_secret("TIMESTAMP_TZ", "America/Chicago"),
            TIMESTAMP_FMT=_get_secret("TIMESTAMP_FMT", "%Y-%m-%d-%I-%M-%p"),
            TO_RECIPIENTS=[s.strip() for s in _get_secret("TO_RECIPIENTS","").split(",") if s.strip()],
            CC_RECIPIENTS=[s.strip() for s in _get_secret("CC_RECIPIENTS","").split(",") if s.strip()],
            USE_ALL_REPORT_KEYS=_get_secret("USE_ALL_REPORT_KEYS","false").lower() in {"1","true","yes","on","y","t"},
            REPORT_KEY_RUN_LIST=[s.strip().upper() for s in _get_secret("REPORT_KEY_RUN_LIST","").split(",") if s.strip()],
            REPORT_KEY_RECIPIENTS=__import__("json").loads(_get_secret("REPORT_KEY_RECIPIENTS","{}") or "{}"),
            DEFAULT_ORDER_RECIPIENTS=[s.strip() for s in _get_secret("DEFAULT_ORDER_RECIPIENTS","").split(",") if s.strip()],
            INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL=_get_secret("INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL","false").lower() in {"1","true","yes","on","y","t"},
            SEND_SEPARATE_FULL_ORDER_EMAIL=_get_secret("SEND_SEPARATE_FULL_ORDER_EMAIL","true").lower() in {"1","true","yes","on","y","t"},
            SCOPES=[s.strip() for s in _get_secret("SCOPES","").split(",") if s.strip()],
            FORCE_REAUTH=False,
            REDIRECT_PORT=int(_get_secret("REDIRECT_PORT","0")),
            HTTP_TIMEOUT_SECONDS=int(_get_secret("HTTP_TIMEOUT_SECONDS","300")),
        )


    def to_env(self) -> str:
        # Serialize to .env format (simple)
        data = asdict(self)
        # Convert list and dict fields
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