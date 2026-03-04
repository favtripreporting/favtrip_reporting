from __future__ import annotations
import os
import json
from dataclasses import dataclass, asdict
from typing import Dict, List, Optional
from pathlib import Path
from dotenv import load_dotenv

_BOOL_TRUE = {"1", "true", "yes", "on", "y", "t"}


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
            CALC_SPREADSHEET_ID=os.getenv("CALC_SPREADSHEET_ID", ""),
            INCOMING_FOLDER_ID=os.getenv("INCOMING_FOLDER_ID", ""),
            MANAGER_REPORT_FOLDER_ID=os.getenv("MANAGER_REPORT_FOLDER_ID", ""),
            ORDER_REPORT_FOLDER_ID=os.getenv("ORDER_REPORT_FOLDER_ID", ""),
            GID_MANAGER_PDF=os.getenv("GID_MANAGER_PDF", "1921812573"),
            GID_ORDER_CSV=os.getenv("GID_ORDER_CSV", "1875928148"),
            LOCATION_SHEET_TITLE=os.getenv("LOCATION_SHEET_TITLE", "REFR: Values"),
            LOCATION_NAMED_RANGE=os.getenv("LOCATION_NAMED_RANGE", "_locations"),
            TIMESTAMP_TZ=os.getenv("TIMESTAMP_TZ", "America/Chicago"),
            TIMESTAMP_FMT=os.getenv("TIMESTAMP_FMT", "%Y-%m-%d-%I-%M-%p"),
            TO_RECIPIENTS=_csv(os.getenv("TO_RECIPIENTS", "")),
            CC_RECIPIENTS=_csv(os.getenv("CC_RECIPIENTS", "")),
            USE_ALL_REPORT_KEYS=_b(os.getenv("USE_ALL_REPORT_KEYS", "false")),
            REPORT_KEY_RUN_LIST=_csv(os.getenv("REPORT_KEY_RUN_LIST", "")),
            REPORT_KEY_RECIPIENTS=_json_or_empty(os.getenv("REPORT_KEY_RECIPIENTS", "")),
            DEFAULT_ORDER_RECIPIENTS=_csv(os.getenv("DEFAULT_ORDER_RECIPIENTS", "")),
            INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL=_b(
                os.getenv("INCLUDE_FULL_ORDER_IN_EACH_REPORT_KEY_EMAIL", "false")
            ),
            SEND_SEPARATE_FULL_ORDER_EMAIL=_b(
                os.getenv("SEND_SEPARATE_FULL_ORDER_EMAIL", "true")
            ),
            SCOPES=_csv(
                os.getenv(
                    "SCOPES",
                    "https://www.googleapis.com/auth/drive,https://www.googleapis.com/auth/spreadsheets,https://www.googleapis.com/auth/gmail.send",
                )
            ),
            FORCE_REAUTH=_b(os.getenv("FORCE_REAUTH", "false")),
            REDIRECT_PORT=int(os.getenv("REDIRECT_PORT", "58285")),
            HTTP_TIMEOUT_SECONDS=int(os.getenv("HTTP_TIMEOUT_SECONDS", "300")),
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