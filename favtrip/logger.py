from dataclasses import dataclass, field
from datetime import datetime
from typing import List, Optional

@dataclass
class LogEvent:
    ts: datetime
    level: str
    message: str

@dataclass
class StatusLogger:
    events: List[LogEvent] = field(default_factory=list)
    print_to_console: bool = True
    file_path: Optional[str] = "last_run.log"
    overwrite: bool = True

    def __post_init__(self):
        # Prepare the file on first use
        self._fh = None
        if self.file_path:
            mode = "w" if self.overwrite else "a"
            try:
                self._fh = open(self.file_path, mode, encoding="utf-8", buffering=1)  # line-buffered
            except Exception:
                # If we cannot open a file, we keep running without file logging
                self._fh = None

    def _emit(self, line: str):
        if self.print_to_console:
            print(line)
        if self._fh:
            try:
                self._fh.write(line + "\n")
                self._fh.flush()  # ensure immediate persistence
            except Exception:
                pass

    def _log(self, level: str, message: str):
        evt = LogEvent(datetime.now(), level, message)
        self.events.append(evt)
        self._emit(f"[{evt.ts:%Y-%m-%d %H:%M:%S}] {level}: {message}")

    def info(self, message: str):
        self._log("INFO", message)

    def warn(self, message: str):
        self._log("WARN", message)

    def error(self, message: str):
        self._log("ERROR", message)

    def as_text(self) -> str:
        return "\n".join(f"[{e.ts:%H:%M:%S}] {e.level}: {e.message}" for e in self.events)

    def last_line(self) -> str:
        if not self.events:
            return "Starting…"
        e = self.events[-1]
        return f"[{e.ts:%H:%M:%S}] {e.level}: {e.message}"

    def close(self):
        try:
            if self._fh:
                self._fh.close()
        except Exception:
            pass