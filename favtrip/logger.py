"""
logger
======================================
This module provides two dataclasses—`LogEvent` and `StatusLogger`—to record simple,
human-readable status messages during a process or script run. It is designed to be:

- **Simple**: minimal API (`info`, `warn`, `error`) and a small in-memory log.
- **Immediate**: console prints occur synchronously; file writes are line-buffered and flushed.
- **Fail-open**: if a log file cannot be opened or written, logging proceeds to console and memory.
- **Portable**: standard library only (dataclasses, datetime, typing).

-------------------------------------------------------------------------------
Data Model
-------------------------------------------------------------------------------
- LogEvent
    - ts (datetime.datetime): Timestamp captured via `datetime.now()` when the event is recorded.
      Note: this is a **naive** datetime in local time.
    - level (str): Log level label (e.g., "INFO", "WARN", "ERROR").
    - message (str): The event text.

- StatusLogger
    - events (list[LogEvent]): In-memory event history in append order.
    - print_to_console (bool): If True (default), each log line is printed to stdout.
    - file_path (str | None): If set, lines are also written to this file. If `None`, file logging
      is disabled. Default is "last_run.log".
    - overwrite (bool): If True (default), the log file is opened in write mode on instantiation;
      otherwise it is appended to.

-------------------------------------------------------------------------------
Output Format
-------------------------------------------------------------------------------
- Console/file lines: `[YYYY-MM-DD HH:MM:SS] LEVEL: message`
- `as_text()`:         `[HH:MM:SS] LEVEL: message` per line (no date, suitable for compact display)
- `last_line()`:       Returns the most recent line in `as_text()` format, or `"Starting…"` if empty.

-------------------------------------------------------------------------------
Behavior & Guarantees
-------------------------------------------------------------------------------
- **File handling**: On initialization, if `file_path` is provided, the file is opened once in
  line-buffered text mode (`buffering=1`) and UTF-8 encoding. If opening fails, the logger
  continues without a file handle.
- **Atomicity**: Each `_emit` call attempts to write a single line and then flush. Any file write
  errors are swallowed; console output and in-memory storage are unaffected.
- **Timestamps**: Timestamps are captured at call time (`datetime.now()`), local time, naive datetimes.
- **Memory growth**: All events are retained in `events`; for long-running processes, consider
  pruning or exporting periodically.
- **Thread-safety**: Not thread-safe. If you need concurrent logging, protect calls with a lock or
  adapt the implementation for multi-thread/process usage.
- **No rotation**: No file rotation or size limiting. Use external tools or extend as needed.
"""

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
