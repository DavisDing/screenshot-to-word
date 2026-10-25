from __future__ import annotations

import datetime as _dt
import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path
from threading import Lock
from typing import Callable, Optional


class AppLogger:
    def __init__(self, log_file: Path, sink: Optional[Callable[[str], None]] = None):
        self.log_file = Path(log_file)
        self.log_file.parent.mkdir(parents=True, exist_ok=True)
        self.sink = sink
        self._lock = Lock()
        self._handler = RotatingFileHandler(
            self.log_file, maxBytes=2 * 1024 * 1024, backupCount=3, encoding="utf-8"
        )
        self._handler.setFormatter(logging.Formatter("%(message)s"))

    def set_sink(self, sink: Optional[Callable[[str], None]]) -> None:
        self.sink = sink

    def close(self) -> None:
        with self._lock:
            self._handler.close()

    def log(self, message: str) -> None:
        timestamp = _dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        line = f"[{timestamp}] {message}"
        with self._lock:
            try:
                self._handler.emit(logging.makeLogRecord({"msg": line, "args": ()}))
            except Exception:
                pass
        if self.sink:
            self.sink(line)
