from __future__ import annotations

import datetime as _dt
from pathlib import Path
from typing import Callable, Optional


class AppLogger:
    def __init__(self, log_file: Path, sink: Optional[Callable[[str], None]] = None):
        self.log_file = Path(log_file)
        self.log_file.parent.mkdir(parents=True, exist_ok=True)
        self.sink = sink

    def set_sink(self, sink: Optional[Callable[[str], None]]) -> None:
        self.sink = sink

    def log(self, message: str) -> None:
        timestamp = _dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        line = f"[{timestamp}] {message}"
        try:
            with self.log_file.open("a", encoding="utf-8") as f:
                f.write(line + "\n")
        except Exception:
            pass
        if self.sink:
            self.sink(line)
