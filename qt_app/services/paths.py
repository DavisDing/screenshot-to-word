from __future__ import annotations

import os
import sys
from pathlib import Path


def get_base_path() -> Path:
    """Return user-visible app directory for both source and frozen builds."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parents[2]


def ensure_app_dirs(base_path: Path) -> None:
    for name in ("excel_input", "word_output", "logs", "Temp"):
        (base_path / name).mkdir(parents=True, exist_ok=True)


def get_resource_path(relative_path: str) -> Path:
    """Return a bundled resource path in source and PyInstaller builds."""
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS) / relative_path
    return Path(__file__).resolve().parents[2] / relative_path
