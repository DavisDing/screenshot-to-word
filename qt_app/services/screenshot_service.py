from __future__ import annotations

import datetime as _dt
from pathlib import Path
from typing import Optional

from PySide6.QtCore import QBuffer, QByteArray, QIODevice
from PySide6.QtGui import QGuiApplication


class ScreenshotService:
    def __init__(self, temp_dir: Path, logger):
        self.temp_dir = Path(temp_dir)
        self.temp_dir.mkdir(parents=True, exist_ok=True)
        self.logger = logger

    def capture_fullscreen(self, stem: str = "screenshot") -> Optional[Path]:
        timestamp = _dt.datetime.now().strftime("%Y%m%d_%H%M%S")
        safe_stem = "".join(ch if ch.isalnum() or ch in "-_" else "_" for ch in stem)[:50] or "screenshot"
        path = self.temp_dir / f"{safe_stem}_{timestamp}.png"
        screen = QGuiApplication.primaryScreen()
        if screen is None:
            self.logger.log("截图失败：未检测到屏幕")
            return None
        pixmap = screen.grabWindow(0)
        if pixmap.isNull():
            self.logger.log("截图失败：系统返回空图像，请检查屏幕录制/截图权限")
            return None
        if not pixmap.save(str(path), "PNG"):
            self.logger.log(f"截图保存失败：{path}")
            return None
        self.logger.log(f"截图保存：{path}")
        return path
