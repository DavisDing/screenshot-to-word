from __future__ import annotations

import datetime as _dt
from pathlib import Path
import time
from typing import Optional
class ScreenshotService:
    def __init__(self, temp_dir: Path, logger):
        self.temp_dir = Path(temp_dir)
        self.temp_dir.mkdir(parents=True, exist_ok=True)
        self.logger = logger
        self.cleanup_stale_files()

    def cleanup_stale_files(self, *, max_age_days: int = 7) -> int:
        """Remove only old PNGs directly in the application's managed Temp folder."""
        threshold = time.time() - max_age_days * 24 * 60 * 60
        removed = 0
        for path in self.temp_dir.glob("*.png"):
            try:
                if path.is_file() and path.stat().st_mtime < threshold:
                    path.unlink()
                    removed += 1
            except OSError as exc:
                self.logger.log(f"[警告] 清理过期临时截图失败：{path}：{exc}")
        if removed:
            self.logger.log(f"已清理 {removed} 张过期临时截图")
        return removed

    def capture_fullscreen(self, stem: str = "screenshot") -> Optional[Path]:
        from PySide6.QtGui import QGuiApplication

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
