from __future__ import annotations

import os
import sys
import threading
import traceback
from pathlib import Path

from qt_app.services.paths import ensure_app_dirs, get_base_path
from qt_app.version import APP_VERSION

APP_TITLE = "Screenshot To Word"


def write_crash_log(base_path: Path, exc: BaseException) -> Path:
    log_dir = base_path / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    crash_path = log_dir / "crash.log"
    with crash_path.open("a", encoding="utf-8") as f:
        f.write("\n" + "=" * 80 + "\n")
        f.write("启动异常\n")
        f.write("".join(traceback.format_exception(type(exc), exc, exc.__traceback__)))
    return crash_path


def main() -> int:
    base_path = get_base_path()
    os.chdir(base_path)
    ensure_app_dirs(base_path)

    try:
        from PySide6.QtWidgets import QApplication, QMessageBox
        from qt_app.services.logger import AppLogger
        from qt_app.ui.icons import app_icon
        from qt_app.ui.main_window import MainWindow
        from qt_app.ui.style import APP_QSS
    except Exception as exc:  # Dependency/import problems should be diagnosable after packaging.
        crash_path = write_crash_log(base_path, exc)
        print(f"启动失败，请检查依赖或打包配置。详情见：{crash_path}", file=sys.stderr)
        return 1

    app = QApplication(sys.argv)
    app.setApplicationName(APP_TITLE)
    app.setApplicationVersion(APP_VERSION)
    icon = app_icon()
    if not icon.isNull():
        app.setWindowIcon(icon)
    app.setStyleSheet(APP_QSS)
    logger = AppLogger(base_path / "logs" / "output.log")

    def report_unhandled(exc_type, exc_value, exc_traceback) -> None:
        details = "".join(traceback.format_exception(exc_type, exc_value, exc_traceback))
        logger.log("[未处理异常]\n" + details)

    # PySide routes uncaught slot exceptions through sys.excepthook; install the
    # thread hook as well so background callback failures reach the same log.
    previous_exception_hook = sys.excepthook
    previous_thread_hook = threading.excepthook
    sys.excepthook = report_unhandled
    threading.excepthook = lambda args: report_unhandled(args.exc_type, args.exc_value, args.exc_traceback)

    try:
        window = MainWindow(base_path, logger)
        window.show()
        return app.exec()
    except Exception as exc:
        crash_path = write_crash_log(base_path, exc)
        try:
            QMessageBox.critical(None, "程序启动失败", f"程序启动失败：{exc}\n\n详情已写入：{crash_path}")
        except Exception:
            pass
        return 1
    finally:
        sys.excepthook = previous_exception_hook
        threading.excepthook = previous_thread_hook
        logger.set_sink(None)
        logger.close()


if __name__ == "__main__":
    raise SystemExit(main())
