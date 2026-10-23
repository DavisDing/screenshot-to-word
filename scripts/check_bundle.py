"""Check a real PyInstaller bundle before CI uploads it; never delete files."""
from __future__ import annotations

import argparse
from pathlib import Path
import re
import sys


REQUIRED_MODULES = {
    "qt_app.services.excel_service",
    "qt_app.services.word_service",
    "qt_app.services.screenshot_service",
    "qt_app.ui.main_window",
    "qt_app.ui.excel_mapping_dialog",
    "qt_app.ui.control_window",
    "qt_app.ui.annotator_window",
    "PySide6.support.deprecated",
    "PIL.Image",
    "openpyxl.reader.drawings",
    "docx",
}
REQUIRED_RESOURCES = (
    "icon.ico",
    "assets/icon.png",
    "assets/chevron-down.svg",
    "docx/templates/default.docx",
)
UNUSED_QT_STACKS = re.compile(r"qt6?(?:webengine|qml|quick|multimedia|3d|virtualkeyboard)", re.I)


def check_contents(paths: list[str], modules: set[str], platform: str) -> list[str]:
    """Validate paths and the embedded PYZ module index, without executing the app."""
    normalized = [path.replace("\\", "/").lower() for path in paths]
    errors = []
    for resource in REQUIRED_RESOURCES:
        if not any(path == resource or path.endswith("/" + resource) for path in normalized):
            errors.append(f"Missing resource: {resource}")

    extensions = {
        "PySide6/QtCore", "PySide6/QtGui", "PySide6/QtWidgets", "PySide6/QtSvg",
        "shiboken6/Shiboken", "lxml/etree",
    }
    for extension in sorted(extensions):
        prefix = extension.lower() + "."
        if not any(
            (path.startswith(prefix) or "/" + prefix in path)
            and path.endswith((".pyd", ".so"))
            for path in normalized
        ):
            errors.append(f"Missing extension: {extension}")

    plugins = ("qwindows.dll", "qsvg.dll") if platform == "win32" else ("libqcocoa.dylib", "libqsvg.dylib")
    for plugin in plugins:
        if not any(path.endswith("/" + plugin) for path in normalized):
            errors.append(f"Missing Qt plugin: {plugin}")
    for module in sorted(REQUIRED_MODULES - modules):
        errors.append(f"Missing bundled module: {module}")
    unused = [path for path in normalized if UNUSED_QT_STACKS.search(path) or "/qml/" in path]
    if unused:
        errors.append("Unused Qt stacks were collected: " + ", ".join(unused[:10]))
    return errors


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("bundle", type=Path, help="onedir directory or macOS .app")
    args = parser.parse_args()
    root = args.bundle.resolve()
    if sys.platform not in {"win32", "darwin"}:
        parser.error("Run on Windows or macOS, matching the built bundle")
    if not root.is_dir():
        parser.error(f"Bundle directory does not exist: {root}")

    # Read the actual frozen archive, rather than assuming source imports prove
    # that delayed service imports made it into the executable.
    from PyInstaller.archive.readers import CArchiveReader

    executable = root / ("Contents/MacOS/DesktopTestTool" if sys.platform == "darwin" else "DesktopTestTool.exe")
    archive = CArchiveReader(str(executable))
    pyz_name = next(name for name in archive.toc if name.endswith(".pyz"))
    modules = set(archive.open_embedded_archive(pyz_name).toc)
    files = [path for path in root.rglob("*") if path.is_file()]
    errors = check_contents([path.relative_to(root).as_posix() for path in files], modules, sys.platform)
    if errors:
        for error in errors:
            print(error, file=sys.stderr)
        return 1

    # macOS frameworks contain symlinks; count each underlying file only once.
    sizes = {}
    for path in files:
        stat = path.stat()
        sizes[stat.st_dev, stat.st_ino] = stat.st_size
    print(f"Bundle checks passed: {len(files)} file entries, {sum(sizes.values()) / 1024**2:.2f} MiB unique file bytes")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
