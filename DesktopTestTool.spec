# Shared Windows/macOS onedir build. Keep Qt's official dependency/plugin hooks;
# collecting every PySide6 submodule also collects unused Qt application stacks.
import sys
from pathlib import Path

from PyInstaller.utils.hooks import collect_data_files


root = Path(SPECPATH)
is_macos = sys.platform == "darwin"

a = Analysis(
    [str(root / "main.py")],
    pathex=[str(root)],
    binaries=[],
    datas=[
        (str(root / "icon.ico"), "."),
        (str(root / "assets"), "assets"),
    ] + collect_data_files("docx", includes=["templates/*"]),
    hiddenimports=[
        # The stylesheet loads an SVG at runtime, without a Python QtSvg import.
        "PySide6.QtSvg",
        "PySide6.support.deprecated",
    ],
    hookspath=[str(root / "pyinstaller_hooks")],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        "PySide6.QtWebEngineCore",
        "PySide6.QtWebEngineWidgets",
        "PySide6.QtWebEngineQuick",
        "PySide6.QtQml",
        "PySide6.QtQuick",
        "PySide6.QtQuickWidgets",
        "PySide6.QtQuickControls2",
        "PySide6.QtQuick3D",
        "PySide6.QtMultimedia",
        "PySide6.QtMultimediaWidgets",
        "PySide6.Qt3DCore",
        "PySide6.Qt3DRender",
        "PySide6.Qt3DInput",
        "PySide6.Qt3DLogic",
        "PySide6.Qt3DAnimation",
        "PySide6.Qt3DExtras",
    ],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="DesktopTestTool",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    icon=str(root / ("icon.icns" if is_macos else "icon.ico")),
)
bundle = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    name="DesktopTestTool",
)
if is_macos:
    app = BUNDLE(
        bundle,
        name="DesktopTestTool.app",
        icon=str(root / "icon.icns"),
        bundle_identifier=None,
    )
