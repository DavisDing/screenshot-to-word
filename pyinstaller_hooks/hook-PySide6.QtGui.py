"""Keep official QtGui collection, except the unused on-screen keyboard."""
from pathlib import Path

from PyInstaller.utils.hooks.qt import add_qt6_dependencies


hiddenimports, binaries, datas = add_qt6_dependencies(__file__)
# Filter before PyInstaller scans plugin dependencies: the virtual keyboard
# pulls in QtQuick/QML even when their Python modules are excluded. Native
# platform/input-method plugins and all image plugins remain collected.
binaries = [
    entry for entry in binaries
    if "qtvirtualkeyboardplugin" not in Path(entry[0]).name.lower()
]
