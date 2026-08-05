from __future__ import annotations

from PySide6.QtGui import QIcon

from qt_app.services.paths import get_resource_path


def app_icon() -> QIcon:
    """Load the bundled app icon for every top-level window/dialog."""
    icon = QIcon()
    for name in ("icon.ico", "assets/icon.png"):
        path = get_resource_path(name)
        if path.exists():
            icon.addFile(str(path))
    return icon
