from __future__ import annotations

import os

DEFAULT_VERSION = "dev"

try:
    from qt_app._build_version import APP_VERSION as _PACKAGED_VERSION
except Exception:
    _PACKAGED_VERSION = ""

APP_VERSION = os.environ.get("APP_VERSION") or _PACKAGED_VERSION or DEFAULT_VERSION
VERSION_LABEL = APP_VERSION if APP_VERSION != DEFAULT_VERSION else "开发版"
