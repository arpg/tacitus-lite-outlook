"""Where the mail folder lives.

From a source checkout it is the checkout itself. From the installed
program (a PyInstaller build, `sys.frozen`) it is `Documents\\OutlookMail`,
the same folder the installer creates. `TACITUS_LITE_HOME` overrides
both, which the tests and the CI install check use.
"""
from __future__ import annotations

import os
import sys
from pathlib import Path

HOME_ENV = "TACITUS_LITE_HOME"
HOME_NAME = "OutlookMail"
CSIDL_PERSONAL = 5


def documents() -> Path:
    """The user's Documents folder, as Windows resolves it (OneDrive included)."""
    if sys.platform == "win32":
        try:
            import ctypes
            import ctypes.wintypes
            buf = ctypes.create_unicode_buffer(ctypes.wintypes.MAX_PATH)
            if ctypes.windll.shell32.SHGetFolderPathW(None, CSIDL_PERSONAL, None, 0, buf) == 0 and buf.value:
                return Path(buf.value)
        except Exception:
            pass
    return Path.home() / "Documents"


def is_frozen() -> bool:
    return bool(getattr(sys, "frozen", False))


def source_root() -> Path:
    return Path(__file__).resolve().parent.parent


def home() -> Path:
    env = os.environ.get(HOME_ENV)
    if env:
        return Path(env).expanduser().resolve()
    if is_frozen():
        return documents() / HOME_NAME
    return source_root()
