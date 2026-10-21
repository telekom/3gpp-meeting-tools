# --- File: src/core/utils/folder_opener.py ---
import os
import webbrowser
from pathlib import Path

from PyQt5.QtWidgets import QMessageBox


def open_folder(p: Path) -> None:
    """Opens a directory or file in the native file explorer or browser."""
    path_obj = Path(p)
    if path_obj.exists():
        if hasattr(os, "startfile"):
            os.startfile(str(path_obj))
        else:
            webbrowser.open(f"file:///{path_obj}")
    else:
        QMessageBox.warning(None, "Not Found", "Target file/folder does not exist yet.")
