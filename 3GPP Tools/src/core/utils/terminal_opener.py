# --- File: src/core/utils/terminal_opener.py ---
import subprocess
from pathlib import Path
from PyQt5.QtWidgets import QMessageBox


def open_terminal(path: Path, parent=None):
    """Open a terminal with its working directory set to path."""
    path = Path(path)
    if not path.is_dir():
        QMessageBox.warning(parent, "Not Found", "The target folder does not exist.")
        return

    try:
        subprocess.Popen(["wt.exe", "-d", str(path)], cwd=str(path))
        return
    except FileNotFoundError:
        # Windows Terminal is not installed or its app execution alias is unavailable.
        pass
    except OSError as exc:
        QMessageBox.critical(parent, "Terminal Error", f"Could not open Windows Terminal:\n{exc}")
        return

    # Keep the fallback useful on Windows installations without Windows Terminal.
    escaped_path = str(path).replace("'", "''")
    try:
        subprocess.Popen(
            ["powershell.exe", "-NoExit", "-Command", f"Set-Location -LiteralPath '{escaped_path}'"],
            cwd=str(path),
        )
    except OSError as exc:
        QMessageBox.critical(parent, "Terminal Error", f"Could not open a terminal:\n{exc}")
