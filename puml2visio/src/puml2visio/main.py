import sys
import logging
import urllib.request
import os
from pathlib import Path

from PyQt5.QtWidgets import QApplication, QDialog

from puml2visio.ui.ui_components import GLOBAL_STYLE, ProxyDialog, create_app_icon
from puml2visio.ui.main_window import DragDropUI
from puml2visio.utils.utils import JAR_NAME, get_best_java
from puml2visio.utils.paths import get_project_root, get_asset_path

# ==========================================
# --- PATH RESOLUTION ---
# ==========================================
# __file__ is src/puml2visio/main.py
PACKAGE_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = PACKAGE_DIR.parent.parent # The outer puml2visio folder

# ==========================================
# --- LOGGING SETUP ---
# ==========================================
# get_project_root() gives us src/puml2visio/.
# We go up one more parent to put the log file next to pyproject.toml
log_file_path = get_project_root().parent / "puml2vsdx.log"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(log_file_path, mode="a", encoding="utf-8"),
        logging.StreamHandler(sys.stdout)
    ]
)

if __name__ == '__main__':
    # This prevents Windows from grouping our app under the generic Python snake logo!
    if os.name == 'nt':
        import ctypes
        myappid = '3gpp.puml2visio.converter.1.0'
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)

    app = QApplication(sys.argv)
    app.setWindowIcon(create_app_icon())
    app.setStyle("Fusion")
    app.setStyleSheet(GLOBAL_STYLE)

    # --- POINT TO THE TEMPLATES FOLDER ---
    jar_path = get_asset_path(JAR_NAME)
    version_file = jar_path.with_suffix('.version')

    # --- SMART PROXY CHECK ---
    needs_download = False
    if not jar_path.exists():
        needs_download = True
    else:
        _, java_major = get_best_java()
        if java_major > 0:
            required_type = "modern" if java_major >= 11 else "legacy"
            current_type = None
            if version_file.exists():
                try:
                    current_type = version_file.read_text(encoding="utf-8").strip()
                except:
                    pass
            if current_type != required_type:
                needs_download = True

    if needs_download:
        proxy_dialog = ProxyDialog()
        if proxy_dialog.exec_() == QDialog.Accepted:
            http_val, https_val = proxy_dialog.get_proxies()
            proxies = {}
            if http_val: proxies['http'] = http_val
            if https_val: proxies['https'] = https_val
            if proxies:
                proxy_handler = urllib.request.ProxyHandler(proxies)
                opener = urllib.request.build_opener(proxy_handler)
                urllib.request.install_opener(opener)

    window = DragDropUI()
    window.show()
    sys.exit(app.exec_())