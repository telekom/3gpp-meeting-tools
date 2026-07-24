import sys
import logging
import urllib.request
import os
from pathlib import Path

from PyQt5.QtWidgets import QApplication, QDialog

from core.ui.ui_components import GLOBAL_STYLE, ProxyDialog, create_app_icon
from core.utils.utils import get_best_java
from core.utils.paths import get_project_root
from modules.meetings.plugin_loader import register_meetings_plugin
from modules.puml2visio.plugin_loader import register_puml2visio_plugin
from main_window import DragDropUI
from modules.puml2visio.config.paths import PLANTUML_JAR_NAME
from modules.puml2visio.utils.paths import get_puml2visio_asset_path
from modules.specifications.plugin_loader import register_specs_plugin

from modules.word_tools.plugin_loader import register_word_plugin

# ==========================================
# --- PATH RESOLUTION ---
# ==========================================
PACKAGE_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = PACKAGE_DIR.parent.parent # The outer 3GPP Tools folder

# ==========================================
# --- LOGGING SETUP ---
# ==========================================
# get_project_root() gives us src/3GPP Tools/.
# We go up one more parent to put the log file next to pyproject.toml
log_file_path = get_project_root() / "3gpp_tools.log"

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
        myappid = '3gpp.3GPP Tools.converter.1.1'
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)

    register_puml2visio_plugin()
    register_word_plugin()
    register_specs_plugin()
    register_meetings_plugin()

    app = QApplication(sys.argv)
    app.setWindowIcon(create_app_icon())
    app.setStyle("Fusion")
    app.setStyleSheet(GLOBAL_STYLE)

    # --- POINT TO THE TEMPLATES FOLDER ---
    jar_path = get_puml2visio_asset_path(PLANTUML_JAR_NAME)
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