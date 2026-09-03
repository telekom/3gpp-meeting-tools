# --- File: src/main_tools.py ---
import sys
import logging
import urllib.request
import os
import faulthandler
from pathlib import Path

from PyQt5.QtWidgets import QApplication, QDialog
from PyQt5.QtCore import QTimer

from core.ui.ui_components import GLOBAL_STYLE, ProxyDialog, create_app_icon
from core.utils.utils import get_best_java
from core.utils.paths import get_project_root
from modules.meetings.plugin_loader import register_meetings_plugin
from modules.puml2visio.plugin_loader import register_puml2visio_plugin
from main_window import DragDropUI
from modules.puml2visio.config.paths import PLANTUML_JAR_NAME
from modules.puml2visio.utils.paths import get_puml2visio_asset_path
from modules.spec_search.plugin_loader import register_spec_search_plugin
from modules.specifications.plugin_loader import register_specs_plugin
from modules.word_tools.plugin_loader import register_word_plugin

# ==========================================
# --- FAULTHANDLER & LOGGING SETUP ---
# ==========================================
# Enables pure C-level thread dumps that work even if Qt locks the Python GIL
faulthandler.enable()
faulthandler.dump_traceback_later(timeout=5.0, repeat=True, file=sys.__stderr__)

log_file_path = get_project_root() / "3gpp_tools.log"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(log_file_path, mode="a", encoding="utf-8"),
        logging.StreamHandler(sys.stdout)
    ]
)

def direct_console_write(message: str):
    """Bypasses Python buffers to guarantee immediate terminal output."""
    sys.__stderr__.write(message + "\n")
    sys.__stderr__.flush()

if __name__ == '__main__':
    if os.name == 'nt':
        import ctypes
        myappid = '3GPP Delegate Tools.1.0'
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)

    logging.info("🏁 [STARTUP] Registering plugins...")
    register_puml2visio_plugin()
    register_word_plugin()
    register_specs_plugin()
    register_meetings_plugin()
    register_spec_search_plugin()
    logging.info("🏁 [STARTUP] Plugins registered successfully.")

    app = QApplication(sys.argv)
    app.setWindowIcon(create_app_icon())
    app.setStyle("Fusion")
    app.setStyleSheet(GLOBAL_STYLE)
    app.setQuitOnLastWindowClosed(True)

    jar_path = get_puml2visio_asset_path(PLANTUML_JAR_NAME)
    version_file = jar_path.with_suffix('.version')

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
                except Exception:
                    pass
            if current_type != required_type:
                needs_download = True

    if needs_download:
        logging.info("🏁 [STARTUP] Prompting for proxy settings...")
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

    logging.info("🏁 [STARTUP] Instantiating main window (DragDropUI)...")
    window = DragDropUI()

    direct_console_write("🏁 [STARTUP] Displaying main window...")
    window.show()
    direct_console_write("🚀 [STARTUP] window.show() completed successfully.")

    # Window displayed successfully; cancel the hang watchdog
    faulthandler.cancel_dump_traceback_later()

    logging.info("🏁 [STARTUP] Entering Qt event loop (app.exec_)...")
    exit_code = app.exec_()
    os._exit(exit_code)