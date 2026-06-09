import sys
import logging
import urllib.request
from pathlib import Path

from PyQt5.QtWidgets import QApplication, QDialog

from ui_components import GLOBAL_STYLE, ProxyDialog
from main_window import DragDropUI
from utils import JAR_NAME, get_best_java

# ==========================================
# --- LOGGING SETUP ---
# ==========================================
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler("puml2vsdx.log", mode="a", encoding="utf-8"),
        logging.StreamHandler(sys.stdout)
    ]
)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    app.setStyleSheet(GLOBAL_STYLE)

    jar_path = Path(__file__).parent.resolve() / JAR_NAME
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