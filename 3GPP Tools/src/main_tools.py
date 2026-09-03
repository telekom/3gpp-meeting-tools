# --- File: src/main_tools.py ---
import sys
import logging
import urllib.request
import os
import threading
import time
import traceback
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
# --- LOGGING SETUP ---
# ==========================================
log_file_path = get_project_root() / "3gpp_tools.log"
dump_file_path = get_project_root() / "startup_freeze_dump.txt"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(log_file_path, mode="a", encoding="utf-8"),
        logging.StreamHandler(sys.stdout)
    ]
)

event_loop_running = False


def direct_console_write(message: str):
    """Bypasses Python logging buffers to guarantee direct terminal output."""
    sys.__stderr__.write(message + "\n")
    sys.__stderr__.flush()


def startup_watchdog(timeout_seconds=7.0):
    def _monitor():
        start_time = time.time()
        dumped = False
        while not event_loop_running:
            time.sleep(2.0)
            elapsed = time.time() - start_time
            if event_loop_running:
                break

            if elapsed < timeout_seconds:
                direct_console_write(f"⏱️ [WATCHDOG HEARTBEAT] Startup in progress ({elapsed:.1f}s elapsed)...")
            elif not dumped:
                dumped = True
                direct_console_write(f"\n🚨 [WATCHDOG] Startup has hung for {elapsed:.1f}s! DUMPING ALL ACTIVE THREADS:")

                dump_lines = [f"=== STARTUP FREEZE THREAD DUMP ({time.ctime()}) ==="]
                for thread_id, frame in sys._current_frames().items():
                    thread_name = next(
                        (t.name for t in threading.enumerate() if t.ident == thread_id),
                        f"Thread-{thread_id}"
                    )
                    stack_summary = "".join(traceback.format_stack(frame))
                    entry = f"\n>>> THREAD: {thread_name} (ID: {thread_id}) <<<\n{stack_summary}"
                    dump_lines.append(entry)
                    direct_console_write(entry)

                try:
                    with open(dump_file_path, "w", encoding="utf-8") as f:
                        f.write("\n".join(dump_lines))
                    direct_console_write(f"\n📁 Stack dump written to: {dump_file_path}\n")
                except Exception as e:
                    direct_console_write(f"⚠️ Could not write dump file: {e}")

    t = threading.Thread(target=_monitor, daemon=True, name="StartupWatchdog")
    t.start()


if __name__ == '__main__':
    if os.name == 'nt':
        import ctypes

        myappid = '3GPP Delegate Tools.1.0'
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)

    startup_watchdog(timeout_seconds=7.0)

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


    def show_main_window():
        global event_loop_running
        logging.info("🏁 [STARTUP] Displaying window inside active event loop...")
        window.show()
        event_loop_running = True
        logging.info("🚀 [STARTUP] Qt Event Loop running smoothly.")


    # Schedule window display on the first tick of the event loop
    QTimer.singleShot(0, show_main_window)

    logging.info("🏁 [STARTUP] Entering Qt event loop (app.exec_)...")
    sys.exit(app.exec_())