# --- File: src/core/ui/main_window/shell_window.py ---
"""
DragDropUI shell: constructor, Qt lifecycle overrides (resize/show/close) and logging sink.
Feature logic lives in the mixins/ package (see README.md for a 'where to edit what' map).
"""
import logging
from PyQt5.QtCore import QTimer
from PyQt5.QtWidgets import QMainWindow, QApplication
from core.queue_manager import QueueManager
from core.ui.right_side_panel import MODE_ASSISTANT
from core.ui.shell_state import restore_shell_state, save_shell_state
from core.ui.window_session import collect_open_windows_state, save_open_windows_state
from core.ui.ui_components import (
    apply_native_window_icon,
    create_app_icon,
    install_window_icon_manager,
)
from core.ui.ui_panels import GuiLogHandler
from core.utils.paths import get_project_root
from modules.puml2visio.config.paths import PLANTUML_JAR_NAME
from modules.puml2visio.core.live_preview import LivePreviewManager
from modules.puml2visio.utils.paths import get_puml2visio_asset_path
from core.ui.main_window.mixins import (
    ShellSetupMixin,
    PanelsNavigationMixin,
    PumlConversionMixin,
    SystemServicesMixin,
)


class DragDropUI(
    ShellSetupMixin,
    PanelsNavigationMixin,
    PumlConversionMixin,
    SystemServicesMixin,
    QMainWindow,
):
    """Main application shell window composing UI setup, panels, conversion and system services."""
    def __init__(self, gui_log_handler=None):
        super().__init__()
        self.setWindowTitle("3GPP Delegate Tools")
        self.setWindowIcon(create_app_icon())
        install_window_icon_manager(QApplication.instance())
        self.setMinimumSize(780, 560)
        self.resize(1160, 780)

        self.jar_path = get_puml2visio_asset_path(PLANTUML_JAR_NAME)
        self.last_out_path = ""
        self.init_thread = None
        self._last_update_check_result = None
        # Background monitors are managed by self.status_bar
        self._services_started = False  # Guard to prevent duplicate background starts
        self._visio_ready = True
        self._core_ready = True
        self._assistant_docked = False

        logging.info("🏁 [STARTUP:UI] Starting _setup_ui()...")
        self._setup_ui()
        restore_shell_state(self)
        if (
            hasattr(self, "right_panel")
            and self.right_panel.is_open
            and self.right_panel.mode == MODE_ASSISTANT
            and hasattr(self, "tabs")
            and hasattr(self, "assistant_tab")
            and self.tabs.currentWidget() is self.assistant_tab
        ):
            self.right_panel.close_panel()
        else:
            self._sync_assistant_dock_slot()
        logging.info("🏁 [STARTUP:UI] _setup_ui() finished successfully.")

        # ==========================================
        # --- GLOBAL GUI LOGGER ---
        # ==========================================
        #
        # main_tools normally creates GuiLogHandler before application startup
        # so that logs generated before ConsolePanel exists are not lost.
        #
        # Keep a fallback here so DragDropUI can still be instantiated directly
        # by development tools or other callers without having to provide a
        # handler explicitly.
        if gui_log_handler is None:
            self.gui_logger = GuiLogHandler()
            self.gui_logger.setLevel(logging.NOTSET)
            self.gui_logger.setFormatter(
                logging.Formatter("%(message)s")
            )
            logging.getLogger().addHandler(self.gui_logger)
        else:
            self.gui_logger = gui_log_handler

        # ConsolePanel now exists because _setup_ui() has completed.
        #
        # attach() connects the Qt signal and immediately replays all startup
        # records that were buffered before the panel was available.
        self.gui_logger.attach(self.console_panel.log_message)
        self.gui_logger.log_emitted.connect(self.bottom_drawer.record_log_event)

        # Wire Queue Manager
        logging.info("🏁 [STARTUP:QUEUE] Initializing QueueManager...")
        app_context = {"jar_path": self.jar_path}
        self.queue_manager = QueueManager(app_context=app_context)
        self.queue_manager.queue_updated.connect(self._on_queue_updated)
        self.queue_manager.processing_state_changed.connect(self._update_system_status)
        self.queue_manager.conversion_success.connect(self.on_conversion_success)
        self.queue_manager.conversion_failed.connect(self.on_conversion_failed)

        self.queue_panel.remove_requested.connect(self.queue_manager.remove_items)
        self.queue_panel.clear_requested.connect(self.queue_manager.clear_queue)
        self.queue_panel.abort_requested.connect(self.queue_manager.abort_current_task)

        self.queue_manager.processing_state_changed.connect(
            lambda is_proc, status: self.queue_panel.abort_btn.setEnabled(is_proc)
        )

        logging.info("🏁 [STARTUP:CACHE] Loading previous editor session cache...")
        self.cache_file = get_project_root() / ".editor_cache.puml"
        self._load_cache()
        logging.info("🏁 [STARTUP:CACHE] Previous editor cache step complete.")

        self.save_timer = QTimer()
        self.save_timer.setSingleShot(True)
        self.save_timer.setInterval(2000)
        self.save_timer.timeout.connect(self.save_cache)
        self.code_tab.text_input.textChanged.connect(self.save_timer.start)

        logging.info("🏁 [STARTUP:PREVIEW] Initializing LivePreviewManager...")
        self.live_preview = LivePreviewManager(self.code_tab.text_input, self.jar_path)
        self.live_preview.log_msg.connect(self.log_message)
        self.live_preview.svg_updated.connect(self.code_tab.svg_preview.load_svg)
        self.live_preview.error_detected.connect(self.code_tab.text_input.set_error_line)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        dock_open = hasattr(self, "right_panel") and self.right_panel.is_open
        if hasattr(self, "nav_rail"):
            self.nav_rail.apply_responsive_width(self.width(), dock_open=dock_open)
        if hasattr(self, "status_bar") and hasattr(self.status_bar, "apply_responsive_width"):
            self.status_bar.apply_responsive_width(self.width())
        if hasattr(self, "toast_manager"):
            self.toast_manager.reposition()

    def log_message(self, message: str, level=logging.INFO):
        logging.log(level, message)

    def showEvent(self, event):
        super().showEvent(event)
        apply_native_window_icon(self, set_class_icon=True)

        if not self._services_started:
            self._services_started = True

            logging.info(
                "🏁 [STARTUP:WINDOW] DragDropUI displayed. "
                "Scheduling background services..."
            )

            QTimer.singleShot(0, self._start_background_services)


    def closeEvent(self, event):
        logging.info("🏁 [SHUTDOWN] Window closeEvent triggered. Saving cache...")
        try:
            self.save_cache()
        except Exception as e:
            logging.warning(f"⚠️ [SHUTDOWN] Could not save cache: {e}")

        try:
            save_open_windows_state(collect_open_windows_state(self))
        except Exception as e:
            logging.warning(f"⚠️ [SHUTDOWN] Could not save open workspace windows: {e}")

        try:
            save_shell_state(self)
        except Exception as e:
            logging.warning(f"⚠️ [SHUTDOWN] Could not save shell layout state: {e}")

        self.status_bar.stop_monitors()
        if hasattr(self, "toast_manager") and self.toast_manager is not None:
            self.toast_manager.dismiss_all()

        # 3. Stop initialization thread independently
        if hasattr(self, 'init_thread') and self.init_thread is not None and self.init_thread.isRunning():
            try:
                self.init_thread.quit()
                self.init_thread.wait(300)
            except Exception:
                pass

        # 4. Cooperatively finish short-lived NAS query workers.
        if hasattr(self, "nas_tab") and self.nas_tab is not None:
            try:
                self.nas_tab.shutdown(wait_ms=1500)
            except Exception as e:
                logging.warning(f"⚠️ [SHUTDOWN] Could not stop NAS workers cleanly: {e}")

        # 5. Cooperatively stop an active Assistant research run.
        if hasattr(self, "assistant_tab") and self.assistant_tab is not None:
            try:
                self.assistant_tab.shutdown(wait_ms=500)
            except Exception as e:
                logging.warning(f"⚠️ [SHUTDOWN] Could not stop Assistant worker cleanly: {e}")

        # 6. Close CodeEditorTab modeless AI dialog if open.
        if hasattr(self, "code_tab") and self.code_tab is not None:
            try:
                ai_dlg = getattr(self.code_tab, "_ai_dialog", None)
                if ai_dlg is not None:
                    ai_dlg.close()
                self.code_tab.close()
            except Exception:
                pass

        # 7. Cooperatively stop any active WordExtractorTab threads.
        if hasattr(self, "word_tab") and self.word_tab is not None:
            for attr in ("_excerpt_scan_thread", "_excerpt_extract_thread"):
                th = getattr(self.word_tab, attr, None)
                if th is not None and th.isRunning():
                    try:
                        th.requestInterruption()
                        th.wait(500)
                    except Exception:
                        pass

        # 8. Shut down remaining feature tabs and docked panels if they define shutdown or close.
        for tab_name in ("quick_search_panel", "spec_search_tab", "work_items_tab", "meetings_tab", "assistant_tab"):
            tab = getattr(self, tab_name, None)
            if tab is not None:
                try:
                    if hasattr(tab, "shutdown") and callable(tab.shutdown):
                        tab.shutdown()
                    elif hasattr(tab, "close") and callable(tab.close):
                        tab.close()
                except Exception as e:
                    logging.warning(f"⚠️ [SHUTDOWN] Could not stop {tab_name} cleanly: {e}")

        # 9. Clean up any remaining top-level windows (e.g. SessionPlanWindow, standalone dialogs)
        try:
            for widget in QApplication.topLevelWidgets():
                if widget is not self and widget.isVisible():
                    try:
                        widget.close()
                    except Exception:
                        pass
        except Exception as e:
            logging.warning(f"⚠️ [SHUTDOWN] Could not close top-level windows cleanly: {e}")

        logging.info("🏁 [SHUTDOWN] Window closing. Releasing Qt window handle...")
        super().closeEvent(event)
        event.accept()
