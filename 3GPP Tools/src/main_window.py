import datetime
import logging
import os
import urllib.request
import webbrowser
from pathlib import Path

from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtGui import QKeySequence, QTextCursor
from PyQt5.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QSplitter, QStatusBar, QApplication,
    QDialog, QTabWidget, QPushButton, QShortcut, QLabel
)

from core.config.config import HELP_URL
from core.network.session import NetworkConfigDialog
from core.network.wifi_monitor import WifiMonitorThread
from core.queue_manager import QueueManager
from core.ui.ui_components import ProxyDialog, create_app_icon
from core.ui.ui_panels import (
    ConsolePanel, QueuePanel, ProcessManagerDialog, DatabaseMaintenanceDialog, GuiLogHandler
)
from core.utils.paths import get_project_root
from modules.meetings.ui.ui_tabs import MeetingsTab
from modules.puml2visio.config.paths import PLANTUML_JAR_NAME
from modules.puml2visio.core.live_preview import LivePreviewManager
from modules.puml2visio.core.visio_converter import VisioReaderThread
from modules.puml2visio.templates.plantuml_templates import PLANTUML_TYPES
from modules.puml2visio.ui.ui_tabs import CodeEditorTab, BatchConvertTab
from modules.puml2visio.utils.paths import get_puml2visio_asset_path
from modules.puml2visio.utils.utils import encode_plantuml, InitializationThread
from modules.specifications.ui.ui_tabs import SpecificationsTab
from modules.spec_search.ui.spec_search_tabs import SpecSearchTab
from modules.word_tools.ui.word_tabs import WordExtractorTab
from modules.work_items.ui.ui_tabs import WorkItemsTab
from modules.nas.ui.nas_tabs import NASTab
from core.ai.ollama_client import OllamaClient, OllamaMonitorThread
from core.ai.ollama_dialog import OllamaConfigDialog


class DragDropUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("3GPP Delegate Tools")
        self.setWindowIcon(create_app_icon())
        self.resize(950, 750)

        self.jar_path = get_puml2visio_asset_path(PLANTUML_JAR_NAME)
        self.last_out_path = ""
        self.init_thread = None
        self.wifi_monitor = None
        self.ollama_monitor = None
        self._services_started = False  # Guard to prevent duplicate background starts

        logging.info("🏁 [STARTUP:UI] Starting _setup_ui()...")
        self._setup_ui()
        logging.info("🏁 [STARTUP:UI] _setup_ui() finished successfully.")

        # Global GUI Logger
        self.gui_logger = GuiLogHandler()
        self.gui_logger.setFormatter(logging.Formatter('%(message)s'))
        logging.getLogger().addHandler(self.gui_logger)
        self.gui_logger.log_emitted.connect(self.console_panel.log_message)

        # Wire Queue Manager
        logging.info("🏁 [STARTUP:QUEUE] Initializing QueueManager...")
        app_context = {"jar_path": self.jar_path}
        self.queue_manager = QueueManager(app_context=app_context)
        self.queue_manager.log_msg.connect(self.log_message)
        self.queue_manager.queue_updated.connect(self.queue_panel.update_list)
        self.queue_manager.processing_state_changed.connect(self._update_system_status)
        self.queue_manager.conversion_success.connect(self.on_conversion_success)

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

    def _setup_ui(self):
        central_widget = QWidget()
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(15, 15, 15, 15)

        self.splitter = QSplitter(Qt.Vertical)

        # --- TOP HALF: TABS ---
        self.tabs = QTabWidget()

        logging.info("🏁 [STARTUP:UI:TAB] 1/8 Initializing CodeEditorTab...")
        self.code_tab = CodeEditorTab()
        self.code_tab.template_requested.connect(self.insert_template)
        self.code_tab.docs_requested.connect(self.open_template_docs)
        self.code_tab.clear_requested.connect(self.clear_editor)
        self.code_tab.undo_requested.connect(self.code_tab.text_input.undo)
        self.code_tab.copy_code_requested.connect(self.copy_editor_code)
        self.code_tab.live_view_toggled.connect(self.toggle_live_view)
        self.code_tab.planttext_requested.connect(self.show_in_planttext)
        self.code_tab.copy_path_requested.connect(self.copy_out_path)
        self.code_tab.open_folder_requested.connect(self.open_export_folder)
        self.code_tab.export_requested.connect(self._save_and_queue_pasted_text)
        self.code_tab.file_dropped.connect(self.extract_code_from_visio)

        logging.info("🏁 [STARTUP:UI:TAB] 2/8 Initializing BatchConvertTab...")
        self.batch_tab = BatchConvertTab()
        self.batch_tab.files_dropped.connect(
            lambda paths, fmt: self.queue_manager.add_batch(paths, target_format=fmt)
        )

        logging.info("🏁 [STARTUP:UI:TAB] 3/8 Initializing WordExtractorTab...")
        self.word_tab = WordExtractorTab()
        self.word_tab.extract_visio_requested.connect(
            lambda fp: self.queue_manager.add_item(Path(fp), "extract_visio")
        )
        self.word_tab.split_doc_requested.connect(
            lambda fp, pref, depth: self.queue_manager.add_item(
                Path(fp), "split_docx", {"prefix": pref, "depth": depth}
            )
        )
        self.word_tab.compare_doc_requested.connect(
            lambda a, b, keep_open: self.queue_manager.add_item(
                Path("Word_Comparison_Task"),
                "compare_docx",
                {"doc_a": a, "doc_b": b, "keep_open": keep_open}
            )
        )
        self.word_tab.convert_doc_requested.connect(
            lambda source_doc, target_fmt: self.queue_manager.add_item(
                Path(source_doc),
                "word_convert",
                {"fmt": target_fmt}
            )
        )

        db_path = get_project_root() / "3gpp_data.db"
        logging.info("🏁 [STARTUP:UI:TAB] 4/8 Initializing SpecificationsTab...")
        self.specs_tab = SpecificationsTab(db_path)
        self.specs_tab.update_db_requested.connect(
            lambda force_meta: self.queue_manager.add_item(
                Path("3GPP_Archive"),
                "update_specs_db",
                {"db_path": db_path, "force_metadata": force_meta, "target_specs": []}
            )
        )
        self.specs_tab.update_specific_requested.connect(
            lambda target_specs, force_meta: self.queue_manager.add_item(
                Path("3GPP_Archive_Targeted"),
                "update_specs_db",
                {"db_path": db_path, "force_metadata": force_meta, "target_specs": target_specs}
            )
        )

        logging.info("🏁 [STARTUP:UI:TAB] 5/8 Initializing MeetingsTab...")
        self.meetings_tab = MeetingsTab(db_path)

        logging.info("🏁 [STARTUP:UI:TAB] 6/8 Initializing WorkItemsTab...")
        self.work_items_tab = WorkItemsTab(db_path)
        self.work_items_tab.global_action_requested.connect(self.meetings_tab._handle_global_action_from_window)

        self.meetings_tab.update_db_requested.connect(
            lambda wg, docs, dyna: self.queue_manager.add_item(
                Path("3GPP_Meetings"),
                "update_meetings_db",
                {"db_path": db_path, "sync_wg": wg, "sync_docs": docs, "sync_dyna": dyna}
            )
        )
        self.meetings_tab.update_specific_requested.connect(
            lambda target_list, wg, docs, dyna: self.queue_manager.add_item(
                Path("3GPP_Meetings_Targeted"),
                "update_meetings_db",
                {
                    "db_path": db_path,
                    "target_meetings": target_list,
                    "sync_wg": wg,
                    "sync_docs": docs,
                    "sync_dyna": dyna
                }
            )
        )

        nas_db_path = get_project_root() / "3gpp_protocol_data.db"
        logging.info("🏁 [STARTUP:UI:TAB] 7/8 Initializing NASTab...")
        self.nas_tab = NASTab(nas_db_path, db_path)

        spec_search_db_path = get_project_root() / "3gpp_spec_search.db"
        logging.info("🏁 [STARTUP:UI:TAB] 8/8 Initializing SpecSearchTab...")
        self.spec_search_tab = SpecSearchTab(spec_search_db_path, db_path)

        self.tabs.addTab(self.code_tab, "📝 PlantUML")
        self.tabs.addTab(self.batch_tab, "🔄 Visio")
        self.tabs.addTab(self.word_tab, "📘 Word")
        self.tabs.addTab(self.specs_tab, "📚 Specifications")
        self.tabs.addTab(self.spec_search_tab, "🔎 Spec Search")
        self.tabs.addTab(self.work_items_tab, "📋 Work Items")
        self.tabs.addTab(self.meetings_tab, "🗓️ Meetings")
        self.tabs.addTab(self.nas_tab, "🔬 Protocols")

        # Connect tab activation for lazy loading
        self.tabs.currentChanged.connect(self._on_tab_changed)

        # Tab corner help link
        self.help_btn = QPushButton("📖 Help (F1)")
        self.help_btn.setStyleSheet("""
            QPushButton { border: none; background: transparent; color: #395396; font-weight: bold; padding: 4px 15px; }
            QPushButton:hover { color: #1E5C99; text-decoration: underline; }
        """)
        self.help_btn.setCursor(Qt.PointingHandCursor)
        self.help_btn.clicked.connect(self.open_documentation)
        self.tabs.setCornerWidget(self.help_btn, Qt.TopRightCorner)

        self.help_shortcut = QShortcut(QKeySequence("F1"), self)
        self.help_shortcut.activated.connect(self.open_documentation)

        self.splitter.addWidget(self.tabs)

        # --- BOTTOM HALF: PANELS ---
        logging.info("🏁 [STARTUP:UI:PANELS] Initializing ConsolePanel and QueuePanel...")
        self.bottom_splitter = QSplitter(Qt.Horizontal)

        self.console_panel = ConsolePanel()
        self.console_panel.proxy_requested.connect(self.open_proxy_settings)
        self.console_panel.network_config_requested.connect(lambda: NetworkConfigDialog(self).exec_())
        self.console_panel.update_requested.connect(self.check_for_jar_updates)
        self.console_panel.task_manager_requested.connect(self.open_task_manager)
        self.console_panel.db_maintenance_requested.connect(self.open_db_maintenance)

        self.specs_tab.log_msg.connect(self.console_panel.log_message)
        self.nas_tab.log_msg.connect(self.console_panel.log_message)
        self.spec_search_tab.log_msg.connect(self.console_panel.log_message)

        self.queue_panel = QueuePanel()

        self.bottom_splitter.addWidget(self.console_panel)
        self.bottom_splitter.addWidget(self.queue_panel)
        self.bottom_splitter.setSizes([650, 250])

        self.splitter.addWidget(self.bottom_splitter)
        self.splitter.setSizes([700, 100])
        main_layout.addWidget(self.splitter)

        central_widget.setLayout(main_layout)
        self.setCentralWidget(central_widget)

        # Bottom Status Bar (single clean instance)
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("⏳ Initializing...")

        # Immediate initialization of Ollama indicator from saved config
        from core.ai.ollama_client import load_ollama_config
        saved_cfg = load_ollama_config()
        saved_model = saved_cfg.get("selected_model", "").strip()
        initial_ollama_text = f"🦙 {saved_model}" if saved_model else "🦙 Ollama"

        self.ollama_indicator = QPushButton(initial_ollama_text)
        self.ollama_indicator.setCursor(Qt.PointingHandCursor)
        self.ollama_indicator.setToolTip("Click to configure Ollama LLM connection")
        self.ollama_indicator.setStyleSheet("""
                    QPushButton {
                        background-color: transparent;
                        border: 1px solid #CBD5E1;
                        border-radius: 4px;
                        padding: 1px 8px;
                        font-size: 11px;
                        font-weight: bold;
                        color: #64748B;
                    }
                    QPushButton:hover {
                        background-color: #F1F5F9;
                        border-color: #94A3B8;
                    }
                """)
        self.ollama_indicator.clicked.connect(self.open_ollama_settings)
        self.status_bar.addPermanentWidget(self.ollama_indicator)

        # Network Indicator (starts as neutral Network label instead of "Checking...")
        self.network_indicator = QLabel("📶 Network")
        self.network_indicator.setStyleSheet("color: gray; padding: 0 10px;")
        self.status_bar.addPermanentWidget(self.network_indicator)

        logging.info("🏁 [STARTUP:UI] _setup_ui() layout complete.")

    def _on_tab_changed(self, index: int):
        """Triggers lazy loading when an inactive tab is selected for the first time."""
        active_widget = self.tabs.widget(index)
        if hasattr(active_widget, 'ensure_loaded'):
            active_widget.ensure_loaded()

    def _launch_init_thread(self, check_updates=False):
        """Thread-safe launcher that prevents duplicate initialization checks."""
        if self.init_thread is not None and self.init_thread.isRunning():
            logging.warning("⚠️ [STARTUP] InitializationThread already running. Skipping duplicate launch.")
            return

        self.init_thread = InitializationThread(self.jar_path, check_updates=check_updates)
        self.init_thread.ui_log_msg.connect(self.log_message)
        self.init_thread.init_complete.connect(self.on_init_complete)
        self.init_thread.network_error.connect(self.open_proxy_settings)
        self.init_thread.start()

        # Failsafe: Ensure UI is fully interactive within 3 seconds regardless of thread status
        QTimer.singleShot(3000, lambda: self.tabs.setEnabled(True))

    # --- DIALOG & THREAD MANAGEMENT ---
    def open_db_maintenance(self):
        """Spawns the Database Maintenance & Compaction dialog."""
        dialog = DatabaseMaintenanceDialog(self)
        dialog.exec_()

    def open_task_manager(self):
        """Spawns the COM Process Manager dialog as a non-blocking floating window."""
        if not hasattr(self, 'task_manager_dialog') or not self.task_manager_dialog.isVisible():
            self.task_manager_dialog = ProcessManagerDialog(self)
            self.task_manager_dialog.show()
        else:
            self.task_manager_dialog.activateWindow()

    def on_init_complete(self, success: bool):
        # Always re-enable tabs so non-Visio modules remain fully functional
        self.tabs.setEnabled(True)

        if success:
            self._update_system_status(False, "🟢 System Idle.")
            self.log_message("🚀 System Ready. Paste code or drop files to begin.\n" + "-" * 45)
        else:
            self.batch_tab.set_state("error", "⚠️ Visio unavailable. Visio batch conversions will be disabled.")
            self._update_system_status(False, "🟡 System Ready (Visio Unavailable).")
            self.log_message("⚠️ System initialized with warnings. Visio is not registered, but other modules are ready to use.\n" + "-" * 45)

    # --- AUTO-SAVE LOGIC ---
    def _load_cache(self):
        if self.cache_file.exists():
            try:
                text = self.cache_file.read_text(encoding="utf-8")
                if text.strip():
                    self.code_tab.text_input.setPlainText(text)
                    self.log_message("♻️ Restored main window previous session.", logging.INFO)
            except Exception as e:
                self.log_message(f"⚠️ Could not load previous session: {e}", logging.WARNING)

    def save_cache(self):
        try:
            self.cache_file.write_text(self.code_tab.get_text(), encoding="utf-8")
        except Exception:
            pass

    # --- UI INTERACTION LOGIC ---
    def open_documentation(self):
        try:
            webbrowser.open(HELP_URL)
            self.log_message("🌐 Opened Documentation (README) in your web browser.", logging.INFO)
        except Exception as e:
            self.log_message(f"❌ Failed to open documentation URL: {e}", logging.ERROR)

    def _update_system_status(self, is_processing: bool, status_text: str):
        self.status_bar.showMessage(status_text)
        if is_processing:
            self.batch_tab.set_state(
                "busy", "⚙️ Processing Queue...\n\nPlease wait until finished or drop more files to queue them."
            )
        else:
            self.batch_tab.set_state(
                "ready", "📥 Drag & Drop your .puml, .txt, .pptx, or .vsdx file(s) here\n\n(Batch converts between Visio & PowerPoint)"
            )

    def _set_editor_text(self, text):
        cursor = self.code_tab.text_input.textCursor()
        cursor.beginEditBlock()
        cursor.select(QTextCursor.Document)
        cursor.insertText(text)
        cursor.endEditBlock()

        cursor.setPosition(0)
        self.code_tab.text_input.setTextCursor(cursor)
        self.live_preview.update_now()

    def insert_template(self, selected_template):
        tpl = PLANTUML_TYPES[selected_template]["template"]
        self._set_editor_text(tpl)
        self.log_message(f"📝 Inserted boilerplate for '{selected_template}' diagram.", logging.INFO)

    def clear_editor(self):
        self._set_editor_text("")
        self.log_message("🗑️ Editor cleared.", logging.INFO)

    def copy_editor_code(self):
        code_text = self.code_tab.get_text()
        if code_text:
            QApplication.clipboard().setText(code_text)
            self.log_message("📄 Source code copied to clipboard.", logging.INFO)
        else:
            self.log_message("⚠️ Editor is empty. Nothing to copy.", logging.WARNING)

    def open_export_folder(self):
        export_dir = get_project_root().parent
        try:
            os.startfile(export_dir)
            self.log_message(f"📂 Opened export directory: {export_dir}", logging.INFO)
        except Exception as e:
            self.log_message(f"❌ Failed to open directory: {e}", logging.ERROR)

    def open_template_docs(self, selected_template):
        url = PLANTUML_TYPES[selected_template]["url"]
        try:
            webbrowser.open(url)
            self.log_message(f"🌐 Opened documentation for '{selected_template}'.", logging.INFO)
        except Exception as e:
            self.log_message(f"❌ Failed to open URL: {e}", logging.ERROR)

    def toggle_live_view(self, checked):
        self.live_preview.toggle(checked)

    def open_proxy_settings(self):
        proxy_dialog = ProxyDialog(self)
        if proxy_dialog.exec_() == QDialog.Accepted:
            http_val, https_val = proxy_dialog.get_proxies()
            proxies = {}
            if http_val:
                proxies['http'] = http_val
                os.environ['HTTP_PROXY'] = http_val
            else:
                os.environ.pop('HTTP_PROXY', None)

            if https_val:
                proxies['https'] = https_val
                os.environ['HTTPS_PROXY'] = https_val
            else:
                os.environ.pop('HTTPS_PROXY', None)

            # Configure standard library opener for urllib
            proxy_handler = urllib.request.ProxyHandler(proxies)
            opener = urllib.request.build_opener(proxy_handler)
            urllib.request.install_opener(opener)

            status = "updated" if proxies else "cleared (direct connection)"
            self.log_message(f"✅ Proxy settings {status}. Retrying system checks...", logging.INFO)

            self.batch_tab.set_state("ready", "⏳ Re-initializing system checks...")
            self.status_bar.showMessage("⏳ Re-initializing...")
            self.tabs.setEnabled(False)

            # Re-launch system initialization checks cleanly
            self._launch_init_thread(check_updates=False)

    def check_for_jar_updates(self):
        self.log_message("\n🔄 Initiating manual update check...", logging.INFO)
        self.batch_tab.set_state("ready", "⏳ Checking online for updates...")
        self.status_bar.showMessage("⏳ Checking for updates...")
        self.tabs.setEnabled(False)
        self._launch_init_thread(check_updates=True)

    def extract_code_from_visio(self, file_path):
        self.code_tab.text_input.clear()
        self.code_tab.text_input.setPlaceholderText(
            f"⏳ Extracting source from {Path(file_path).name}...\nPlease wait..."
        )
        self.code_tab.text_input.setEnabled(False)
        self.log_message(f"📂 Reading embedded source from: {Path(file_path).name}")
        self.reader_thread = VisioReaderThread(file_path)
        self.reader_thread.text_extracted.connect(self.on_visio_code_read)
        self.reader_thread.error_occurred.connect(self.on_visio_code_error)
        self.reader_thread.start()

    def on_visio_code_read(self, source_code):
        self.code_tab.text_input.setEnabled(True)
        self._set_editor_text(source_code)
        self.log_message("✅ Successfully extracted PlantUML source from Visio file.")

    def on_visio_code_error(self, error_msg):
        self.code_tab.text_input.setEnabled(True)
        self.code_tab.text_input.setPlaceholderText(
            "Paste PlantUML code OR drop a generated .vsdx file here to extract its source..."
        )
        self.log_message(f"❌ {error_msg}")

    def show_in_planttext(self):
        raw_text = self.code_tab.get_text()
        if not raw_text: return
        try:
            url = f"https://www.planttext.com/?text={encode_plantuml(raw_text)}"
            webbrowser.open(url)
            self.log_message("🌐 Opened code in planttext.com")
        except Exception as e:
            self.log_message(f"❌ Failed to open: {e}")

    def copy_out_path(self):
        if self.last_out_path:
            QApplication.clipboard().setText(self.last_out_path)
            self.log_message(f"📋 Copied to clipboard: {self.last_out_path}")

    def _save_and_queue_pasted_text(self, target_format):
        raw_text = self.code_tab.get_text()
        if not raw_text: return

        base_dir = get_project_root().parent
        timestamp = datetime.datetime.now().strftime("%Y.%m.%d %H-%M-%S")
        base_name = f"{timestamp} diagram"
        puml_path = base_dir / f"{base_name}.puml"

        counter = 1
        while puml_path.exists() or puml_path.with_suffix(".vsdx").exists() or puml_path.with_suffix(".svg").exists():
            puml_path = base_dir / f"{base_name}_{counter}.puml"
            counter += 1

        with open(puml_path, "w", encoding="utf-8") as f:
            f.write(raw_text)

        self.queue_manager.add_item(puml_path, target_format)

    def on_conversion_success(self, out_path: str):
        if out_path == "SPECS_DB_PASS_ONE":
            self.log_message("🔄 Basic files loaded! You can now search while metadata downloads.", logging.INFO)
            if hasattr(self, 'specs_tab'):
                self.specs_tab.refresh_table()
                if hasattr(self.specs_tab, 'set_bg_sync_active'):
                    self.specs_tab.set_bg_sync_active(True)

        elif out_path == "SPECS_DB_PASS_TWO":
            self.log_message("✅ Deep metadata fully updated!", logging.INFO)
            if hasattr(self, 'specs_tab'):
                if hasattr(self.specs_tab, 'set_bg_sync_active'):
                    self.specs_tab.set_bg_sync_active(False)
                else:
                    self.specs_tab.refresh_table()

        elif out_path == "OPENED_IN_PPT":
            self.log_message("👁️ PowerPoint is open with your new slide. You can copy it directly.")

        elif out_path == "SPECS_DB_UPDATED":
            self.log_message("🔄 Reloading specifications table...")
            if hasattr(self, 'specs_tab'):
                self.specs_tab.refresh_table()

        elif out_path and out_path.startswith("MEETINGS_DB"):
            self.log_message("🔄 Reloading meetings table...", logging.INFO)
            if hasattr(self, 'meetings_tab'):
                self.meetings_tab._populate_filters()
                self.meetings_tab.refresh_table()

        elif out_path:
            self.last_out_path = out_path
            self.code_tab.set_copy_path_enabled(True, out_path)

            if out_path.lower().endswith(('.svg', '.txt')):
                try:
                    os.startfile(out_path)
                    ext = Path(out_path).suffix[1:].upper()
                    self.log_message(f"👁️ Opened {ext} in default system viewer.")
                except Exception as e:
                    self.log_message(f"⚠️ Could not automatically open file: {e}")
            else:
                file_name = Path(out_path).name
                self.log_message(f"✅ Successfully saved/downloaded: {file_name}", logging.INFO)

    def log_message(self, message: str, level=logging.INFO):
        logging.log(level, message)

    def showEvent(self, event):
        """Launches background workers independently once the window is rendered."""
        super().showEvent(event)
        if not getattr(self, '_services_started', False):
            self._services_started = True
            logging.info("🏁 [STARTUP:WINDOW] DragDropUI displayed. Scheduling independent background services...")

            # Decoupled launch: Each service starts on its own event-loop tick
            QTimer.singleShot(50, self._start_ollama_monitor)
            QTimer.singleShot(150, self._start_wifi_monitor)
            QTimer.singleShot(300, self._start_system_init_check)

    def _start_background_services(self):
        """Launches all background monitors independently so one never blocks another."""
        logging.info("🏁 [STARTUP:SERVICES] Launching background services...")

        # 1. System JAR & Visio Engine check (independent)
        self._start_system_init_check()

        # 2. Network & 3GPP Wi-Fi Monitor (independent)
        self._start_wifi_monitor()

        # 3. Ollama Local LLM Monitor (independent)
        self._start_ollama_monitor()

        logging.info("🏁 [STARTUP:SERVICES] All background service threads active.")

    def _start_system_init_check(self):
        """Starts Visio / PlantUML engine verification independently."""
        try:
            self._launch_init_thread(check_updates=False)
        except Exception as e:
            logging.error(f"❌ Failed to start InitializationThread: {e}")

    def _start_wifi_monitor(self):
        """Starts Wi-Fi monitor independently."""
        try:
            if self.wifi_monitor is None or not self.wifi_monitor.isRunning():
                logging.info("🏁 [STARTUP:SERVICES] Starting WifiMonitorThread...")
                self.wifi_monitor = WifiMonitorThread(parent=None)
                self.wifi_monitor.status_updated.connect(self._update_network_indicator)
                self.wifi_monitor.start()
        except Exception as e:
            logging.error(f"❌ Failed to start WifiMonitorThread: {e}")
            self.network_indicator.setText("🌐 Offline")

    def _start_ollama_monitor(self):
        """Starts Ollama monitor independently (pure socket check, <1ms)."""
        try:
            if self.ollama_monitor is None or not self.ollama_monitor.isRunning():
                logging.info("🏁 [STARTUP:SERVICES] Starting OllamaMonitorThread...")
                self.ollama_monitor = OllamaMonitorThread(parent=None)
                self.ollama_monitor.status_updated.connect(self._update_ollama_indicator)
                self.ollama_monitor.start()
        except Exception as e:
            logging.error(f"❌ Failed to start OllamaMonitorThread: {e}")
            self.ollama_indicator.setText("🦙 Offline")

    def _update_network_indicator(self, ssid: str, is_3gpp: bool, server_reachable: bool):
        if is_3gpp and server_reachable:
            self.network_indicator.setText(f"🟢 {ssid} (Local Server Active)")
            self.network_indicator.setStyleSheet("color: #2e7d32; font-weight: bold; padding: 0 10px;")
        elif is_3gpp and not server_reachable:
            self.network_indicator.setText(f"🟡 {ssid} (No Local Server)")
            self.network_indicator.setStyleSheet("color: #b8860b; font-weight: bold; padding: 0 10px;")
        else:
            display_name = ssid if ssid else "Offline"
            self.network_indicator.setText(f"🌐 {display_name}")
            self.network_indicator.setStyleSheet("color: gray; padding: 0 10px;")

    def _update_ollama_indicator(self, is_online: bool, selected_model: str, models: list, error: str):
        if is_online:
            model_display = selected_model if selected_model else (models[0] if models else "Connected")
            self.ollama_indicator.setText(f"🦙 {model_display}")
            self.ollama_indicator.setToolTip(f"Ollama Online ({len(models)} models available)\nClick to configure")
            self.ollama_indicator.setStyleSheet("""
                QPushButton {
                    background-color: #F0FDF4;
                    border: 1px solid #BBF7D0;
                    border-radius: 4px;
                    padding: 1px 8px;
                    font-size: 11px;
                    font-weight: bold;
                    color: #15803D;
                }
                QPushButton:hover {
                    background-color: #DCFCE7;
                    border-color: #86EFAC;
                }
            """)
        else:
            self.ollama_indicator.setText("🦙 Offline")
            self.ollama_indicator.setToolTip(f"Ollama Unreachable: {error}\nClick to configure")
            self.ollama_indicator.setStyleSheet("""
                QPushButton {
                    background-color: #FEF2F2;
                    border: 1px solid #FECACA;
                    border-radius: 4px;
                    padding: 1px 8px;
                    font-size: 11px;
                    font-weight: bold;
                    color: #DC2626;
                }
                QPushButton:hover {
                    background-color: #FEE2E2;
                    border-color: #FCA5A5;
                }
            """)

    def open_ollama_settings(self):
        """Opens the Ollama connection and model preferences dialog."""
        dialog = OllamaConfigDialog(self)
        if dialog.exec_() == QDialog.Accepted:
            # Re-read status immediately
            if self.ollama_monitor and self.ollama_monitor.isRunning():
                cfg = dialog.cfg
                self.ollama_monitor.client.reconfigure(cfg.get("host"), cfg.get("proxy_mode"))

    def closeEvent(self, event):
        logging.info("🏁 [SHUTDOWN] Window closeEvent triggered. Saving cache...")
        try:
            self.save_cache()
        except Exception as e:
            logging.warning(f"⚠️ [SHUTDOWN] Could not save cache: {e}")

        # 1. Stop WiFi Monitor independently
        if self.wifi_monitor is not None and self.wifi_monitor.isRunning():
            try:
                self.wifi_monitor.stop()
            except Exception:
                pass

        # 2. Stop Ollama Monitor independently
        if self.ollama_monitor is not None and self.ollama_monitor.isRunning():
            try:
                self.ollama_monitor.stop()
            except Exception:
                pass

        # 3. Stop initialization thread independently
        if hasattr(self, 'init_thread') and self.init_thread is not None and self.init_thread.isRunning():
            try:
                self.init_thread.quit()
                self.init_thread.wait(300)
            except Exception:
                pass

        logging.info("🏁 [SHUTDOWN] Window closing. Releasing Qt window handle...")
        super().closeEvent(event)
        event.accept()