# --- File: src/main_window.py ---
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
    QDialog, QTabWidget, QPushButton, QShortcut, QLabel, QMessageBox, QMenu
)

from core.config.config import HELP_URL
from core.config.useful_links import populate_help_menu, get_useful_links_config_path
from core.network.session import NetworkConfigDialog, NetworkSession, ProxyProfileManager
from core.network.wifi_monitor import WifiMonitorThread
from core.queue_manager import QueueManager
from core.ui.ui_components import ProxyDialog, create_app_icon
from core.ui.ui_panels import (
    ConsolePanel, QueuePanel, ProcessManagerDialog, DatabaseMaintenanceDialog, GuiLogHandler
)
from core.ui.system_status_bar import SystemStatusBar
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
from modules.assistant.ui.assistant_tab import AssistantTab
from core.ai.ollama_client import OllamaClient, OllamaMonitorThread
from core.ai.ollama_dialog import OllamaConfigDialog
from modules.puml2visio.ui.java_dialog import JavaMaintenanceDialog


class DragDropUI(QMainWindow):
    def __init__(self, gui_log_handler=None):
        super().__init__()
        self.setWindowTitle("3GPP Delegate Tools")
        self.setWindowIcon(create_app_icon())
        self.resize(950, 750)

        self.jar_path = get_puml2visio_asset_path(PLANTUML_JAR_NAME)
        self.last_out_path = ""
        self.init_thread = None
        # Background monitors are managed by self.status_bar
        self._services_started = False  # Guard to prevent duplicate background starts
        self._visio_ready = True
        self._core_ready = True

        logging.info("🏁 [STARTUP:UI] Starting _setup_ui()...")
        self._setup_ui()
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

        # Wire Queue Manager
        logging.info("🏁 [STARTUP:QUEUE] Initializing QueueManager...")
        app_context = {"jar_path": self.jar_path}
        self.queue_manager = QueueManager(app_context=app_context)
        self.queue_manager.queue_updated.connect(self.queue_panel.update_list)
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
        self.code_tab.redo_requested.connect(self.code_tab.text_input.redo)
        self.code_tab.copy_code_requested.connect(self.copy_editor_code)
        self.code_tab.live_view_toggled.connect(self.toggle_live_view)
        self.code_tab.browser_preview_requested.connect(lambda: self.live_preview.open_in_browser())
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
        self.batch_tab.send_to_editor_requested.connect(self._on_send_code_to_editor)
        self.batch_tab.insert_snippet_requested.connect(self._on_insert_snippet_from_visio)
        self.batch_tab.java_maintenance_requested.connect(self.open_java_maintenance)
        self.batch_tab.open_folder_requested.connect(self.open_export_folder)

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
        self.word_tab.convert_markdown_requested.connect(
            lambda doc, rev, media, out_dir, hdr: self.queue_manager.add_item(
                Path(doc),
                "word_to_markdown",
                {"fmt": "md", "revisions": rev, "extract_media": media, "output_dir": out_dir, "include_header": hdr}
            )
        )

        db_path = get_project_root() / "db" / "3gpp_data.db"
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

        nas_db_path = get_project_root() / "db" / "3gpp_protocol_data.db"
        logging.info("🏁 [STARTUP:UI:TAB] 7/9 Initializing NASTab...")
        self.nas_tab = NASTab(nas_db_path, db_path)

        spec_search_db_path = get_project_root() / "db" / "3gpp_spec_search.db"
        logging.info("🏁 [STARTUP:UI:TAB] 8/9 Initializing SpecSearchTab...")
        self.spec_search_tab = SpecSearchTab(spec_search_db_path, db_path)

        # Wire bidirectional deep-linking between Specifications and Spec Search tabs
        self.spec_search_tab.navigate_to_specs_tab_requested.connect(self._navigate_from_search_to_specs)
        self.specs_tab.search_in_spec_search_requested.connect(self._navigate_from_specs_to_search)
        self.specs_tab.index_in_spec_search_requested.connect(self._index_spec_from_specs_tab)
        self.specs_tab.batch_index_in_spec_search_requested.connect(self._batch_index_specs_from_specs_tab)

        # Wire deep-linking from Protocols tab to Specifications and Spec Search tabs
        self.nas_tab.navigate_to_specs_tab_requested.connect(self._navigate_from_search_to_specs)
        self.nas_tab.search_in_spec_search_requested.connect(self._navigate_from_specs_to_search)

        logging.info("🏁 [STARTUP:UI:TAB] 9/9 Initializing AssistantTab...")
        self.assistant_tab = AssistantTab(
            specs_db_path=db_path,
            spec_search_db_path=spec_search_db_path,
            protocol_db_path=nas_db_path,
        )
        self.assistant_tab.open_ollama_settings_requested.connect(self.open_ollama_settings)
        self.assistant_tab.open_spec_search_requested.connect(self._navigate_from_assistant_to_spec_search)
        self.assistant_tab.open_specs_requested.connect(self._navigate_from_assistant_to_specs)
        self.assistant_tab.open_protocol_requested.connect(self._navigate_from_assistant_to_protocols)

        self.tabs.addTab(self.code_tab, "📝 PlantUML")
        self.tabs.addTab(self.batch_tab, "🔄 Visio && Office")
        self.tabs.addTab(self.word_tab, "📘 Word")
        self.tabs.addTab(self.specs_tab, "📚 Specifications")
        self.tabs.addTab(self.spec_search_tab, "🔎 Spec Search")
        self.tabs.addTab(self.work_items_tab, "📋 Work Items")
        self.tabs.addTab(self.meetings_tab, "🗓️ Meetings")
        self.tabs.addTab(self.nas_tab, "🔬 Protocols")
        self.tabs.addTab(self.assistant_tab, "🤖 Assistant")

        # Connect tab activation for lazy loading
        self.tabs.currentChanged.connect(self._on_tab_changed)

        # Tab corner help dropdown menu
        self.help_btn = QPushButton("📖 Help ▾")
        self.help_btn.setObjectName("helpDropdownBtn")
        self.help_btn.setStyleSheet("""
            QPushButton#helpDropdownBtn {
                border: 1px solid #CBD5E1;
                background-color: #F8FAFC;
                color: #1E293B;
                font-weight: 600;
                font-size: 12px;
                padding: 0px 14px;
                min-height: 26px;
                max-height: 26px;
                border-radius: 5px;
                margin-right: 8px;
                margin-top: 3px;
                margin-bottom: 4px;
            }
            QPushButton#helpDropdownBtn:hover {
                background-color: #EEF2F6;
                border-color: #94A3B8;
                color: #0F172A;
            }
            QPushButton#helpDropdownBtn::menu-indicator {
                image: none;
                width: 0px;
            }
        """)
        self.help_btn.setCursor(Qt.PointingHandCursor)
        self.help_btn.setToolTip("Useful 3GPP tools, external portals, and application documentation (F1 for Docs)")

        # Create dynamic Help & Resources dropdown menu
        self.help_menu = QMenu(self)
        self.help_menu.setToolTipsVisible(True)
        self.help_menu.aboutToShow.connect(self._rebuild_help_menu)
        self.help_btn.setMenu(self.help_menu)
        self._rebuild_help_menu()

        self.tabs.setCornerWidget(self.help_btn, Qt.TopRightCorner)

        self.help_shortcut = QShortcut(QKeySequence("F1"), self)
        self.help_shortcut.activated.connect(self.open_documentation)

        self.splitter.addWidget(self.tabs)

        # --- BOTTOM HALF: PANELS ---
        logging.info("🏁 [STARTUP:UI:PANELS] Initializing ConsolePanel and QueuePanel...")
        self.bottom_splitter = QSplitter(Qt.Horizontal)

        # Wire Console Panel buttons in _setup_ui():
        self.console_panel = ConsolePanel()
        self.console_panel.proxy_requested.connect(self.open_proxy_settings)
        self.console_panel.network_config_requested.connect(lambda: NetworkConfigDialog(self).exec_())
        self.console_panel.java_info_requested.connect(self.open_java_maintenance)
        self.console_panel.update_requested.connect(self.open_java_maintenance)
        self.console_panel.task_manager_requested.connect(self.open_task_manager)
        self.console_panel.db_maintenance_requested.connect(self.open_db_maintenance)

        self.queue_panel = QueuePanel()

        self.bottom_splitter.addWidget(self.console_panel)
        self.bottom_splitter.addWidget(self.queue_panel)
        self.bottom_splitter.setSizes([650, 250])

        self.splitter.addWidget(self.bottom_splitter)
        self.splitter.setSizes([700, 100])
        main_layout.addWidget(self.splitter)

        central_widget.setLayout(main_layout)
        self.setCentralWidget(central_widget)

        # Bottom System Status Bar (decoupled component)
        self.status_bar = SystemStatusBar(self)
        self.setStatusBar(self.status_bar)
        self.status_bar.reinit_requested.connect(self._on_status_reinit_requested)
        self.status_bar.update_proxy_indicator(self.console_panel)

        logging.info("🏁 [STARTUP:UI] _setup_ui() layout complete.")

    def _on_tab_changed(self, index: int):
        """Triggers lazy loading when an inactive tab is selected for the first time."""
        active_widget = self.tabs.widget(index)
        if hasattr(active_widget, 'ensure_loaded'):
            active_widget.ensure_loaded()

    def _navigate_from_search_to_specs(self, spec_number: str, version: str):
        """Switches to Specifications tab and selects the requested spec and version."""
        self.tabs.setCurrentWidget(self.specs_tab)
        self.specs_tab.select_and_show_spec(spec_number, version)

    def _navigate_from_specs_to_search(self, spec_number: str, version: str):
        """Switches to Spec Search tab and focuses on the requested spec and version."""
        self.tabs.setCurrentWidget(self.spec_search_tab)
        self.spec_search_tab.focus_spec(spec_number, version)

    def _navigate_from_assistant_to_spec_search(self, spec_number: str, version: str, clause: str):
        """Switches to Spec Search tab and loads the requested spec, version, and clause."""
        self.tabs.setCurrentWidget(self.spec_search_tab)
        if spec_number:
            self.spec_search_tab.focus_spec(spec_number, version)
        if clause and hasattr(self.spec_search_tab, "clause_filter_input"):
            self.spec_search_tab.clause_filter_input.setText(clause)

    def _navigate_from_assistant_to_specs(self, spec_number: str, version: str):
        """Switches to Specifications tab and selects the requested spec and version."""
        self.tabs.setCurrentWidget(self.specs_tab)
        self.specs_tab.select_and_show_spec(spec_number, version)

    def _navigate_from_assistant_to_protocols(self, protocol: str, message: str):
        """Switches to Protocols tab and selects the requested message."""
        self.tabs.setCurrentWidget(self.nas_tab)
        if message and hasattr(self.nas_tab, "_jump_to_message"):
            self.nas_tab._jump_to_message(message)

    def _index_spec_from_specs_tab(self, spec_number: str, version: str):
        """Directly indexes a single specification version from Specifications tab into Spec Search."""
        self._batch_index_specs_from_specs_tab([(spec_number, version)])

    def _batch_index_specs_from_specs_tab(self, spec_items: list):
        """Directly indexes one or more specification versions from Specifications tab into Spec Search."""
        from modules.nas.core.parsing.protocol_parser_common import ProtocolDocxDispatcher
        from modules.spec_search.core.spec_search_threads import is_change_mark_file

        if not spec_items:
            return

        tasks = []
        for spec_number, version in spec_items:
            clean_spec = str(spec_number).strip()
            clean_ver = str(version).lstrip("vV").strip()

            # 1. Exact lookup via get_spec_files
            files = self.specs_tab.db.get_spec_files(clean_spec)
            matching_file = None
            if clean_ver:
                matching_file = next(
                    (f for f in files if str(f.get("version", "")).lstrip("vV").strip() == clean_ver),
                    None
                )
            elif files:
                matching_file = files[0]
                clean_ver = str(matching_file.get("version", "")).lstrip("vV").strip()

            filename = ""
            file_url = ""
            release_date = ""

            if matching_file:
                filename = matching_file.get("filename", "")
                file_url = matching_file.get("url", "")
                release_date = matching_file.get("upload_date", "")
            else:
                rows = self.specs_tab.db.search_files(spec_number=clean_spec, release_version=clean_ver)
                if rows:
                    filename = rows[0][4]
                    file_url = rows[0][6]
                    release_date = rows[0][7]
                    if not clean_ver and len(rows[0]) > 5:
                        clean_ver = str(rows[0][5]).lstrip("vV").strip()

            if not filename and not file_url:
                continue

            task = {
                "spec_number": clean_spec,
                "version": clean_ver,
                "filename": filename,
                "file_url": file_url,
                "release_date": release_date,
            }

            # Check if Word document is already extracted locally in download_dir
            spec_dir = Path(self.specs_tab.download_dir) / clean_spec
            stem = filename.replace(".zip", "")
            if spec_dir.exists() and stem:
                local_docs = [
                    f for f in spec_dir.iterdir()
                    if f.is_file()
                    and f.stem.lower().startswith(stem.lower())
                    and f.suffix.lower() in [".docx", ".doc"]
                    and not f.name.startswith("~$")
                    and not is_change_mark_file(f.name)
                ]
                if local_docs:
                    local_docs.sort(key=ProtocolDocxDispatcher._extract_part_index)
                    task["local_docx_paths"] = local_docs

            tasks.append(task)

        if not tasks:
            QMessageBox.warning(
                self,
                "Spec Not Found",
                "Could not find specification metadata or archive records for the requested version(s)."
            )
            return

        self.tabs.setCurrentWidget(self.spec_search_tab)
        self.spec_search_tab._start_batch_ingestion(tasks)

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
        """Spawns the Database Maintenance & Compaction dialog as a non-blocking modeless window."""
        if not hasattr(self, 'db_maintenance_dialog') or not self.db_maintenance_dialog.isVisible():
            self.db_maintenance_dialog = DatabaseMaintenanceDialog(self)
            self.db_maintenance_dialog.show()
        else:
            self.db_maintenance_dialog.raise_()
            self.db_maintenance_dialog.activateWindow()

    def open_task_manager(self):
        """Spawns the COM Process Manager dialog as a non-blocking floating window."""
        if not hasattr(self, 'task_manager_dialog') or not self.task_manager_dialog.isVisible():
            self.task_manager_dialog = ProcessManagerDialog(self)
            self.task_manager_dialog.show()
        else:
            self.task_manager_dialog.activateWindow()

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
    def _rebuild_help_menu(self):
        populate_help_menu(
            self.help_menu,
            on_open_url=self.open_external_url,
            on_edit_config=self.edit_useful_links_config,
        )

    def open_external_url(self, url: str, title: str = "Web Resource"):
        try:
            webbrowser.open(url)
            self.log_message(f"🌐 Opened {title} in your web browser: {url}", logging.INFO)
        except Exception as e:
            self.log_message(f"❌ Failed to open URL '{url}': {e}", logging.ERROR)

    def edit_useful_links_config(self):
        config_path = get_useful_links_config_path()
        try:
            if hasattr(os, "startfile"):
                os.startfile(str(config_path))
            else:
                webbrowser.open(f"file:///{config_path}")
            self.log_message(f"📝 Opened useful links configuration: {config_path.name}", logging.INFO)
        except Exception as e:
            self.log_message(f"❌ Failed to open config file '{config_path}': {e}", logging.ERROR)

    def open_documentation(self):
        try:
            webbrowser.open(HELP_URL)
            self.log_message("🌐 Opened Documentation (README) in your web browser.", logging.INFO)
        except Exception as e:
            self.log_message(f"❌ Failed to open documentation URL: {e}", logging.ERROR)

    def _update_system_status(self, is_processing: bool, status_text: str):
        self.status_bar.showMessage(status_text)
        if hasattr(self, 'word_tab') and hasattr(self.word_tab, 'on_processing_state_changed'):
            self.word_tab.on_processing_state_changed(is_processing)
        if is_processing:
            self.batch_tab.set_state(
                "busy", "⚙️ Processing Queue...\n\nPlease wait until finished or drop more files to queue them."
            )
        else:
            if not getattr(self, "_core_ready", True):
                self.batch_tab.set_state("error", "❌ Core engine error. Java or PlantUML engine is unavailable.")
            elif not getattr(self, "_visio_ready", True):
                self.batch_tab.set_state("error", "⚠️ Visio unavailable. Visio batch conversions will be disabled.")
            else:
                self.batch_tab.set_state(
                    "ready", "📥 Drag & Drop your .puml, .txt, .pptx, or .vsdx file(s) here\n\n(Batch converts between Visio & PowerPoint)"
                )

    def _on_send_code_to_editor(self, code: str):
        self.tabs.setCurrentWidget(self.code_tab)
        self._set_editor_text(code)
        self.log_message("📝 Loaded extracted PlantUML code into Editor tab.", logging.INFO)

    def _on_insert_snippet_from_visio(self, snippet):
        self.tabs.setCurrentWidget(self.code_tab)
        if hasattr(self.code_tab, "text_input"):
            self.code_tab.text_input.insert_snippet(snippet)
        label = getattr(snippet, 'label', 'snippet')
        self.log_message(f"🧩 Inserted '{label}' into PlantUML editor.", logging.INFO)

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
        curr_text = self.code_tab.get_text()
        if curr_text and curr_text.strip() != tpl.strip():
            reply = QMessageBox.question(
                self,
                "Replace Diagram?",
                f"The editor already contains PlantUML code.\n\n"
                f"Do you want to replace it completely with the '{selected_template}' boilerplate?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No
            )
            if reply != QMessageBox.Yes:
                return

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
        export_dir = get_project_root().parent / "output"
        export_dir.mkdir(parents=True, exist_ok=True)
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
        open_browser = self.code_tab.live_view_btn.isChecked()
        self.live_preview.toggle(checked, open_browser=open_browser)

    def _on_status_reinit_requested(self):
        self.log_message("✅ Proxy configuration updated. Retrying system checks...", logging.INFO)
        self.batch_tab.set_state("ready", "⏳ Re-initializing system checks...")
        self.status_bar.showMessage("⏳ Re-initializing...")
        self.tabs.setEnabled(False)
        self._launch_init_thread(check_updates=False)

    def open_proxy_settings(self):
        self.status_bar.open_proxy_settings()


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

        params = {}
        if target_format == "pptx_append":
            from modules.puml2visio.core.powerpoint_converter import get_open_powerpoint_presentations
            from modules.puml2visio.ui.select_presentation_dialog import SelectPresentationDialog

            open_presentations = get_open_powerpoint_presentations()
            if not open_presentations:
                reply = QMessageBox.question(
                    self,
                    "No Open PowerPoint Presentation",
                    "No open Microsoft PowerPoint presentation could be detected.\n\n"
                    "Would you like to export to a new PowerPoint presentation instead?",
                    QMessageBox.Yes | QMessageBox.No,
                    QMessageBox.Yes
                )
                if reply == QMessageBox.Yes:
                    target_format = "pptx"
                else:
                    return
            elif len(open_presentations) == 1:
                target_pres = open_presentations[0]
                params = {
                    "target_presentation": target_pres["name"],
                    "insert_at": "end"
                }
            else:
                dialog = SelectPresentationDialog(open_presentations, parent=self)
                if dialog.exec_() != QDialog.Accepted:
                    return
                if getattr(dialog, "is_new_presentation_requested", lambda: False)():
                    target_format = "pptx"
                else:
                    chosen = dialog.get_selected_presentation()
                    if not chosen:
                        return
                    params = {
                        "target_presentation": chosen["name"],
                        "insert_at": chosen.get("insert_at", "end")
                    }

        out_dir = get_project_root().parent / "output"
        out_dir.mkdir(parents=True, exist_ok=True)
        timestamp = datetime.datetime.now().strftime("%Y.%m.%d %H-%M-%S")
        base_name = f"{timestamp} diagram"
        puml_path = out_dir / f"{base_name}.puml"

        counter = 1
        while (
            puml_path.exists()
            or puml_path.with_suffix(".vsdx").exists()
            or puml_path.with_suffix(".pptx").exists()
            or puml_path.with_suffix(".svg").exists()
        ):
            puml_path = out_dir / f"{base_name}_{counter}.puml"
            counter += 1

        with open(puml_path, "w", encoding="utf-8") as f:
            f.write(raw_text)

        self.queue_manager.add_item(puml_path, target_format, params=params)

    def on_conversion_failed(self, error_msg: str):
        if hasattr(self, 'word_tab') and hasattr(self.word_tab, 'on_conversion_failed'):
            self.word_tab.on_conversion_failed(error_msg)

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
                # Only Phase 1 can discover/add meeting folders and therefore
                # potentially change the available WG choices. Phases 2 and 3
                # update Docs/DynaReport metadata only.
                if out_path == "MEETINGS_DB_PHASE_1":
                    self.meetings_tab._populate_filters()
                self.meetings_tab.refresh_table()

                # Phase 2 updates the Docs-derived first/last TDoc ranges.
                # Re-evaluate an unchanged Quick TDoc Jump immediately so a
                # newly resolved meeting appears without requiring retyping.
                if out_path == "MEETINGS_DB_PHASE_2":
                    self.meetings_tab.refresh_quick_tdoc_lookup()

        elif out_path:
            self.last_out_path = out_path
            self.code_tab.set_copy_path_enabled(True, out_path)

            if hasattr(self, 'word_tab') and hasattr(self.word_tab, 'on_conversion_success'):
                self.word_tab.on_conversion_success(out_path)

            if hasattr(self, 'batch_tab') and hasattr(self.batch_tab, 'on_conversion_success'):
                self.batch_tab.on_conversion_success(out_path)

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
        super().showEvent(event)

        if not self._services_started:
            self._services_started = True

            logging.info(
                "🏁 [STARTUP:WINDOW] DragDropUI displayed. "
                "Scheduling background services..."
            )

            QTimer.singleShot(0, self._start_background_services)

    def _start_background_services(self):
        """Launches all background monitors independently so one never blocks another."""
        logging.info("🏁 [STARTUP:SERVICES] Launching background services...")

        logging.info("🏁 [STARTUP:SERVICES] 1/3 Starting system init...")
        self._start_system_init_check()
        logging.info("🏁 [STARTUP:SERVICES] 1/3 System init launched.")

        logging.info("🏁 [STARTUP:SERVICES] 2/3 Starting Ollama monitor...")
        self._start_ollama_monitor()
        logging.info("🏁 [STARTUP:SERVICES] 2/3 Ollama monitor launched.")

        logging.info("🏁 [STARTUP:SERVICES] 3/3 Starting Wi-Fi monitor...")
        self._start_wifi_monitor()
        logging.info("🏁 [STARTUP:SERVICES] 3/3 Wi-Fi monitor launched.")

        logging.info("🏁 [STARTUP:SERVICES] All background service threads active.")

    def _start_system_init_check(self):
        """Starts Visio / PlantUML engine verification independently."""
        try:
            self._launch_init_thread(check_updates=False)
        except Exception as e:
            logging.error(f"❌ Failed to start InitializationThread: {e}")

    def _start_wifi_monitor(self):
        self.status_bar.start_wifi_monitor()

    def _start_ollama_monitor(self):
        self.status_bar.start_ollama_monitor()

    def open_ollama_settings(self):
        self.status_bar.open_ollama_settings()


    def closeEvent(self, event):
        logging.info("🏁 [SHUTDOWN] Window closeEvent triggered. Saving cache...")
        try:
            self.save_cache()
        except Exception as e:
            logging.warning(f"⚠️ [SHUTDOWN] Could not save cache: {e}")

        self.status_bar.stop_monitors()

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

        # 8. Shut down remaining feature tabs if they define shutdown or close.
        for tab_name in ("spec_search_tab", "work_items_tab", "meetings_tab"):
            tab = getattr(self, tab_name, None)
            if tab is not None:
                try:
                    if hasattr(tab, "shutdown") and callable(tab.shutdown):
                        tab.shutdown()
                    elif hasattr(tab, "close") and callable(tab.close):
                        tab.close()
                except Exception as e:
                    logging.warning(f"⚠️ [SHUTDOWN] Could not stop {tab_name} cleanly: {e}")

        logging.info("🏁 [SHUTDOWN] Window closing. Releasing Qt window handle...")
        super().closeEvent(event)
        event.accept()

    def display_java_info(self):
        """Inspects and logs full Java runtime details to the terminal output."""
        from core.utils.utils import get_detailed_java_info
        from modules.puml2visio.utils.utils import get_plantuml_version

        info = get_detailed_java_info()
        self.log_message("\n" + "=" * 55)
        if not info["installed"]:
            self.log_message("❌ Java Check: No working Java runtime found!", logging.ERROR)
            self.log_message(
                "   Please install Java (JRE/JDK 11+ recommended) from https://adoptium.net or oracle.com.",
                logging.ERROR
            )
        else:
            self.log_message("☕ Java Runtime Diagnostics:", logging.INFO)
            self.log_message(f"   • Version:      Java {info['major']} ({info['version_str']})", logging.INFO)
            self.log_message(f"   • Architecture: {info['arch']}", logging.INFO)
            self.log_message(f"   • VM / Vendor:  {info['vendor']}", logging.INFO)
            self.log_message(f"   • Executable:   {info['path']}", logging.INFO)

            ver_file = self.jar_path.with_suffix('.version')
            req_type = "modern" if info['major'] >= 11 else "legacy"
            cur_type = ver_file.read_text(encoding="utf-8").strip() if ver_file.exists() else "unspecified"
            self.log_message(f"   • Profile Mode: {req_type.capitalize()} (Active profile: {cur_type})", logging.INFO)
            if self.jar_path.exists():
                jar_v = get_plantuml_version(info['path'], self.jar_path)
                self.log_message(f"   • PlantUML JAR: Version {jar_v or 'Detected'} ({self.jar_path.name})",
                                 logging.INFO)
            else:
                self.log_message(f"   • PlantUML JAR: ⚠️ Missing ({self.jar_path.name})", logging.WARNING)
        self.log_message("=" * 55 + "\n")

    def on_init_complete(self, core_ready: bool, visio_ready: bool = True):
        # Always re-enable tabs so non-Visio modules remain fully functional
        self._core_ready = core_ready
        self._visio_ready = visio_ready
        self.tabs.setEnabled(True)

        if core_ready and visio_ready:
            self.batch_tab.set_state(
                "ready",
                "📥 Drag & Drop your .puml, .txt, .pptx, or .vsdx file(s) here\n\n(Batch converts between Visio & PowerPoint)"
            )
            self._update_system_status(False, "🟢 System Idle.")
            self.log_message("🚀 System Ready. Paste code or drop files to begin.\n" + "-" * 45)
        elif core_ready and not visio_ready:
            self.batch_tab.set_state("error", "⚠️ Visio unavailable. Visio batch conversions will be disabled.")
            self._update_system_status(False, "🟡 System Ready (Visio Unavailable).")
            self.log_message(
                "⚠️ System initialized: Visio is not registered (Visio batch conversions disabled), "
                "but PlantUML, Specs, Protocols, and Word tools are ready to use.\n" + "-" * 45
            )
        else:
            self.batch_tab.set_state("error", "❌ Core engine error. Java or PlantUML engine is unavailable.")
            self._update_system_status(False, "🔴 Engine Error (Java Unavailable).")
            self.log_message(
                "❌ System initialized with errors. Java runtime was not detected. Please verify your Java installation.\n" + "-" * 45
            )

    def open_java_maintenance(self):
        """Spawns the consolidated Java Runtime & PlantUML Maintenance dialog."""
        dialog = JavaMaintenanceDialog(self.jar_path, parent=self)
        dialog.configuration_changed.connect(lambda: self._launch_init_thread(check_updates=False))
        dialog.exec_()