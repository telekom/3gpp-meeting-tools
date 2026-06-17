import logging
import datetime
import urllib.request
import webbrowser
import os
from pathlib import Path

from PyQt5.QtWidgets import QMainWindow, QWidget, QVBoxLayout, QSplitter, QStatusBar, QApplication, QDialog, QTabWidget, QPushButton, QShortcut
from PyQt5.QtGui import QKeySequence
from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtGui import QTextCursor

from core.ui.ui_components import ProxyDialog
from modules.puml2visio.ui.ui_tabs import CodeEditorTab, BatchConvertTab
from core.queue_manager import QueueManager

from core.config.config import HELP_URL
from core.utils.paths import get_project_root
from core.ui.ui_panels import ConsolePanel, QueuePanel, ProcessManagerDialog

from modules.puml2visio.utils.utils import encode_plantuml, InitializationThread
from modules.puml2visio.config.paths import PLANTUML_JAR_NAME
from modules.puml2visio.core.visio_converter import VisioReaderThread
from modules.puml2visio.core.live_preview import LivePreviewManager
from modules.puml2visio.templates.plantuml_templates import PLANTUML_TYPES
from modules.puml2visio.utils.paths import get_puml2visio_asset_path
from modules.specifications.ui.ui_tabs import SpecificationsTab

from modules.word_tools.ui.word_tabs import WordExtractorTab


class DragDropUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PlantUML to Visio Converter (3GPP)")
        self.resize(950, 750)

        self.jar_path = get_puml2visio_asset_path(PLANTUML_JAR_NAME)
        self.last_out_path = ""

        self._setup_ui()

        # --- Wire the Queue Manager directly to the UI ---
        app_context = {
            "jar_path": self.jar_path
        }
        self.queue_manager = QueueManager(app_context=app_context)
        self.queue_manager.log_msg.connect(self.log_message)
        self.queue_manager.queue_updated.connect(self.queue_panel.update_list)
        self.queue_manager.processing_state_changed.connect(self._update_system_status)
        self.queue_manager.conversion_success.connect(self.on_conversion_success)

        self.queue_panel.remove_requested.connect(self.queue_manager.remove_items)
        self.queue_panel.clear_requested.connect(self.queue_manager.clear_queue)

        self.cache_file = get_project_root().parent / ".editor_cache.puml"
        self._load_cache()

        self.save_timer = QTimer()
        self.save_timer.setSingleShot(True)
        self.save_timer.setInterval(2000)
        self.save_timer.timeout.connect(self.save_cache)
        self.code_tab.text_input.textChanged.connect(self.save_timer.start)

        self.live_preview = LivePreviewManager(self.code_tab.text_input, self.jar_path)
        self.live_preview.log_msg.connect(self.log_message)

        self._launch_init_thread(check_updates=False)

    def _setup_ui(self):
        central_widget = QWidget()
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(15, 15, 15, 15)

        self.splitter = QSplitter(Qt.Vertical)

        # --- TOP HALF: TABS ---
        self.tabs = QTabWidget()

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

        self.batch_tab = BatchConvertTab()
        self.batch_tab.files_dropped.connect(lambda paths: self.queue_manager.add_batch(paths))

        self.word_tab = WordExtractorTab()

        self.word_tab.extract_visio_requested.connect(
            lambda fp: self.queue_manager.add_item(Path(fp), "extract_visio")
        )
        self.word_tab.split_doc_requested.connect(
            lambda fp, pref, depth: self.queue_manager.add_item(Path(fp), "split_docx",
                                                                {"prefix": pref, "depth": depth})
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

        db_path = get_project_root().parent / "3gpp_specs.db"

        # Pass the db_path into the tab so the table can read it!
        self.specs_tab = SpecificationsTab(db_path)

        # Existing Full Update Signal
        self.specs_tab.update_db_requested.connect(
            lambda force_meta: self.queue_manager.add_item(
                Path("3GPP_Archive"),
                "update_specs_db",
                {
                    "db_path": db_path,
                    "force_metadata": force_meta,
                    "target_specs": []  # Empty list triggers full update
                }
            )
        )

        # NEW: Targeted Update Signal from Right-Click Menu
        self.specs_tab.update_specific_requested.connect(
            lambda target_specs, force_meta: self.queue_manager.add_item(
                Path("3GPP_Archive_Targeted"),
                "update_specs_db",
                {
                    "db_path": db_path,
                    "force_metadata": force_meta,
                    "target_specs": target_specs  # Pass the targeted list!
                }
            )
        )

        self.tabs.addTab(self.code_tab, "📝 Code Editor")
        self.tabs.addTab(self.batch_tab, "📂 Batch Convert")
        self.tabs.addTab(self.word_tab, "📄 Word Tools")
        self.tabs.addTab(self.specs_tab, "📚 3GPP Specifications")
        self.tabs.setEnabled(False)

        # --- TAB CORNER HELP WIDGET ---
        self.help_btn = QPushButton("📖 Help (F1)")
        # Make it look like a clean, borderless link instead of a chunky button
        self.help_btn.setStyleSheet("""
                    QPushButton { border: none; background: transparent; color: #395396; font-weight: bold; padding: 4px 15px; }
                    QPushButton:hover { color: #1E5C99; text-decoration: underline; }
                """)
        self.help_btn.setCursor(Qt.PointingHandCursor)
        self.help_btn.clicked.connect(self.open_documentation)

        # Inject it into the top right corner of the Tab Bar
        self.tabs.setCornerWidget(self.help_btn, Qt.TopRightCorner)

        # Restore the F1 Global Shortcut
        self.help_shortcut = QShortcut(QKeySequence("F1"), self)
        self.help_shortcut.activated.connect(self.open_documentation)

        self.splitter.addWidget(self.tabs)

        # --- BOTTOM HALF: PANELS ---
        self.bottom_splitter = QSplitter(Qt.Horizontal)

        self.console_panel = ConsolePanel()
        self.console_panel.proxy_requested.connect(self.open_proxy_settings)
        self.console_panel.update_requested.connect(self.check_for_jar_updates)

        self.console_panel.task_manager_requested.connect(self.open_task_manager)

        self.queue_panel = QueuePanel()

        self.bottom_splitter.addWidget(self.console_panel)
        self.bottom_splitter.addWidget(self.queue_panel)
        self.bottom_splitter.setSizes([650, 250])

        self.splitter.addWidget(self.bottom_splitter)
        self.splitter.setSizes([700, 100])
        main_layout.addWidget(self.splitter)

        central_widget.setLayout(main_layout)
        self.setCentralWidget(central_widget)

        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("⏳ Initializing...")

    # --- THREAD MANAGEMENT ---
    def _launch_init_thread(self, check_updates=False):
        self.init_thread = InitializationThread(self.jar_path, check_updates=check_updates)
        self.init_thread.ui_log_msg.connect(self.log_message)
        self.init_thread.init_complete.connect(self.on_init_complete)
        self.init_thread.network_error.connect(self.open_proxy_settings)
        self.init_thread.start()

    def on_init_complete(self, success: bool):
        if success:
            self.tabs.setEnabled(True)
            self._update_system_status(False, "🟢 System Idle.")
            self.log_message("🚀 System Ready. Paste code or drop files to begin.\n" + "-" * 45)
        else:
            self.batch_tab.set_state("error", "❌ Initialization Failed.")
            self.status_bar.showMessage("❌ Initialization Failed. Check log for details.")

    # --- AUTO-SAVE LOGIC ---
    def _load_cache(self):
        if self.cache_file.exists():
            try:
                text = self.cache_file.read_text(encoding="utf-8")
                if text.strip():
                    self.code_tab.text_input.setPlainText(text)
                    self.log_message("♻️ Restored previous session.", logging.INFO)
            except Exception as e:
                self.log_message(f"⚠️ Could not load previous session: {e}", logging.WARNING)

    def save_cache(self):
        try:
            self.cache_file.write_text(self.code_tab.get_text(), encoding="utf-8")
        except Exception:
            pass

    def closeEvent(self, event):
        self.save_cache()
        super().closeEvent(event)

    # --- UI INTERACTION LOGIC ---
    def open_documentation(self):
        """Opens the configured README URL in the default web browser."""
        try:
            webbrowser.open(HELP_URL)
            self.log_message("🌐 Opened Documentation (README) in your web browser.", logging.INFO)
        except Exception as e:
            self.log_message(f"❌ Failed to open documentation URL: {e}", logging.ERROR)

    def _update_system_status(self, is_processing: bool, status_text: str):
        """Called dynamically by the QueueManager."""
        self.status_bar.showMessage(status_text)
        if is_processing:
            self.batch_tab.set_state("busy",
                                     "⚙️ Processing Queue...\n\nPlease wait until finished or drop more files to queue them.")
        else:
            self.batch_tab.set_state("ready",
                                     "📥 Drag & Drop your .puml or .txt file(s) here\n\n(Batch exports as Visio files)")

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
        proxy_dialog = ProxyDialog()
        if proxy_dialog.exec_() == QDialog.Accepted:
            http_val, https_val = proxy_dialog.get_proxies()
            proxies = {}
            if http_val: proxies['http'] = http_val
            if https_val: proxies['https'] = https_val

            proxy_handler = urllib.request.ProxyHandler(proxies)
            opener = urllib.request.build_opener(proxy_handler)
            urllib.request.install_opener(opener)

            status = "updated" if proxies else "cleared (direct connection)"
            self.log_message(f"✅ Proxy settings {status}. Retrying system checks...", logging.INFO)

            self.batch_tab.set_state("ready", "⏳ Re-initializing system checks...")
            self.status_bar.showMessage("⏳ Re-initializing...")
            self.tabs.setEnabled(False)

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
            f"⏳ Extracting source from {Path(file_path).name}...\nPlease wait...")
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
            "Paste PlantUML code OR drop a generated .vsdx file here to extract its source...")
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

    def open_task_manager(self):
        """Spawns the process manager as a non-blocking floating window."""
        if not hasattr(self, 'task_manager_dialog') or not self.task_manager_dialog.isVisible():
            self.task_manager_dialog = ProcessManagerDialog(self)
            self.task_manager_dialog.show()
        else:
            self.task_manager_dialog.activateWindow()

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
        if out_path == "OPENED_IN_PPT":
            self.log_message("👁️ PowerPoint is open with your new slide. You can copy it directly.")
        elif out_path == "SPECS_DB_UPDATED":
            self.log_message("🔄 Reloading specifications table...")
            if hasattr(self, 'specs_tab'):
                self.specs_tab.refresh_table()
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

    def log_message(self, message: str, level=logging.INFO):
        self.console_panel.log_message(message, level)