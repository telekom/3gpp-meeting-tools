import logging
import datetime
import urllib.request
import webbrowser
import os
import subprocess
from pathlib import Path

from PyQt5.QtWidgets import QMainWindow, QWidget, QVBoxLayout, QSplitter, QStatusBar, QApplication, QDialog
from PyQt5.QtCore import Qt, QTimer, QThread, pyqtSignal
from PyQt5.QtGui import QTextCursor

from ui_components import ProxyDialog
from ui_tabs import CodeEditorTab, BatchConvertTab, WordExtractorTab
# --- NEW: Import our abstracted UI Panels! ---
from ui_panels import ConsolePanel, QueuePanel

from utils import JAR_NAME, encode_plantuml, InitializationThread, get_best_java
from word_extractor import WordExtractorThread
from visio_converter import VisioReaderThread, ConverterThread, SvgConverterThread
from powerpoint_converter import PptxConverterThread
from live_preview import LivePreviewManager
from plantuml_templates import PLANTUML_TYPES


# ==========================================
# --- ASCII CONVERTER THREAD ---
# ==========================================
class AsciiConverterThread(QThread):
    ui_log_msg = pyqtSignal(str, int)
    finished_path = pyqtSignal(str)
    finished = pyqtSignal()

    def __init__(self, puml_path: Path, jar_path: Path):
        super().__init__()
        self.puml_path = puml_path
        self.jar_path = jar_path

    def run(self):
        try:
            java_exe, _ = get_best_java()
            cmd = [java_exe, "-jar", str(self.jar_path), "-tutxt", str(self.puml_path)]
            kwargs = {'creationflags': 0x08000000} if os.name == 'nt' else {}

            subprocess.run(cmd, check=True, capture_output=True, text=True, cwd=self.puml_path.parent, **kwargs)

            utxt_path = self.puml_path.with_suffix(".utxt")
            txt_path = self.puml_path.with_name(self.puml_path.stem + "_ascii.txt")

            if utxt_path.exists():
                if txt_path.exists(): txt_path.unlink()
                utxt_path.rename(txt_path)
                self.ui_log_msg.emit("✅ Unicode Text Art generated successfully.", logging.INFO)
                self.finished_path.emit(str(txt_path))
            else:
                self.ui_log_msg.emit(
                    "❌ Failed to generate text art. Format may not be supported for this diagram type.", logging.ERROR)
                self.finished_path.emit("")

        except Exception as e:
            self.ui_log_msg.emit(f"❌ Text Conversion Error: {e}", logging.ERROR)
            self.finished_path.emit("")
        finally:
            self.finished.emit()


class DragDropUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PlantUML to Visio Converter (3GPP)")
        self.resize(950, 750)

        self.jar_path = Path(__file__).parent.resolve() / JAR_NAME
        self.file_queue = []
        self.is_processing = False
        self.last_out_path = ""

        self._setup_ui()

        self.cache_file = Path(__file__).parent.resolve() / ".editor_cache.puml"
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

        # ==========================================
        # --- TOP HALF: TABS ---
        # ==========================================
        from PyQt5.QtWidgets import QTabWidget
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
        self.batch_tab.files_dropped.connect(self.handle_batch_drop)

        self.word_tab = WordExtractorTab()
        self.word_tab.file_dropped.connect(self.start_word_extraction)

        self.tabs.addTab(self.code_tab, "📝 Code Editor")
        self.tabs.addTab(self.batch_tab, "📂 Batch Convert")
        self.tabs.addTab(self.word_tab, "📄 Word Extractor")
        self.tabs.setEnabled(False)

        self.splitter.addWidget(self.tabs)

        # ==========================================
        # --- BOTTOM HALF: PANELS ---
        # ==========================================
        self.bottom_splitter = QSplitter(Qt.Horizontal)

        # 1. Console Panel
        self.console_panel = ConsolePanel()
        self.console_panel.proxy_requested.connect(self.open_proxy_settings)
        self.console_panel.update_requested.connect(self.check_for_jar_updates)

        # 2. Queue Panel
        self.queue_panel = QueuePanel()
        self.queue_panel.remove_requested.connect(self.remove_selected_from_queue)
        self.queue_panel.clear_requested.connect(self.clear_queue)

        self.bottom_splitter.addWidget(self.console_panel)
        self.bottom_splitter.addWidget(self.queue_panel)
        self.bottom_splitter.setSizes([650, 250])

        self.splitter.addWidget(self.bottom_splitter)
        self.splitter.setSizes([450, 250])
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
        export_dir = Path(__file__).parent.resolve()
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

    def _get_display_name(self, file_path):
        import re
        name = file_path.name
        if re.match(r"^\d{4}\.\d{2}\.\d{2} \d{2}-\d{2}-\d{2} ", name):
            return name[20:]
        return name

    # --- QUEUE MANAGER DELEGATION ---
    def _refresh_queue_list(self):
        display_items = []
        for index, (file_path, target_format) in enumerate(self.file_queue, start=1):
            display_name = self._get_display_name(file_path)
            display_items.append(f"{index}. {display_name} → .{target_format.upper()}")

        self.queue_panel.update_list(display_items)

    def remove_selected_from_queue(self, rows):
        """Receives a list of integers (row indices) from the QueuePanel."""
        for row in rows:
            del self.file_queue[row]
        self._refresh_queue_list()
        self._update_queue_ui_text()

    def clear_queue(self):
        self.file_queue.clear()
        self._refresh_queue_list()
        self._update_queue_ui_text()

    def on_init_complete(self, success: bool):
        if success:
            self.tabs.setEnabled(True)
            self._set_drop_zone_ready()
            self.log_message("🚀 System Ready. Paste code or drop files to begin.\n" + "-" * 45)
        else:
            self.batch_tab.set_state("error", "❌ Initialization Failed.")
            self.status_bar.showMessage("❌ Initialization Failed. Check log for details.")

    def _set_drop_zone_ready(self):
        self.batch_tab.set_state("ready",
                                 "📥 Drag && Drop your .puml or .txt file(s) here\n\n(Batch exports as Visio files)")
        self.status_bar.showMessage("🟢 System Idle.")

    def _set_drop_zone_busy(self):
        self.batch_tab.set_state("busy",
                                 "⚙️ Processing Queue...\n\nPlease wait until finished or drop more files to queue them.")

    def handle_batch_drop(self, file_paths):
        for file_path in file_paths:
            self.file_queue.append((Path(file_path), "vsdx"))
        self._refresh_queue_list()
        if not self.is_processing:
            self.process_next_in_queue()
        else:
            self._update_queue_ui_text()

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

    def start_word_extraction(self, file_path):
        self.word_extractor_thread = WordExtractorThread(file_path)
        self.word_extractor_thread.ui_log_msg.connect(self.log_message)
        self.word_extractor_thread.start()

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

    def _save_and_queue_pasted_text(self, target_format):
        raw_text = self.code_tab.get_text()
        if not raw_text: return

        base_dir = Path(__file__).parent.resolve()
        timestamp = datetime.datetime.now().strftime("%Y.%m.%d %H-%M-%S")
        base_name = f"{timestamp} diagram"
        puml_path = base_dir / f"{base_name}.puml"

        counter = 1
        while puml_path.exists() or puml_path.with_suffix(".vsdx").exists() or puml_path.with_suffix(".svg").exists():
            puml_path = base_dir / f"{base_name}_{counter}.puml"
            counter += 1

        with open(puml_path, "w", encoding="utf-8") as f:
            f.write(raw_text)

        self.file_queue.append((puml_path, target_format))
        self._refresh_queue_list()

        if self.is_processing:
            self._update_queue_ui_text()
        else:
            self.process_next_in_queue()

    def _update_queue_ui_text(self, current_file_name=""):
        remaining = len(self.file_queue)
        rem_text = f" | {remaining} items waiting in queue." if remaining > 0 else ""
        if current_file_name:
            self.status_bar.showMessage(f"⚙️ Processing: {current_file_name}{rem_text}")
        else:
            curr_text = self.status_bar.currentMessage().split("|")[0].strip()
            if "Idle" in curr_text and remaining == 0:
                self.status_bar.showMessage("🟢 System Idle.")
            else:
                self.status_bar.showMessage(f"{curr_text}{rem_text}")

    def process_next_in_queue(self):
        if not self.file_queue:
            self.is_processing = False
            self._set_drop_zone_ready()
            return

        self.is_processing = True
        self._set_drop_zone_busy()

        next_file, target_format = self.file_queue.pop(0)
        display_name = self._get_display_name(next_file)

        self._refresh_queue_list()
        self._update_queue_ui_text(f"{display_name} (to .{target_format.upper()})")

        if target_format == "svg":
            self.conv_thread = SvgConverterThread(next_file, self.jar_path)
        elif target_format == "pptx":
            self.conv_thread = PptxConverterThread(next_file, self.jar_path)
        elif target_format == "ascii":
            self.conv_thread = AsciiConverterThread(next_file, self.jar_path)
        else:
            self.conv_thread = ConverterThread(next_file, self.jar_path)

        self.conv_thread.ui_log_msg.connect(self.log_message)
        self.conv_thread.finished_path.connect(self.on_conversion_success)
        self.conv_thread.finished.connect(self.process_next_in_queue)
        self.conv_thread.start()

    def on_conversion_success(self, out_path: str):
        if out_path == "OPENED_IN_PPT":
            self.log_message("👁️ PowerPoint is open with your new slide. You can copy it directly.")
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

    # --- WRAPPER METHOD ---
    def log_message(self, message: str, level=logging.INFO):
        """Passes logs to the newly decoupled ConsolePanel."""
        self.console_panel.log_message(message, level)