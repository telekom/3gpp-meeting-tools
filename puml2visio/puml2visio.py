import sys
import logging
import datetime
import urllib.request
import webbrowser
from pathlib import Path

from PyQt5.QtWidgets import (QApplication, QMainWindow, QLabel, QVBoxLayout,
                             QWidget, QTextEdit, QDialog, QLineEdit, QPushButton,
                             QFormLayout, QHBoxLayout, QTabWidget, QCheckBox,
                             QSplitter, QStatusBar, QListWidget)
from PyQt5.QtCore import Qt, pyqtSignal

# --- IMPORT FROM OUR MODULAR BACKEND ---
from utils import JAR_NAME, encode_plantuml, InitializationThread, get_best_java
from word_extractor import WordExtractorThread
from visio_converter import VisioReaderThread, ConverterThread, SvgConverterThread
from powerpoint_converter import PptxConverterThread

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

# ==========================================
# --- GLOBAL STYLESHEET (ALL-BLUE THEME) ---
# ==========================================
GLOBAL_STYLE = """
    QWidget {
        font-family: "Segoe UI", Arial, sans-serif;
        font-size: 13px;
        color: #333333;
    }
    QToolTip {
        color: #333333;
        background-color: #F8F8F8;
        border: 1px solid #D0D0D0;
        border-radius: 4px;
        padding: 4px;
    }
    QTabWidget::pane {
        border: 1px solid #D0D0D0;
        border-radius: 8px;
        background: #FFFFFF;
        top: -1px;
    }
    QTabBar::tab {
        background: #EAEAEA;
        border: 1px solid #D0D0D0;
        padding: 8px 16px;
        margin-right: 2px;
        border-top-left-radius: 6px;
        border-top-right-radius: 6px;
    }
    QTabBar::tab:selected {
        background: #FFFFFF;
        border-bottom-color: #FFFFFF;
        font-weight: bold;
        color: #395396;
    }
    QTabBar::tab:hover:!selected {
        background: #F0F0F0;
    }
    QPushButton {
        padding: 8px 16px;
        border-radius: 6px;
        border: 1px solid #CCCCCC;
        background-color: #F8F8F8;
        font-weight: bold;
    }
    QPushButton:hover {
        background-color: #EAEAEA;
    }
    QPushButton:disabled {
        background-color: #F0F0F0;
        color: #A0A0A0;
        border: 1px solid #DFDFDF;
    }

    /* Primary Action Buttons */
    QPushButton#primaryBtn, QPushButton#pptBtn, QPushButton#svgBtn {
        background-color: #1E5C99; 
        color: white; 
        border: none;
    }
    QPushButton#primaryBtn:hover, QPushButton#pptBtn:hover, QPushButton#svgBtn:hover {
        background-color: #15426E;
    }

    /* Splitter Handle */
    QSplitter::handle {
        background-color: #E0E0E0;
        height: 2px;
        margin: 4px 0px;
    }
    QSplitter::handle:hover {
        background-color: #395396;
    }

    /* Status Bar */
    QStatusBar {
        background-color: #F0F0F0;
        border-top: 1px solid #D0D0D0;
        color: #333333;
    }

    /* Dark Theme for Console & Queue List */
    QTextEdit#console, QListWidget#queueList {
        background-color: #1E1E1E; 
        color: #D4D4D4; 
        font-family: Consolas, 'Courier New', monospace; 
        font-size: 13px; 
        border-radius: 8px; 
        padding: 8px;
        border: 1px solid #444444;
    }
    QListWidget#queueList::item {
        padding: 4px;
        border-bottom: 1px solid #333333;
    }
    QListWidget#queueList::item:selected {
        background-color: #264F78;
        color: #FFFFFF;
        border-radius: 4px;
    }
"""


# ==========================================
# --- GUI COMPONENTS ---
# ==========================================
class ProxyDialog(QDialog):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Network Configuration")
        self.setModal(True)
        self.resize(520, 250)

        layout = QVBoxLayout()
        title = QLabel("📡 Proxy Configuration")
        title.setStyleSheet("font-size: 16px; font-weight: bold; margin-bottom: 5px;")
        layout.addWidget(title)

        desc = QLabel("Leave blank to connect directly. Required only for initial JAR download.")
        desc.setStyleSheet("color: #666; margin-bottom: 15px;")
        layout.addWidget(desc)

        form = QFormLayout()
        self.http_input = QLineEdit()
        self.http_input.setStyleSheet("padding: 5px; border: 1px solid #ccc; border-radius: 4px;")
        self.https_input = QLineEdit()
        self.https_input.setStyleSheet("padding: 5px; border: 1px solid #ccc; border-radius: 4px;")
        self.sync_checkbox = QCheckBox("Use the same proxy for HTTPS")
        self.sync_checkbox.stateChanged.connect(self.on_sync_changed)
        self.http_input.textChanged.connect(self.on_http_changed)

        form.addRow("HTTP Proxy:", self.http_input)
        form.addRow("", self.sync_checkbox)
        form.addRow("HTTPS Proxy:", self.https_input)
        layout.addLayout(form)

        # --- NEW: Live Testing Status Label ---
        self.status_lbl = QLabel("")
        self.status_lbl.setWordWrap(True)
        self.status_lbl.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.status_lbl)

        btn_layout = QHBoxLayout()
        self.skip_btn = QPushButton("Skip")
        self.skip_btn.clicked.connect(self.skip)

        # --- NEW: Test Connection Button ---
        self.test_btn = QPushButton("🔄 Test Connection")
        self.test_btn.setToolTip("Ping GitHub to verify if your proxy settings are working.")
        self.test_btn.clicked.connect(self.test_connection)

        self.save_btn = QPushButton("Save && Continue")
        self.save_btn.setObjectName("primaryBtn")
        self.save_btn.clicked.connect(self.accept)

        btn_layout.addWidget(self.skip_btn)
        btn_layout.addStretch()
        btn_layout.addWidget(self.test_btn)
        btn_layout.addWidget(self.save_btn)

        layout.addSpacing(10)
        layout.addLayout(btn_layout)
        self.setLayout(layout)

    def test_connection(self):
        """Attempts to ping GitHub using the provided settings."""
        self.status_lbl.setText("⏳ Testing connection to GitHub... Please wait.")
        self.status_lbl.setStyleSheet("color: #D83B01; font-weight: bold;")
        QApplication.processEvents()  # Force UI update while blocking

        http_val, https_val = self.get_proxies()
        proxies = {}
        if http_val: proxies['http'] = http_val
        if https_val: proxies['https'] = https_val

        try:
            if proxies:
                proxy_handler = urllib.request.ProxyHandler(proxies)
                opener = urllib.request.build_opener(proxy_handler)
            else:
                # Test direct connection without proxy
                opener = urllib.request.build_opener(urllib.request.ProxyHandler({}))

            # Use method="HEAD" to ping the server without downloading the whole page
            req = urllib.request.Request("https://github.com", method="HEAD")
            opener.open(req, timeout=5)

            self.status_lbl.setText("✅ Connection Successful! You can now save.")
            self.status_lbl.setStyleSheet("color: #6A9955; font-weight: bold;")
        except Exception as e:
            self.status_lbl.setText(f"❌ Connection Failed: {str(e)}")
            self.status_lbl.setStyleSheet("color: #D32F2F; font-weight: bold;")

    def on_sync_changed(self, state):
        if state == Qt.Checked:
            self.https_input.setEnabled(False)
            self.https_input.setText(self.http_input.text())
        else:
            self.https_input.setEnabled(True)

    def on_http_changed(self, text):
        if self.sync_checkbox.isChecked():
            self.https_input.setText(text)

    def skip(self):
        self.http_input.clear()
        self.https_input.clear()
        self.accept()

    def get_proxies(self):
        return self.http_input.text().strip(), self.https_input.text().strip()


class CodeDropTextEdit(QTextEdit):
    """Text editor that accepts dropped Visio files for reverse-extraction."""
    file_dropped = pyqtSignal(str)

    def __init__(self):
        super().__init__()
        self.default_style = """
            QTextEdit {
                font-family: Consolas, Courier New, monospace; 
                font-size: 13px; 
                border: 2px solid #E0E0E0; 
                border-radius: 8px; 
                padding: 10px;
                background-color: #FAFAFA;
            }
            QTextEdit:focus {
                border: 2px solid #395396;
                background-color: #FFFFFF;
            }
        """
        self.hover_style = self.default_style.replace("border: 2px solid #E0E0E0;",
                                                      "border: 2px dashed #395396; background-color: #EBF3FC;")
        self.setStyleSheet(self.default_style)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            if any(url.toLocalFile().lower().endswith('.vsdx') for url in urls):
                self.setStyleSheet(self.hover_style)
                event.acceptProposedAction()
                return
        super().dragEnterEvent(event)

    def dragLeaveEvent(self, event):
        self.setStyleSheet(self.default_style)
        super().dragLeaveEvent(event)

    def dropEvent(self, event):
        self.setStyleSheet(self.default_style)
        if event.mimeData().hasUrls():
            for url in event.mimeData().urls():
                file_path = url.toLocalFile()
                if file_path.lower().endswith('.vsdx'):
                    self.file_dropped.emit(file_path)
                    event.acceptProposedAction()
                    return
        super().dropEvent(event)


class InteractiveDropLabel(QLabel):
    """A generic Drop Area label that highlights on hover."""
    file_dropped = pyqtSignal(list)

    def __init__(self, text, accepted_extensions):
        super().__init__(text)
        self.accepted_extensions = accepted_extensions
        self.setAlignment(Qt.AlignCenter)
        self.setAcceptDrops(True)
        self.default_style = "border: 3px dashed #B0B0B0; border-radius: 10px; font-size: 15px; font-weight: bold; color: #777; background-color: #FAFAFA;"
        self.hover_style = "border: 3px dashed #395396; border-radius: 10px; font-size: 15px; font-weight: bold; color: #395396; background-color: #EBF3FC;"
        self.busy_style = "border: 3px dashed #D83B01; border-radius: 10px; font-size: 15px; font-weight: bold; color: #D83B01; background-color: #FDF4F0;"
        self.error_style = "border: 3px dashed #D32F2F; border-radius: 10px; font-size: 15px; font-weight: bold; color: #D32F2F; background-color: #FDEDED;"
        self.setStyleSheet(self.default_style)

    def set_state(self, state, text=None):
        if text: self.setText(text)
        if state == "ready":
            self.setStyleSheet(self.default_style)
        elif state == "busy":
            self.setStyleSheet(self.busy_style)
        elif state == "error":
            self.setStyleSheet(self.error_style)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            if any(url.toLocalFile().lower().endswith(ext) for url in urls for ext in self.accepted_extensions):
                self.setStyleSheet(self.hover_style)
                event.accept()
                return
        event.ignore()

    def dragLeaveEvent(self, event):
        self.setStyleSheet(self.default_style)

    def dropEvent(self, event):
        self.setStyleSheet(self.default_style)
        valid_files = []
        for url in event.mimeData().urls():
            file_path = url.toLocalFile()
            if any(file_path.lower().endswith(ext) for ext in self.accepted_extensions):
                valid_files.append(file_path)
        if valid_files:
            self.file_dropped.emit(valid_files)


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

        self.init_thread = InitializationThread(self.jar_path)
        self.init_thread.ui_log_msg.connect(self.log_message)
        self.init_thread.init_complete.connect(self.on_init_complete)
        self.init_thread.start()

    def _setup_ui(self):
        central_widget = QWidget()
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(15, 15, 15, 15)

        self.splitter = QSplitter(Qt.Vertical)

        # --- TOP HALF: TABS ---
        self.tabs = QTabWidget()

        # Tab 1: Paste Code
        self.tab_text = QWidget()
        tab_text_layout = QVBoxLayout()
        tab_text_layout.setContentsMargins(15, 15, 15, 15)

        self.text_input = CodeDropTextEdit()
        self.text_input.setPlaceholderText(
            "Paste PlantUML code OR drop a generated .vsdx file here to extract its source...")
        self.text_input.file_dropped.connect(self.extract_code_from_visio)
        self.text_input.setToolTip(
            "Type or paste PlantUML code here. Drag & drop a generated .vsdx file to retrieve its source code.")

        btn_layout = QHBoxLayout()

        self.clear_btn = QPushButton("🗑️ Clear")
        self.clear_btn.clicked.connect(self.text_input.clear)
        self.clear_btn.setToolTip("Clear the text editor.")

        self.planttext_btn = QPushButton("🌐 Show in planttext")
        self.planttext_btn.clicked.connect(self.show_in_planttext)
        self.planttext_btn.setToolTip("Open your code in PlantText.com for a quick web preview.")

        self.copy_btn = QPushButton("📋 Copy Path")
        self.copy_btn.setEnabled(False)
        self.copy_btn.clicked.connect(self.copy_out_path)
        self.copy_btn.setToolTip("Copy the file path of the last generated diagram.")

        self.convert_svg_btn = QPushButton("Export SVG")
        self.convert_svg_btn.setObjectName("svgBtn")
        self.convert_svg_btn.clicked.connect(lambda: self._save_and_queue_pasted_text("svg"))
        self.convert_svg_btn.setToolTip("Generate a standard, scalable vector graphic (.svg) and open it.")

        self.convert_pptx_btn = QPushButton("Export PPTX")
        self.convert_pptx_btn.setObjectName("pptBtn")
        self.convert_pptx_btn.clicked.connect(lambda: self._save_and_queue_pasted_text("pptx"))
        self.convert_pptx_btn.setToolTip("Generate a PowerPoint slide containing natively editable Office shapes.")

        self.convert_vsdx_btn = QPushButton("Export Visio")
        self.convert_vsdx_btn.setObjectName("primaryBtn")
        self.convert_vsdx_btn.clicked.connect(lambda: self._save_and_queue_pasted_text("vsdx"))
        self.convert_vsdx_btn.setToolTip("Generate a perfectly aligned, natively editable Visio diagram (.vsdx).")

        btn_layout.addWidget(self.clear_btn)
        btn_layout.addWidget(self.planttext_btn)
        btn_layout.addStretch()
        btn_layout.addWidget(self.copy_btn)
        btn_layout.addWidget(self.convert_svg_btn)
        btn_layout.addWidget(self.convert_pptx_btn)
        btn_layout.addWidget(self.convert_vsdx_btn)

        tab_text_layout.addWidget(self.text_input)
        tab_text_layout.addLayout(btn_layout)
        self.tab_text.setLayout(tab_text_layout)

        # Tab 2: Drag & Drop Batch Files
        self.tab_file = QWidget()
        tab_file_layout = QVBoxLayout()
        tab_file_layout.setContentsMargins(15, 15, 15, 15)
        self.drop_label = InteractiveDropLabel("⏳ Initializing system checks... Please wait.", ['.puml', '.txt'])
        self.drop_label.file_dropped.connect(self.handle_batch_drop)
        tab_file_layout.addWidget(self.drop_label)
        self.tab_file.setLayout(tab_file_layout)

        # Tab 3: Word Extractor
        self.tab_word = QWidget()
        tab_word_layout = QVBoxLayout()
        tab_word_layout.setContentsMargins(15, 15, 15, 15)
        self.word_drop_label = InteractiveDropLabel(
            "📥 Drag && Drop your Microsoft Word (.docx) file here\n\nExtracts all embedded Visio diagrams to the file's folder.",
            ['.docx'])
        self.word_drop_label.file_dropped.connect(lambda files: self.start_word_extraction(files[0]))
        tab_word_layout.addWidget(self.word_drop_label)
        self.tab_word.setLayout(tab_word_layout)

        self.tabs.addTab(self.tab_text, "📝 Code Editor")
        self.tabs.addTab(self.tab_file, "📂 Batch Convert")
        self.tabs.addTab(self.tab_word, "📄 Word Extractor")
        self.tabs.setEnabled(False)

        self.splitter.addWidget(self.tabs)

        # --- BOTTOM HALF: CONSOLE AND QUEUE ---
        self.bottom_splitter = QSplitter(Qt.Horizontal)

        # 1. Console Terminal
        console_container = QWidget()
        console_layout = QVBoxLayout()
        console_layout.setContentsMargins(0, 5, 0, 0)

        console_header = QHBoxLayout()
        terminal_lbl = QLabel("Terminal Output")
        terminal_lbl.setStyleSheet("font-weight: bold; color: #555;")

        # Proxy Update Button
        self.proxy_btn = QPushButton("📡 Proxy")
        self.proxy_btn.setFixedSize(70, 24)
        self.proxy_btn.setStyleSheet("padding: 2px; font-size: 11px;")
        self.proxy_btn.setToolTip("Update network proxy settings and retry system initialization.")
        self.proxy_btn.clicked.connect(self.open_proxy_settings)

        clear_log_btn = QPushButton("Clear")
        clear_log_btn.setFixedSize(60, 24)
        clear_log_btn.setStyleSheet("padding: 2px; font-size: 11px;")
        clear_log_btn.clicked.connect(lambda: self.console.clear())

        console_header.addWidget(terminal_lbl)
        console_header.addStretch()
        console_header.addWidget(self.proxy_btn)
        console_header.addWidget(clear_log_btn)

        self.console = QTextEdit()
        self.console.setReadOnly(True)
        self.console.setObjectName("console")

        console_layout.addLayout(console_header)
        console_layout.addWidget(self.console)
        console_container.setLayout(console_layout)
        self.bottom_splitter.addWidget(console_container)

        # 2. Queue Viewer
        queue_container = QWidget()
        queue_layout = QVBoxLayout()
        queue_layout.setContentsMargins(0, 5, 0, 0)

        queue_header = QHBoxLayout()
        queue_lbl = QLabel("Queue")
        queue_lbl.setStyleSheet("font-weight: bold; color: #555;")

        self.remove_btn = QPushButton("Remove")
        self.remove_btn.setFixedSize(60, 24)
        self.remove_btn.setStyleSheet("padding: 2px; font-size: 11px;")
        self.remove_btn.setToolTip("Remove selected item(s) from the waiting queue.")
        self.remove_btn.clicked.connect(self.remove_selected_from_queue)

        self.clear_q_btn = QPushButton("Clear All")
        self.clear_q_btn.setFixedSize(60, 24)
        self.clear_q_btn.setStyleSheet("padding: 2px; font-size: 11px;")
        self.clear_q_btn.setToolTip("Remove all waiting items from the queue.")
        self.clear_q_btn.clicked.connect(self.clear_queue)

        queue_header.addWidget(queue_lbl)
        queue_header.addStretch()
        queue_header.addWidget(self.remove_btn)
        queue_header.addWidget(self.clear_q_btn)

        self.queue_list = QListWidget()
        self.queue_list.setObjectName("queueList")
        self.queue_list.setSelectionMode(QListWidget.ExtendedSelection)

        queue_layout.addLayout(queue_header)
        queue_layout.addWidget(self.queue_list)
        queue_container.setLayout(queue_layout)

        self.bottom_splitter.addWidget(queue_container)
        self.bottom_splitter.setSizes([650, 250])

        self.splitter.addWidget(self.bottom_splitter)
        self.splitter.setSizes([450, 250])
        main_layout.addWidget(self.splitter)

        central_widget.setLayout(main_layout)
        self.setCentralWidget(central_widget)

        # --- STATUS BAR ---
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("⏳ Initializing...")

    # --- APPLICATION LOGIC ---
    def open_proxy_settings(self):
        """Opens the proxy dialog and retries initialization if settings are updated."""
        proxy_dialog = ProxyDialog()
        if proxy_dialog.exec_() == QDialog.Accepted:
            http_val, https_val = proxy_dialog.get_proxies()
            proxies = {}
            if http_val: proxies['http'] = http_val
            if https_val: proxies['https'] = https_val

            # Apply the new proxy settings globally to the application
            proxy_handler = urllib.request.ProxyHandler(proxies)
            opener = urllib.request.build_opener(proxy_handler)
            urllib.request.install_opener(opener)

            status = "updated" if proxies else "cleared (direct connection)"
            self.log_message(f"✅ Proxy settings {status}. Retrying system checks...", logging.INFO)

            # Reset UI state
            self.drop_label.set_state("ready", "⏳ Re-initializing system checks...")
            self.status_bar.showMessage("⏳ Re-initializing...")
            self.tabs.setEnabled(False)

            # Relaunch the background initialization thread
            self.init_thread = InitializationThread(self.jar_path)
            self.init_thread.ui_log_msg.connect(self.log_message)
            self.init_thread.init_complete.connect(self.on_init_complete)
            self.init_thread.start()

    def _get_display_name(self, file_path):
        import re
        name = file_path.name
        if re.match(r"^\d{4}\.\d{2}\.\d{2} \d{2}-\d{2}-\d{2} ", name):
            return name[20:]
        return name

    def _refresh_queue_list(self):
        self.queue_list.clear()
        for index, (file_path, target_format) in enumerate(self.file_queue, start=1):
            display_name = self._get_display_name(file_path)
            self.queue_list.addItem(f"{index}. {display_name} → .{target_format.upper()}")

    def remove_selected_from_queue(self):
        selected_items = self.queue_list.selectedItems()
        if not selected_items: return
        rows = sorted([self.queue_list.row(item) for item in selected_items], reverse=True)
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
            self.drop_label.set_state("error", "❌ Initialization Failed.")
            self.status_bar.showMessage("❌ Initialization Failed. Check log for details.")

    def _set_drop_zone_ready(self):
        self.drop_label.set_state("ready",
                                  "📥 Drag && Drop your .puml or .txt file(s) here\n\n(Batch exports as Visio files)")
        self.status_bar.showMessage("🟢 System Idle.")

    def _set_drop_zone_busy(self):
        self.drop_label.set_state("busy",
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
        self.text_input.clear()
        self.text_input.setPlaceholderText(f"⏳ Extracting source from {Path(file_path).name}...\nPlease wait...")
        self.text_input.setEnabled(False)
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
        self.text_input.setEnabled(True)
        self.text_input.setPlainText(source_code)
        self.log_message("✅ Successfully extracted PlantUML source from Visio file.")

    def on_visio_code_error(self, error_msg):
        self.text_input.setEnabled(True)
        self.text_input.setPlaceholderText(
            "Paste PlantUML code OR drop a generated .vsdx file here to extract its source...")
        self.log_message(f"❌ {error_msg}")

    def show_in_planttext(self):
        raw_text = self.text_input.toPlainText().strip()
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
        raw_text = self.text_input.toPlainText().strip()
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
            self.copy_btn.setEnabled(True)
            self.copy_btn.setStyleSheet("background-color: #395396; color: white; border: none;")

            if out_path.lower().endswith('.svg'):
                try:
                    import os
                    os.startfile(out_path)
                    ext = Path(out_path).suffix[1:].upper()
                    self.log_message(f"👁️ Opened {ext} in default system viewer.")
                except Exception as e:
                    self.log_message(f"⚠️ Could not automatically open file: {e}")

    def log_message(self, message: str, level=logging.INFO):
        color = "#D4D4D4"
        if "❌" in message or "Error" in message:
            color = "#F44747"
        elif "⚠️" in message or "Warning" in message:
            color = "#D7BA7D"
        elif "✅" in message or "Success" in message or "Ready" in message:
            color = "#6A9955"

        html_msg = f'<span style="color: {color};">{message.replace(chr(10), "<br>")}</span>'
        self.console.append(html_msg)

        QApplication.processEvents()
        scrollbar = self.console.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())


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