import sys
import logging
import datetime
import urllib.request
import webbrowser
from pathlib import Path

from PyQt5.QtWidgets import (QApplication, QMainWindow, QLabel, QVBoxLayout,
                             QWidget, QTextEdit, QDialog, QLineEdit, QPushButton,
                             QFormLayout, QHBoxLayout, QTabWidget, QCheckBox)
from PyQt5.QtCore import Qt, pyqtSignal

# --- IMPORT THE LOGIC FROM OUR NEW BACKEND FILE ---
from visio_backend import (
    InitializationThread,
    VisioReaderThread,
    WordExtractorThread,
    ConverterThread,
    SvgConverterThread,  # NEW
    encode_plantuml,
    JAR_NAME
)

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
# --- GUI COMPONENTS ---
# ==========================================
class ProxyDialog(QDialog):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Network Configuration")
        self.setModal(True)
        self.resize(500, 220)

        layout = QVBoxLayout()
        title = QLabel("📡 Proxy Configuration")
        title.setStyleSheet("font-size: 16px; font-weight: bold; margin-bottom: 5px;")
        layout.addWidget(title)

        desc = QLabel("Leave blank to connect directly. Required only for initial JAR download.")
        desc.setStyleSheet("color: #555; margin-bottom: 15px;")
        layout.addWidget(desc)

        form = QFormLayout()
        self.http_input = QLineEdit()
        self.https_input = QLineEdit()
        self.sync_checkbox = QCheckBox("Use the same proxy for HTTPS")
        self.sync_checkbox.stateChanged.connect(self.on_sync_changed)
        self.http_input.textChanged.connect(self.on_http_changed)

        form.addRow("HTTP Proxy:", self.http_input)
        form.addRow("", self.sync_checkbox)
        form.addRow("HTTPS Proxy:", self.https_input)
        layout.addLayout(form)

        btn_layout = QHBoxLayout()
        self.skip_btn = QPushButton("Continue Without Proxy")
        self.skip_btn.clicked.connect(self.skip)
        self.save_btn = QPushButton("Set Proxy && Continue")
        self.save_btn.setStyleSheet("background-color: #0078d7; color: white; font-weight: bold;")
        self.save_btn.clicked.connect(self.accept)

        btn_layout.addWidget(self.skip_btn)
        btn_layout.addStretch()
        btn_layout.addWidget(self.save_btn)

        layout.addSpacing(10)
        layout.addLayout(btn_layout)
        self.setLayout(layout)

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
    file_dropped = pyqtSignal(str)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            if any(url.toLocalFile().lower().endswith('.vsdx') for url in urls):
                event.acceptProposedAction()
                return
        super().dragEnterEvent(event)

    def dropEvent(self, event):
        if event.mimeData().hasUrls():
            for url in event.mimeData().urls():
                file_path = url.toLocalFile()
                if file_path.lower().endswith('.vsdx'):
                    self.file_dropped.emit(file_path)
                    event.acceptProposedAction()
                    return
        super().dropEvent(event)


class WordDropLabel(QLabel):
    file_dropped = pyqtSignal(str)

    def __init__(self):
        super().__init__(
            "📥 Drag && Drop your Microsoft Word (.docx) file here\n\nExtracts all embedded Visio diagrams to the file's folder.")
        self.setAlignment(Qt.AlignCenter)
        self.setStyleSheet("""
            QLabel {
                border: 3px dashed #2b579a;
                border-radius: 10px;
                background-color: #f3f8fd;
                font-size: 15px;
                font-weight: bold;
                color: #2b579a;
                padding: 20px;
            }
        """)
        self.setAcceptDrops(True)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            if any(url.toLocalFile().lower().endswith('.docx') for url in urls):
                event.accept()
                return
        event.ignore()

    def dropEvent(self, event):
        for url in event.mimeData().urls():
            file_path = url.toLocalFile()
            if file_path.lower().endswith('.docx'):
                self.file_dropped.emit(file_path)


class DragDropUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PlantUML to Visio Converter (3GPP)")
        self.resize(850, 650)
        self.setAcceptDrops(True)

        self.jar_path = Path(__file__).parent.resolve() / JAR_NAME
        self.file_queue = []  # Stores tuples: (Path, 'vsdx' or 'svg')
        self.is_processing = False
        self.last_out_path = ""

        self._setup_ui()

        self.init_thread = InitializationThread(self.jar_path)
        self.init_thread.ui_log_msg.connect(self.log_message)
        self.init_thread.init_complete.connect(self.on_init_complete)
        self.init_thread.start()

    def _setup_ui(self):
        main_layout = QVBoxLayout()
        self.tabs = QTabWidget()

        # Tab 1: Paste Code
        self.tab_text = QWidget()
        tab_text_layout = QVBoxLayout()
        self.text_input = CodeDropTextEdit()
        self.text_input.setPlaceholderText(
            "Paste PlantUML code OR drop a generated .vsdx file here to extract its source...")
        self.text_input.setStyleSheet(
            "font-family: Consolas, Courier New, monospace; font-size: 13px; border: 1px solid #ccc;")
        self.text_input.file_dropped.connect(self.extract_code_from_visio)

        btn_layout = QHBoxLayout()
        self.clear_btn = QPushButton("🗑️ Clear")
        self.clear_btn.clicked.connect(self.text_input.clear)
        self.planttext_btn = QPushButton("🌐 Show in planttext.com")
        self.planttext_btn.clicked.connect(self.show_in_planttext)
        self.copy_btn = QPushButton("📋 Copy File Path")
        self.copy_btn.setEnabled(False)
        self.copy_btn.clicked.connect(self.copy_out_path)

        # Format Export Buttons
        self.convert_svg_btn = QPushButton("Export to SVG")
        self.convert_svg_btn.setStyleSheet("background-color: #ffaa00; color: #000; font-weight: bold;")
        self.convert_svg_btn.clicked.connect(self.convert_pasted_text_to_svg)

        self.convert_vsdx_btn = QPushButton("Export to Visio")
        self.convert_vsdx_btn.setStyleSheet("background-color: #4af626; color: #000; font-weight: bold;")
        self.convert_vsdx_btn.clicked.connect(self.convert_pasted_text_to_visio)

        btn_layout.addWidget(self.clear_btn)
        btn_layout.addWidget(self.planttext_btn)
        btn_layout.addWidget(self.copy_btn)
        btn_layout.addWidget(self.convert_svg_btn)
        btn_layout.addWidget(self.convert_vsdx_btn)

        tab_text_layout.addWidget(self.text_input)
        tab_text_layout.addLayout(btn_layout)
        self.tab_text.setLayout(tab_text_layout)

        # Tab 2: Drag & Drop Files
        self.tab_file = QWidget()
        tab_file_layout = QVBoxLayout()
        self.drop_label = QLabel("⏳ Initializing system checks... Please wait.")
        self.drop_label.setAlignment(Qt.AlignCenter)
        self.drop_label.setStyleSheet(
            "border: 3px dashed #888; border-radius: 10px; font-size: 15px; font-weight: bold; color: #555;")
        tab_file_layout.addWidget(self.drop_label)
        self.tab_file.setLayout(tab_file_layout)

        # Tab 3: Word Extractor
        self.tab_word = QWidget()
        tab_word_layout = QVBoxLayout()
        self.word_drop_label = WordDropLabel()
        self.word_drop_label.file_dropped.connect(self.start_word_extraction)
        tab_word_layout.addWidget(self.word_drop_label)
        self.tab_word.setLayout(tab_word_layout)

        self.tabs.addTab(self.tab_text, "📝 Paste Code")
        self.tabs.addTab(self.tab_file, "📂 Drag && Drop Files")
        self.tabs.addTab(self.tab_word, "📄 Word Extractor")
        self.tabs.setEnabled(False)

        main_layout.addWidget(self.tabs, stretch=1)

        # Console
        self.console = QTextEdit()
        self.console.setReadOnly(True)
        self.console.setStyleSheet(
            "background-color: #1e1e1e; color: #4af626; font-family: Consolas, monospace; font-size: 13px;")
        main_layout.addWidget(self.console, stretch=1)

        container = QWidget()
        container.setLayout(main_layout)
        self.setCentralWidget(container)

    def on_init_complete(self, success: bool):
        if success:
            self.tabs.setEnabled(True)
            self._set_drop_zone_ready()
            self.log_message("\n🚀 System Ready. Paste code or drop files to begin.\n" + "-" * 45)
        else:
            self.drop_label.setText("❌ Initialization Failed.")

    def _set_drop_zone_ready(self):
        self.drop_label.setText("📥 Drag && Drop your .puml or .txt file(s) here\n\n(Batch exports as Visio files)")
        self.drop_label.setStyleSheet(
            "border: 3px dashed #4af626; background-color: #f4f4f4; font-size: 15px; font-weight: bold;")
        self.convert_vsdx_btn.setEnabled(True)
        self.convert_vsdx_btn.setText("Export to Visio")

    def _set_drop_zone_busy(self):
        self.drop_label.setText("⚙️ Processing Queue...\n\nPlease wait until finished.")
        self.drop_label.setStyleSheet(
            "border: 3px dashed #f6a826; background-color: #fff4e5; font-size: 15px; font-weight: bold; color: #d37e00;")
        self.convert_vsdx_btn.setEnabled(False)
        self.convert_vsdx_btn.setText("Processing...")

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

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event):
        if not event.mimeData().urls(): return

        puml_added = 0
        for url in event.mimeData().urls():
            file_path = Path(url.toLocalFile())
            suffix = file_path.suffix.lower()

            if suffix in [".puml", ".txt"]:
                # Drag and drop defaults to VSDX
                self.file_queue.append((file_path, "vsdx"))
                puml_added += 1
            elif suffix == ".docx":
                self.start_word_extraction(str(file_path))
            else:
                self.log_message(f"⚠️ Skipped unsupported file: {file_path.name}")

        if puml_added > 0 and not self.is_processing:
            self.process_next_in_queue()

    def convert_pasted_text_to_visio(self):
        self._save_and_queue_pasted_text("vsdx")

    def convert_pasted_text_to_svg(self):
        self._save_and_queue_pasted_text("svg")

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
        if not self.is_processing:
            self.process_next_in_queue()

    def process_next_in_queue(self):
        if not self.file_queue:
            self.is_processing = False
            self._set_drop_zone_ready()
            return

        self.is_processing = True
        self._set_drop_zone_busy()

        next_file, target_format = self.file_queue.pop(0)

        # Launch appropriate thread based on requested format
        if target_format == "svg":
            self.conv_thread = SvgConverterThread(next_file, self.jar_path)
        else:
            self.conv_thread = ConverterThread(next_file, self.jar_path)

        self.conv_thread.ui_log_msg.connect(self.log_message)
        self.conv_thread.finished_path.connect(self.on_conversion_success)
        self.conv_thread.finished.connect(self.process_next_in_queue)
        self.conv_thread.start()

    def on_conversion_success(self, out_path: str):
        if out_path:
            self.last_out_path = out_path
            self.copy_btn.setEnabled(True)
            self.copy_btn.setStyleSheet("background-color: #0078d7; color: white; font-weight: bold;")

    def log_message(self, message: str):
        self.console.append(message)
        QApplication.processEvents()
        scrollbar = self.console.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())


if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setStyle("Fusion")

    jar_path = Path(__file__).parent.resolve() / JAR_NAME

    if not jar_path.exists():
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