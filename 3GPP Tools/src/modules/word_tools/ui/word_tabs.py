import webbrowser
from pathlib import Path

from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtWidgets import (
    QCheckBox,
    QComboBox,
    QFormLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QSpinBox,
    QStackedWidget,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)
import win32com.client
import pythoncom

from core.ui.ui_components import InteractiveDropLabel
from modules.word_tools.core.libreoffice_converter import (
    LIBREOFFICE_DOWNLOAD_URL,
    is_libreoffice_available,
)


class DocumentSelectorPane(QWidget):
    """A symmetric, reusable widget handling Local, Open, and URL inputs."""

    def __init__(self, title: str):
        super().__init__()
        self.title = title
        self.selected_files = []
        self._setup_ui()

    def _setup_ui(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)

        lbl = QLabel(f"<b>{self.title}</b>")
        lbl.setStyleSheet("color: #444; margin-bottom: 5px;")
        layout.addWidget(lbl)

        self.tabs = QTabWidget()
        self.tabs.setObjectName("selector_tabs")

        # Tab 1: Local File Drop
        self.drop_tab = QWidget()
        drop_layout = QVBoxLayout(self.drop_tab)
        self.drop_zone = InteractiveDropLabel("Drop .docx here", [".docx"])
        self.drop_zone.file_dropped.connect(self._on_drop)
        drop_layout.addWidget(self.drop_zone)
        self.tabs.addTab(self.drop_tab, "📁 Local")

        # Tab 2: Open Documents
        self.open_tab = QWidget()
        open_layout = QVBoxLayout(self.open_tab)

        self.open_combo = QComboBox()
        self.refresh_btn = QPushButton("↻ Refresh Active Documents")
        self.refresh_btn.clicked.connect(self.poll_open_documents)

        open_layout.addWidget(QLabel("Select an open Word document:"))
        open_layout.addWidget(self.open_combo)
        open_layout.addWidget(self.refresh_btn)
        open_layout.addStretch()
        self.tabs.addTab(self.open_tab, "🖥️ Open Docs")

        # Tab 3: URL
        self.url_tab = QWidget()
        url_layout = QVBoxLayout(self.url_tab)

        self.url_input = QLineEdit()
        self.url_input.setPlaceholderText("https://...")

        url_layout.addWidget(QLabel("Paste document URL:"))
        url_layout.addWidget(self.url_input)
        url_layout.addStretch()
        self.tabs.addTab(self.url_tab, "🌐 URL")

        layout.addWidget(self.tabs)
        self.setLayout(layout)
        self.tabs.currentChanged.connect(self._on_tab_changed)

    def _on_drop(self, files):
        if files:
            self.selected_files = files
            if len(files) == 1:
                self.drop_zone.set_state("ready", f"Ready:\n{Path(files[0]).name}")
            else:
                self.drop_zone.set_state("ready", f"Ready:\n{len(files)} files queued for batch")

    def poll_open_documents(self):
        self.open_combo.clear()
        try:
            pythoncom.CoInitialize()
            word = win32com.client.GetActiveObject("Word.Application")
            for doc in word.Documents:
                doc_name = str(doc.Name).strip()
                if doc_name:
                    self.open_combo.addItem(doc_name, doc.FullName)
            if self.open_combo.count() == 0:
                self.open_combo.addItem("No open documents detected.", "")
        except Exception:
            self.open_combo.addItem("No open documents detected.", "")
        finally:
            pythoncom.CoUninitialize()

    def get_inputs(self) -> list:
        idx = self.tabs.currentIndex()
        if idx == 0:
            return self.selected_files
        elif idx == 1:
            doc = self.open_combo.currentData()
            return [doc] if doc else []
        elif idx == 2:
            url = self.url_input.text().strip()
            return [url] if url else []
        return []

    def _on_tab_changed(self, index):
        if index == 1:
            self.poll_open_documents()


class WordExtractorTab(QWidget):
    extract_visio_requested = pyqtSignal(str)
    split_doc_requested = pyqtSignal(str, str, int)
    compare_doc_requested = pyqtSignal(str, str, bool)
    convert_doc_requested = pyqtSignal(str, str)

    def __init__(self):
        super().__init__()
        self._setup_ui()
        self._check_libreoffice_status()

    def _setup_ui(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(15, 15, 15, 15)

        # LibreOffice Missing Notification Banner
        self.status_banner = QWidget()
        banner_layout = QHBoxLayout(self.status_banner)
        banner_layout.setContentsMargins(10, 8, 10, 8)

        self.banner_icon = QLabel("⚠️")
        self.banner_text = QLabel(
            "<b>LibreOffice is not installed.</b> Direct conversion and fallback conversion of legacy .doc files "
            "require LibreOffice."
        )
        self.banner_text.setWordWrap(True)
        self.banner_text.setStyleSheet("color: #B78103;")

        self.install_btn = QPushButton("📥 Install LibreOffice")
        self.install_btn.setStyleSheet(
            "font-weight: bold; background-color: #2e7d32; color: white; padding: 4px 10px; border-radius: 4px;"
        )
        self.install_btn.clicked.connect(lambda: webbrowser.open(LIBREOFFICE_DOWNLOAD_URL))

        banner_layout.addWidget(self.banner_icon)
        banner_layout.addWidget(self.banner_text, 1)
        banner_layout.addWidget(self.install_btn)
        layout.addWidget(self.status_banner)

        # Operation Switcher
        switcher_layout = QHBoxLayout()
        switcher_layout.addWidget(QLabel("<b>⚙️ Operation Type:</b>"))

        self.op_combo = QComboBox()
        self.op_combo.setStyleSheet("padding: 4px; font-weight: bold; font-size: 13px;")
        self.op_combo.addItems([
            "Convert Legacy .doc to .docx (Explicit LibreOffice)",
            "Extract Embedded Visio Diagrams",
            "Subtractive Slicing (Split by Clause)",
            "Compare Documents (Native Word Diff)",
            "Convert Document Format (Auto / Word)",
        ])
        switcher_layout.addWidget(self.op_combo)
        switcher_layout.addStretch()
        layout.addLayout(switcher_layout)

        # Stack
        self.stack = QStackedWidget()

        # Card 0: Explicit LibreOffice Drag & Drop Converter
        self.card_lo_convert = QWidget()
        lo_layout = QVBoxLayout(self.card_lo_convert)
        self.lo_drop = InteractiveDropLabel(
            "📥 Drag & Drop legacy Word 97-2003 (.doc) file(s) here to convert to .docx explicitly via LibreOffice",
            [".doc"],
        )
        self.lo_drop.file_dropped.connect(self._on_lo_dropped)
        lo_layout.addWidget(self.lo_drop)
        self.stack.addWidget(self.card_lo_convert)

        # Card 1: Visio Extractor
        self.card_visio = QWidget()
        visio_layout = QVBoxLayout(self.card_visio)
        self.visio_drop = InteractiveDropLabel(
            "📥 Drag & Drop a .docx file here to extract its Visio components", [".docx"]
        )
        self.visio_drop.file_dropped.connect(
            lambda files: [self.extract_visio_requested.emit(f) for f in files]
        )
        visio_layout.addWidget(self.visio_drop)
        self.stack.addWidget(self.card_visio)

        # Card 2: Splitter
        self.card_split = QWidget()
        split_layout = QVBoxLayout(self.card_split)

        form = QFormLayout()
        self.prefix_input = QLineEdit("6.")
        self.prefix_input.setStyleSheet("padding: 4px; border: 1px solid #CCC; border-radius: 4px;")

        self.depth_input = QSpinBox()
        self.depth_input.setRange(1, 6)
        self.depth_input.setValue(2)
        self.depth_input.setStyleSheet("padding: 4px; border: 1px solid #CCC; border-radius: 4px;")

        form.addRow("Target Clause Prefix:", self.prefix_input)
        form.addRow("Heading Depth Hierarchy:", self.depth_input)
        split_layout.addLayout(form)

        self.split_drop = InteractiveDropLabel(
            "📥 Drag & Drop a .docx file here to slice it into chapters", [".docx"]
        )
        self.split_drop.file_dropped.connect(
            lambda files: [
                self.split_doc_requested.emit(
                    f, self.prefix_input.text().strip(), self.depth_input.value()
                )
                for f in files
            ]
        )
        split_layout.addWidget(self.split_drop)
        self.stack.addWidget(self.card_split)

        # Card 3: The Comparator
        self.card_compare = QWidget()
        compare_layout = QVBoxLayout(self.card_compare)

        panes_layout = QHBoxLayout()
        self.pane_a = DocumentSelectorPane("📄 DOCUMENT A (Original)")
        self.pane_b = DocumentSelectorPane("📄 DOCUMENT B (Revised)")
        panes_layout.addWidget(self.pane_a)
        panes_layout.addWidget(self.pane_b)
        compare_layout.addLayout(panes_layout)

        self.keep_open_cb = QCheckBox("Keep source documents (A and B) open after comparison")
        self.keep_open_cb.setStyleSheet("color: #444; margin-top: 5px;")
        self.keep_open_cb.setChecked(True)
        compare_layout.addWidget(self.keep_open_cb)

        self.run_compare_btn = QPushButton("⚖️ Run Word Comparison")
        self.run_compare_btn.setStyleSheet(
            "font-weight: bold; padding: 10px; background-color: #395396; color: white; border-radius: 4px;"
        )
        self.run_compare_btn.clicked.connect(self._trigger_comparison)
        compare_layout.addWidget(self.run_compare_btn)

        self.stack.addWidget(self.card_compare)

        # Card 4: Generic Word Converter (Auto / Word Fallback)
        self.card_convert = QWidget()
        convert_layout = QVBoxLayout(self.card_convert)

        self.pane_convert = DocumentSelectorPane("📄 DOCUMENT TO CONVERT")
        convert_layout.addWidget(self.pane_convert)

        conv_form = QFormLayout()
        self.format_combo = QComboBox()
        self.format_combo.setStyleSheet("padding: 4px; font-weight: bold;")
        self.format_combo.addItems(["PDF", "DOCX", "HTML", "XPS", "RTF", "TXT"])
        conv_form.addRow("Target Format:", self.format_combo)
        convert_layout.addLayout(conv_form)

        self.run_convert_btn = QPushButton("🔄 Convert Document")
        self.run_convert_btn.setStyleSheet(
            "font-weight: bold; padding: 10px; background-color: #395396; color: white; border-radius: 4px;"
        )
        self.run_convert_btn.clicked.connect(self._trigger_conversion)
        convert_layout.addWidget(self.run_convert_btn)

        self.stack.addWidget(self.card_convert)

        layout.addWidget(self.stack)
        self.setLayout(layout)
        self.op_combo.currentIndexChanged.connect(self.stack.setCurrentIndex)

    def _check_libreoffice_status(self):
        if is_libreoffice_available():
            self.status_banner.setVisible(False)
        else:
            self.status_banner.setStyleSheet(
                "background-color: #FFF3E0; border: 1px solid #FFE082; border-radius: 4px;"
            )
            self.status_banner.setVisible(True)

    def _on_lo_dropped(self, files):
        if not is_libreoffice_available():
            self._check_libreoffice_status()
        for f in files:
            if f:
                # "docx_libreoffice" routes strictly to the LibreOffice CLI engine
                self.convert_doc_requested.emit(f, "docx_libreoffice")

    def _trigger_comparison(self):
        val_a_list = self.pane_a.get_inputs()
        val_b_list = self.pane_b.get_inputs()

        val_a = val_a_list[0] if val_a_list else ""
        val_b = val_b_list[0] if val_b_list else ""
        keep_open = self.keep_open_cb.isChecked()

        if val_a and val_b:
            self.compare_doc_requested.emit(val_a, val_b, keep_open)

    def _trigger_conversion(self):
        source_docs = self.pane_convert.get_inputs()
        target_fmt = self.format_combo.currentText().lower()

        for doc in source_docs:
            if doc:
                self.convert_doc_requested.emit(doc, target_fmt)