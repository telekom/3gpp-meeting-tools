from pathlib import Path

from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
                             QLabel, QComboBox, QLineEdit, QSpinBox,
                             QFormLayout, QStackedWidget, QTabWidget, QCheckBox)
from PyQt5.QtCore import pyqtSignal
import win32com.client
import pythoncom

from core.ui.ui_components import InteractiveDropLabel

class DocumentSelectorPane(QWidget):
    """A symmetric, reusable widget handling Local, Open, and URL inputs."""

    def __init__(self, title: str):
        super().__init__()
        self.title = title
        self.selected_file = ""
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
        self.drop_zone = InteractiveDropLabel("Drop .docx here", ['.docx'])
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
            self.selected_file = files[0]
            self.drop_zone.set_state("ready", f"Ready:\n{Path(self.selected_file).name}")

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

    def get_input(self) -> str:
        idx = self.tabs.currentIndex()
        if idx == 0:
            return self.selected_file
        elif idx == 1:
            return self.open_combo.currentData() or ""
        elif idx == 2:
            return self.url_input.text().strip()
        return ""

    def _on_tab_changed(self, index):
        if index == 1:
            self.poll_open_documents()

class WordExtractorTab(QWidget):
    extract_visio_requested = pyqtSignal(str)
    split_doc_requested = pyqtSignal(str, str, int)
    compare_doc_requested = pyqtSignal(str, str, bool)

    def __init__(self):
        super().__init__()
        self._setup_ui()

    def _setup_ui(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(15, 15, 15, 15)

        # Switcher
        switcher_layout = QHBoxLayout()
        switcher_layout.addWidget(QLabel("<b>⚙️ Operation Type:</b>"))

        self.op_combo = QComboBox()
        self.op_combo.setStyleSheet("padding: 4px; font-weight: bold; font-size: 13px;")
        self.op_combo.addItems([
            "Extract Embedded Visio Diagrams",
            "Subtractive Slicing (Split by Clause)",
            "Compare Documents (Native Word Diff)"
        ])
        switcher_layout.addWidget(self.op_combo)
        switcher_layout.addStretch()
        layout.addLayout(switcher_layout)

        # Stack
        self.stack = QStackedWidget()

        # Card 1: Visio Extractor
        self.card_visio = QWidget()
        visio_layout = QVBoxLayout(self.card_visio)
        self.visio_drop = InteractiveDropLabel("📥 Drag & Drop a .docx file here to extract its Visio components", ['.docx'])
        self.visio_drop.file_dropped.connect(lambda files: [self.extract_visio_requested.emit(f) for f in files])
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

        self.split_drop = InteractiveDropLabel("📥 Drag & Drop a .docx file here to slice it into chapters", ['.docx'])
        self.split_drop.file_dropped.connect(
            lambda files: [self.split_doc_requested.emit(f, self.prefix_input.text().strip(), self.depth_input.value()) for f in files])
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
            "font-weight: bold; padding: 10px; background-color: #395396; color: white; border-radius: 4px;")
        self.run_compare_btn.clicked.connect(self._trigger_comparison)
        compare_layout.addWidget(self.run_compare_btn)

        self.stack.addWidget(self.card_compare)

        layout.addWidget(self.stack)
        self.setLayout(layout)
        self.op_combo.currentIndexChanged.connect(self.stack.setCurrentIndex)

    def _trigger_comparison(self):
        val_a = self.pane_a.get_input()
        val_b = self.pane_b.get_input()
        keep_open = self.keep_open_cb.isChecked()

        if val_a and val_b:
            self.compare_doc_requested.emit(val_a, val_b, keep_open)