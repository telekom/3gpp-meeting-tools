# --- File: src/modules/word_tools/ui/document_selector_pane.py ---
import logging
from pathlib import Path
from typing import List, Optional

from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtWidgets import (
    QComboBox,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)
import pythoncom
import win32com.client

from core.ui.ui_components import (
    BUTTON_STYLE_TOOLBAR_SECONDARY,
    CARD_FRAME_STYLE,
    SEARCH_INPUT_STYLE, InteractiveDropLabel,
)

logger = logging.getLogger(__name__)

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
        layout.setSpacing(6)

        lbl = QLabel(f"<b>{self.title}</b>")
        lbl.setStyleSheet("color: #1E293B; margin-bottom: 2px;")
        layout.addWidget(lbl)

        self.tabs = QTabWidget()
        self.tabs.setObjectName("selector_tabs")
        self.drop_tab = QWidget()
        drop_layout = QVBoxLayout(self.drop_tab)
        self.drop_zone = InteractiveDropLabel("Drop .docx here", [".docx"])
        self.drop_zone.file_dropped.connect(self._on_drop)
        drop_layout.addWidget(self.drop_zone)
        self.tabs.addTab(self.drop_tab, "📁 Local")

        self.open_tab = QWidget()
        open_layout = QVBoxLayout(self.open_tab)
        open_layout.setSpacing(6)
        self.open_combo = QComboBox()
        self.refresh_btn = QPushButton("↻ Refresh Active Documents")
        self.refresh_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.refresh_btn.clicked.connect(self.poll_open_documents)

        open_layout.addWidget(QLabel("Select an open Word document:"))
        open_layout.addWidget(self.open_combo)
        open_layout.addWidget(self.refresh_btn)
        open_layout.addStretch()
        self.tabs.addTab(self.open_tab, "🖥️ Open Docs")

        self.url_tab = QWidget()
        url_layout = QVBoxLayout(self.url_tab)
        url_layout.setSpacing(6)

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

