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
    QMessageBox,
    QPushButton,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

from core.ui.ui_components import (
    BUTTON_STYLE_TOOLBAR_SECONDARY,
    CARD_FRAME_STYLE,
    SEARCH_INPUT_STYLE,
    InteractiveDropLabel,
)
from modules.word_tools.core.document_opener import (
    get_active_word_document,
    list_open_word_documents,
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

        header_row = QHBoxLayout()
        lbl = QLabel(f"<b>{self.title}</b>")
        lbl.setStyleSheet("color: #1E293B; margin-bottom: 2px;")
        header_row.addWidget(lbl)
        header_row.addStretch()

        self.quick_active_btn = QPushButton("🪟 Use Active Word")
        self.quick_active_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.quick_active_btn.setCursor(Qt.PointingHandCursor)
        self.quick_active_btn.setToolTip(
            "Select the document currently active in Microsoft Word (supports unsaved documents via automatic snapshot)."
        )
        self.quick_active_btn.clicked.connect(self.select_active_word_document)
        header_row.addWidget(self.quick_active_btn)
        layout.addLayout(header_row)

        self.tabs = QTabWidget()
        self.tabs.setObjectName("selector_tabs")
        self.drop_tab = QWidget()
        drop_layout = QVBoxLayout(self.drop_tab)
        self.drop_zone = InteractiveDropLabel("Drop .docx here", [".docx", ".doc"])
        self.drop_zone.file_dropped.connect(self._on_drop)
        drop_layout.addWidget(self.drop_zone)
        self.tabs.addTab(self.drop_tab, "📁 Local")

        self.open_tab = QWidget()
        open_layout = QVBoxLayout(self.open_tab)
        open_layout.setSpacing(6)
        self.open_combo = QComboBox()

        open_btn_row = QHBoxLayout()
        self.select_active_btn = QPushButton("🪟 Select Active Document")
        self.select_active_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.select_active_btn.setCursor(Qt.PointingHandCursor)
        self.select_active_btn.setToolTip("Automatically select the foreground document in Microsoft Word.")
        self.select_active_btn.clicked.connect(self.select_active_word_document)

        self.refresh_btn = QPushButton("↻ Refresh Open Documents")
        self.refresh_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.refresh_btn.setCursor(Qt.PointingHandCursor)
        self.refresh_btn.clicked.connect(self.poll_open_documents)

        open_btn_row.addWidget(self.select_active_btn)
        open_btn_row.addWidget(self.refresh_btn)

        open_layout.addWidget(QLabel("Select an open Word document:"))
        open_layout.addWidget(self.open_combo)
        open_layout.addLayout(open_btn_row)
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
        docs = list_open_word_documents()
        if not docs:
            self.open_combo.addItem("No open documents detected.", "")
            return

        active_idx = 0
        for idx, d in enumerate(docs):
            name = d.get("name", "Untitled")
            full_name = d.get("full_name") or name
            is_act = d.get("is_active", False)
            is_mod = d.get("modified", False)
            prefix = "🟢 [Active] " if is_act else "📄 "
            suffix = " *" if is_mod else ""
            self.open_combo.addItem(f"{prefix}{name}{suffix}", full_name)
            if is_act:
                active_idx = idx

        self.open_combo.setCurrentIndex(active_idx)

    def select_active_word_document(self):
        """Switches to the Open Docs tab, refreshes, and selects the foreground Word document."""
        self.tabs.setCurrentIndex(1)
        self.poll_open_documents()
        active_doc = get_active_word_document(allow_unsaved_snapshot=True)
        if not active_doc:
            QMessageBox.warning(
                self,
                "No Active Word Document",
                "No active Microsoft Word document could be detected.\n\n"
                "Please ensure Microsoft Word is running and a document is open.",
            )

    def get_inputs(self) -> list:
        idx = self.tabs.currentIndex()
        if idx == 0:
            return self.selected_files
        elif idx == 1:
            ident = self.open_combo.currentData()
            if not ident:
                return []
            # Resolve through get_active_word_document so unsaved or modified open docs
            # are safely snapshotted to a temporary file before comparison/conversion.
            resolved = get_active_word_document(
                allow_unsaved_snapshot=True,
                target_identifier=str(ident),
            )
            if resolved and resolved.get("path"):
                return [resolved["path"]]
            return [ident] if Path(str(ident)).exists() else []
        elif idx == 2:
            url = self.url_input.text().strip()
            return [url] if url else []
        return []

    def _on_tab_changed(self, index):
        if index == 1:
            self.poll_open_documents()
