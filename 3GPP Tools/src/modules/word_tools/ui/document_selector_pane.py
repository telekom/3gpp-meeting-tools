# --- File: src/modules/word_tools/ui/document_selector_pane.py ---
import logging
import os
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


class ClickableDropLabel(InteractiveDropLabel):
    """An InteractiveDropLabel that also emits a clicked signal on mouse press."""

    clicked = pyqtSignal()

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.clicked.emit()
        super().mousePressEvent(event)


class DocumentSelectorPane(QWidget):
    """
    A symmetric, reusable widget handling Local, Open, and URL inputs.
    Includes click-to-browse, explicit Browse buttons, metadata preview, and active Word staging.
    """

    file_selected = pyqtSignal(list)

    def __init__(self, title: str, accepted_extensions=None):
        super().__init__()
        self.title = title
        self.accepted_extensions = accepted_extensions or [".docx", ".doc", ".docm"]
        self.selected_files = []
        self._setup_ui()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(6)

        header_row = QHBoxLayout()
        self.lbl_title = QLabel(f"<b>{self.title}</b>")
        self.lbl_title.setStyleSheet("color: #1E293B; margin-bottom: 2px;")
        header_row.addWidget(self.lbl_title)
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

        # Tab 1: Local File
        self.drop_tab = QWidget()
        drop_layout = QVBoxLayout(self.drop_tab)
        drop_layout.setContentsMargins(8, 8, 8, 8)
        drop_layout.setSpacing(6)

        self.drop_zone = ClickableDropLabel(
            "📥 Drag & Drop document here\n(or click to browse)",
            self.accepted_extensions,
        )
        self.drop_zone.setCursor(Qt.PointingHandCursor)
        self.drop_zone.file_dropped.connect(self._on_drop)
        self.drop_zone.clicked.connect(self._browse_local_file)
        drop_layout.addWidget(self.drop_zone)

        # File actions bar (Browse, Clear, Info)
        local_bar = QHBoxLayout()
        self.btn_browse = QPushButton("📂 Browse File...")
        self.btn_browse.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_browse.setCursor(Qt.PointingHandCursor)
        self.btn_browse.clicked.connect(self._browse_local_file)
        local_bar.addWidget(self.btn_browse)

        self.btn_clear = QPushButton("✕ Clear")
        self.btn_clear.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_clear.setCursor(Qt.PointingHandCursor)
        self.btn_clear.setEnabled(False)
        self.btn_clear.clicked.connect(self.clear_selection)
        local_bar.addWidget(self.btn_clear)

        self.lbl_meta = QLabel()
        self.lbl_meta.setStyleSheet("font-size: 11px; color: #475569;")
        local_bar.addWidget(self.lbl_meta, 1)

        drop_layout.addLayout(local_bar)
        self.tabs.addTab(self.drop_tab, "📁 Local")

        # Tab 2: Open Word Documents
        self.open_tab = QWidget()
        open_layout = QVBoxLayout(self.open_tab)
        open_layout.setContentsMargins(8, 8, 8, 8)
        open_layout.setSpacing(6)
        self.open_combo = QComboBox()
        self.open_combo.currentIndexChanged.connect(self._on_open_combo_changed)

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
        open_btn_row.addStretch()

        open_layout.addWidget(QLabel("Select an open Word document:"))
        open_layout.addWidget(self.open_combo)
        open_layout.addLayout(open_btn_row)
        open_layout.addStretch()
        self.tabs.addTab(self.open_tab, "🖥️ Open Docs")

        # Tab 3: URL
        self.url_tab = QWidget()
        url_layout = QVBoxLayout(self.url_tab)
        url_layout.setContentsMargins(8, 8, 8, 8)
        url_layout.setSpacing(6)

        self.url_input = QLineEdit()
        self.url_input.setStyleSheet(SEARCH_INPUT_STYLE)
        self.url_input.setPlaceholderText("https://... (Word document or SharePoint link)")

        url_layout.addWidget(QLabel("Paste document URL:"))
        url_layout.addWidget(self.url_input)
        url_layout.addStretch()
        self.tabs.addTab(self.url_tab, "🌐 URL")

        layout.addWidget(self.tabs)
        self.tabs.currentChanged.connect(self._on_tab_changed)

    def _browse_local_file(self):
        filter_str = f"Word Documents (*{' *'.join(self.accepted_extensions)});;All Files (*.*)"
        files, _ = QFileDialog.getOpenFileNames(
            self,
            f"Select {self.title}",
            "",
            filter_str,
        )
        if files:
            self._set_files(files)

    def _on_drop(self, files):
        if files:
            self._set_files(files)

    def _set_files(self, files: List[str]):
        self.selected_files = [str(Path(f).resolve()) for f in files]
        self.tabs.setCurrentIndex(0)
        if len(files) == 1:
            p = Path(files[0])
            size_str = ""
            if p.exists():
                size_kb = p.stat().st_size / 1024
                size_str = f" ({size_kb:.1f} KB)" if size_kb < 1024 else f" ({size_kb/1024:.2f} MB)"

            label_str = ""
            try:
                from modules.word_tools.core.sensitivity_label import get_document_sensitivity_label
                sl = get_document_sensitivity_label(p)
                if sl and sl.get("name"):
                    label_str = f" | 🛡️ {sl['name']}"
            except Exception:
                pass

            self.drop_zone.set_state("ready", f"Ready:\n{p.name}{size_str}{label_str}")
            self.lbl_meta.setText(f"Selected: {p.name}{label_str}")
        else:
            self.drop_zone.set_state("ready", f"Ready:\n{len(files)} files queued for batch")
            self.lbl_meta.setText(f"{len(files)} files selected")
        self.btn_clear.setEnabled(True)
        self.file_selected.emit(self.selected_files)

    def clear_selection(self):
        self.selected_files = []
        self.drop_zone.set_state("idle", "📥 Drag & Drop document here\n(or click to browse)")
        self.lbl_meta.setText("")
        self.btn_clear.setEnabled(False)

    def set_file_path(self, path_str: str):
        """Programmatically set a file path."""
        if path_str and Path(path_str).exists():
            self._set_files([path_str])

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
            return
        if active_doc.get("path"):
            self.file_selected.emit([active_doc["path"]])

    def _on_open_combo_changed(self, index: int):
        if self.tabs.currentIndex() == 1:
            inputs = self.get_inputs()
            if inputs:
                self.file_selected.emit(inputs)

    def get_inputs(self) -> list:
        idx = self.tabs.currentIndex()
        if idx == 0:
            return self.selected_files
        elif idx == 1:
            ident = self.open_combo.currentData()
            if not ident:
                return []
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
            inputs = self.get_inputs()
            if inputs:
                self.file_selected.emit(inputs)
