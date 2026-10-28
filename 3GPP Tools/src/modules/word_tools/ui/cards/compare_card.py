"""Card: Document Comparison (Native Word Diff)."""

import logging
from pathlib import Path
from typing import Optional

from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtWidgets import (
    QCheckBox,
    QFrame,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from core.ui.ui_components import (
    BUTTON_STYLE_TOOLBAR_SECONDARY,
    CARD_FRAME_STYLE,
)
from modules.word_tools.ui.document_selector_pane import DocumentSelectorPane

logger = logging.getLogger(__name__)


class WordCompareCard(QWidget):
    """Card: Compare two Word documents using Microsoft Word native Track Changes comparison."""

    compare_doc_requested = pyqtSignal(str, str, dict)

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self._setup_ui()

    def _setup_ui(self):
        compare_layout = QVBoxLayout(self)
        compare_layout.setContentsMargins(10, 10, 10, 10)
        compare_layout.setSpacing(10)

        cmp_hdr = QFrame()
        cmp_hdr.setObjectName("cardFrame")
        cmp_hdr.setStyleSheet(CARD_FRAME_STYLE)
        ch_layout = QVBoxLayout(cmp_hdr)
        ch_layout.setContentsMargins(14, 12, 14, 12)
        ch_layout.setSpacing(4)

        cmp_header = QHBoxLayout()
        cmp_title = QLabel("<b>⚖️ Document Revision Comparison (Microsoft Word COM Diff)</b>")
        cmp_title.setStyleSheet("font-size: 13px; color: #1E293B;")
        cmp_header.addWidget(cmp_title)
        cmp_header.addStretch()

        self.btn_swap_slots = QPushButton("⇄ Swap A ↔ B")
        self.btn_swap_slots.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_swap_slots.setToolTip("Swap Original (Document A) and Revised (Document B) documents")
        self.btn_swap_slots.clicked.connect(self._swap_compare_slots)
        cmp_header.addWidget(self.btn_swap_slots)
        ch_layout.addLayout(cmp_header)

        cmp_desc = QLabel(
            "Performs native Track Changes document comparison using Microsoft Word. "
            "Preserves character-level formatting changes, table modifications, and author attributions. "
            "Works seamlessly with local files, remote URLs, and currently active Microsoft Word windows."
        )
        cmp_desc.setWordWrap(True)
        cmp_desc.setStyleSheet("color: #475569; font-size: 11px;")
        ch_layout.addWidget(cmp_desc)
        compare_layout.addWidget(cmp_hdr)

        panes_layout = QHBoxLayout()
        panes_layout.setSpacing(12)
        self.pane_a = DocumentSelectorPane("📄 DOCUMENT A (ORIGINAL BASELINE)")
        self.pane_b = DocumentSelectorPane("📄 DOCUMENT B (REVISED VERSION)")
        panes_layout.addWidget(self.pane_a)
        panes_layout.addWidget(self.pane_b)
        compare_layout.addLayout(panes_layout)

        cmp_opts_frame = QFrame()
        cmp_opts_frame.setObjectName("cardFrame")
        cmp_opts_frame.setStyleSheet(CARD_FRAME_STYLE)
        opts_layout = QHBoxLayout(cmp_opts_frame)
        opts_layout.setContentsMargins(12, 8, 12, 8)
        opts_layout.setSpacing(16)

        opts_layout.addWidget(QLabel("<b>Comparison Options:</b>"))
        self.keep_open_cb = QCheckBox("Keep source documents open in Word after comparison")
        self.keep_open_cb.setChecked(True)
        opts_layout.addWidget(self.keep_open_cb)

        self.chk_cmp_formatting = QCheckBox("Detect formatting changes (fonts, styles, indentation)")
        self.chk_cmp_formatting.setChecked(True)
        opts_layout.addWidget(self.chk_cmp_formatting)
        opts_layout.addStretch()
        compare_layout.addWidget(cmp_opts_frame)

        self.run_compare_btn = QPushButton("🚀 Run Word Comparison")
        self.run_compare_btn.setObjectName("primaryBtn")
        self.run_compare_btn.setCursor(Qt.PointingHandCursor)
        self.run_compare_btn.clicked.connect(self._trigger_comparison)
        compare_layout.addWidget(self.run_compare_btn)
        compare_layout.addStretch()

    def _swap_compare_slots(self):
        inputs_a = self.pane_a.get_inputs()
        inputs_b = self.pane_b.get_inputs()

        self.pane_a.clear_selection()
        self.pane_b.clear_selection()

        if inputs_b:
            self.pane_a.set_file_path(inputs_b[0])
        if inputs_a:
            self.pane_b.set_file_path(inputs_a[0])

    def _trigger_comparison(self):
        inputs_a = self.pane_a.get_inputs()
        inputs_b = self.pane_b.get_inputs()

        if not inputs_a or not inputs_b:
            QMessageBox.warning(
                self,
                "Incomplete Selection",
                "Please select or drop both Document A (Original) and Document B (Revised) to run comparison.",
            )
            return

        options = {
            "keep_open": self.keep_open_cb.isChecked(),
            "detect_formatting": self.chk_cmp_formatting.isChecked(),
        }
        self.compare_doc_requested.emit(inputs_a[0], inputs_b[0], options)
