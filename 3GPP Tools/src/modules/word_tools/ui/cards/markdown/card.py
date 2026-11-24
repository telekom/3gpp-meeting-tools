"""Bidirectional Markdown <-> Word Studio card composed from modular mixins."""

from __future__ import annotations

from typing import Optional

from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtWidgets import (
    QFrame,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QStackedWidget,
    QVBoxLayout,
    QWidget,
)

from core.ui.theme import C
from core.ui.ui_components import (
    BUTTON_STYLE_TOOLBAR_SECONDARY,
    CARD_FRAME_STYLE,
    TOGGLE_PILL_STYLE,
)
from .mixins import MarkdownActionsMixin, MarkdownM2wUiMixin, MarkdownW2mUiMixin


class WordMarkdownCard(
    MarkdownW2mUiMixin,
    MarkdownM2wUiMixin,
    MarkdownActionsMixin,
    QWidget,
):
    """
    Card: Bidirectional Markdown ⇄ Word Studio.
    Supports fast headless conversion of Word documents to Markdown (with revision tracking & media extraction)
    and compilation of Markdown documents to Word (.docx) with embedded media, PlantUML diagrams, and 3GPP styles.
    """

    convert_markdown_requested = pyqtSignal(str, dict)
    convert_markdown_to_word_requested = pyqtSignal(str, dict)

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self._last_md_output_path: Optional[str] = None
        self._last_converted_md: Optional[str] = None
        self._last_docx_output_path: Optional[str] = None
        self._awaiting_w2m: bool = False
        self._awaiting_m2w: bool = False
        self._routing_files: bool = False
        self._setup_ui()

        # Backward compatibility aliases
        self.pane_md = self.pane_markdown

    def _setup_ui(self):
        root_layout = QVBoxLayout(self)
        root_layout.setContentsMargins(10, 10, 10, 10)
        root_layout.setSpacing(10)

        # Direction Navigation Strip
        nav_frame = QFrame()
        nav_frame.setObjectName("cardFrame")
        nav_frame.setStyleSheet(CARD_FRAME_STYLE)
        nav_layout = QHBoxLayout(nav_frame)
        nav_layout.setContentsMargins(12, 6, 12, 6)
        nav_layout.setSpacing(8)

        lbl_dir = QLabel("<b>Direction:</b>")
        lbl_dir.setStyleSheet(f"color: {C.TEXT_SECONDARY}; font-size: 11px;")
        nav_layout.addWidget(lbl_dir)

        self.btn_mode_w2m = QPushButton("📄 Word → Markdown")
        self.btn_mode_m2w = QPushButton("📝 Markdown → Word")
        for btn in (self.btn_mode_w2m, self.btn_mode_m2w):
            btn.setCheckable(True)
            btn.setStyleSheet(TOGGLE_PILL_STYLE)
            btn.setCursor(Qt.PointingHandCursor)

        self.btn_mode_w2m.setChecked(True)
        self.btn_mode_w2m.clicked.connect(lambda: self.set_studio_mode(0))
        self.btn_mode_m2w.clicked.connect(lambda: self.set_studio_mode(1))
        nav_layout.addWidget(self.btn_mode_w2m)
        nav_layout.addWidget(self.btn_mode_m2w)
        nav_layout.addStretch()

        self.btn_switch_direction = QPushButton("⇄ Switch Direction")
        self.btn_switch_direction.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_switch_direction.setCursor(Qt.PointingHandCursor)
        self.btn_switch_direction.setToolTip(
            "Quickly toggle between Word-to-Markdown and Markdown-to-Word modes."
        )
        self.btn_switch_direction.clicked.connect(self._toggle_direction)
        nav_layout.addWidget(self.btn_switch_direction)

        root_layout.addWidget(nav_frame)

        # Stacked Pages: [0] Word -> Markdown, [1] Markdown -> Word
        self.studio_stack = QStackedWidget()

        page_w2m = QWidget()
        self._setup_page_w2m(page_w2m)
        self.studio_stack.addWidget(page_w2m)

        page_m2w = QWidget()
        self._setup_page_m2w(page_m2w)
        self.studio_stack.addWidget(page_m2w)

        root_layout.addWidget(self.studio_stack)
