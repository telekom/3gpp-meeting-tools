"""UI setup mixin for SpecClauseExcerptDialog."""

from __future__ import annotations

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QCheckBox,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QProgressBar,
    QPushButton,
    QTreeWidget,
    QVBoxLayout,
)

from core.ui.theme import C
from core.ui.ui_components import (
    BADGE_STYLE_INFO,
    BADGE_STYLE_MUTED,
    BADGE_STYLE_PRIMARY,
    BUTTON_STYLE_TOOLBAR_SECONDARY,
    SEARCH_INPUT_STYLE,
)


class SpecClauseUiSetupMixin:
    """Builds the header, progress strip, filter bar, clause tree, and 2-tier footer."""

    def _setup_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(16, 16, 16, 16)
        main_layout.setSpacing(10)

        # 1. Header Card
        header_card = QFrame()
        header_card.setObjectName("cardFrame")
        header_layout = QVBoxLayout(header_card)
        header_layout.setContentsMargins(14, 12, 14, 12)
        header_layout.setSpacing(6)

        title_row = QHBoxLayout()
        type_badge = QLabel(f"<b>{self.spec_type}</b>")
        type_badge.setStyleSheet(BADGE_STYLE_INFO)

        num_label = QLabel(f"<b>{self.spec_num}</b>")
        num_label.setStyleSheet(f"font-size: 16px; color: {C.TEXT_STRONG}; font-weight: bold;")

        ver_badge = QLabel(f"<b>v{self.version}</b>")
        ver_badge.setStyleSheet(BADGE_STYLE_PRIMARY)

        title_row.addWidget(type_badge)
        title_row.addWidget(num_label)
        title_row.addWidget(ver_badge)
        title_row.addStretch()

        self.source_summary_badge = QLabel("⏳ Initializing...")
        self.source_summary_badge.setStyleSheet(BADGE_STYLE_MUTED)
        title_row.addWidget(self.source_summary_badge)

        header_layout.addLayout(title_row)

        desc_label = QLabel(
            "Select specific clauses or subclauses below to extract into standalone Word (.docx) or Markdown (.md) documents."
        )
        desc_label.setStyleSheet(f"color: {C.TEXT_SECONDARY}; font-size: 12px;")
        header_layout.addWidget(desc_label)
        main_layout.addWidget(header_card)

        # 2. Progress / Status Strip
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setFixedHeight(14)
        self.progress_bar.setTextVisible(False)
        self.progress_bar.setStyleSheet(f"""
            QProgressBar {{
                border: 1px solid {C.BORDER_SUBTLE};
                border-radius: 4px;
                background-color: {C.BG_MUTED};
            }}
            QProgressBar::chunk {{
                background-color: {C.ACCENT_BRIGHT};
                border-radius: 3px;
            }}
        """)

        self.status_label = QLabel("Preparing specification documents...")
        self.status_label.setStyleSheet(
            f"color: {C.TEXT_SECONDARY}; font-size: 11px; font-style: italic;"
        )

        progress_box = QVBoxLayout()
        progress_box.setSpacing(4)
        progress_box.addWidget(self.status_label)
        progress_box.addWidget(self.progress_bar)
        main_layout.addLayout(progress_box)

        # 3. Filter & Tree Controls
        filter_layout = QHBoxLayout()
        filter_layout.setSpacing(6)

        self.filter_input = QLineEdit()
        self.filter_input.setPlaceholderText(
            "🔍 Filter clauses by number or title (e.g. 5.2, registration, NWDAF)..."
        )
        self.filter_input.setStyleSheet(SEARCH_INPUT_STYLE)
        self.filter_input.setClearButtonEnabled(True)
        self.filter_input.textChanged.connect(self._filter_tree)
        filter_layout.addWidget(self.filter_input, 1)

        self.lbl_filter_matches = QLabel()
        self.lbl_filter_matches.setStyleSheet(
            f"font-size: 11px; color: {C.TEXT_MUTED}; padding-right: 4px;"
        )
        filter_layout.addWidget(self.lbl_filter_matches)

        self.btn_expand_all = QPushButton("Expand All")
        self.btn_expand_all.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_expand_all.setCursor(Qt.PointingHandCursor)
        self.btn_expand_all.clicked.connect(lambda: self.clause_tree.expandAll())
        filter_layout.addWidget(self.btn_expand_all)

        self.btn_collapse_all = QPushButton("Collapse All")
        self.btn_collapse_all.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_collapse_all.setCursor(Qt.PointingHandCursor)
        self.btn_collapse_all.clicked.connect(lambda: self.clause_tree.collapseAll())
        filter_layout.addWidget(self.btn_collapse_all)

        main_layout.addLayout(filter_layout)

        # 4. Hierarchical Clause Tree
        self.clause_tree = QTreeWidget()
        self.clause_tree.setHeaderLabels(["Clause / Heading", "Source Document"])
        header = self.clause_tree.header()
        header.setStretchLastSection(False)
        header.setSectionResizeMode(0, QHeaderView.Stretch)
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.clause_tree.setStyleSheet(f"""
            QTreeWidget {{
                background-color: white;
                border: 1px solid {C.BORDER_STRONG};
                border-radius: 6px;
                padding: 4px;
                font-size: 12px;
            }}
            QTreeWidget::item {{
                padding: 3px 0;
            }}
            QTreeWidget::item:hover {{
                background-color: {C.BG_CANVAS};
            }}
            QTreeWidget::item:selected {{
                background-color: {C.BORDER_SUBTLE};
                color: {C.TEXT_PRIMARY};
            }}
        """)
        self.clause_tree.itemChanged.connect(self._on_tree_item_changed)
        main_layout.addWidget(self.clause_tree, 1)

        # 5. Bottom Footer (2-Tier Options & Actions)
        footer_layout = QVBoxLayout()
        footer_layout.setSpacing(10)
        footer_layout.setContentsMargins(0, 4, 0, 0)

        # Row 1: Selection Status & Export Options
        options_layout = QHBoxLayout()
        options_layout.setSpacing(10)

        self.btn_clear = QPushButton("Clear Selection")
        self.btn_clear.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_clear.setCursor(Qt.PointingHandCursor)
        self.btn_clear.clicked.connect(self._clear_selection)
        self.btn_clear.setEnabled(False)
        options_layout.addWidget(self.btn_clear)

        self.lbl_selection_count = QLabel("No clauses selected")
        self.lbl_selection_count.setStyleSheet(
            f"color: {C.TEXT_MUTED}; font-size: 12px; font-weight: 500;"
        )
        options_layout.addWidget(self.lbl_selection_count)

        options_layout.addStretch()

        self.open_after_cb = QCheckBox("Open when finished")
        self.open_after_cb.setChecked(True)
        self.open_after_cb.setStyleSheet(f"color: {C.TEXT_BODY}; font-size: 12px;")
        self.open_after_cb.setToolTip(
            "Open the extracted document in Microsoft Word or default viewer when extraction completes."
        )
        options_layout.addWidget(self.open_after_cb)

        self.chk_also_markdown = QCheckBox("Also export as Markdown (with figures)")
        self.chk_also_markdown.setStyleSheet(f"color: {C.TEXT_BODY}; font-size: 12px;")
        self.chk_also_markdown.setToolTip(
            "Automatically convert the extracted clause(s) to Markdown (.md) and extract all embedded figures and diagrams into a companion media folder."
        )
        options_layout.addWidget(self.chk_also_markdown)

        footer_layout.addLayout(options_layout)

        # Row 2: Primary & Secondary Action Buttons
        action_layout = QHBoxLayout()
        action_layout.setSpacing(8)
        action_layout.addStretch()

        self.btn_cr_pcr = QPushButton("📝 Draft CR / pCR...")
        self.btn_cr_pcr.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_cr_pcr.setCursor(Qt.PointingHandCursor)
        self.btn_cr_pcr.setEnabled(False)
        self.btn_cr_pcr.clicked.connect(self._open_cr_pcr_from_selection)
        action_layout.addWidget(self.btn_cr_pcr)

        self.btn_extract = QPushButton("✂️ Extract Selected Clauses...")
        self.btn_extract.setCursor(Qt.PointingHandCursor)
        self.btn_extract.setEnabled(False)
        self.btn_extract.setStyleSheet(f"""
            QPushButton {{
                font-weight: bold;
                background-color: {C.ACCENT_BRIGHT};
                color: white;
                padding: 7px 20px;
                border-radius: 4px;
                border: 1px solid {C.ACCENT_BRIGHT_BORDER};
            }}
            QPushButton:hover {{
                background-color: {C.ACCENT_BRIGHT_BORDER};
            }}
            QPushButton:disabled {{
                background-color: {C.BORDER_STRONG};
                color: {C.TEXT_FAINT};
                border: 1px solid {C.BORDER_STRONG};
            }}
        """)
        self.btn_extract.clicked.connect(self._extract_selected_clauses)
        action_layout.addWidget(self.btn_extract)

        self.btn_close = QPushButton("Close")
        self.btn_close.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_close.setCursor(Qt.PointingHandCursor)
        self.btn_close.clicked.connect(self.close)
        action_layout.addWidget(self.btn_close)

        footer_layout.addLayout(action_layout)
        main_layout.addLayout(footer_layout)
