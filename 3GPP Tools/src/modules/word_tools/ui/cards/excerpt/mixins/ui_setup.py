"""UI setup mixin for `WordExcerptCard`."""

from PyQt5.QtGui import QKeySequence
from PyQt5.QtWidgets import (
    QAbstractItemView,
    QCheckBox,
    QComboBox,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QListWidget,
    QProgressBar,
    QPushButton,
    QShortcut,
    QTreeWidget,
    QVBoxLayout,
)

from core.ui.theme import C
from core.ui.ui_components import (
    BUTTON_STYLE_TOOLBAR_SECONDARY,
    InteractiveDropLabel,
    SEARCH_INPUT_STYLE,
)
from modules.word_tools.core.word_config import WordConfig
from modules.word_tools.core.word_excerpt_naming import NAMING_SCHEMES


class ExcerptCardUiSetupMixin:
    """Builds the heading-scanner drop zone, tree view, naming strip, and extraction actions."""

    def _setup_ui(self):
        excerpt_layout = QVBoxLayout(self)
        excerpt_layout.setContentsMargins(12, 10, 12, 10)
        excerpt_layout.setSpacing(8)

        excerpt_top = QHBoxLayout()
        self.excerpt_order_combo = QComboBox()
        self.excerpt_order_combo.addItem("Filename (natural alphabetical)", "filename")
        self.excerpt_order_combo.addItem("Drop order", "drop")
        excerpt_top.addWidget(QLabel("<b>Source order:</b>"))
        excerpt_top.addWidget(self.excerpt_order_combo)
        excerpt_top.addStretch()
        excerpt_layout.addLayout(excerpt_top)

        self.excerpt_drop = InteractiveDropLabel(
            "📥 Drop one or more .docx files here to scan their headings", [".docx"]
        )
        self.excerpt_drop.file_dropped.connect(self._scan_excerpt_files)
        excerpt_layout.addWidget(self.excerpt_drop)

        source_row = QHBoxLayout()
        self.excerpt_source_list = QListWidget()
        self.excerpt_source_list.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.excerpt_source_list.setMaximumHeight(105)
        self.excerpt_source_list.setToolTip("Select one or more source documents. Press Delete to remove them.")
        source_row.addWidget(self.excerpt_source_list, 1)

        source_buttons = QVBoxLayout()
        self.excerpt_remove_btn = QPushButton("Remove Selected")
        self.excerpt_remove_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.excerpt_remove_btn.setEnabled(False)
        self.excerpt_remove_btn.clicked.connect(self._remove_excerpt_sources)
        self.excerpt_clear_sources_btn = QPushButton("Clear All")
        self.excerpt_clear_sources_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.excerpt_clear_sources_btn.setEnabled(False)
        self.excerpt_clear_sources_btn.clicked.connect(self._clear_excerpt_sources)
        source_buttons.addWidget(self.excerpt_remove_btn)
        source_buttons.addWidget(self.excerpt_clear_sources_btn)
        source_buttons.addStretch()
        source_row.addLayout(source_buttons)
        excerpt_layout.addLayout(source_row)

        self.excerpt_progress_label = QLabel("Ready to scan.")
        self.excerpt_progress_label.setVisible(False)
        self.excerpt_progress = QProgressBar()
        self.excerpt_progress.setRange(0, 100)
        self.excerpt_progress.setValue(0)
        self.excerpt_progress.setTextVisible(True)
        self.excerpt_progress.setVisible(False)
        excerpt_layout.addWidget(self.excerpt_progress_label)
        excerpt_layout.addWidget(self.excerpt_progress)

        self.excerpt_source_list.itemSelectionChanged.connect(
            lambda: self.excerpt_remove_btn.setEnabled(bool(self.excerpt_source_list.selectedItems()))
        )
        self._excerpt_delete_shortcut = QShortcut(QKeySequence.Delete, self.excerpt_source_list)
        self._excerpt_delete_shortcut.activated.connect(self._remove_excerpt_sources)
        self.excerpt_order_combo.currentIndexChanged.connect(self._rescan_excerpt_sources)

        filter_row = QHBoxLayout()
        self.excerpt_filter = QLineEdit()
        self.excerpt_filter.setStyleSheet(SEARCH_INPUT_STYLE)
        self.excerpt_filter.setPlaceholderText("Filter headings (e.g. registration, NWDAF, 6.3.2)...")
        self.excerpt_filter.textChanged.connect(self._filter_excerpt_tree)
        filter_row.addWidget(self.excerpt_filter, 1)

        self.lbl_excerpt_match_count = QLabel()
        self.lbl_excerpt_match_count.setStyleSheet(f"font-size: 12px; color: {C.TEXT_SECONDARY}; padding-right: 4px;")
        filter_row.addWidget(self.lbl_excerpt_match_count)

        self.excerpt_expand_btn = QPushButton("Expand All")
        self.excerpt_expand_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.excerpt_collapse_btn = QPushButton("Collapse All")
        self.excerpt_collapse_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.excerpt_expand_btn.clicked.connect(lambda: self.excerpt_tree.expandAll())
        self.excerpt_collapse_btn.clicked.connect(lambda: self.excerpt_tree.collapseAll())
        filter_row.addWidget(self.excerpt_expand_btn)
        filter_row.addWidget(self.excerpt_collapse_btn)
        excerpt_layout.addLayout(filter_row)

        self.excerpt_tree = QTreeWidget()
        self.excerpt_tree.setHeaderLabels(["Heading", "Source file"])
        header = self.excerpt_tree.header()
        header.setStretchLastSection(False)
        header.setSectionResizeMode(0, QHeaderView.Stretch)
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.excerpt_tree.itemChanged.connect(self._on_tree_item_changed)
        excerpt_layout.addWidget(self.excerpt_tree, 1)

        naming_frame = QFrame()
        naming_frame.setStyleSheet(f"""
            QFrame {{
                background-color: {C.BG_SURFACE};
                border: 1px solid {C.BORDER_SUBTLE};
                border-radius: 6px;
            }}
        """)
        naming_layout = QHBoxLayout(naming_frame)
        naming_layout.setContentsMargins(10, 6, 10, 6)
        naming_layout.setSpacing(8)

        lbl_naming = QLabel("<b>🏷️ Naming:</b>")
        lbl_naming.setStyleSheet(f"color: {C.TEXT_BODY}; font-size: 12px;")
        naming_layout.addWidget(lbl_naming)

        self.excerpt_naming_combo = QComboBox()
        self.excerpt_naming_combo.setStyleSheet("font-size: 12px;")
        for scheme_id, scheme_label in NAMING_SCHEMES:
            self.excerpt_naming_combo.addItem(scheme_label, scheme_id)
        current_scheme = WordConfig.get_excerpt_naming_scheme()
        idx = self.excerpt_naming_combo.findData(current_scheme)
        if idx >= 0:
            self.excerpt_naming_combo.setCurrentIndex(idx)
        self.excerpt_naming_combo.currentIndexChanged.connect(self._on_naming_scheme_changed)
        naming_layout.addWidget(self.excerpt_naming_combo)

        lbl_file = QLabel("<b>Export File:</b>")
        lbl_file.setStyleSheet(f"color: {C.TEXT_SECONDARY}; font-size: 12px; margin-left: 4px;")
        naming_layout.addWidget(lbl_file)

        self.excerpt_filename_edit = QLineEdit()
        self.excerpt_filename_edit.setStyleSheet(SEARCH_INPUT_STYLE + "; font-size: 12px; padding: 4px 8px;")
        self.excerpt_filename_edit.setPlaceholderText("Auto-generated differentiated filename (editable)...")
        self.excerpt_filename_edit.textEdited.connect(self._on_filename_text_edited)
        naming_layout.addWidget(self.excerpt_filename_edit, 1)

        self.excerpt_reset_name_btn = QPushButton("🔄 Reset")
        self.excerpt_reset_name_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY + "; font-size: 12px; padding: 4px 10px;")
        self.excerpt_reset_name_btn.setToolTip("Reset filename to auto-generated differentiated name based on selected clauses")
        self.excerpt_reset_name_btn.clicked.connect(self._reset_filename_to_auto)
        naming_layout.addWidget(self.excerpt_reset_name_btn)

        excerpt_layout.addWidget(naming_frame)

        action_row = QHBoxLayout()
        self.excerpt_select_all_btn = QPushButton("Select All")
        self.excerpt_select_all_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.excerpt_select_all_btn.clicked.connect(self._select_all_excerpt_headings)
        action_row.addWidget(self.excerpt_select_all_btn)

        self.excerpt_select_l1_btn = QPushButton("Select Level 1 Only")
        self.excerpt_select_l1_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.excerpt_select_l1_btn.clicked.connect(self._select_level1_excerpt_headings)
        action_row.addWidget(self.excerpt_select_l1_btn)

        self.excerpt_clear_btn = QPushButton("Clear Selection")
        self.excerpt_clear_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.excerpt_clear_btn.clicked.connect(self._clear_excerpt_selection)
        action_row.addWidget(self.excerpt_clear_btn)

        action_row.addStretch()

        self.chk_also_markdown = QCheckBox("Also export as Markdown (with figures)")
        self.chk_also_markdown.setToolTip(
            "Automatically convert the extracted sections to Markdown (.md) and extract all embedded figures and diagrams into a companion media folder."
        )
        self.chk_also_markdown.setStyleSheet(f"color: {C.TEXT_BODY}; font-weight: 500; margin-right: 8px;")
        action_row.addWidget(self.chk_also_markdown)

        self.excerpt_extract_btn = QPushButton("✂️ Extract Selected Sections")
        self.excerpt_extract_btn.setObjectName("primaryBtn")
        self.excerpt_extract_btn.setEnabled(False)
        self.excerpt_extract_btn.clicked.connect(self._extract_selected_sections)
        action_row.addWidget(self.excerpt_extract_btn)
        excerpt_layout.addLayout(action_row)
