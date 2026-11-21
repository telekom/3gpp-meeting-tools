"""UI layout construction mixin for CrPcrDialog."""

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QButtonGroup,
    QCheckBox,
    QComboBox,
    QFrame,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QProgressBar,
    QPushButton,
    QRadioButton,
    QScrollArea,
    QSplitter,
    QStackedWidget,
    QTextEdit,
    QTreeWidget,
    QVBoxLayout,
    QWidget,
)

from core.ui.theme import C
from core.ui.ui_components import (
    BADGE_STYLE_MUTED,
    BADGE_STYLE_PRIMARY,
    BUTTON_STYLE_TOOLBAR_SECONDARY,
    SEARCH_INPUT_STYLE,
)
from modules.word_tools.core.word_config import WordConfig
from modules.word_tools.core.word_excerpt_naming import CR_NAMING_SCHEMES
from modules.word_tools.ui.cr_pcr.styles import BUTTON_STYLE_PRIMARY


class CrPcrUiSetupMixin:
    """Builds the contribution mode selector, template banner, form, clause tree, and actions."""

    def _setup_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(14, 14, 14, 14)
        main_layout.setSpacing(10)

        # -------------------------------------------------------------
        # 1. Top Mode Selector Card
        # -------------------------------------------------------------
        top_card = QFrame()
        top_card.setObjectName("cardFrame")
        top_layout = QHBoxLayout(top_card)
        top_layout.setContentsMargins(14, 10, 14, 10)
        top_layout.setSpacing(16)

        mode_label = QLabel("<b>Contribution Type:</b>")
        mode_label.setStyleSheet(f"font-size: 13px; color: {C.TEXT_STRONG};")
        top_layout.addWidget(mode_label)

        self.radio_pcr = QRadioButton("Pseudo-CR (pCR / Discussion && Proposal)")
        self.radio_pcr.setStyleSheet(f"font-size: 13px; font-weight: bold; color: {C.OFFICE_BLUE};")
        self.radio_cr = QRadioButton("Formal Change Request (CR Template)")
        self.radio_cr.setStyleSheet(f"font-size: 13px; font-weight: bold; color: {C.OFFICE_BLUE};")

        self.mode_group = QButtonGroup(self)
        self.mode_group.addButton(self.radio_pcr)
        self.mode_group.addButton(self.radio_cr)

        if self.current_mode == "cr":
            self.radio_cr.setChecked(True)
        else:
            self.radio_pcr.setChecked(True)

        self.radio_pcr.toggled.connect(self._on_mode_changed)
        top_layout.addWidget(self.radio_pcr)
        top_layout.addWidget(self.radio_cr)
        top_layout.addStretch()

        spec_badge = QLabel(f"📄 {self.spec_type} {self.spec_num} v{self.version}")
        spec_badge.setStyleSheet(BADGE_STYLE_PRIMARY)
        top_layout.addWidget(spec_badge)

        self.status_badge = QLabel("⏳ Initializing...")
        self.status_badge.setStyleSheet(BADGE_STYLE_MUTED)
        top_layout.addWidget(self.status_badge)

        main_layout.addWidget(top_card)

        # -------------------------------------------------------------
        # 1b. 3GPP Word Template Status Banner (3gpp_70.dotx)
        # -------------------------------------------------------------
        tpl_card = QFrame()
        tpl_card.setObjectName("cardFrame")
        tpl_layout = QHBoxLayout(tpl_card)
        tpl_layout.setContentsMargins(14, 8, 14, 8)
        tpl_layout.setSpacing(10)

        tpl_title = QLabel("<b>Word 3GPP Template:</b>")
        tpl_title.setStyleSheet(f"font-size: 12px; color: {C.TEXT_STRONG};")
        tpl_layout.addWidget(tpl_title)

        self.dotx_badge = QLabel()
        tpl_layout.addWidget(self.dotx_badge)

        self.btn_dotx_action = QPushButton()
        self.btn_dotx_action.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_dotx_action.clicked.connect(self._on_toggle_dotx_clicked)
        tpl_layout.addWidget(self.btn_dotx_action)

        tpl_layout.addStretch()

        tpl_hint = QLabel(
            "3gpp_70.dotx activates the official 3GPP macro ribbon and style catalog in Microsoft Word."
        )
        tpl_hint.setStyleSheet(f"font-size: 11px; color: {C.TEXT_MUTED};")
        tpl_layout.addWidget(tpl_hint)

        main_layout.addWidget(tpl_card)
        self._update_dotx_indicator()

        # -------------------------------------------------------------
        # 2. Splitter: Left = Metadata Form | Right = Clause Selector
        # -------------------------------------------------------------
        splitter = QSplitter(Qt.Horizontal)
        splitter.setChildrenCollapsible(False)

        # Left: Metadata Scroll Area
        meta_scroll = QScrollArea()
        meta_scroll.setWidgetResizable(True)
        meta_scroll.setFrameShape(QFrame.NoFrame)
        meta_scroll.setStyleSheet("background: transparent;")

        meta_container = QWidget()
        meta_layout = QVBoxLayout(meta_container)
        meta_layout.setContentsMargins(0, 0, 8, 0)
        meta_layout.setSpacing(10)

        # A. Meeting Context Group
        mtg_group = QGroupBox("Meeting Context")
        mtg_group.setStyleSheet(f"font-weight: bold; color: {C.TEXT_BODY};")
        mtg_grid = QGridLayout(mtg_group)
        mtg_grid.setSpacing(6)

        mtg_grid.addWidget(QLabel("TSG / WG:"), 0, 0)
        self.input_wg = QLineEdit(self.meeting_info.get("wg_name", "SA2"))
        mtg_grid.addWidget(self.input_wg, 0, 1)

        mtg_grid.addWidget(QLabel("Meeting #:"), 0, 2)
        self.input_mtg_num = QLineEdit(self.meeting_info.get("meeting_number", ""))
        mtg_grid.addWidget(self.input_mtg_num, 0, 3)

        mtg_grid.addWidget(QLabel("Location:"), 1, 0)
        self.input_location = QLineEdit(self.meeting_info.get("location", ""))
        mtg_grid.addWidget(self.input_location, 1, 1, 1, 3)

        mtg_grid.addWidget(QLabel("Dates:"), 2, 0)
        dates_val = self.meeting_info.get("dates", "")
        if not dates_val and self.meeting_info.get("start_date"):
            dates_val = f"{self.meeting_info.get('start_date')} – {self.meeting_info.get('end_date')}"
        self.input_dates = QLineEdit(dates_val)
        mtg_grid.addWidget(self.input_dates, 2, 1, 1, 3)

        meta_layout.addWidget(mtg_group)

        # B. Contribution Core Metadata Group
        contrib_group = QGroupBox("Contribution Details")
        contrib_group.setStyleSheet(f"font-weight: bold; color: {C.TEXT_BODY};")
        contrib_grid = QGridLayout(contrib_group)
        contrib_grid.setSpacing(6)

        contrib_grid.addWidget(QLabel("TDoc ID:"), 0, 0)
        self.input_tdoc_id = QLineEdit(self.tdoc_info.get("TDoc", ""))
        self.input_tdoc_id.setPlaceholderText("e.g. S2-2609420 (or blank for draft)")
        contrib_grid.addWidget(self.input_tdoc_id, 0, 1)

        contrib_grid.addWidget(QLabel("Source:"), 0, 2)
        initial_source = self.tdoc_info.get("Source") or self.default_company
        self.input_source = QLineEdit(initial_source)
        self.input_source.setPlaceholderText("e.g. Ericsson, T-Mobile")
        self.input_source.textChanged.connect(self._save_default_company)
        contrib_grid.addWidget(self.input_source, 0, 3)

        contrib_grid.addWidget(QLabel("Title:"), 1, 0)
        self.input_title = QLineEdit(self.tdoc_info.get("Title", ""))
        self.input_title.setPlaceholderText("Contribution title or proposed solution")
        contrib_grid.addWidget(self.input_title, 1, 1, 1, 3)

        contrib_grid.addWidget(QLabel("Work Item:"), 2, 0)
        self.input_work_item = QLineEdit(self.tdoc_info.get("Work Item", ""))
        self.input_work_item.setPlaceholderText("e.g. FS_ARCH_enIMS, 5GS")
        contrib_grid.addWidget(self.input_work_item, 2, 1)

        contrib_grid.addWidget(QLabel("Agenda Item:"), 2, 2)
        self.input_ai = QLineEdit(self.tdoc_info.get("Agenda Item", ""))
        self.input_ai.setPlaceholderText("e.g. 20.10.1")
        contrib_grid.addWidget(self.input_ai, 2, 3)

        meta_layout.addWidget(contrib_group)

        # C. Stacked Mode-Specific Parameters (pCR vs CR)
        self.mode_stack = QStackedWidget()

        # Page 0: pCR Parameters
        pcr_page = QWidget()
        pcr_vbox = QVBoxLayout(pcr_page)
        pcr_vbox.setContentsMargins(0, 0, 0, 0)
        pcr_vbox.setSpacing(6)

        pcr_meta_row = QHBoxLayout()
        pcr_meta_row.addWidget(QLabel("Document for:"))
        self.combo_doc_for = QComboBox()
        self.combo_doc_for.addItems(["Approval", "Discussion and Decision", "Information"])
        initial_for = self.tdoc_info.get("For", "Approval")
        idx_for = self.combo_doc_for.findText(initial_for, Qt.MatchContains)
        if idx_for >= 0:
            self.combo_doc_for.setCurrentIndex(idx_for)
        pcr_meta_row.addWidget(self.combo_doc_for)
        pcr_meta_row.addStretch()
        pcr_vbox.addLayout(pcr_meta_row)

        pcr_vbox.addWidget(QLabel("Abstract of the contribution:"))
        self.text_abstract = QTextEdit()
        self.text_abstract.setMaximumHeight(65)
        self.text_abstract.setText(self.tdoc_info.get("Abstract", ""))
        self.text_abstract.setPlaceholderText("Brief abstract of the solution or proposal...")
        pcr_vbox.addWidget(self.text_abstract)

        pcr_vbox.addWidget(QLabel("1. Discussion (Background & Motivation):"))
        self.text_discussion = QTextEdit()
        self.text_discussion.setMaximumHeight(75)
        default_disc = f"This document proposes updates to {self.spec_type} {self.spec_num} v{self.version}."
        self.text_discussion.setText(default_disc)
        pcr_vbox.addWidget(self.text_discussion)

        self.mode_stack.addWidget(pcr_page)

        # Page 1: CR Parameters
        cr_page = QWidget()
        cr_vbox = QVBoxLayout(cr_page)
        cr_vbox.setContentsMargins(0, 0, 0, 0)
        cr_vbox.setSpacing(6)

        cr_grid = QGridLayout()
        cr_grid.addWidget(QLabel("Category:"), 0, 0)
        self.combo_cat = QComboBox()
        self.combo_cat.addItems([
            "F - Correction",
            "A - Corresp. change to an earlier release",
            "B - Addition of feature",
            "C - Functional modification of feature",
            "D - Editorial modification",
        ])
        cr_grid.addWidget(self.combo_cat, 0, 1)

        cr_grid.addWidget(QLabel("Release:"), 0, 2)
        self.combo_release = QComboBox()
        self.combo_release.addItems([
            "Rel-19", "Rel-18", "Rel-17", "Rel-16", "Rel-15", "Rel-14", "Rel-8"
        ])
        if self.release:
            idx_rel = self.combo_release.findText(self.release)
            if idx_rel >= 0:
                self.combo_release.setCurrentIndex(idx_rel)
        cr_grid.addWidget(self.combo_release, 0, 3)

        cr_grid.addWidget(QLabel("CR Number:"), 1, 0)
        self.input_cr_num = QLineEdit(self.tdoc_info.get("CR", "xxxx"))
        cr_grid.addWidget(self.input_cr_num, 1, 1)

        cr_grid.addWidget(QLabel("Revision:"), 1, 2)
        self.input_cr_rev = QLineEdit(self.tdoc_info.get("CR Rev", "-"))
        cr_grid.addWidget(self.input_cr_rev, 1, 3)

        cr_vbox.addLayout(cr_grid)

        # Proposed change affects checkboxes
        affects_box = QHBoxLayout()
        affects_box.addWidget(QLabel("Affects: "))
        self.chk_uicc = QCheckBox("UICC")
        self.chk_me = QCheckBox("ME")
        self.chk_ran = QCheckBox("RAN")
        self.chk_cn = QCheckBox("CN")
        self.chk_cn.setChecked(True)
        affects_box.addWidget(self.chk_uicc)
        affects_box.addWidget(self.chk_me)
        affects_box.addWidget(self.chk_ran)
        affects_box.addWidget(self.chk_cn)
        affects_box.addStretch()
        cr_vbox.addLayout(affects_box)

        cr_vbox.addWidget(QLabel("Reason for Change:"))
        self.text_reason = QTextEdit()
        self.text_reason.setMaximumHeight(65)
        self.text_reason.setPlaceholderText("Why the current specification requires correction or modification...")
        cr_vbox.addWidget(self.text_reason)

        cr_vbox.addWidget(QLabel("Summary of Change:"))
        self.text_summary = QTextEdit()
        self.text_summary.setMaximumHeight(65)
        self.text_summary.setPlaceholderText("High-level summary of the proposed technical modifications...")
        cr_vbox.addWidget(self.text_summary)

        cr_vbox.addWidget(QLabel("Consequences if not approved:"))
        self.text_consequences = QTextEdit()
        self.text_consequences.setMaximumHeight(55)
        self.text_consequences.setPlaceholderText("Potential issues, ambiguities, or interoperability risks...")
        cr_vbox.addWidget(self.text_consequences)

        self.mode_stack.addWidget(cr_page)
        meta_layout.addWidget(self.mode_stack)

        meta_scroll.setWidget(meta_container)
        splitter.addWidget(meta_scroll)

        # Right: Clause Selector & Tree View
        right_container = QWidget()
        right_layout = QVBoxLayout(right_container)
        right_layout.setContentsMargins(8, 0, 0, 0)
        right_layout.setSpacing(6)

        filter_row = QHBoxLayout()
        filter_row.setSpacing(6)
        self.filter_input = QLineEdit()
        self.filter_input.setPlaceholderText("🔍 Filter clauses by number or title (e.g. 5.1, NRF)...")
        self.filter_input.setStyleSheet(SEARCH_INPUT_STYLE)
        self.filter_input.setClearButtonEnabled(True)
        self.filter_input.textChanged.connect(self._filter_tree)
        filter_row.addWidget(self.filter_input, 1)

        btn_expand = QPushButton("Expand All")
        btn_expand.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        btn_expand.clicked.connect(lambda: self.clause_tree.expandAll())
        filter_row.addWidget(btn_expand)

        btn_collapse = QPushButton("Collapse All")
        btn_collapse.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        btn_collapse.clicked.connect(lambda: self.clause_tree.collapseAll())
        filter_row.addWidget(btn_collapse)
        right_layout.addLayout(filter_row)

        # Clause Tree
        self.clause_tree = QTreeWidget()
        self.clause_tree.setHeaderLabels(["Clause / Subclause", "Source Document"])
        self.clause_tree.setColumnWidth(0, 360)
        self.clause_tree.setStyleSheet(f"""
            QTreeWidget {{
                background-color: white;
                border: 1px solid {C.BORDER_STRONG};
                border-radius: 6px;
                padding: 4px;
                font-size: 12px;
            }}
            QTreeWidget::item:hover {{ background-color: {C.BG_CANVAS}; }}
            QTreeWidget::item:selected {{ background-color: {C.BORDER_SUBTLE}; color: {C.TEXT_PRIMARY}; }}
        """)
        self.clause_tree.itemChanged.connect(self._on_tree_item_changed)
        right_layout.addWidget(self.clause_tree, 1)

        # Real-time Change Block Summary Strip
        self.block_summary_label = QLabel("No clauses selected.")
        self.block_summary_label.setWordWrap(True)
        self.block_summary_label.setStyleSheet(
            f"color: {C.TEXT_SECONDARY}; font-size: 11px; padding: 4px; background: {C.BG_CANVAS}; border-radius: 4px;"
        )
        right_layout.addWidget(self.block_summary_label)

        splitter.addWidget(right_container)
        splitter.setStretchFactor(0, 4)
        splitter.setStretchFactor(1, 5)
        main_layout.addWidget(splitter, 1)

        # -------------------------------------------------------------
        # 3. Progress Bar & Status
        # -------------------------------------------------------------
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setFixedHeight(12)
        self.progress_bar.setTextVisible(False)
        self.progress_bar.setStyleSheet(f"""
            QProgressBar {{ border: 1px solid {C.BORDER_SUBTLE}; border-radius: 4px; background-color: {C.BG_CANVAS}; }}
            QProgressBar::chunk {{ background-color: {C.ACCENT_BRIGHT}; border-radius: 3px; }}
        """)
        self.status_label = QLabel("Ready to select clauses and generate contribution.")
        self.status_label.setStyleSheet(f"color: {C.TEXT_SECONDARY}; font-size: 11px; font-style: italic;")

        # -------------------------------------------------------------
        # 3b. Export Naming Scheme & Filename Preview Strip
        # -------------------------------------------------------------
        naming_layout = QHBoxLayout()
        naming_layout.setSpacing(6)
        naming_layout.setContentsMargins(0, 0, 0, 0)

        lbl_naming = QLabel("<b>Naming:</b>")
        lbl_naming.setStyleSheet(f"color: {C.TEXT_SECONDARY}; font-size: 11px;")
        naming_layout.addWidget(lbl_naming)

        self.naming_combo = QComboBox()
        self.naming_combo.setStyleSheet("font-size: 11px;")
        for scheme_id, scheme_label in CR_NAMING_SCHEMES:
            self.naming_combo.addItem(scheme_label, scheme_id)
        current_scheme = WordConfig.get_cr_naming_scheme()
        idx = self.naming_combo.findData(current_scheme)
        if idx >= 0:
            self.naming_combo.setCurrentIndex(idx)
        self.naming_combo.currentIndexChanged.connect(self._on_naming_scheme_changed)
        naming_layout.addWidget(self.naming_combo)

        lbl_file = QLabel("<b>Export File:</b>")
        lbl_file.setStyleSheet(f"color: {C.TEXT_SECONDARY}; font-size: 11px; margin-left: 4px;")
        naming_layout.addWidget(lbl_file)

        self.filename_edit = QLineEdit()
        self.filename_edit.setStyleSheet(SEARCH_INPUT_STYLE + "; font-size: 11px; padding: 4px 6px;")
        self.filename_edit.setPlaceholderText("Auto-generated differentiated filename (editable)...")
        self.filename_edit.textEdited.connect(self._on_filename_text_edited)
        naming_layout.addWidget(self.filename_edit, 1)

        self.btn_reset_name = QPushButton("🔄 Reset")
        self.btn_reset_name.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY + "; font-size: 11px; padding: 3px 8px;")
        self.btn_reset_name.setToolTip("Reset filename to auto-generated differentiated name based on metadata and selected clauses")
        self.btn_reset_name.clicked.connect(self._reset_filename_to_auto)
        naming_layout.addWidget(self.btn_reset_name)

        main_layout.addLayout(naming_layout)

        # Wire live reactive metadata updates to filename preview
        self.input_tdoc_id.textChanged.connect(lambda: self._update_preview_filename())
        self.input_cr_num.textChanged.connect(lambda: self._update_preview_filename())
        self.input_cr_rev.textChanged.connect(lambda: self._update_preview_filename())
        self.input_work_item.textChanged.connect(lambda: self._update_preview_filename())
        self.input_title.textChanged.connect(lambda: self._update_preview_filename())

        progress_box = QVBoxLayout()
        progress_box.setSpacing(3)
        progress_box.addWidget(self.status_label)
        progress_box.addWidget(self.progress_bar)
        main_layout.addLayout(progress_box)

        # -------------------------------------------------------------
        # 4. Bottom Action Bar
        # -------------------------------------------------------------
        bottom_layout = QHBoxLayout()
        bottom_layout.setSpacing(8)

        btn_clear = QPushButton("Clear Selection")
        btn_clear.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        btn_clear.clicked.connect(self._clear_selection)
        bottom_layout.addWidget(btn_clear)

        self.open_word_cb = QCheckBox("Open generated document in Word")
        self.open_word_cb.setChecked(True)
        self.open_word_cb.setStyleSheet(f"color: {C.TEXT_SECONDARY}; font-size: 12px; margin-left: 8px;")
        bottom_layout.addWidget(self.open_word_cb)

        bottom_layout.addStretch()

        self.btn_generate = QPushButton("🚀 Generate Contribution Document...")
        self.btn_generate.setStyleSheet(BUTTON_STYLE_PRIMARY)
        self.btn_generate.setEnabled(False)
        self.btn_generate.clicked.connect(self._generate_contribution)
        bottom_layout.addWidget(self.btn_generate)

        btn_close = QPushButton("Close")
        btn_close.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        btn_close.clicked.connect(self.close)
        bottom_layout.addWidget(btn_close)

        main_layout.addLayout(bottom_layout)
