# --- File: src/modules/word_tools/ui/cr_pcr_dialog.py ---
"""
Interactive 3GPP Contribution Dialog for CR and pCR Generation.

Enables telecommunications delegates to rapidly compose formal Change Requests (CR)
or Pseudo-CRs (pCR) directly from specification sections. Features automatic
meeting metadata pre-fill, dynamic 3GPP style enforcement, contiguous change-block
grouping, and seamless Word COM assembly with enterprise sensitivity labeling.
"""

import logging
import os
from pathlib import Path
from typing import Dict, List, Optional, Sequence

from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtWidgets import (
    QButtonGroup,
    QCheckBox,
    QComboBox,
    QDialog,
    QFileDialog,
    QFrame,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QRadioButton,
    QScrollArea,
    QSplitter,
    QStackedWidget,
    QTextEdit,
    QTreeWidget,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
)

from core.ui.ui_components import (
    BADGE_STYLE_INFO,
    BADGE_STYLE_MUTED,
    BADGE_STYLE_PRIMARY,
    BUTTON_STYLE_TOOLBAR_SECONDARY,
    CARD_FRAME_STYLE,
    SEARCH_INPUT_STYLE,
)

BUTTON_STYLE_PRIMARY = """
    QPushButton {
        font-weight: bold;
        background-color: #0066CC;
        color: white;
        padding: 7px 18px;
        border-radius: 4px;
        border: 1px solid #0055AA;
    }
    QPushButton:hover {
        background-color: #0052A3;
    }
    QPushButton:disabled {
        background-color: #CBD5E1;
        color: #94A3B8;
        border: 1px solid #CBD5E1;
    }
"""
from modules.specifications.ui.threads import SpecExcerptPrepThread
from modules.word_tools.core.cr_pcr_generator import (
    ChangeBlock,
    ContributionMetadata,
    CrPcrGeneratorThread,
    extract_clause_summary,
    group_selected_headings_into_blocks,
)
from modules.word_tools.core.document_opener import open_document
from modules.word_tools.core.styles_3gpp import (
    install_3gpp_dotx_template,
    is_3gpp_dotx_installed,
    uninstall_3gpp_dotx_template,
)
from modules.word_tools.core.word_config import WordConfig
from modules.word_tools.core.word_excerpt_extractor import HeadingInfo

logger = logging.getLogger(__name__)


class CrPcrDialog(QDialog):
    """
    Comprehensive dialog for generating 3GPP CRs and pCRs.
    Can be launched from Specification tab or Meetings tab.
    """

    contribution_generated = pyqtSignal(Path)

    def __init__(
        self,
        spec_num: str,
        spec_type: str = "TS",
        version: str = "",
        release: str = "",
        fname: str = "",
        url: str = "",
        spec_dl_dir: Optional[Path] = None,
        ordered_paths: Optional[Sequence[str]] = None,
        headings: Optional[Sequence[HeadingInfo]] = None,
        meeting_info: Optional[Dict[str, str]] = None,
        tdoc_info: Optional[Dict[str, str]] = None,
        initial_mode: str = "pcr",
        parent=None,
    ):
        super().__init__(parent)
        self.spec_num = spec_num.strip()
        self.spec_type = (spec_type or "TS").upper().strip()
        self.version = version.strip()
        self.release = release.strip()
        self.fname = fname.strip()
        self.url = url.strip()
        self.spec_dl_dir = Path(spec_dl_dir) if spec_dl_dir else Path.cwd()
        self.meeting_info = meeting_info or {}
        self.tdoc_info = tdoc_info or {}
        self.current_mode = initial_mode.lower()

        self._prep_thread: Optional[SpecExcerptPrepThread] = None
        self._gen_thread: Optional[CrPcrGeneratorThread] = None

        self._ordered_paths: List[str] = list(ordered_paths) if ordered_paths else []
        self._headings: List[HeadingInfo] = list(headings) if headings else []
        self._all_new_map: Dict[str, bool] = {}
        self._is_updating_checks = False

        cfg = WordConfig.load()
        self.default_company = cfg.get("default_source_company", "")

        self.setWindowTitle(
            f"📝 Draft 3GPP Contribution: {self.spec_type} {self.spec_num} v{self.version}"
        )
        self.setWindowFlags(self.windowFlags() | Qt.Window)
        self.resize(1020, 780)
        self.setMinimumWidth(850)
        self.setMinimumHeight(600)

        self.setStyleSheet(f"""
            QDialog {{ background-color: #F8F9FA; }}
            {CARD_FRAME_STYLE}
        """)

        self._setup_ui()

        # If headings are already provided, populate tree immediately
        if self._ordered_paths and self._headings:
            self._on_prep_finished(self._ordered_paths, self._headings)
        elif self.spec_num and self.version:
            self._start_spec_preparation()

    @property
    def headings(self) -> List[HeadingInfo]:
        return self._headings

    @property
    def ordered_paths(self) -> List[str]:
        return self._ordered_paths

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
        mode_label.setStyleSheet("font-size: 13px; color: #1A202C;")
        top_layout.addWidget(mode_label)

        self.radio_pcr = QRadioButton("Pseudo-CR (pCR / Discussion && Proposal)")
        self.radio_pcr.setStyleSheet("font-size: 13px; font-weight: bold; color: #005A9E;")
        self.radio_cr = QRadioButton("Formal Change Request (CR Template)")
        self.radio_cr.setStyleSheet("font-size: 13px; font-weight: bold; color: #005A9E;")

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
        tpl_title.setStyleSheet("font-size: 12px; color: #1E293B;")
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
        tpl_hint.setStyleSheet("font-size: 11px; color: #64748B;")
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
        mtg_group.setStyleSheet("font-weight: bold; color: #2D3748;")
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
        contrib_group.setStyleSheet("font-weight: bold; color: #2D3748;")
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
        self.clause_tree.setStyleSheet("""
            QTreeWidget {
                background-color: white;
                border: 1px solid #CBD5E1;
                border-radius: 6px;
                padding: 4px;
                font-size: 12px;
            }
            QTreeWidget::item:hover { background-color: #F1F5F9; }
            QTreeWidget::item:selected { background-color: #E2E8F0; color: #0F172A; }
        """)
        self.clause_tree.itemChanged.connect(self._on_tree_item_changed)
        right_layout.addWidget(self.clause_tree, 1)

        # Real-time Change Block Summary Strip
        self.block_summary_label = QLabel("No clauses selected.")
        self.block_summary_label.setWordWrap(True)
        self.block_summary_label.setStyleSheet(
            "color: #4A5568; font-size: 11px; padding: 4px; background: #EDF2F7; border-radius: 4px;"
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
        self.progress_bar.setStyleSheet("""
            QProgressBar { border: 1px solid #E2E8F0; border-radius: 4px; background-color: #EDF2F7; }
            QProgressBar::chunk { background-color: #0066CC; border-radius: 3px; }
        """)
        self.status_label = QLabel("Ready to select clauses and generate contribution.")
        self.status_label.setStyleSheet("color: #4A5568; font-size: 11px; font-style: italic;")

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
        self.open_word_cb.setStyleSheet("color: #4A5568; font-size: 12px; margin-left: 8px;")
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

    # -----------------------------------------------------------------
    # UI State Handlers & Pre-population
    # -----------------------------------------------------------------

    def _on_mode_changed(self):
        is_cr = self.radio_cr.isChecked()
        self.current_mode = "cr" if is_cr else "pcr"
        self.mode_stack.setCurrentIndex(1 if is_cr else 0)
        self._update_change_block_summary()

    def _save_default_company(self, text: str):
        comp = text.strip()
        if comp:
            cfg = WordConfig.load()
            cfg["default_source_company"] = comp
            WordConfig.save(cfg)

    def _start_spec_preparation(self):
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.status_label.setText("Preparing specification document for extraction...")
        self.status_badge.setText("⏳ Preparing Spec...")
        self.status_badge.setStyleSheet(BADGE_STYLE_MUTED)

        self._prep_thread = SpecExcerptPrepThread(
            spec_num=self.spec_num,
            version=self.version,
            fname=self.fname,
            url=self.url,
            spec_dl_dir=self.spec_dl_dir,
        )
        self._prep_thread.progress.connect(lambda pct, msg: (self.progress_bar.setValue(pct), self.status_label.setText(msg)))
        self._prep_thread.prepared.connect(self._on_prep_finished)
        self._prep_thread.error.connect(self._on_prep_error)
        self._prep_thread.start()

    def _on_prep_finished(self, ordered_paths: list, headings: list):
        self._ordered_paths = list(ordered_paths)
        self._headings = list(headings)

        parts_count = len(self._ordered_paths)
        parts_text = f"{parts_count} part" if parts_count == 1 else f"{parts_count} parts"
        self.status_badge.setText(f"📄 {parts_text} ({len(self._headings)} clauses)")
        self.status_badge.setStyleSheet(BADGE_STYLE_PRIMARY)

        self.progress_bar.setValue(100)
        self.status_label.setText("Ready: Select clauses to generate contribution.")
        self._populate_tree()

    def _on_prep_error(self, err_msg: str):
        self.progress_bar.setVisible(False)
        self.status_label.setText("⚠️ Specification preparation failed.")
        self.status_badge.setText("Failed")
        self.status_badge.setStyleSheet("color: #DC2626;")
        QMessageBox.critical(self, "Preparation Error", f"Could not prepare specification:\n{err_msg}")

    def _populate_tree(self):
        self._is_updating_checks = True
        self.clause_tree.clear()

        by_source = {}
        for h in self._headings:
            by_source.setdefault(h.source_path, []).append(h)

        multi_part = len(self._ordered_paths) > 1

        for source_path in self._ordered_paths:
            source_name = Path(source_path).name
            if multi_part:
                source_item = QTreeWidgetItem([f"📄 {source_name}", source_name])
                source_item.setFlags(source_item.flags() & ~Qt.ItemIsUserCheckable)
                font = source_item.font(0)
                font.setBold(True)
                source_item.setFont(0, font)
                self.clause_tree.addTopLevelItem(source_item)
                parent_root = source_item
            else:
                parent_root = None

            parents = {}
            for h in by_source.get(source_path, []):
                item = QTreeWidgetItem([h.text, source_name])
                item.setData(0, Qt.UserRole, h.id)
                item.setCheckState(0, Qt.Unchecked)

                parent = None
                for level in range(h.level - 1, 0, -1):
                    if level in parents:
                        parent = parents[level]
                        break

                if parent is not None:
                    parent.addChild(item)
                elif parent_root is not None:
                    parent_root.addChild(item)
                else:
                    self.clause_tree.addTopLevelItem(item)

                parents[h.level] = item
                for level in list(parents):
                    if level > h.level:
                        del parents[level]

        self.clause_tree.expandToDepth(1 if multi_part else 0)
        self._is_updating_checks = False
        self._update_change_block_summary()

    def _on_tree_item_changed(self, item: QTreeWidgetItem, column: int):
        if self._is_updating_checks or column != 0:
            return

        self._is_updating_checks = True
        try:
            state = item.checkState(0)

            def set_child_states(parent_item):
                for i in range(parent_item.childCount()):
                    child = parent_item.child(i)
                    if child.data(0, Qt.UserRole):
                        child.setCheckState(0, state)
                    set_child_states(child)

            set_child_states(item)
        finally:
            self._is_updating_checks = False

        self._update_change_block_summary()

    def _filter_tree(self, text: str):
        query = text.strip().lower()

        def visit(item: QTreeWidgetItem) -> bool:
            own_match = not query or query in item.text(0).lower()
            child_match = False
            for i in range(item.childCount()):
                child_match = visit(item.child(i)) or child_match

            is_visible = own_match or child_match
            item.setHidden(not is_visible)
            if query and child_match:
                item.setExpanded(True)
            return is_visible

        for i in range(self.clause_tree.topLevelItemCount()):
            visit(self.clause_tree.topLevelItem(i))

    def _clear_selection(self):
        self._is_updating_checks = True
        try:
            def clear_node(item: QTreeWidgetItem):
                if item.data(0, Qt.UserRole):
                    item.setCheckState(0, Qt.Unchecked)
                for i in range(item.childCount()):
                    clear_node(item.child(i))

            for i in range(self.clause_tree.topLevelItemCount()):
                clear_node(self.clause_tree.topLevelItem(i))
        finally:
            self._is_updating_checks = False

        self._update_change_block_summary()

    def _get_selected_ids(self) -> List[str]:
        selected = []

        def collect(item: QTreeWidgetItem):
            h_id = item.data(0, Qt.UserRole)
            if h_id and item.checkState(0) == Qt.Checked:
                selected.append(h_id)
            for i in range(item.childCount()):
                collect(item.child(i))

        for i in range(self.clause_tree.topLevelItemCount()):
            collect(self.clause_tree.topLevelItem(i))

        return selected

    def _update_change_block_summary(self):
        selected_ids = self._get_selected_ids()
        count = len(selected_ids)
        self.btn_generate.setEnabled(count > 0)

        if count == 0:
            self.block_summary_label.setText("No clauses selected.")
            self.btn_generate.setText("🚀 Generate Contribution Document...")
            return

        blocks = group_selected_headings_into_blocks(self._headings, selected_ids, self._all_new_map)
        summary_str = extract_clause_summary(blocks)

        block_desc = []
        for b in blocks:
            lbl = "First change" if b.index == 0 else "Next change"
            clauses = ", ".join(h.text[:25] for h in b.headings[:2])
            if len(b.headings) > 2:
                clauses += f" (+{len(b.headings) - 2} more)"
            block_desc.append(f"<b>{lbl}:</b> {clauses}")

        info_text = f"<b>{len(blocks)} Change Section(s)</b> ({summary_str})<br>" + " | ".join(block_desc)
        self.block_summary_label.setText(info_text)

        btn_type = "pCR" if self.current_mode == "pcr" else "CR"
        self.btn_generate.setText(f"🚀 Generate {btn_type} ({len(blocks)} Section{'s' if len(blocks) != 1 else ''})...")

    # -----------------------------------------------------------------
    # Generation Action
    # -----------------------------------------------------------------

    def _generate_contribution(self):
        selected_ids = self._get_selected_ids()
        if not selected_ids:
            QMessageBox.information(self, "No Clauses Selected", "Please select at least one clause.")
            return

        # Prepare ContributionMetadata
        is_cr = self.current_mode == "cr"
        cat_code = self.combo_cat.currentText().split(" - ")[0].strip() if is_cr else "F"

        meta = ContributionMetadata(
            mode=self.current_mode,
            meeting_wg=self.input_wg.text().strip(),
            meeting_num=self.input_mtg_num.text().strip(),
            meeting_dates=self.input_dates.text().strip(),
            meeting_location=self.input_location.text().strip(),
            tdoc_id=self.input_tdoc_id.text().strip() or "Draft-TDoc",
            title=self.input_title.text().strip() or f"Change to {self.spec_type} {self.spec_num}",
            source_company=self.input_source.text().strip() or "Delegate",
            doc_for=self.combo_doc_for.currentText() if not is_cr else "Approval",
            agenda_item=self.input_ai.text().strip(),
            work_item=self.input_work_item.text().strip(),
            abstract=self.text_abstract.toPlainText().strip() if not is_cr else "",
            discussion_text=self.text_discussion.toPlainText().strip() if not is_cr else "",
            spec_type=self.spec_type,
            spec_num=self.spec_num,
            spec_version=self.version,
            spec_release=self.combo_release.currentText() if is_cr else self.release,
            cr_num=self.input_cr_num.text().strip() if is_cr else "<CR#>",
            cr_rev=self.input_cr_rev.text().strip() if is_cr else "-",
            category=cat_code,
            reason_for_change=self.text_reason.toPlainText().strip() if is_cr else "",
            summary_of_change=self.text_summary.toPlainText().strip() if is_cr else "",
            consequences=self.text_consequences.toPlainText().strip() if is_cr else "",
            affects_uicc=self.chk_uicc.isChecked() if is_cr else False,
            affects_me=self.chk_me.isChecked() if is_cr else False,
            affects_ran=self.chk_ran.isChecked() if is_cr else False,
            affects_cn=self.chk_cn.isChecked() if is_cr else False,
        )

        clean_spec = self.spec_num.replace(".", "_")
        prefix = f"{meta.tdoc_id}_" if meta.tdoc_id and meta.tdoc_id != "Draft-TDoc" else ""
        default_out_name = f"{prefix}{self.current_mode.upper()}_{self.spec_type}{clean_spec}_v{self.version}.docx"
        default_out_path = str(self.spec_dl_dir / default_out_name)

        output_path, _ = QFileDialog.getSaveFileName(
            self,
            f"Save {self.current_mode.upper()} Document",
            default_out_path,
            "Word Document (*.docx)",
        )

        if not output_path:
            return

        if not output_path.lower().endswith(".docx"):
            output_path += ".docx"

        self.btn_generate.setEnabled(False)
        self.btn_generate.setText("⏳ Generating...")
        self.progress_bar.setValue(0)
        self.progress_bar.setVisible(True)
        self.status_label.setText("Starting contribution generation...")

        self._gen_thread = CrPcrGeneratorThread(
            metadata=meta,
            source_paths=self._ordered_paths,
            headings=self._headings,
            selected_ids=selected_ids,
            output_path=output_path,
            all_new_map=self._all_new_map,
        )
        self._gen_thread.progress.connect(self._on_gen_progress)
        self._gen_thread.finished_path.connect(self._on_gen_finished)
        self._gen_thread.failed.connect(self._on_gen_failed)
        self._gen_thread.finished.connect(lambda: self.btn_generate.setEnabled(True))
        self._gen_thread.start()

    def _on_gen_progress(self, pct: int, msg: str):
        self.progress_bar.setValue(pct)
        self.status_label.setText(msg)

    def _on_gen_finished(self, out_path_str: str):
        p = Path(out_path_str)
        self.progress_bar.setValue(100)
        self.status_label.setText(f"✅ Generated: {p.name}")

        self.contribution_generated.emit(p)

        if self.open_word_cb.isChecked() and p.exists():
            open_document(p)

        QMessageBox.information(
            self,
            "Generation Complete",
            f"Successfully generated {self.current_mode.upper()} document:\n{p}",
        )

    def _on_gen_failed(self, err_msg: str):
        self.progress_bar.setValue(100)
        self.status_label.setText("⚠️ Generation failed.")
        QMessageBox.critical(self, "Generation Error", f"Contribution generation failed:\n{err_msg}")

    def _update_dotx_indicator(self):
        installed = is_3gpp_dotx_installed()
        if installed:
            self.dotx_badge.setText("🟢 Installed (3gpp_70.dotx)")
            self.dotx_badge.setStyleSheet("""
                QLabel {
                    font-size: 11px;
                    font-weight: bold;
                    color: #065F46;
                    background-color: #D1FAE5;
                    border: 1px solid #A7F3D0;
                    border-radius: 10px;
                    padding: 3px 8px;
                }
            """)
            self.dotx_badge.setToolTip("3gpp_70.dotx is installed in %APPDATA%\\Microsoft\\Templates")
            self.btn_dotx_action.setText("🗑️ Remove Template")
            self.btn_dotx_action.setToolTip("Remove 3gpp_70.dotx from your Microsoft Word templates folder")
        else:
            self.dotx_badge.setText("⚪ Not Installed")
            self.dotx_badge.setStyleSheet(BADGE_STYLE_MUTED)
            self.dotx_badge.setToolTip("3gpp_70.dotx is not installed in %APPDATA%\\Microsoft\\Templates")
            self.btn_dotx_action.setText("📥 Install Template")
            self.btn_dotx_action.setToolTip("Download and install 3gpp_70.dotx into %APPDATA%\\Microsoft\\Templates")

    def _on_toggle_dotx_clicked(self):
        if is_3gpp_dotx_installed():
            reply = QMessageBox.question(
                self,
                "Remove 3GPP Word Template",
                "Do you want to remove 3gpp_70.dotx from your Microsoft Word templates folder?\n\n"
                "(%APPDATA%\\Microsoft\\Templates\\3gpp_70.dotx)\n\n"
                "Word will no longer automatically attach the 3GPP macro ribbon to newly created documents.",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No,
            )
            if reply == QMessageBox.Yes:
                ok, msg = uninstall_3gpp_dotx_template()
                if ok:
                    QMessageBox.information(self, "Template Removed", msg)
                else:
                    QMessageBox.warning(self, "Removal Failed", msg)
                self._update_dotx_indicator()
        else:
            ok, msg = install_3gpp_dotx_template(force=True)
            if ok:
                QMessageBox.information(self, "Template Installed", msg)
            else:
                QMessageBox.warning(self, "Installation Failed", msg)
            self._update_dotx_indicator()

    def closeEvent(self, event):
        if self._prep_thread and self._prep_thread.isRunning():
            self._prep_thread.requestInterruption()
            self._prep_thread.wait(200)
        if self._gen_thread and self._gen_thread.isRunning():
            self._gen_thread.requestInterruption()
            self._gen_thread.wait(500)
        super().closeEvent(event)
