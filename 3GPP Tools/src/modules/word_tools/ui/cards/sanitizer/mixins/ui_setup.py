"""UI setup mixin for WordSanitizerCard."""

from __future__ import annotations

from PyQt5.QtWidgets import (
    QCheckBox,
    QComboBox,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QPushButton,
    QScrollArea,
    QTableWidget,
    QVBoxLayout,
    QWidget,
)

from core.ui.flow_layout import FlowLayout
from core.ui.theme import C
from core.ui.ui_components import (
    BUTTON_STYLE_TOOLBAR_SECONDARY,
    CARD_FRAME_STYLE,
    TABLE_STYLE_CLEAN,
)
from modules.word_tools.core.document_sanitizer import LANGUAGE_PRESETS
from modules.word_tools.core.word_config import WordConfig
from modules.word_tools.ui.document_selector_pane import DocumentSelectorPane


class SanitizerUiSetupMixin:
    """Builds the scrollable Sanitizer card UI (audit KPIs, presets, language/author panels, table)."""

    def _setup_ui(self):
        outer_layout = QVBoxLayout(self)
        outer_layout.setContentsMargins(0, 0, 0, 0)
        outer_layout.setSpacing(0)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setStyleSheet("background: transparent;")

        container = QWidget()
        card_layout = QVBoxLayout(container)
        card_layout.setContentsMargins(10, 10, 10, 10)
        card_layout.setSpacing(10)

        # 1. Header & Description Frame
        hdr_frame = QFrame()
        hdr_frame.setObjectName("cardFrame")
        hdr_frame.setStyleSheet(CARD_FRAME_STYLE)
        hf_layout = QVBoxLayout(hdr_frame)
        hf_layout.setContentsMargins(14, 12, 14, 12)
        hf_layout.setSpacing(6)

        hf_title = QLabel("<b>🛠️ Document Sanitizer &amp; Revision Hygiene Hub</b>")
        hf_title.setStyleSheet(f"font-size: 14px; color: {C.TEXT_STRONG};")
        hf_layout.addWidget(hf_title)

        hf_desc = QLabel(
            "Comprehensive pre-flight sanitation for 3GPP Word contributions: unify proofing language across all "
            "stories and styles, surgically purge Track Changes language marks ('Formatted: Language...'), "
            "consolidate reviewer identities under a single corporate signature, and clean private metadata."
        )
        hf_desc.setWordWrap(True)
        hf_desc.setStyleSheet(f"color: {C.TEXT_SECONDARY}; font-size: 12px;")
        hf_layout.addWidget(hf_desc)
        card_layout.addWidget(hdr_frame)

        # 2. Document Selector Pane
        self.pane_sanitizer = DocumentSelectorPane(
            "📄 WORD CONTRIBUTION (.DOCX / .DOCM)",
            accepted_extensions=[".docx", ".docm"],
        )
        self.pane_sanitizer.file_selected.connect(self._on_sanitizer_file_selected)
        card_layout.addWidget(self.pane_sanitizer)

        # 3. Executive Audit Summary Frame
        self.audit_frame = QFrame()
        self.audit_frame.setObjectName("cardFrame")
        self.audit_frame.setStyleSheet(CARD_FRAME_STYLE)
        audit_layout = QVBoxLayout(self.audit_frame)
        audit_layout.setContentsMargins(14, 10, 14, 10)
        audit_layout.setSpacing(8)

        self.lbl_audit_summary = QLabel(
            "No document selected. Drop or select a .docx file above to audit."
        )
        self.lbl_audit_summary.setStyleSheet(
            f"color: {C.TEXT_MUTED}; font-size: 12px; font-style: italic;"
        )
        audit_layout.addWidget(self.lbl_audit_summary)

        # KPI Badges Row
        kpi_widget = QWidget(self.audit_frame)
        kpi_widget.setStyleSheet("background: transparent; border: none;")
        kpi_row = FlowLayout(kpi_widget, margin=0, h_spacing=6, v_spacing=6)

        self.lbl_kpi_lang = QLabel("🌐 Languages: —")
        self.lbl_kpi_lang_rev = QLabel("🏷️ Lang Changes: 0")
        self.lbl_kpi_authors = QLabel("👥 Authors: 0")
        self.lbl_kpi_text_rev = QLabel("📝 Ins: 0 · Del: 0")
        self.lbl_kpi_fmt = QLabel("🎨 Formatting: 0")
        self.lbl_kpi_comments = QLabel("💬 Comments: 0")

        pill_style = (
            f"background-color: {C.BG_CANVAS}; color: {C.TEXT_BODY}; "
            f"border: 1px solid {C.BORDER_STRONG}; border-radius: 4px; "
            f"padding: 3px 8px; font-size: 11px; font-weight: 500;"
        )
        for lbl in (
            self.lbl_kpi_lang,
            self.lbl_kpi_lang_rev,
            self.lbl_kpi_authors,
            self.lbl_kpi_text_rev,
            self.lbl_kpi_fmt,
            self.lbl_kpi_comments,
        ):
            lbl.setStyleSheet(pill_style)
            kpi_row.addWidget(lbl)
        audit_layout.addWidget(kpi_widget)

        # 1-Click Configuration Presets
        presets_widget = QWidget(self.audit_frame)
        presets_widget.setStyleSheet("background: transparent; border: none;")
        presets_row = FlowLayout(presets_widget, margin=0, h_spacing=6, v_spacing=6)
        lbl_presets = QLabel("<b>1-Click Presets:</b>")
        lbl_presets.setStyleSheet(
            f"font-size: 12px; color: {C.TEXT_SECONDARY}; padding-top: 4px;"
        )
        presets_row.addWidget(lbl_presets)

        btn_preset_3gpp = QPushButton("⚡ 3GPP Contribution Prep (en-GB + DT + Purge Lang)")
        btn_preset_3gpp.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        btn_preset_3gpp.setToolTip(
            "Sets en-GB, purges language clutter balloons, and unifies authors to corporate signature."
        )
        btn_preset_3gpp.clicked.connect(self._apply_preset_3gpp)

        btn_preset_lang_only = QPushButton("🌐 Fix Language Only (en-GB + Purge Lang)")
        btn_preset_lang_only.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        btn_preset_lang_only.setToolTip(
            "Unifies language to en-GB and purges language balloons without touching author names."
        )
        btn_preset_lang_only.clicked.connect(self._apply_preset_lang_only)

        btn_preset_anon = QPushButton("🎭 Full Anonymize (en-GB + Contributor + Clear Metadata)")
        btn_preset_anon.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        btn_preset_anon.setToolTip(
            "Replaces authors with 'Contributor', clears comments and private metadata."
        )
        btn_preset_anon.clicked.connect(self._apply_preset_anonymize)

        presets_row.addWidget(btn_preset_3gpp)
        presets_row.addWidget(btn_preset_lang_only)
        presets_row.addWidget(btn_preset_anon)
        audit_layout.addWidget(presets_widget)

        card_layout.addWidget(self.audit_frame)

        # 4. Operations Configuration Grid
        ops_grid = QGridLayout()
        ops_grid.setSpacing(10)
        ops_grid.setColumnStretch(0, 1)
        ops_grid.setColumnStretch(1, 1)

        # Section A: Language Unification & Purge
        sec_lang_frame = QFrame()
        sec_lang_frame.setObjectName("cardFrame")
        sec_lang_frame.setStyleSheet(CARD_FRAME_STYLE)
        sl_layout = QVBoxLayout(sec_lang_frame)
        sl_layout.setContentsMargins(12, 10, 12, 10)
        sl_layout.setSpacing(6)

        self.chk_sec_lang = QCheckBox("🌐 Proofing Language Unification && Purge")
        self.chk_sec_lang.setStyleSheet(
            f"QCheckBox {{ font-weight: bold; font-size: 13px; color: {C.TEXT_STRONG}; }}"
        )
        self.chk_sec_lang.setChecked(True)
        sl_layout.addWidget(self.chk_sec_lang)

        lang_target_row = QHBoxLayout()
        lang_target_row.addWidget(QLabel("Target Language:"))
        self.combo_target_lang = QComboBox()
        self.combo_target_lang.setSizeAdjustPolicy(
            QComboBox.AdjustToMinimumContentsLengthWithIcon
        )
        self.combo_target_lang.setMinimumContentsLength(18)
        for code, label in LANGUAGE_PRESETS:
            self.combo_target_lang.addItem(label, code)
        self.combo_target_lang.addItem("Other / Custom BCP-47 Code...", "custom")
        self.combo_target_lang.currentIndexChanged.connect(self._on_target_lang_combo_changed)
        lang_target_row.addWidget(self.combo_target_lang, 1)

        self.edit_custom_lang = QLineEdit()
        self.edit_custom_lang.setPlaceholderText("e.g. en-US, de-DE")
        self.edit_custom_lang.setMaximumWidth(100)
        self.edit_custom_lang.setVisible(False)
        lang_target_row.addWidget(self.edit_custom_lang)
        sl_layout.addLayout(lang_target_row)

        self.chk_lang_all_stories = QCheckBox(
            "Apply to all stories (body, headers, footers, footnotes)"
        )
        self.chk_lang_all_stories.setToolTip(
            "Apply proofing language normalization across body, headers, footers, footnotes, and comments."
        )
        self.chk_lang_all_stories.setChecked(True)
        sl_layout.addWidget(self.chk_lang_all_stories)

        self.chk_lang_styles = QCheckBox(
            "Normalize styles.xml && docDefaults (prevent reverting)"
        )
        self.chk_lang_styles.setChecked(True)
        sl_layout.addWidget(self.chk_lang_styles)

        self.chk_lang_purge_revisions = QCheckBox(
            "Purge language Track Changes (<w:rPrChange> / <w:lang>)"
        )
        self.chk_lang_purge_revisions.setStyleSheet(
            f"QCheckBox {{ font-weight: 600; color: {C.TEXT_STRONG}; }}"
        )
        self.chk_lang_purge_revisions.setChecked(True)
        self.chk_lang_purge_revisions.setToolTip(
            "Eliminates trivial 'Formatted: Language...' Track Changes balloons while cleanly preserving "
            "all substantive text insertions and deletions intact."
        )
        sl_layout.addWidget(self.chk_lang_purge_revisions)

        self.chk_clear_no_proof = QCheckBox("Clear <w:noProof> spell-check suppression marks")
        self.chk_clear_no_proof.setChecked(False)
        sl_layout.addWidget(self.chk_clear_no_proof)

        ops_grid.addWidget(sec_lang_frame, 0, 0)

        # Section B: Revision Author Consolidation
        sec_author_frame = QFrame()
        sec_author_frame.setObjectName("cardFrame")
        sec_author_frame.setStyleSheet(CARD_FRAME_STYLE)
        sa_layout = QVBoxLayout(sec_author_frame)
        sa_layout.setContentsMargins(12, 10, 12, 10)
        sa_layout.setSpacing(6)

        self.chk_sec_authors = QCheckBox("👥 Revision Author Consolidation")
        self.chk_sec_authors.setStyleSheet(
            f"QCheckBox {{ font-weight: bold; font-size: 13px; color: {C.TEXT_STRONG}; }}"
        )
        self.chk_sec_authors.setChecked(True)
        sa_layout.addWidget(self.chk_sec_authors)

        author_target_row = QHBoxLayout()
        author_target_row.addWidget(QLabel("Target Author Signature:"))
        self.edit_target_author = QLineEdit()
        default_company = WordConfig.load().get("default_source_company", "Deutsche Telekom")
        self.edit_target_author.setText(default_company)
        self.edit_target_author.textChanged.connect(self._on_target_author_changed)
        author_target_row.addWidget(self.edit_target_author, 1)

        author_target_row.addWidget(QLabel("Initials:"))
        self.edit_target_initials = QLineEdit()
        self.edit_target_initials.setMaximumWidth(60)
        self.edit_target_initials.setText(self._derive_initials(default_company))
        author_target_row.addWidget(self.edit_target_initials)
        sa_layout.addLayout(author_target_row)

        # Author Presets Quick Bar
        ap_widget = QWidget(sec_author_frame)
        ap_widget.setStyleSheet("background: transparent; border: none;")
        ap_bar = FlowLayout(ap_widget, margin=0, h_spacing=4, v_spacing=4)
        ap_lbl = QLabel(
            f"<span style='color:{C.TEXT_MUTED}; font-size:11px;'>Quick Signatures:</span>"
        )
        ap_bar.addWidget(ap_lbl)
        for sig in ["Deutsche Telekom", "Editor", "Rapporteur", "Contributor", "Anonymous"]:
            btn_sig = QPushButton(sig)
            btn_sig.setStyleSheet(
                f"font-size: 11px; padding: 2px 7px; background-color: {C.BG_CANVAS}; "
                f"border: 1px solid {C.BORDER_STRONG}; border-radius: 3px;"
            )
            btn_sig.clicked.connect(lambda checked, s=sig: self._apply_author_preset(s))
            ap_bar.addWidget(btn_sig)
        sa_layout.addWidget(ap_widget)

        self.chk_author_initials = QCheckBox("Update reviewer initials in comments and callouts")
        self.chk_author_initials.setChecked(True)
        sa_layout.addWidget(self.chk_author_initials)

        self.chk_author_metadata = QCheckBox(
            "Update Dublin Core metadata (creator / lastModifiedBy)"
        )
        self.chk_author_metadata.setChecked(True)
        sa_layout.addWidget(self.chk_author_metadata)

        ops_grid.addWidget(sec_author_frame, 0, 1)
        card_layout.addLayout(ops_grid)

        # 5. Detected Authors Table
        table_hdr = QHBoxLayout()
        table_hdr.addWidget(QLabel("<b>Detected Reviewers in Selected Document:</b>"))
        table_hdr.addStretch()

        btn_sel_all = QPushButton("Select All")
        btn_sel_all.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        btn_sel_all.clicked.connect(lambda: self._select_all_authors(True))
        btn_desel_all = QPushButton("Deselect All")
        btn_desel_all.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        btn_desel_all.clicked.connect(lambda: self._select_all_authors(False))
        table_hdr.addWidget(btn_sel_all)
        table_hdr.addWidget(btn_desel_all)
        card_layout.addLayout(table_hdr)

        self.author_table = QTableWidget(0, 7)
        self.author_table.setHorizontalHeaderLabels(
            ["", "Author Name", "Ins", "Del", "Fmt", "Comments", "Initials"]
        )
        self.author_table.setStyleSheet(TABLE_STYLE_CLEAN)
        self.author_table.setMaximumHeight(150)
        self.author_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.author_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        for col in (2, 3, 4, 5, 6):
            self.author_table.horizontalHeader().setSectionResizeMode(
                col, QHeaderView.ResizeToContents
            )
        card_layout.addWidget(self.author_table)

        # 6. Advanced Document Hygiene
        sec_cleanup_frame = QFrame()
        sec_cleanup_frame.setObjectName("cardFrame")
        sec_cleanup_frame.setStyleSheet(CARD_FRAME_STYLE)
        sc_layout = FlowLayout(sec_cleanup_frame, margin=12, h_spacing=14, v_spacing=6)

        self.chk_sec_cleanup = QCheckBox("🧹 Advanced Cleanup:")
        self.chk_sec_cleanup.setStyleSheet(
            f"QCheckBox {{ font-weight: bold; font-size: 13px; color: {C.TEXT_STRONG}; }}"
        )
        self.chk_sec_cleanup.setChecked(False)
        sc_layout.addWidget(self.chk_sec_cleanup)

        self.chk_purge_all_fmt = QCheckBox(
            "Purge ALL formatting revisions (keep substantive text)"
        )
        self.chk_remove_comments = QCheckBox("Remove all reviewer comments")
        self.chk_sanitizer_backup = QCheckBox(
            "Create safety backup (.bak) before updating in-place"
        )
        self.chk_sanitizer_backup.setChecked(True)

        sc_layout.addWidget(self.chk_purge_all_fmt)
        sc_layout.addWidget(self.chk_remove_comments)
        sc_layout.addWidget(self.chk_sanitizer_backup)
        card_layout.addWidget(sec_cleanup_frame)

        # 7. Action Execution Bar
        act_box = QHBoxLayout()
        act_box.setSpacing(8)

        self.btn_run_in_place = QPushButton("⚡ Sanitize && Update In-Place")
        self.btn_run_in_place.setObjectName("primaryBtn")
        self.btn_run_in_place.setToolTip(
            "Applies selected sanitizations directly to the source document (creates .bak backup)."
        )
        self.btn_run_in_place.clicked.connect(lambda: self._trigger_sanitizer(in_place=True))

        self.btn_run_save_as = QPushButton("💾 Save As New Sanitized Document...")
        self.btn_run_save_as.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_run_save_as.clicked.connect(lambda: self._trigger_sanitizer(in_place=False))

        act_box.addWidget(self.btn_run_in_place)
        act_box.addWidget(self.btn_run_save_as)
        act_box.addStretch()
        card_layout.addLayout(act_box)

        # 8. Result Indicator Frame
        self.author_result_frame = QFrame()
        self.author_result_frame.setObjectName("cardFrame")
        self.author_result_frame.setStyleSheet(
            CARD_FRAME_STYLE
            + f"""
            QFrame#cardFrame {{
                background-color: {C.WG_RAN_BG};
                border: 1px solid {C.EMERALD_BORDER};
                border-radius: 6px;
                padding: 6px 12px;
            }}
        """
        )
        self.author_result_frame.setVisible(False)
        rf_layout = QHBoxLayout(self.author_result_frame)
        rf_layout.setContentsMargins(4, 4, 4, 4)

        self.author_result_lbl = QLabel()
        self.author_result_lbl.setStyleSheet(
            f"color: {C.WG_RAN_TEXT}; font-size: 11px; font-weight: bold;"
        )
        rf_layout.addWidget(self.author_result_lbl, 1)

        self.btn_open_author_file = QPushButton("👁️ Open Document")
        self.btn_open_author_file.setStyleSheet("font-size: 11px; padding: 2px 8px;")
        self.btn_open_author_file.clicked.connect(self._open_author_result_file)
        rf_layout.addWidget(self.btn_open_author_file)

        self.btn_open_author_folder = QPushButton("📂 Open Folder")
        self.btn_open_author_folder.setStyleSheet("font-size: 11px; padding: 2px 8px;")
        self.btn_open_author_folder.clicked.connect(self._open_author_result_folder)
        rf_layout.addWidget(self.btn_open_author_folder)

        card_layout.addWidget(self.author_result_frame)
        card_layout.addStretch()

        scroll.setWidget(container)
        outer_layout.addWidget(scroll)
