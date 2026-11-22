"""UI layout and filter toolbar setup mixin for SpecVersionHistoryDialog."""

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QAbstractItemView, QCheckBox, QComboBox, QFrame, QHBoxLayout, QHeaderView,
    QLabel, QLineEdit, QMenu, QPushButton, QSplitter, QTableWidget, QTextBrowser,
    QVBoxLayout,
)

from core.ui.font_scale import register_scaled_row_height
from core.ui.theme import C
from core.ui.ui_components import (
    BADGE_STYLE_PRIMARY,
    BUTTON_STYLE_TOOLBAR_SECONDARY,
    COMBOBOX_STYLE_TOOLBAR,
    SEARCH_INPUT_STYLE,
    TABLE_STYLE_CLEAN,
)


class _HistoryUiMixin:
    """Builds the header filters, change history table, notes drawer, and footer."""

    def _setup_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(14, 14, 14, 14)
        main_layout.setSpacing(10)

        # -------------------------------------------------------------
        # 1. HEADER CARD (Spec badge, Title, Version Filter, Search Box)
        # -------------------------------------------------------------
        header_card = QFrame()
        header_card.setObjectName("cardFrame")
        header_layout = QVBoxLayout(header_card)
        header_layout.setContentsMargins(12, 10, 12, 10)
        header_layout.setSpacing(8)

        top_row = QHBoxLayout()
        self.spec_badge = QLabel(f"<b>{self.spec_number}</b>")
        self.spec_badge.setStyleSheet(BADGE_STYLE_PRIMARY + " font-size: 14px; padding: 3px 10px;")

        # Fetch formal title if available
        spec_title = "Specification Change History"
        if self.specs_db and hasattr(self.specs_db, "get_spec_details"):
            try:
                details = self.specs_db.get_spec_details(self.spec_number)
                if details and details.get("title"):
                    spec_title = details["title"]
            except Exception:
                pass

        self.title_lbl = QLabel(spec_title)
        self.title_lbl.setStyleSheet(f"font-size: 13px; font-weight: bold; color: {C.TEXT_STRONG};")
        self.title_lbl.setTextInteractionFlags(Qt.TextSelectableByMouse)

        top_row.addWidget(self.spec_badge)
        top_row.addWidget(self.title_lbl)
        top_row.addStretch()

        # TDoc Metadata Report Action Button
        self.btn_tdoc_report = QPushButton("📑 TDoc Dossier Report ▾")
        self.btn_tdoc_report.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_tdoc_report.setCursor(Qt.PointingHandCursor)
        self.btn_tdoc_report.setToolTip("Generate an on-demand technical report showing what all referenced TDocs are about, their authors, and abstracts.")
        report_menu = QMenu(self)
        report_menu.setToolTipsVisible(True)

        act_ver = report_menu.addAction("📑 Report for Active Version")
        act_ver.setToolTip("Extract metadata and generate report for all TDocs in the currently selected version.")
        act_ver.triggered.connect(lambda: self._generate_tdoc_report(scope="version"))

        act_sel = report_menu.addAction("📑 Report for Selected Rows")
        act_sel.setToolTip("Extract metadata and generate report for TDocs in highlighted table rows.")
        act_sel.triggered.connect(lambda: self._generate_tdoc_report(scope="selection"))

        act_all = report_menu.addAction("📑 Full Specification Dossier (All Versions)")
        act_all.setToolTip("Extract metadata and generate report for all TDocs across all indexed releases of this specification.")
        act_all.triggered.connect(lambda: self._generate_tdoc_report(scope="all"))

        self.btn_tdoc_report.setMenu(report_menu)
        top_row.addWidget(self.btn_tdoc_report)

        self.rescan_btn = QPushButton("🔄 Re-scan from Local Doc")
        self.rescan_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.rescan_btn.setCursor(Qt.PointingHandCursor)
        self.rescan_btn.setToolTip("Re-extracts Change History from the latest local Word document.")
        self.rescan_btn.clicked.connect(self._on_rescan_clicked)
        top_row.addWidget(self.rescan_btn)

        header_layout.addLayout(top_row)

        # Controls Row: Version selector + Live search filter
        filter_row = QHBoxLayout()
        filter_row.setSpacing(8)

        ver_lbl = QLabel("<b>Version:</b>")
        ver_lbl.setStyleSheet(f"color: {C.TEXT_SECONDARY}; font-size: 12px;")

        self.version_combo = QComboBox()
        self.version_combo.setStyleSheet(COMBOBOX_STYLE_TOOLBAR)
        self.version_combo.setMinimumWidth(160)
        self.version_combo.currentIndexChanged.connect(self._on_version_changed)

        self.search_input = QLineEdit()
        self.search_input.setStyleSheet(SEARCH_INPUT_STYLE)
        self.search_input.setPlaceholderText("🔍 Filter by TDoc (e.g. S2-2606625), CR #, meeting, or text...")
        self.search_input.setClearButtonEnabled(True)
        self.search_input.textChanged.connect(self._apply_filters)

        self.stats_lbl = QLabel("0 record(s)")
        self.stats_lbl.setStyleSheet(f"color: {C.TEXT_MUTED}; font-size: 12px; font-weight: bold;")

        filter_row.addWidget(ver_lbl)
        filter_row.addWidget(self.version_combo)
        filter_row.addWidget(self.search_input, stretch=1)
        filter_row.addWidget(self.stats_lbl)

        header_layout.addLayout(filter_row)

        # Controls Row 2: Affected Clause Filter + Origin CR Relevance Sorting
        clause_row = QHBoxLayout()
        clause_row.setSpacing(8)

        self.chk_clause_filter = QCheckBox("🎯 Affected Clause:")
        self.chk_clause_filter.setStyleSheet(f"font-size: 11px; font-weight: 600; color: {C.TEXT_BODY};")
        self.chk_clause_filter.toggled.connect(self._apply_filters)

        self.clause_input = QLineEdit()
        self.clause_input.setStyleSheet(SEARCH_INPUT_STYLE)
        self.clause_input.setPlaceholderText("Clause # (e.g. 5.15.11)...")
        self.clause_input.setFixedWidth(160)
        self.clause_input.setClearButtonEnabled(True)
        self.clause_input.textChanged.connect(self._apply_filters)

        self.clause_mode_combo = QComboBox()
        self.clause_mode_combo.setStyleSheet(COMBOBOX_STYLE_TOOLBAR)
        self.clause_mode_combo.addItems([
            "Rank by Clause (Additive)",
            "Strict Clause Filter"
        ])
        self.clause_mode_combo.setToolTip(
            "Rank by Clause (Additive): Boosts matching CRs to the top while preserving CRs without cover sheets.\n"
            "Strict Clause Filter: Hides CRs that do not explicitly list this clause."
        )
        self.clause_mode_combo.currentIndexChanged.connect(self._apply_filters)

        self.chk_origin_ranking = QCheckBox("⭐ Origin CR Ranking")
        self.chk_origin_ranking.setStyleSheet(f"font-size: 11px; font-weight: bold; color: {C.ACCENT_SKY};")
        self.chk_origin_ranking.setToolTip("Ranks candidate CRs by origin relevance using Category B, Title/Intro keywords, Affected Clauses, and chronological release primacy.")
        self.chk_origin_ranking.toggled.connect(self._apply_filters)

        clause_row.addWidget(self.chk_clause_filter)
        clause_row.addWidget(self.clause_input)
        clause_row.addWidget(self.clause_mode_combo)
        clause_row.addSpacing(12)
        clause_row.addWidget(self.chk_origin_ranking)
        clause_row.addStretch()

        header_layout.addLayout(clause_row)

        main_layout.addWidget(header_card)

        # -------------------------------------------------------------
        # 2. SPLITTER: TABLE (TOP) + DETAILS / CHIPS DRAWER (BOTTOM)
        # -------------------------------------------------------------
        splitter = QSplitter(Qt.Vertical)
        splitter.setHandleWidth(6)
        splitter.setStyleSheet(f"""
            QSplitter::handle {{
                background-color: {C.BORDER_SUBTLE};
                border-radius: 2px;
            }}
            QSplitter::handle:hover {{
                background-color: {C.ACCENT_SKY};
            }}
        """)

        # Table Container Card
        table_card = QFrame()
        table_card.setObjectName("cardFrame")
        table_layout = QVBoxLayout(table_card)
        table_layout.setContentsMargins(2, 2, 2, 2)

        self.table = QTableWidget()
        self.table.setStyleSheet(TABLE_STYLE_CLEAN)
        self.table.setColumnCount(8)
        self.table.setHorizontalHeaderLabels([
            "Date", "Meeting", "Version", "Type", "CR #", "Cat", "Primary TDoc", "Subject / Rapporteur Notes"
        ])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        register_scaled_row_height(self.table, 26)
        self.table.verticalHeader().setVisible(False)
        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self._show_context_menu)
        self.table.itemSelectionChanged.connect(self._on_table_selection_changed)
        self.table.itemDoubleClicked.connect(self._on_row_double_clicked)

        table_layout.addWidget(self.table)
        splitter.addWidget(table_card)

        # Details / Rapporteur Notes / Interactive TDoc View Card
        self.details_card = QFrame()
        self.details_card.setObjectName("cardFrame")
        details_layout = QVBoxLayout(self.details_card)
        details_layout.setContentsMargins(12, 10, 12, 10)
        details_layout.setSpacing(6)

        details_header = QHBoxLayout()
        self.drawer_title_lbl = QLabel("<b>Selected Entry Details:</b>")
        self.drawer_title_lbl.setStyleSheet(f"font-size: 12px; color: {C.TEXT_BODY};")

        self.btn_copy_entry_tdocs = QPushButton("📋 Copy Entry TDocs")
        self.btn_copy_entry_tdocs.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_copy_entry_tdocs.setCursor(Qt.PointingHandCursor)
        self.btn_copy_entry_tdocs.setToolTip("Copies all TDoc identifiers for the selected entry to the clipboard.")
        self.btn_copy_entry_tdocs.clicked.connect(self._copy_selected_entry_tdocs)
        self.btn_copy_entry_tdocs.setEnabled(False)

        self.btn_open_primary = QPushButton("📄 Open TDoc")
        self.btn_open_primary.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_open_primary.setCursor(Qt.PointingHandCursor)
        self.btn_open_primary.setToolTip("Open the primary TDoc Word document or ZIP archive.")
        self.btn_open_primary.clicked.connect(self._open_selected_tdoc)
        self.btn_open_primary.setEnabled(False)

        details_header.addWidget(self.drawer_title_lbl)
        details_header.addStretch()
        details_header.addWidget(self.btn_copy_entry_tdocs)
        details_header.addWidget(self.btn_open_primary)
        details_layout.addLayout(details_header)

        # Unified Interactive Notes & TDoc Browser
        self.notes_browser = QTextBrowser()
        self.notes_browser.setReadOnly(True)
        self.notes_browser.setOpenLinks(False)
        self.notes_browser.setStyleSheet(f"""
            QTextBrowser {{
                background-color: {C.BG_SURFACE};
                border: 1px solid {C.BORDER_SUBTLE};
                border-radius: 4px;
                color: {C.TEXT_STRONG};
                padding: 6px 8px;
            }}
        """)
        self.notes_browser.anchorClicked.connect(self._on_notes_anchor_clicked)
        self.notes_browser.setContextMenuPolicy(Qt.CustomContextMenu)
        self.notes_browser.customContextMenuRequested.connect(self._show_notes_context_menu)
        details_layout.addWidget(self.notes_browser)

        splitter.addWidget(self.details_card)
        splitter.setSizes([400, 200])
        main_layout.addWidget(splitter, stretch=1)

        # -------------------------------------------------------------
        # 3. ACTION BUTTONS FOOTER
        # -------------------------------------------------------------
        footer_layout = QHBoxLayout()
        footer_layout.setSpacing(8)

        self.btn_copy_list = QPushButton("📋 Copy TDocs List")
        self.btn_copy_list.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_copy_list.setCursor(Qt.PointingHandCursor)
        self.btn_copy_list.setToolTip("Copies all TDoc identifiers for the active version filter as a comma-separated list.")
        self.btn_copy_list.clicked.connect(self._copy_tdocs_list)

        self.btn_open_doc = QPushButton("📂 Open Spec Document")
        self.btn_open_doc.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_open_doc.setCursor(Qt.PointingHandCursor)
        self.btn_open_doc.setToolTip("Opens the baseline specification Word document.")
        self.btn_open_doc.clicked.connect(self._open_spec_document)

        self.btn_footer_report = QPushButton("📑 TDoc Dossier Report...")
        self.btn_footer_report.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_footer_report.setCursor(Qt.PointingHandCursor)
        self.btn_footer_report.setToolTip("Generate technical dossier report for referenced TDocs.")
        self.btn_footer_report.clicked.connect(lambda: self._generate_tdoc_report(scope="version"))

        footer_layout.addWidget(self.btn_copy_list)
        footer_layout.addWidget(self.btn_open_doc)
        footer_layout.addWidget(self.btn_footer_report)
        footer_layout.addStretch()

        close_btn = QPushButton("Close")
        close_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        close_btn.setCursor(Qt.PointingHandCursor)
        close_btn.clicked.connect(self.accept)
        footer_layout.addWidget(close_btn)

        main_layout.addLayout(footer_layout)

        # Preset values now that all widgets are fully constructed
        if self.initial_search:
            self.search_input.setText(self.initial_search)
        if self.initial_clause:
            self.clause_input.setText(self.initial_clause)
            self.chk_clause_filter.setChecked(True)
        if self.sort_by_origin_relevance:
            self.chk_origin_ranking.setChecked(True)
