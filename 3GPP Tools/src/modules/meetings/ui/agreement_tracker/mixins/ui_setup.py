"""Main splitter layout, query ribbon, KPI cards, table, and inspector drawer setup."""

from PyQt5.QtCore import QDate, Qt
from PyQt5.QtWidgets import (
    QCheckBox, QDateEdit, QFrame, QHBoxLayout, QHeaderView, QLabel, QLineEdit,
    QProgressBar, QPushButton, QSplitter, QTableView, QTextBrowser, QTextEdit,
    QVBoxLayout, QWidget,
)

from core.ui.theme import C
from core.ui.ui_components import (
    BUTTON_STYLE_TOOLBAR_DANGER,
    BUTTON_STYLE_TOOLBAR_SECONDARY,
    SEARCH_INPUT_STYLE,
)
from modules.meetings.ui.contributions.contribution_widgets import (
    AITokenInputWidget,
    KPICard,
    WITokenInputWidget,
)
from modules.meetings.ui.tdocs_components import CheckableComboBox

from ..model import AgreementResultsTableModel


class _TrackerUiMixin:
    """Builds the two-pane agreement tracker layout, filters, KPIs, and detail drawer."""

    def _setup_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(12, 12, 12, 12)
        main_layout.setSpacing(8)

        splitter = QSplitter(Qt.Horizontal)

        # =====================================================================
        # LEFT SIDEBAR: Filters & Study Item Scope
        # =====================================================================
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(0, 0, 8, 0)
        left_layout.setSpacing(6)

        title_lbl = QLabel("<b>⚖️ Study Item & Agreement Filters</b>")
        title_lbl.setStyleSheet(f"font-size: 13px; color: {C.TEXT_STRONG};")
        left_layout.addWidget(title_lbl)

        # Working Groups + Quick Presets
        wg_header = QHBoxLayout()
        wg_header.addWidget(QLabel("Working Groups:"))
        wg_header.addStretch()

        btn_ran23 = QPushButton("RAN2+3")
        btn_ran23.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        btn_ran23.setToolTip("Select RAN2 and RAN3")
        btn_ran23.clicked.connect(lambda: self.set_active_wgs(["RAN2", "RAN3"]))

        btn_all_ran = QPushButton("All RAN")
        btn_all_ran.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        btn_all_ran.setToolTip("Select RAN1, RAN2, RAN3, RAN4")
        btn_all_ran.clicked.connect(lambda: self.set_active_wgs(["RAN1", "RAN2", "RAN3", "RAN4"]))

        btn_sa2 = QPushButton("SA2")
        btn_sa2.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        btn_sa2.setToolTip("Select SA2")
        btn_sa2.clicked.connect(lambda: self.set_active_wgs(["SA2"]))

        wg_header.addWidget(btn_ran23)
        wg_header.addWidget(btn_all_ran)
        wg_header.addWidget(btn_sa2)
        left_layout.addLayout(wg_header)

        self.wg_filter = CheckableComboBox("Working Groups")
        wgs = self.meetings_db.get_working_groups()
        self.wg_filter.updateItems(wgs)
        left_layout.addWidget(self.wg_filter)
        # Default to RAN2 & RAN3 since those are the primary Chair Notes Study Item WGs
        self.set_active_wgs(["RAN2", "RAN3"])

        # Date Range
        left_layout.addWidget(QLabel("Meeting Date Range:"))
        date_box = QHBoxLayout()
        self.date_from = QDateEdit()
        self.date_from.setCalendarPopup(True)
        self.date_from.setDate(QDate.currentDate().addMonths(-12))

        self.date_to = QDateEdit()
        self.date_to.setCalendarPopup(True)
        self.date_to.setDate(QDate.currentDate().addMonths(1))

        date_box.addWidget(self.date_from)
        date_box.addWidget(QLabel("to"))
        date_box.addWidget(self.date_to)
        left_layout.addLayout(date_box)

        # Study Items / Work Items + Followed/Starred Quick Loaders
        wi_header = QHBoxLayout()
        wi_header.addWidget(QLabel("Study Items / Work Items:"))
        wi_header.addStretch()

        btn_followed = QPushButton("🎯 Followed")
        btn_followed.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        btn_followed.setToolTip("Add your Followed Work Items")
        btn_followed.clicked.connect(self._load_followed_wis)

        btn_starred = QPushButton("⭐ Starred")
        btn_starred.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        btn_starred.setToolTip("Add your Starred Work Items")
        btn_starred.clicked.connect(self._load_starred_wis)

        btn_clear_wi = QPushButton("✕ Clear")
        btn_clear_wi.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        btn_clear_wi.setToolTip("Clear selected Work Items")
        btn_clear_wi.clicked.connect(self._clear_wi_tokens)

        wi_header.addWidget(btn_followed)
        wi_header.addWidget(btn_starred)
        wi_header.addWidget(btn_clear_wi)
        left_layout.addLayout(wi_header)

        self.wi_token_widget = WITokenInputWidget(self.wi_db, self)
        self.wi_token_widget.input_edit.setPlaceholderText(
            "e.g. FS_6G_Radio, NR_AIML, XR, LTM..."
        )
        left_layout.addWidget(self.wi_token_widget)

        # Agenda Items (supports prefix matching!)
        left_layout.addWidget(QLabel("Agenda Items (exact or prefix, e.g. 9.3, 10):"))
        self.ai_token_widget = AITokenInputWidget(self)
        self.ai_token_widget.input_edit.setPlaceholderText("Add AI or prefix (e.g. 9.3, 10.2)...")
        left_layout.addWidget(self.ai_token_widget)

        self.chk_wi_ai_or = QCheckBox("Match Work Item OR Agenda Item")
        self.chk_wi_ai_or.setChecked(False)
        self.chk_wi_ai_or.setToolTip(
            "When checked, an agreement matches if it belongs to EITHER a selected Work Item OR a selected Agenda Item."
        )
        left_layout.addWidget(self.chk_wi_ai_or)

        # Topic / Keyword Query
        left_layout.addWidget(QLabel("Topic / Technical Keyword Filter:"))
        self.keyword_edit = QLineEdit()
        self.keyword_edit.setPlaceholderText("e.g. PDCP, paging, trajectory, Ha, sensing (comma for OR)...")
        self.keyword_edit.setStyleSheet(SEARCH_INPUT_STYLE)
        self.keyword_edit.returnPressed.connect(self._start_search)
        left_layout.addWidget(self.keyword_edit)

        # Agreement Types to Include
        types_lbl = QLabel("<b>Include Item Types:</b>")
        types_lbl.setStyleSheet(f"font-size: 11px; color: {C.TEXT_BODY}; margin-top: 4px;")
        left_layout.addWidget(types_lbl)

        self.chk_agreements = QCheckBox("🟢 Agreements, Working Assumptions && Conclusions")
        self.chk_agreements.setChecked(True)
        left_layout.addWidget(self.chk_agreements)

        self.chk_ffs = QCheckBox("🔵 Open Issues / FFS ('To be continued')")
        self.chk_ffs.setChecked(True)
        left_layout.addWidget(self.chk_ffs)

        self.chk_agreed_tdocs = QCheckBox("📄 Agreed TPs / pCRs / CRs")
        self.chk_agreed_tdocs.setChecked(True)
        left_layout.addWidget(self.chk_agreed_tdocs)

        self.chk_prev_recaps = QCheckBox("🕰️ Include 'From previous meeting' banner recaps")
        self.chk_prev_recaps.setChecked(False)
        self.chk_prev_recaps.setToolTip(
            "RAN3 Chair Notes often repeat previous meeting agreements under 'From RAN3#...:' in Agenda Item headers. "
            "Leave unchecked to avoid duplicate agreements across consecutive meetings."
        )
        left_layout.addWidget(self.chk_prev_recaps)

        # Source & Cache Options
        src_lbl = QLabel("<b>FTP & Cache Options:</b>")
        src_lbl.setStyleSheet(f"font-size: 11px; color: {C.TEXT_BODY}; margin-top: 4px;")
        left_layout.addWidget(src_lbl)

        self.chk_auto_download = QCheckBox("📥 Auto-download Final/EOM Chair Notes from 3GPP FTP")
        self.chk_auto_download.setChecked(True)
        self.chk_auto_download.setToolTip(
            "Automatically locates /Inbox/Chairs_Notes (or /Inbox/Chair_Notes), selects the Final/EOM Word or ZIP file, "
            "extracts the clean .docx, and parses agreements."
        )
        left_layout.addWidget(self.chk_auto_download)

        self.chk_bypass_cache = QCheckBox("🔄 Bypass local cache (re-fetch && re-parse)")
        self.chk_bypass_cache.setChecked(False)
        left_layout.addWidget(self.chk_bypass_cache)

        left_layout.addStretch()

        # Search / Stop Buttons
        btn_box = QHBoxLayout()
        self.btn_search = QPushButton("🔍 Track Agreements")
        self.btn_search.setObjectName("primaryBtn")
        self.btn_search.clicked.connect(self._start_search)

        self.btn_cancel = QPushButton("🛑 Stop")
        self.btn_cancel.setStyleSheet(BUTTON_STYLE_TOOLBAR_DANGER)
        self.btn_cancel.setEnabled(False)
        self.btn_cancel.clicked.connect(self._cancel_search)

        btn_box.addWidget(self.btn_search)
        btn_box.addWidget(self.btn_cancel)
        left_layout.addLayout(btn_box)

        splitter.addWidget(left_widget)

        # =====================================================================
        # RIGHT AREA: KPIs, Live Filter, Results Table & Rich Detail Viewer
        # =====================================================================
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_layout.setContentsMargins(8, 0, 0, 0)
        right_layout.setSpacing(6)

        # KPI Row
        kpi_layout = QHBoxLayout()
        kpi_layout.setSpacing(8)
        self.kpi_total = KPICard("Total Items", "0 Items", "Filtered dataset")
        self.kpi_agree = KPICard("Agreements & WAs", "0", "Official conclusions")
        self.kpi_ffs = KPICard("Open Issues / FFS", "0", "For further study")
        self.kpi_tp = KPICard("Agreed TPs / pCRs", "0", "Endorsed documents")
        self.kpi_mtgs = KPICard("Meetings Covered", "0 Mtgs", "With matches")

        kpi_layout.addWidget(self.kpi_total)
        kpi_layout.addWidget(self.kpi_agree)
        kpi_layout.addWidget(self.kpi_ffs)
        kpi_layout.addWidget(self.kpi_tp)
        kpi_layout.addWidget(self.kpi_mtgs)
        right_layout.addLayout(kpi_layout)

        # Filter Summary + Instant Live Filter Bar
        filter_bar = QHBoxLayout()
        self.filter_summary_lbl = QLabel("Select Study Items / WGs on the left and click 'Track Agreements'.")
        self.filter_summary_lbl.setWordWrap(True)
        self.filter_summary_lbl.setTextInteractionFlags(Qt.TextSelectableByMouse)
        self.filter_summary_lbl.setStyleSheet(
            f"QLabel {{ background-color: {C.BG_SURFACE}; color: {C.TEXT_SECONDARY}; "
            f"border: 1px solid {C.BORDER_SUBTLE}; border-radius: 4px; "
            "padding: 5px 8px; font-size: 11px; }"
        )
        filter_bar.addWidget(self.filter_summary_lbl, stretch=3)

        self.live_filter_edit = QLineEdit()
        self.live_filter_edit.setPlaceholderText("⚡ Instant filter within loaded results (e.g. PDCP, AI 9.3, R2-260)...")
        self.live_filter_edit.setClearButtonEnabled(True)
        self.live_filter_edit.setStyleSheet(SEARCH_INPUT_STYLE)
        self.live_filter_edit.textChanged.connect(self._on_live_filter_changed)
        filter_bar.addWidget(self.live_filter_edit, stretch=2)
        right_layout.addLayout(filter_bar)

        # Progress Bar & Compact Log Viewer
        self.progress_bar = QProgressBar()
        self.progress_bar.setFixedHeight(12)
        self.progress_bar.setTextVisible(False)
        right_layout.addWidget(self.progress_bar)

        self.log_viewer = QTextEdit()
        self.log_viewer.setReadOnly(True)
        self.log_viewer.setMaximumHeight(68)
        self.log_viewer.setStyleSheet(
            f"background-color: {C.TEXT_PRIMARY}; color: {C.SKY_BORDER_STRONG}; font-family: Consolas, monospace; font-size: 11px;"
        )
        right_layout.addWidget(self.log_viewer)

        # Vertical Splitter: Table (Top) + Rich Agreement Detail Inspector (Bottom)
        v_splitter = QSplitter(Qt.Vertical)

        self.table_model = AgreementResultsTableModel(self)
        self.table_view = QTableView()
        self.table_view.setModel(self.table_model)
        self.table_view.setAlternatingRowColors(True)
        self.table_view.setSelectionBehavior(QTableView.SelectRows)
        self.table_view.setSelectionMode(QTableView.SingleSelection)
        self.table_view.setStyleSheet(
            f"QTableView {{ border: 1px solid {C.BORDER_STRONG}; gridline-color: {C.BG_CANVAS}; background-color: {C.BG_CARD}; }} "
            f"QTableView::item:selected {{ background-color: {C.ACCENT_SOFT}; color: {C.TEXT_STRONG}; }}"
        )
        self.table_view.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.table_view.horizontalHeader().setStretchLastSection(True)
        self.table_view.selectionModel().currentRowChanged.connect(self._on_current_row_changed)
        self.table_view.doubleClicked.connect(self._on_table_double_clicked)
        self.table_view.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table_view.customContextMenuRequested.connect(self._on_table_context_menu)
        self._apply_column_widths()
        v_splitter.addWidget(self.table_view)

        # Bottom Detail Inspector Card
        detail_card = QFrame()
        detail_card.setStyleSheet(
            f"QFrame {{ background-color: {C.BG_CARD}; border: 1px solid {C.BORDER_STRONG}; border-radius: 6px; }}"
        )
        detail_layout = QVBoxLayout(detail_card)
        detail_layout.setContentsMargins(10, 8, 10, 8)
        detail_layout.setSpacing(6)

        detail_header = QHBoxLayout()
        self.lbl_detail_header = QLabel("<b>🔍 Selected Agreement / Conclusion Details</b>")
        self.lbl_detail_header.setStyleSheet(f"color: {C.TEXT_STRONG}; font-size: 12px; border: none;")
        detail_header.addWidget(self.lbl_detail_header)
        detail_header.addStretch()

        self.btn_open_tdoc = QPushButton("📄 Open Related TDoc")
        self.btn_open_tdoc.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_open_tdoc.setEnabled(False)
        self.btn_open_tdoc.clicked.connect(self._open_selected_tdoc)

        self.btn_open_notes = QPushButton("📝 Open Chair Notes File")
        self.btn_open_notes.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_open_notes.setEnabled(False)
        self.btn_open_notes.clicked.connect(self._open_selected_chair_notes_file)

        self.btn_open_mtg = QPushButton("🗓️ Open Meeting Window")
        self.btn_open_mtg.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_open_mtg.setEnabled(False)
        self.btn_open_mtg.clicked.connect(self._open_selected_meeting)

        self.btn_copy_ag = QPushButton("📋 Copy Text")
        self.btn_copy_ag.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_copy_ag.setEnabled(False)
        self.btn_copy_ag.clicked.connect(self._copy_selected_agreement)

        detail_header.addWidget(self.btn_open_tdoc)
        detail_header.addWidget(self.btn_open_notes)
        detail_header.addWidget(self.btn_open_mtg)
        detail_header.addWidget(self.btn_copy_ag)
        detail_layout.addLayout(detail_header)

        self.detail_browser = QTextBrowser()
        self.detail_browser.setOpenExternalLinks(True)
        self.detail_browser.setStyleSheet(
            f"QTextBrowser {{ background-color: {C.BG_SURFACE}; border: 1px solid {C.BORDER_SUBTLE}; border-radius: 4px; "
            f"padding: 8px; font-size: 12px; color: {C.TEXT_PRIMARY}; }}"
        )
        self.detail_browser.setHtml(
            f"<span style='color:{C.TEXT_MUTED};'>Click any row in the table above to inspect the full agreement text, "
            "Agenda Item hierarchy, Work Item, and linked TDoc metadata.</span>"
        )
        detail_layout.addWidget(self.detail_browser)
        v_splitter.addWidget(detail_card)
        v_splitter.setSizes([450, 220])
        right_layout.addWidget(v_splitter)

        # Footer Bar
        footer_layout = QHBoxLayout()
        self.lbl_count = QLabel("0 agreements loaded.")
        self.lbl_count.setStyleSheet(f"color: {C.TEXT_MUTED}; font-weight: bold;")

        self.btn_export_md = QPushButton("📑 Export Study Item Report (.md / .html)...")
        self.btn_export_md.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_export_md.setEnabled(False)
        self.btn_export_md.setToolTip(
            "Export a structured Study Item / Work Item Agreements Dossier in Markdown & HTML, "
            "grouped by Work Item, Agenda Item, and chronological Meeting."
        )
        self.btn_export_md.clicked.connect(self._export_study_item_report)

        self.btn_export_excel = QPushButton("📥 Export to Excel (.xlsx)...")
        self.btn_export_excel.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_export_excel.setEnabled(False)
        self.btn_export_excel.clicked.connect(self._export_to_excel)

        footer_layout.addWidget(self.lbl_count)
        footer_layout.addStretch()
        footer_layout.addWidget(self.btn_export_md)
        footer_layout.addWidget(self.btn_export_excel)
        right_layout.addLayout(footer_layout)

        splitter.addWidget(right_widget)
        splitter.setSizes([380, 1020])
        main_layout.addWidget(splitter)

    def _apply_column_widths(self):
        header = self.table_view.horizontalHeader()
        widths = {
            0: 60,   # WG
            1: 95,   # Meeting
            2: 85,   # Date
            3: 70,   # AI
            4: 135,  # Work Item
            5: 135,  # Type
            6: 360,  # Agreement / Conclusion
            7: 180,  # AI / Topic Title
            8: 95,   # Related TDoc
            9: 120,  # Source
            10: 160, # Chair Notes File
        }
        for col_idx, width in widths.items():
            if col_idx < len(self.table_model.COLUMNS):
                header.resizeSection(col_idx, width)
