"""UI setup, sidebar filters, KPI strip, results table, and column widths for `ContributionReportDialog`."""

from __future__ import annotations

from PyQt5.QtCore import QDate, Qt
from PyQt5.QtWidgets import (
    QCheckBox,
    QDateEdit,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QProgressBar,
    QPushButton,
    QSplitter,
    QTableView,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from core.ui.theme import C
from core.ui.ui_components import (
    BUTTON_STYLE_TOOLBAR_DANGER,
    BUTTON_STYLE_TOOLBAR_SECONDARY,
)
from modules.meetings.ui.contributions.contribution_models import (
    ContributionResultsTableModel,
)
from modules.meetings.ui.contributions.contribution_widgets import (
    AITokenInputWidget,
    CompanySelectorWidget,
    KPICard,
    WITokenInputWidget,
)
from modules.meetings.ui.tdoc_delegates import SourceCompanyDelegate
from modules.meetings.ui.tdocs_components import CheckableComboBox


class ContributionDialogUiSetupMixin:
    """Mixin building the filter sidebar, KPI cards, results table, and column sizing."""

    def _setup_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(12, 12, 12, 12)
        main_layout.setSpacing(8)

        splitter = QSplitter(Qt.Horizontal)

        # Left: Filter Sidebar
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(0, 0, 8, 0)
        left_layout.setSpacing(6)

        title_lbl = QLabel("<b>Search & Audit Filters</b>")
        title_lbl.setStyleSheet(f"font-size: 13px; color: {C.TEXT_STRONG};")
        left_layout.addWidget(title_lbl)

        left_layout.addWidget(QLabel("Working Groups:"))
        self.wg_filter = CheckableComboBox("Working Groups")
        wgs = self.meetings_db.get_working_groups()
        self.wg_filter.updateItems(wgs)
        left_layout.addWidget(self.wg_filter)

        left_layout.addWidget(QLabel("Date Range:"))
        date_box = QHBoxLayout()
        self.date_from = QDateEdit()
        self.date_from.setCalendarPopup(True)
        self.date_from.setDate(QDate.currentDate().addYears(-1))

        self.date_to = QDateEdit()
        self.date_to.setCalendarPopup(True)
        self.date_to.setDate(QDate.currentDate())

        date_box.addWidget(self.date_from)
        date_box.addWidget(QLabel("to"))
        date_box.addWidget(self.date_to)
        left_layout.addLayout(date_box)

        left_layout.addWidget(QLabel("Work Items (Acronym / Code):"))
        self.wi_token_widget = WITokenInputWidget(self.wi_db, self)
        left_layout.addWidget(self.wi_token_widget)

        left_layout.addWidget(QLabel("Agenda Items:"))
        self.ai_token_widget = AITokenInputWidget(self)
        left_layout.addWidget(self.ai_token_widget)

        self.chk_wi_ai_or = QCheckBox("Match WI OR Agenda Item")
        self.chk_wi_ai_or.setChecked(False)
        self.chk_wi_ai_or.setToolTip(
            "When enabled, a contribution may match either a selected WI or a selected Agenda Item. "
            "When disabled, it must match both when both filters are populated."
        )
        left_layout.addWidget(self.chk_wi_ai_or)

        left_layout.addWidget(QLabel("Companies (from Sanitizer):"))
        self.company_widget = CompanySelectorWidget(self)
        left_layout.addWidget(self.company_widget)

        self.chk_bypass_cache = QCheckBox("Bypass local cache (re-fetch from 3GPP)")
        self.chk_bypass_cache.setToolTip(
            "Force downloading and parsing the latest TDoc list sheets from the 3GPP portal."
        )
        left_layout.addWidget(self.chk_bypass_cache)

        btn_box = QHBoxLayout()
        self.btn_search = QPushButton("🔍 Search Contributions")
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

        # Right: Results, KPIs & Logs
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_layout.setContentsMargins(8, 0, 0, 0)
        right_layout.setSpacing(6)

        kpi_layout = QHBoxLayout()
        kpi_layout.setSpacing(8)
        self.kpi_total = KPICard("Total Contributions", "0 TDocs", "Filtered dataset")
        self.kpi_agree = KPICard("Agreed / Approved", "0 TDocs", "0.0% of total")
        self.kpi_joint = KPICard("Collaboration", "0% Joint", "Co-authored")
        self.kpi_partner = KPICard("Top Co-signer", "None", "Most frequent ally")
        self.kpi_wi = KPICard("Top Work Item", "None", "Highest volume WI")

        kpi_layout.addWidget(self.kpi_total)
        kpi_layout.addWidget(self.kpi_agree)
        kpi_layout.addWidget(self.kpi_joint)
        kpi_layout.addWidget(self.kpi_partner)
        kpi_layout.addWidget(self.kpi_wi)
        right_layout.addLayout(kpi_layout)

        self.filter_summary_lbl = QLabel("No search run yet.")
        self.filter_summary_lbl.setWordWrap(True)
        self.filter_summary_lbl.setTextInteractionFlags(Qt.TextSelectableByMouse)
        self.filter_summary_lbl.setStyleSheet(
            f"QLabel {{ background-color: {C.BG_SURFACE}; color: {C.TEXT_SECONDARY}; "
            f"border: 1px solid {C.BORDER_SUBTLE}; border-radius: 4px; "
            "padding: 5px 8px; font-size: 11px; }"
        )
        right_layout.addWidget(self.filter_summary_lbl)

        self.progress_bar = QProgressBar()
        self.progress_bar.setFixedHeight(14)
        self.progress_bar.setTextVisible(False)
        right_layout.addWidget(self.progress_bar)

        self.log_viewer = QTextEdit()
        self.log_viewer.setReadOnly(True)
        self.log_viewer.setMaximumHeight(80)
        self.log_viewer.setStyleSheet(
            f"background-color: {C.TEXT_PRIMARY}; color: {C.SKY_BORDER_STRONG}; "
            "font-family: monospace; font-size: 11px;"
        )
        right_layout.addWidget(self.log_viewer)

        self.table_model = ContributionResultsTableModel()
        self.table_view = QTableView()
        self.table_view.setModel(self.table_model)
        self.table_view.setAlternatingRowColors(True)
        self.table_view.setSelectionBehavior(QTableView.SelectRows)
        self.table_view.setStyleSheet(
            f"QTableView {{ border: 1px solid {C.BORDER_STRONG}; gridline-color: {C.BG_CANVAS}; "
            f"background-color: {C.BG_CARD}; }} "
            f"QTableView::item:selected {{ background-color: {C.ACCENT_SOFT}; color: {C.TEXT_STRONG}; }}"
        )
        self.table_view.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.table_view.horizontalHeader().setStretchLastSection(True)

        self.source_delegate = SourceCompanyDelegate(self.table_view)
        self.table_view.setItemDelegateForColumn(4, self.source_delegate)

        self.table_view.doubleClicked.connect(self._on_table_double_clicked)
        self.table_view.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table_view.customContextMenuRequested.connect(self._on_table_context_menu)

        self._apply_column_widths()
        right_layout.addWidget(self.table_view)

        footer_layout = QHBoxLayout()
        self.lbl_count = QLabel("0 contributions found.")
        self.lbl_count.setStyleSheet(f"color: {C.TEXT_MUTED}; font-weight: bold;")

        self.btn_stats = QPushButton("📊 Statistics Dashboard...")
        self.btn_stats.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_stats.setEnabled(False)
        self.btn_stats.setToolTip("Compile an interactive HTML analytics report with graphs and timelines.")
        self.btn_stats.clicked.connect(self._generate_statistics)

        self.btn_export = QPushButton("📥 Export to Excel...")
        self.btn_export.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_export.setEnabled(False)
        self.btn_export.clicked.connect(self._export_to_excel)

        footer_layout.addWidget(self.lbl_count)
        footer_layout.addStretch()
        footer_layout.addWidget(self.btn_stats)
        footer_layout.addWidget(self.btn_export)
        right_layout.addLayout(footer_layout)

        splitter.addWidget(right_widget)
        splitter.setSizes([380, 960])
        main_layout.addWidget(splitter)

    def _apply_column_widths(self):
        header = self.table_view.horizontalHeader()
        widths = {
            0: 65,   # WG
            1: 95,   # Meeting
            2: 95,   # TDoc
            3: 280,  # Title
            4: 160,  # Source
            5: 130,  # Matched Company
            6: 60,   # Type
            7: 75,   # For
            8: 85,   # Agenda Item
            9: 95,   # TDoc Status
            10: 140, # Related WIs
            11: 250, # Abstract
        }
        for col_idx, width in widths.items():
            if col_idx < len(self.table_model.COLUMNS):
                header.resizeSection(col_idx, width)

    def set_active_wgs(self, wgs: list):
        for i in range(1, self.wg_filter.model().rowCount()):
            item = self.wg_filter.model().item(i)
            if item:
                item.setCheckState(Qt.Checked if item.text() in wgs else Qt.Unchecked)
        self.wg_filter.update()
