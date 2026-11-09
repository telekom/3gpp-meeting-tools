# --- File: src/modules/meetings/ui/agreement_tracker_dialog.py ---
"""
Cross-Meeting Study Item & Topic Agreement Tracker Dialog (`AgreementTrackerDialog`).

Enables 3GPP delegates (especially in RAN2, RAN3, RAN1, and SA2) to track, search, inspect,
and export all agreements, working assumptions, conclusions, open issues (FFS), and agreed
TPs/pCRs/CRs across meetings and study items without manually digging through FTP folders.
"""
import html
import logging
import os
import re
import webbrowser
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional

import pandas as pd
from PyQt5.QtCore import QAbstractTableModel, QDate, QModelIndex, QPoint, Qt, pyqtSignal
from PyQt5.QtGui import QBrush, QColor, QFont
from PyQt5.QtWidgets import (
    QApplication,
    QCheckBox,
    QDateEdit,
    QDialog,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMenu,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QSplitter,
    QTableView,
    QTextBrowser,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from core.ui.ui_components import (
    BUTTON_STYLE_TOOLBAR_DANGER,
    BUTTON_STYLE_TOOLBAR_SECONDARY,
    SEARCH_INPUT_STYLE,
)
from core.utils.folder_opener import open_local_file
from core.utils.paths import to_extended_path
from core.utils.portal_utils import open_portal_tdoc
from modules.meetings.core.agreement_tracker_worker import AgreementTrackerWorker
from modules.meetings.core.meetings_db import MeetingsDatabase
from modules.meetings.core.settings import MeetingsSettings
from modules.meetings.core.tdocs.tdocs_threads import TDocActionThread
from modules.meetings.core.url_router import URLRouter
from modules.meetings.ui.contributions.contribution_widgets import (
    AITokenInputWidget,
    FilterTokenChip,
    KPICard,
    WITokenInputWidget,
)
from modules.meetings.ui.tdocs_components import CheckableComboBox
from modules.work_items.core.wi_database import WorkItemsDatabase
from modules.work_items.core.wi_settings import WorkItemsSettings

logger = logging.getLogger(__name__)


class AgreementResultsTableModel(QAbstractTableModel):
    """Table model displaying extracted agreements, working assumptions, FFS, and agreed TPs/pCRs."""

    COLUMNS = [
        "WG",
        "Meeting",
        "Date",
        "AI",
        "Work Item",
        "Type",
        "Agreement / Conclusion",
        "AI / Topic Title",
        "Related TDoc",
        "Source",
        "Chair Notes File",
    ]

    def __init__(self, parent=None):
        super().__init__(parent)
        self._all_rows: List[dict] = []
        self._filtered_rows: List[dict] = []
        self._live_filter: str = ""
        self._bold_font = QFont()
        self._bold_font.setBold(True)

    def update_data(self, rows: List[dict]):
        self.beginResetModel()
        self._all_rows = list(rows or [])
        self._apply_live_filter_internal()
        self.endResetModel()

    def set_live_filter(self, query: str):
        self.beginResetModel()
        self._live_filter = (query or "").strip().lower()
        self._apply_live_filter_internal()
        self.endResetModel()

    def _apply_live_filter_internal(self):
        if not self._live_filter:
            self._filtered_rows = list(self._all_rows)
            return
        tokens = [t for t in self._live_filter.split() if t]
        matched = []
        for r in self._all_rows:
            haystack = " ".join([
                str(r.get("WG", "")),
                str(r.get("Meeting", "")),
                str(r.get("AI", "")),
                str(r.get("AI Title", "")),
                str(r.get("Work Item", "")),
                str(r.get("Kind", "")),
                str(r.get("Agreement Text", "")),
                str(r.get("TDoc", "")),
                str(r.get("TDoc Title", "")),
                str(r.get("Source", "")),
            ]).lower()
            if all(tok in haystack for tok in tokens):
                matched.append(r)
        self._filtered_rows = matched

    def get_filtered_rows(self) -> List[dict]:
        return self._filtered_rows

    def get_all_rows(self) -> List[dict]:
        return self._all_rows

    def rowCount(self, parent=QModelIndex()) -> int:
        if parent.isValid():
            return 0
        return len(self._filtered_rows)

    def columnCount(self, parent=QModelIndex()) -> int:
        if parent.isValid():
            return 0
        return len(self.COLUMNS)

    def headerData(self, section: int, orientation: Qt.Orientation, role: int = Qt.DisplayRole):
        if orientation == Qt.Horizontal and role == Qt.DisplayRole:
            if 0 <= section < len(self.COLUMNS):
                return self.COLUMNS[section]
        return None

    def data(self, index: QModelIndex, role: int = Qt.DisplayRole):
        if not index.isValid() or not (0 <= index.row() < len(self._filtered_rows)):
            return None

        row = self._filtered_rows[index.row()]
        col_name = self.COLUMNS[index.column()]

        if role == Qt.UserRole:
            return row

        if role == Qt.DisplayRole:
            if col_name == "WG":
                return row.get("WG", "")
            if col_name == "Meeting":
                return row.get("Meeting", "")
            if col_name == "Date":
                return row.get("start_date", "")
            if col_name == "AI":
                return row.get("AI", "")
            if col_name == "Work Item":
                return row.get("Work Item", "")
            if col_name == "Type":
                kind = str(row.get("Kind", "Agreement"))
                low = kind.lower()
                if "ffs" in low or "open issue" in low:
                    return f"🔵 {kind}"
                if "working assumption" in low:
                    return f"🟡 {kind}"
                if "agreed tdoc" in low or "pcr" in low:
                    return f"📄 {kind}"
                return f"🟢 {kind}"
            if col_name == "Agreement / Conclusion":
                txt = str(row.get("Agreement Text", "")).replace("\n", " ⏎ ")
                return txt if len(txt) <= 220 else txt[:217] + "..."
            if col_name == "AI / Topic Title":
                return row.get("AI Title", "")
            if col_name == "Related TDoc":
                return row.get("TDoc", "")
            if col_name == "Source":
                return row.get("Source", "")
            if col_name == "Chair Notes File":
                return row.get("Chair Notes File", "")

        if role == Qt.ToolTipRole:
            if col_name in ("Agreement / Conclusion", "Type"):
                return str(row.get("Agreement Text", ""))
            if col_name == "Related TDoc":
                tid = row.get("TDoc", "")
                ttitle = row.get("TDoc Title", "")
                return f"{tid}: {ttitle}" if ttitle else tid
            if col_name == "AI / Topic Title":
                return str(row.get("AI Title", ""))

        if role == Qt.ForegroundRole:
            kind = str(row.get("Kind", "")).lower()
            if col_name == "Type":
                if "ffs" in kind or "open issue" in kind:
                    return QBrush(QColor("#1D4ED8"))
                if "working assumption" in kind:
                    return QBrush(QColor("#B45309"))
                if "agreed tdoc" in kind:
                    return QBrush(QColor("#4338CA"))
                return QBrush(QColor("#15803D"))
            if col_name in ("Work Item", "Related TDoc"):
                return QBrush(QColor("#1E5C99"))

        if role == Qt.BackgroundRole:
            kind = str(row.get("Kind", "")).lower()
            if col_name == "Type":
                if "ffs" in kind or "open issue" in kind:
                    return QBrush(QColor("#EFF6FF"))
                if "working assumption" in kind:
                    return QBrush(QColor("#FEF3C7"))
                if "agreed tdoc" in kind:
                    return QBrush(QColor("#EEF2FF"))
                return QBrush(QColor("#F0FDF4"))

        if role == Qt.FontRole:
            if col_name in ("Type", "Work Item", "AI"):
                return self._bold_font

        return None


class AgreementTrackerDialog(QDialog):
    """
    Independent modeless Cross-Meeting Study Item & Topic Agreement Tracker.
    """

    open_meeting_requested = pyqtSignal(str, str)  # (wg_name, meeting_number)

    def __init__(self, meetings_db: MeetingsDatabase, parent=None):
        super().__init__(None, Qt.Window | Qt.WindowMinMaxButtonsHint | Qt.WindowCloseButtonHint)
        self.setModal(False)

        self.meetings_db = meetings_db
        self.settings = MeetingsSettings()
        self.wi_db = WorkItemsDatabase(self.meetings_db.db_path)
        self.wi_settings = WorkItemsSettings()
        self.worker: Optional[AgreementTrackerWorker] = None
        self.active_tdoc_threads: Dict[str, TDocActionThread] = {}

        self.setWindowTitle("⚖️ 3GPP Cross-Meeting Study Item & Topic Agreement Tracker")
        self.resize(1420, 860)
        self._setup_ui()

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
        title_lbl.setStyleSheet("font-size: 13px; color: #1E293B;")
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
        types_lbl.setStyleSheet("font-size: 11px; color: #334155; margin-top: 4px;")
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
        src_lbl.setStyleSheet("font-size: 11px; color: #334155; margin-top: 4px;")
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
            "QLabel { background-color: #F8FAFC; color: #475569; "
            "border: 1px solid #E2E8F0; border-radius: 4px; "
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
            "background-color: #0F172A; color: #38BDF8; font-family: Consolas, monospace; font-size: 11px;"
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
            "QTableView { border: 1px solid #CBD5E1; gridline-color: #F1F5F9; background-color: #FFFFFF; } "
            "QTableView::item:selected { background-color: #EBF3FC; color: #1E293B; }"
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
            "QFrame { background-color: #FFFFFF; border: 1px solid #CBD5E1; border-radius: 6px; }"
        )
        detail_layout = QVBoxLayout(detail_card)
        detail_layout.setContentsMargins(10, 8, 10, 8)
        detail_layout.setSpacing(6)

        detail_header = QHBoxLayout()
        self.lbl_detail_header = QLabel("<b>🔍 Selected Agreement / Conclusion Details</b>")
        self.lbl_detail_header.setStyleSheet("color: #1E293B; font-size: 12px; border: none;")
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
            "QTextBrowser { background-color: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 4px; "
            "padding: 8px; font-size: 12px; color: #0F172A; }"
        )
        self.detail_browser.setHtml(
            "<span style='color:#64748B;'>Click any row in the table above to inspect the full agreement text, "
            "Agenda Item hierarchy, Work Item, and linked TDoc metadata.</span>"
        )
        detail_layout.addWidget(self.detail_browser)
        v_splitter.addWidget(detail_card)
        v_splitter.setSizes([450, 220])
        right_layout.addWidget(v_splitter)

        # Footer Bar
        footer_layout = QHBoxLayout()
        self.lbl_count = QLabel("0 agreements loaded.")
        self.lbl_count.setStyleSheet("color: #64748B; font-weight: bold;")

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

    # =========================================================================
    # Pre-populating Filters & Quick Loaders
    # =========================================================================

    def set_active_wgs(self, wgs: List[str]):
        target_upper = {w.strip().upper() for w in (wgs or []) if w.strip()}
        for i in range(1, self.wg_filter.model().rowCount()):
            item = self.wg_filter.model().item(i)
            if item:
                item.setCheckState(
                    Qt.Checked if item.text().strip().upper() in target_upper else Qt.Unchecked
                )
        self.wg_filter.update()

    def set_date_range(self, d_from: QDate, d_to: QDate):
        if d_from and d_from.isValid():
            self.date_from.setDate(d_from)
        if d_to and d_to.isValid():
            self.date_to.setDate(d_to)

    def add_wi_tokens(self, tokens: List[str]):
        for tok in (tokens or []):
            clean = str(tok).strip()
            if clean and clean not in self.wi_token_widget.selected_tokens:
                self.wi_token_widget.selected_tokens.add(clean)
                chip = FilterTokenChip(clean, self.wi_token_widget)
                chip.removed.connect(self.wi_token_widget._remove_token)
                self.wi_token_widget.chips_layout.addWidget(chip)
        self.wi_token_widget.tokens_changed.emit()

    def _clear_wi_tokens(self):
        for tok in list(self.wi_token_widget.selected_tokens):
            self.wi_token_widget._remove_token(tok)

    def _load_followed_wis(self):
        followed = self.wi_settings.get_followed_wis()
        if not followed:
            QMessageBox.information(
                self, "No Followed WIs", "You have not marked any Work Items as Followed yet."
            )
            return
        self.add_wi_tokens(followed)

    def _load_starred_wis(self):
        starred = self.wi_settings.get_starred_wis()
        if not starred:
            QMessageBox.information(
                self, "No Starred WIs", "You have not starred any Work Items yet."
            )
            return
        self.add_wi_tokens(starred)

    # =========================================================================
    # Search Execution & Lifecycle
    # =========================================================================

    def _cleanup_threads(self):
        if self.worker and self.worker.isRunning():
            self.worker.cancel()
            self.worker.requestInterruption()
            self.worker.quit()
            if not self.worker.wait(1000):
                try:
                    self.worker.terminate()
                    self.worker.wait(500)
                except Exception:
                    pass
        for th in list(self.active_tdoc_threads.values()):
            if th and th.isRunning():
                th.requestInterruption()
                th.quit()
                if not th.wait(500):
                    try:
                        th.terminate()
                        th.wait(200)
                    except Exception:
                        pass
        self.active_tdoc_threads.clear()

    def closeEvent(self, event):
        self._cleanup_threads()
        super().closeEvent(event)

    def reject(self):
        self._cleanup_threads()
        super().reject()

    def _start_search(self):
        selected_wgs = self.wg_filter.getCheckedItems()
        target_wis = self.wi_token_widget.get_tokens()
        target_ais = self.ai_token_widget.get_tokens()
        keyword_query = self.keyword_edit.text().strip()
        wi_ai_or = self.chk_wi_ai_or.isChecked()
        date_from = self.date_from.date().toString("yyyy-MM-dd")
        date_to = self.date_to.date().toString("yyyy-MM-dd")

        if not (
            self.chk_agreements.isChecked()
            or self.chk_ffs.isChecked()
            or self.chk_agreed_tdocs.isChecked()
        ):
            QMessageBox.warning(
                self, "No Item Types Selected", "Please select at least one item type to include."
            )
            return

        summary_parts = []
        if selected_wgs:
            summary_parts.append("WG: " + ", ".join(selected_wgs))
        summary_parts.append(f"{date_from} → {date_to}")
        if target_wis:
            summary_parts.append("WI/SI: " + ", ".join(target_wis))
        if target_ais:
            summary_parts.append("AI: " + ", ".join(target_ais))
        if keyword_query:
            summary_parts.append(f"Keyword: '{keyword_query}'")
        self.filter_summary_lbl.setText("  •  ".join(summary_parts))

        self.btn_search.setEnabled(False)
        self.btn_cancel.setEnabled(True)
        self.btn_export_md.setEnabled(False)
        self.btn_export_excel.setEnabled(False)
        self.log_viewer.clear()
        self.progress_bar.setValue(0)

        self.worker = AgreementTrackerWorker(
            meetings_db=self.meetings_db,
            wgs=selected_wgs,
            date_from=date_from,
            date_to=date_to,
            target_wis=target_wis,
            target_ais=target_ais,
            keyword_query=keyword_query,
            wi_ai_or=wi_ai_or,
            include_agreements=self.chk_agreements.isChecked(),
            include_ffs=self.chk_ffs.isChecked(),
            include_agreed_tdocs=self.chk_agreed_tdocs.isChecked(),
            include_prev_recaps=self.chk_prev_recaps.isChecked(),
            auto_download_missing=self.chk_auto_download.isChecked(),
            bypass_cache=self.chk_bypass_cache.isChecked(),
            cache_dir=self.settings.cache_dir,
            parent=self,
        )
        self.worker.stage_progress.connect(self._on_stage_progress)
        self.worker.log_msg.connect(self._append_log)
        self.worker.results_ready.connect(self._on_results_ready)
        self.worker.finished.connect(self._on_search_finished)
        self.worker.start()

    def _cancel_search(self):
        if self.worker and self.worker.isRunning():
            self.worker.cancel()
            self.btn_cancel.setEnabled(False)

    def _on_stage_progress(self, desc: str, current: int, total: int):
        self.progress_bar.setMaximum(max(total, 1))
        self.progress_bar.setValue(current)
        self.setWindowTitle(
            f"[{current}/{total}] {desc} - 3GPP Study Item & Topic Agreement Tracker"
        )

    def _append_log(self, msg: str):
        self.log_viewer.append(msg)
        self.log_viewer.verticalScrollBar().setValue(self.log_viewer.verticalScrollBar().maximum())

    def _on_results_ready(self, results: List[dict]):
        self.table_model.update_data(results)
        self._refresh_kpis_and_counts()
        self._apply_column_widths()
        if self.table_model.rowCount() > 0:
            self.table_view.selectRow(0)

    def _on_live_filter_changed(self, text: str):
        self.table_model.set_live_filter(text)
        self._refresh_kpis_and_counts()
        if self.table_model.rowCount() > 0:
            self.table_view.selectRow(0)

    def _refresh_kpis_and_counts(self):
        rows = self.table_model.get_filtered_rows()
        all_count = len(self.table_model.get_all_rows())
        total = len(rows)

        if self.live_filter_edit.text().strip():
            self.lbl_count.setText(f"Showing {total} of {all_count} agreement(s).")
        else:
            self.lbl_count.setText(f"Found {total} matching agreement(s)/item(s).")

        if total == 0:
            self.kpi_total.set_value("0 Items", "No matches")
            self.kpi_agree.set_value("0", "0.0% of total")
            self.kpi_ffs.set_value("0", "0 open issues")
            self.kpi_tp.set_value("0", "0 documents")
            self.kpi_mtgs.set_value("0 Mtgs", "No meetings")
            self.btn_export_md.setEnabled(False)
            self.btn_export_excel.setEnabled(False)
            return

        wgs_set = {r.get("WG", "") for r in rows if r.get("WG")}
        mtgs_set = {r.get("Meeting", "") for r in rows if r.get("Meeting")}
        wis_set = {r.get("Work Item", "") for r in rows if r.get("Work Item")}

        agree_cnt = 0
        ffs_cnt = 0
        tp_cnt = 0
        for r in rows:
            k = str(r.get("Kind", "")).lower()
            if "ffs" in k or "open issue" in k:
                ffs_cnt += 1
            elif "agreed tdoc" in k or "pcr" in k:
                tp_cnt += 1
            else:
                agree_cnt += 1

        self.kpi_total.set_value(f"{total} Items", f"Across {len(wis_set)} WI(s)")
        self.kpi_agree.set_value(f"🟢 {agree_cnt}", "Agreements & WAs")
        self.kpi_ffs.set_value(f"🔵 {ffs_cnt}", "Open Issues / FFS")
        self.kpi_tp.set_value(f"📄 {tp_cnt}", "Agreed TPs / pCRs")
        self.kpi_mtgs.set_value(f"{len(mtgs_set)} Mtgs", f"In {', '.join(sorted(wgs_set))}")

        self.btn_export_md.setEnabled(True)
        self.btn_export_excel.setEnabled(True)

    def _on_search_finished(self, success: bool, msg: str):
        self.btn_search.setEnabled(True)
        self.btn_cancel.setEnabled(False)
        self.setWindowTitle("⚖️ 3GPP Cross-Meeting Study Item & Topic Agreement Tracker")
        if not success and "cancelled" not in msg.lower():
            QMessageBox.warning(self, "Tracker Notice", msg)

    # =========================================================================
    # Detail Inspector & Row Actions
    # =========================================================================

    def _get_selected_row(self) -> Optional[dict]:
        idx = self.table_view.currentIndex()
        if not idx.isValid():
            return None
        return self.table_model.data(idx, Qt.UserRole)

    def _on_current_row_changed(self, current: QModelIndex, _previous: QModelIndex):
        if not current.isValid():
            self.btn_open_tdoc.setEnabled(False)
            self.btn_open_notes.setEnabled(False)
            self.btn_open_mtg.setEnabled(False)
            self.btn_copy_ag.setEnabled(False)
            return

        row = self.table_model.data(current, Qt.UserRole)
        if not row:
            return

        tdoc_id = str(row.get("TDoc", "")).strip()
        chair_file = str(row.get("Chair Notes File", "")).strip()

        self.btn_open_tdoc.setEnabled(bool(tdoc_id))
        self.btn_open_tdoc.setText(f"📄 Open {tdoc_id}" if tdoc_id else "📄 Open Related TDoc")
        self.btn_open_notes.setEnabled(bool(chair_file))
        self.btn_open_mtg.setEnabled(True)
        self.btn_copy_ag.setEnabled(True)

        kind = str(row.get("Kind", "Agreement"))
        low_k = kind.lower()
        badge_bg = "#DCFCE7"
        badge_fg = "#166534"
        if "ffs" in low_k or "open issue" in low_k:
            badge_bg = "#DBEAFE"
            badge_fg = "#1E40AF"
        elif "working assumption" in low_k:
            badge_bg = "#FEF3C7"
            badge_fg = "#92400E"
        elif "agreed tdoc" in low_k:
            badge_bg = "#E0E7FF"
            badge_fg = "#3730A3"

        wg = html.escape(str(row.get("WG", "")))
        mtg = html.escape(str(row.get("Meeting", "")))
        dt = html.escape(str(row.get("start_date", "")))
        ai = html.escape(str(row.get("AI", "") or "N/A"))
        ai_title = html.escape(str(row.get("AI Title", "") or ""))
        wi = html.escape(str(row.get("Work Item", "") or "Unspecified"))
        rel = html.escape(str(row.get("Release", "") or ""))
        tdoc_title = html.escape(str(row.get("TDoc Title", "") or ""))
        src = html.escape(str(row.get("Source", "") or ""))
        cfile = html.escape(chair_file)

        raw_text = str(row.get("Agreement Text", ""))
        formatted_lines = []
        for line in raw_text.split("\n"):
            esc_line = html.escape(line)
            # Highlight TDoc IDs inside agreement text
            esc_line = re.sub(
                r'\b([A-Z][A-Z0-9]{1,3}-\d{4,8}(?:[rR]\d{1,2}[a-zA-Z]?)?)\b',
                r"<b><span style='color:#1E5C99;'>\1</span></b>",
                esc_line,
            )
            formatted_lines.append(esc_line)
        body_html = "<br>".join(formatted_lines)

        tdoc_meta_html = ""
        if tdoc_id:
            tdoc_meta_html = (
                f"<div style='margin-top:6px; padding:5px 8px; background:#EFF6FF; border-left:3px solid #3B82F6;'>"
                f"<b>Related TDoc:</b> <span style='color:#1E5C99; font-weight:bold;'>{html.escape(tdoc_id)}</span>"
                f"{f' — {tdoc_title}' if tdoc_title else ''}"
                f"{f' <i>({src})</i>' if src else ''}"
                f"</div>"
            )

        html_doc = f"""
        <div style='font-family: Segoe UI, sans-serif;'>
            <div style='margin-bottom:6px;'>
                <span style='background:{badge_bg}; color:{badge_fg}; padding:2px 8px; border-radius:4px; font-weight:bold;'>
                    {html.escape(kind)}
                </span>
                &nbsp;&nbsp;<b>{mtg}</b> ({dt})
                &nbsp;&nbsp;|&nbsp;&nbsp;<b>Work Item:</b> <span style='color:#1E5C99;'>{wi}</span> {f'({rel})' if rel else ''}
                &nbsp;&nbsp;|&nbsp;&nbsp;<b>AI {ai}:</b> {ai_title}
                &nbsp;&nbsp;|&nbsp;&nbsp;<span style='color:#64748B;'>Source: {cfile}</span>
            </div>
            <div style='padding:8px 10px; background:#FFFFFF; border:1px solid #E2E8F0; border-radius:4px; line-height:1.45;'>
                {body_html}
            </div>
            {tdoc_meta_html}
        </div>
        """
        self.detail_browser.setHtml(html_doc)

    def _open_selected_tdoc(self):
        row = self._get_selected_row()
        if row and row.get("TDoc"):
            self._open_tdoc(row)

    def _open_selected_chair_notes_file(self):
        row = self._get_selected_row()
        if not row:
            return
        folder_name = row.get("folder_name") or row.get("meeting_number", "")
        cfile = row.get("Chair Notes File", "")
        if not folder_name or not cfile:
            return
        mtg_dir = Path(self.settings.cache_dir) / folder_name
        candidates = [
            mtg_dir / "Agenda" / "ChairNotes" / cfile,
            mtg_dir / "Agenda" / cfile,
        ]
        for cand in candidates:
            if os.path.exists(to_extended_path(cand)):
                open_local_file(str(cand))
                return
        # If it was a .xlsx fallback, look for any .xlsx in Agenda/
        agenda_dir = mtg_dir / "Agenda"
        if cfile.endswith(".xlsx") and os.path.exists(to_extended_path(agenda_dir)):
            for f in agenda_dir.iterdir():
                if f.suffix.lower() == ".xlsx":
                    open_local_file(str(f))
                    return
        QMessageBox.information(
            self, "File Not Found Locally", f"Could not find '{cfile}' in {mtg_dir / 'Agenda'}."
        )

    def _open_selected_meeting(self):
        row = self._get_selected_row()
        if row:
            self._request_open_meeting(row)

    def _copy_selected_agreement(self):
        row = self._get_selected_row()
        if not row:
            return
        header = (
            f"[{row.get('Meeting', '')} | AI {row.get('AI', '')} ({row.get('AI Title', '')}) | "
            f"WI: {row.get('Work Item', '')} | {row.get('Kind', '')}]"
        )
        tdoc = row.get("TDoc", "")
        if tdoc:
            header += f" (TDoc: {tdoc})"
        QApplication.clipboard().setText(f"{header}\n{row.get('Agreement Text', '')}")

    def _request_open_meeting(self, row_data: dict):
        wg = str(row_data.get("WG", "")).strip()
        mtg_num = str(row_data.get("meeting_number", "")).strip()
        if not mtg_num:
            mtg_raw = str(row_data.get("Meeting", "")).strip()
            mtg_num = mtg_raw.split("#")[-1].strip() if "#" in mtg_raw else mtg_raw
        if wg and mtg_num:
            self.open_meeting_requested.emit(wg, mtg_num)

    def _open_tdoc(self, row_data: dict):
        tdoc_id = str(row_data.get("TDoc", "")).strip()
        if not tdoc_id:
            return

        match = re.search(r'^(.*?)-?(?:r|rev)(\d{1,2}[a-zA-Z]?)$', tdoc_id, re.IGNORECASE)
        base_tdoc = match.group(1).upper() if match else tdoc_id.upper()
        target_filename = tdoc_id.upper()

        if target_filename in self.active_tdoc_threads:
            th = self.active_tdoc_threads[target_filename]
            if th and th.isRunning():
                return

        wg_name = row_data.get("WG") or row_data.get("wg_name", "")
        folder_name = row_data.get("folder_name") or row_data.get("meeting_number", "")
        raw_url = row_data.get("url_key", "")
        docs_url = row_data.get("docs_folder_url", "")

        main_ftp = (
            raw_url
            if raw_url.startswith("http")
            else (f"https://www.3gpp.org/ftp/{raw_url.lstrip('/')}" if raw_url else "")
        )
        is_active_sync = self.meetings_db.is_active_sync_meeting(
            wg_name,
            row_data.get("start_date", ""),
            row_data.get("end_date", ""),
            int(row_data.get("is_electronic", 0) or 0),
        )
        candidate_urls = URLRouter.build_priority_url_list(
            wg_name, folder_name, main_ftp, is_active_sync, target_filename=target_filename
        )
        if docs_url:
            clean_docs = (
                docs_url
                if docs_url.startswith("http")
                else f"https://www.3gpp.org/ftp/{docs_url.lstrip('/')}"
            ).rstrip("/")
            if clean_docs not in candidate_urls:
                candidate_urls.append(clean_docs)

        if not candidate_urls:
            QMessageBox.warning(
                self, "Missing URL", f"No candidate download routes found for {tdoc_id}."
            )
            return

        meeting_dir = Path(self.settings.cache_dir) / folder_name
        meeting_dir.mkdir(parents=True, exist_ok=True)

        thread = TDocActionThread(
            base_tdoc, target_filename, candidate_urls, meeting_dir, open_file=True
        )
        self.active_tdoc_threads[target_filename] = thread
        thread.finished_action.connect(
            lambda t, s, m, th=thread: self._on_tdoc_action_finished(t, s, m, th)
        )
        thread.start()

    def _on_tdoc_action_finished(
        self, tdoc: str, success: bool, msg: str, thread: TDocActionThread
    ):
        if thread.target_filename in self.active_tdoc_threads:
            del self.active_tdoc_threads[thread.target_filename]
        if not success:
            QMessageBox.warning(self, f"Action Failed: {tdoc}", msg)

    def _on_table_double_clicked(self, index: QModelIndex):
        row = self.table_model.data(index, Qt.UserRole)
        if not row:
            return
        col_name = self.table_model.COLUMNS[index.column()]
        if col_name in ("WG", "Meeting"):
            self._request_open_meeting(row)
        elif col_name == "Chair Notes File":
            self._open_selected_chair_notes_file()
        elif row.get("TDoc"):
            self._open_tdoc(row)

    def _on_table_context_menu(self, pos: QPoint):
        index = self.table_view.indexAt(pos)
        if not index.isValid():
            return
        row = self.table_model.data(index, Qt.UserRole)
        if not row:
            return

        menu = QMenu(self)
        mtg = row.get("Meeting", "")
        tdoc = row.get("TDoc", "")

        act_copy_ag = menu.addAction("📋 Copy Agreement with Metadata")
        act_copy_ag.triggered.connect(self._copy_selected_agreement)

        act_open_notes = menu.addAction(f"📝 Open Chair Notes File ({row.get('Chair Notes File', '')})")
        act_open_notes.triggered.connect(self._open_selected_chair_notes_file)

        act_open_mtg = menu.addAction(f"🗓️ Open Meeting Window ({mtg})")
        act_open_mtg.triggered.connect(lambda: self._request_open_meeting(row))

        if tdoc:
            menu.addSeparator()
            act_open_td = menu.addAction(f"📄 Open Related TDoc ({tdoc})")
            act_open_td.triggered.connect(lambda: self._open_tdoc(row))

            act_portal = menu.addAction(f"🌐 View {tdoc} on 3GU Portal")
            act_portal.triggered.connect(lambda: open_portal_tdoc(tdoc))

        menu.exec_(self.table_view.viewport().mapToGlobal(pos))

    # =========================================================================
    # Exports: Study Item Report (.md & .html) & Excel (.xlsx)
    # =========================================================================

    def _export_study_item_report(self):
        rows = self.table_model.get_filtered_rows()
        if not rows:
            return

        wis = self.wi_token_widget.get_tokens()
        wi_slug = "_".join(wis[:2]) if wis else "All_Topics"
        wi_slug = re.sub(r'[^A-Za-z0-9_\-]', '_', wi_slug)
        default_name = f"3GPP_Agreements_Report_{wi_slug}_{datetime.now().strftime('%Y%m%d')}.md"

        out_path, _ = QFileDialog.getSaveFileName(
            self,
            "Export Study Item & Topic Agreements Report",
            default_name,
            "Markdown & HTML Dossier (*.md);;HTML Report Only (*.html)",
        )
        if not out_path:
            return

        try:
            md_text, html_text = self._compile_study_item_report_content(rows)
            p = Path(out_path)
            if p.suffix.lower() == ".html":
                with open(to_extended_path(p), "w", encoding="utf-8") as f:
                    f.write(html_text)
                open_local_file(str(p))
            else:
                with open(to_extended_path(p), "w", encoding="utf-8") as f:
                    f.write(md_text)
                html_companion = p.with_suffix(".html")
                with open(to_extended_path(html_companion), "w", encoding="utf-8") as f:
                    f.write(html_text)
                open_local_file(str(html_companion))
        except Exception as e:
            QMessageBox.critical(self, "Export Error", f"Failed to export report:\n{e}")

    def _compile_study_item_report_content(self, rows: List[dict]):
        """Builds both a structured Markdown dossier and a rich standalone HTML report."""
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M")
        filter_desc = self.filter_summary_lbl.text()

        # Group by Work Item -> Meeting -> Agenda Item
        grouped: Dict[str, Dict[str, Dict[str, List[dict]]]] = {}
        for r in rows:
            wi = str(r.get("Work Item", "")).strip() or "Unspecified Work Item"
            mtg = f"{r.get('Meeting', '')} ({r.get('start_date', '')})"
            ai = str(r.get("AI", "")).strip() or "General"
            ai_title = str(r.get("AI Title", "")).strip()
            ai_header = f"AI {ai}: {ai_title}" if ai_title else f"AI {ai}"
            grouped.setdefault(wi, {}).setdefault(mtg, {}).setdefault(ai_header, []).append(r)

        md_lines = [
            "# ⚖️ 3GPP Cross-Meeting Study Item & Topic Agreements Dossier",
            "",
            f"- **Generated:** {now_str}",
            f"- **Filter Scope:** {filter_desc}",
            f"- **Total Matched Items:** {len(rows)}",
            "",
            "---",
            "",
        ]

        html_sections = []
        for wi, mtg_map in grouped.items():
            md_lines.append(f"## 🎯 Work Item / Study Item: `{wi}`")
            md_lines.append("")
            html_sections.append(
                f"<div class='wi-card'><h2>🎯 Work Item / Study Item: <code>{html.escape(wi)}</code></h2>"
            )

            for mtg_label, ai_map in mtg_map.items():
                md_lines.append(f"### 🗓️ {mtg_label}")
                md_lines.append("")
                html_sections.append(f"<h3>🗓️ {html.escape(mtg_label)}</h3>")

                for ai_header, items in ai_map.items():
                    md_lines.append(f"#### 📂 {ai_header}")
                    md_lines.append("")
                    html_sections.append(f"<h4>📂 {html.escape(ai_header)}</h4>")

                    for item in items:
                        kind = str(item.get("Kind", "Agreement"))
                        tdoc = str(item.get("TDoc", "")).strip()
                        ttitle = str(item.get("TDoc Title", "")).strip()
                        src = str(item.get("Source", "")).strip()
                        txt = str(item.get("Agreement Text", "")).strip()

                        icon = "🟢"
                        css_cls = "ag-green"
                        if "ffs" in kind.lower() or "open issue" in kind.lower():
                            icon = "🔵"
                            css_cls = "ag-blue"
                        elif "working assumption" in kind.lower():
                            icon = "🟡"
                            css_cls = "ag-amber"
                        elif "agreed tdoc" in kind.lower():
                            icon = "📄"
                            css_cls = "ag-indigo"

                        tdoc_suffix = ""
                        if tdoc:
                            tdoc_suffix = f" *(Related TDoc: **{tdoc}**{f' — {ttitle}' if ttitle else ''}{f' [{src}]' if src else ''})*"

                        md_lines.append(f"- {icon} **[{kind}]**{tdoc_suffix}")
                        for line in txt.split("\n"):
                            md_lines.append(f"  > {line}")
                        md_lines.append("")

                        body_html = "<br>".join(html.escape(l) for l in txt.split("\n"))
                        tdoc_html = ""
                        if tdoc:
                            tdoc_html = (
                                f"<div class='tdoc-ref'>Related TDoc: <b>{html.escape(tdoc)}</b>"
                                f"{f' — {html.escape(ttitle)}' if ttitle else ''}"
                                f"{f' <i>({html.escape(src)})</i>' if src else ''}</div>"
                            )
                        html_sections.append(
                            f"<div class='ag-box {css_cls}'>"
                            f"<div class='ag-badge'>{icon} {html.escape(kind)}</div>"
                            f"<div class='ag-text'>{body_html}</div>"
                            f"{tdoc_html}"
                            f"</div>"
                        )

            html_sections.append("</div>")

        html_doc = f"""<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>3GPP Study Item & Topic Agreements Dossier</title>
<style>
  body {{ font-family: 'Segoe UI', system-ui, sans-serif; background: #F8FAFC; color: #0F172A; margin: 24px auto; max-width: 1120px; padding: 0 20px; }}
  h1 {{ color: #0F172A; border-bottom: 2px solid #CBD5E1; padding-bottom: 8px; }}
  .meta {{ background: #FFFFFF; border: 1px solid #E2E8F0; border-radius: 6px; padding: 12px 16px; margin-bottom: 20px; color: #475569; }}
  .wi-card {{ background: #FFFFFF; border: 1px solid #CBD5E1; border-radius: 8px; padding: 16px 20px; margin-bottom: 24px; box-shadow: 0 1px 2px rgba(0,0,0,0.03); }}
  h2 {{ color: #1E3A8A; margin-top: 0; }}
  h3 {{ color: #334155; border-bottom: 1px solid #E2E8F0; padding-bottom: 4px; margin-top: 18px; }}
  h4 {{ color: #475569; margin: 12px 0 8px 0; }}
  .ag-box {{ border-radius: 6px; padding: 10px 14px; margin-bottom: 10px; border-left: 4px solid #16A34A; background: #F0FDF4; }}
  .ag-blue {{ border-left-color: #2563EB; background: #EFF6FF; }}
  .ag-amber {{ border-left-color: #D97706; background: #FFFBEB; }}
  .ag-indigo {{ border-left-color: #4F46E5; background: #EEF2FF; }}
  .ag-badge {{ font-weight: bold; font-size: 12px; margin-bottom: 4px; color: #1E293B; }}
  .ag-text {{ font-size: 13px; line-height: 1.5; color: #0F172A; }}
  .tdoc-ref {{ margin-top: 6px; font-size: 12px; color: #334155; }}
</style>
</head>
<body>
  <h1>⚖️ 3GPP Cross-Meeting Study Item & Topic Agreements Dossier</h1>
  <div class="meta">
    <b>Generated:</b> {html.escape(now_str)}<br>
    <b>Filter Scope:</b> {html.escape(filter_desc)}<br>
    <b>Total Matched Items:</b> {len(rows)}
  </div>
  {''.join(html_sections)}
</body>
</html>"""

        return "\n".join(md_lines), html_doc

    def _export_to_excel(self):
        rows = self.table_model.get_filtered_rows()
        if not rows:
            return

        wis = self.wi_token_widget.get_tokens()
        wi_slug = "_".join(wis[:2]) if wis else "All_Topics"
        wi_slug = re.sub(r'[^A-Za-z0-9_\-]', '_', wi_slug)
        default_name = f"3GPP_Agreements_{wi_slug}_{datetime.now().strftime('%Y%m%d')}.xlsx"

        out_path, _ = QFileDialog.getSaveFileName(
            self,
            "Export Agreements to Excel",
            default_name,
            "Excel Workbook (*.xlsx)",
        )
        if not out_path:
            return

        try:
            export_records = []
            for r in rows:
                export_records.append({
                    "WG": r.get("WG", ""),
                    "Meeting": r.get("Meeting", ""),
                    "Date": r.get("start_date", ""),
                    "Agenda Item": r.get("AI", ""),
                    "AI / Topic Title": r.get("AI Title", ""),
                    "Work Item / Study Item": r.get("Work Item", ""),
                    "Release": r.get("Release", ""),
                    "Type": r.get("Kind", ""),
                    "Agreement / Conclusion Text": r.get("Agreement Text", ""),
                    "Related TDoc": r.get("TDoc", ""),
                    "TDoc Title": r.get("TDoc Title", ""),
                    "Source": r.get("Source", ""),
                    "Chair Notes File": r.get("Chair Notes File", ""),
                })
            df = pd.DataFrame(export_records)
            df.to_excel(to_extended_path(out_path), index=False)
            open_local_file(str(out_path))
        except Exception as e:
            QMessageBox.critical(self, "Excel Export Error", f"Failed to export Excel workbook:\n{e}")
