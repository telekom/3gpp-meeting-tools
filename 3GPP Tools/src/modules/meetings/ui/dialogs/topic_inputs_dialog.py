# --- File: src/modules/meetings/ui/dialogs/topic_inputs_dialog.py ---
"""
Considered Inputs & Decisions Viewer Dialog (`TopicInputsViewerDialog`).

Interactive visual inspector displaying all candidate technical inputs
(CRs, pCRs, Chairman's Notes text agreements, Working Assumptions, Liaison Statements,
and discussion papers) correlated with a 3GPP Cross-WG Meta-Topic across meeting rounds.
"""

import csv
import logging
import os
import re
from typing import Any, Dict, List, Optional, Set, Tuple

from PyQt5.QtCore import Qt, QThread, pyqtSignal, QTimer
from PyQt5.QtGui import QColor, QFont
from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QLineEdit,
    QTabWidget, QTableWidget, QTableWidgetItem, QHeaderView, QComboBox,
    QFrame, QMessageBox, QWidget, QAbstractItemView, QProgressBar,
    QFileDialog, QApplication, QButtonGroup, QSplitter, QTextEdit,
)

from core.utils.company_sanitizer import CompanySanitizer
from modules.meetings.core.cohort_resolver import MeetingCohort, format_round_label
from modules.meetings.core.cross_wg_ls_tracker import CrossWgLiaisonTracker
from modules.meetings.core.meetings_db import MeetingsDatabase
from modules.meetings.core.tdocs.chair_notes_agreements_parser import clean_chair_notes_text
from modules.meetings.core.tdocs.tdoc_url_resolver import resolve_tdoc_download_url
from modules.meetings.core.topic_decision import (
    TopicDecision,
    classify_tdoc_decision_type,
    normalize_chair_notes_decision_type,
)
from modules.meetings.core.topic_manager import MetaTopicDefinition
from modules.meetings.ui.topic_pills import FilterPillButton, get_wg_colors

logger = logging.getLogger(__name__)

STYLE_PRIMARY_BTN = """
    QPushButton {
        background-color: #10B981;
        color: #FFFFFF;
        font-weight: bold;
        border: 1px solid #059669;
        border-radius: 6px;
        padding: 6px 14px;
        font-size: 11px;
    }
    QPushButton:hover {
        background-color: #059669;
    }
    QPushButton:disabled {
        background-color: #94A3B8;
        border-color: #94A3B8;
    }
"""

STYLE_SECONDARY_BTN = """
    QPushButton {
        background-color: #F8FAFC;
        color: #1E293B;
        border: 1px solid #CBD5E1;
        border-radius: 6px;
        padding: 5px 12px;
        font-size: 11px;
        font-weight: 500;
    }
    QPushButton:hover {
        background-color: #F1F5F9;
        border-color: #94A3B8;
    }
    QPushButton:disabled {
        color: #94A3B8;
        background-color: #F8FAFC;
    }
"""

STYLE_TABLE = """
    QTableWidget {
        background-color: #FFFFFF;
        border: 1px solid #CBD5E1;
        border-radius: 6px;
        gridline-color: #F1F5F9;
        font-size: 11px;
    }
    QHeaderView::section {
        background-color: #F8FAFC;
        color: #475569;
        font-weight: bold;
        font-size: 11px;
        border: none;
        border-bottom: 1px solid #E2E8F0;
        padding: 6px;
    }
    QTableWidget::item {
        padding: 4px 6px;
    }
    QTableWidget::item:selected {
        background-color: #EFF6FF;
        color: #1D4ED8;
    }
"""

BADGE_COLORS = {
    "Text Agreement": ("#DCFCE7", "#166534"),
    "Approved pCR": ("#ECFDF5", "#065F46"),
    "Approved CR": ("#D1FAE5", "#047857"),
    "Endorsed WF": ("#DBEAFE", "#1E40AF"),
    "Working Assumption": ("#FEF3C7", "#92400E"),
    "Agreed LS": ("#EDE9FE", "#5B21B6"),
    "LS in": ("#F3E8FF", "#6B21A8"),
    "LS out": ("#E0E7FF", "#3730A3"),
    "Discussion": ("#F1F5F9", "#334155"),
    "Work Plan": ("#FEE2E2", "#991B1B"),
}


class InputsCollectorThread(QThread):
    """Background worker resolving matching decisions, contributions, and liaison statements."""
    progress_status = pyqtSignal(str, int, int)
    collection_finished = pyqtSignal(list, list, object)  # decisions, contributions, ls_tracker
    collection_failed = pyqtSignal(str)

    def __init__(
        self,
        meetings_db: MeetingsDatabase,
        topic: MetaTopicDefinition,
        rounds: List[MeetingCohort],
        scope_mode: str = "full_dossier",
        parent=None,
    ):
        super().__init__(parent)
        self.meetings_db = meetings_db
        self.topic = topic
        self.rounds = list(rounds)
        self.scope_mode = scope_mode
        self._is_cancelled = False

    def cancel(self):
        self._is_cancelled = True
        self.requestInterruption()

    def run(self):
        try:
            from modules.meetings.core.exporters.topic_dossier_exporter import TopicDossierExporterThread
            helper_thread = TopicDossierExporterThread(
                meetings_db=self.meetings_db,
                topic=self.topic,
                selected_cohorts=self.rounds,
                export_scope_mode=self.scope_mode,
            )

            all_decisions: List[TopicDecision] = []
            all_contributions: List[Dict[str, Any]] = []
            ls_tracker = CrossWgLiaisonTracker()

            total_rounds = len(self.rounds)
            for r_idx, m_round in enumerate(self.rounds, start=1):
                if self._is_cancelled or self.isInterruptionRequested():
                    return

                self.progress_status.emit(f"Scanning {m_round.label} ({r_idx}/{total_rounds})...", r_idx, total_rounds)

                for mtg in m_round.meetings:
                    wg_name = str(mtg.get("wg_name", "")).strip().upper()
                    mtg_num = str(mtg.get("meeting_number", "")).strip()
                    folder_name = str(mtg.get("folder_name", "")).strip()
                    mtg_date = str(mtg.get("start_date", "")).strip()

                    # 1. Parse Chairman's Notes text agreements
                    chair_agreements = helper_thread._load_meeting_agreements(mtg)
                    for ca in chair_agreements:
                        wi_str = str(ca.get("wi", "") or ca.get("study_item", "")).strip()
                        ai_str = str(ca.get("ai", "") or ca.get("agenda_item", "")).strip()
                        text_str = clean_chair_notes_text(ca.get("text", "") or ca.get("summary", ""))

                        if helper_thread._matches_topic_rule(wg_name, ai_str, text_str, wi_str):
                            dec_type = normalize_chair_notes_decision_type(ca.get("decision_type", "Agreement"))
                            all_decisions.append(TopicDecision(
                                cohort_key=m_round.key,
                                wg=wg_name,
                                meeting_number=mtg_num,
                                date=mtg_date,
                                decision_type=dec_type,
                                title_or_summary=text_str,
                                tdoc_id=ca.get("tdoc_id"),
                                source=ca.get("source") or f"Chairman's Notes AI {ai_str}",
                                target_spec=ca.get("spec", ""),
                                ai=ai_str,
                                wi=wi_str,
                                status=ca.get("status", "Agreed"),
                                notes=ca.get("notes", ""),
                            ))

                    # 2. Parse Meeting TDocs
                    tdocs_rows = helper_thread._load_meeting_tdocs(mtg)
                    for row in tdocs_rows:
                        tid = str(row.get("TDoc", "") or row.get("tdoc_id", "") or row.get("TD#", "")).strip().upper()
                        if not tid:
                            continue
                        title = str(row.get("Title", "") or row.get("title", "")).strip()
                        ai_val = str(row.get("Agenda Item", "") or row.get("AI", "") or "").strip()
                        wi_val = str(row.get("Related WIs", "") or row.get("wi", "") or "").strip()
                        t_type = str(row.get("Type", "") or row.get("type", "") or "").strip()
                        status = str(row.get("TDoc Status", "") or row.get("Status", "") or "").strip()
                        source = str(row.get("Source", "") or row.get("source", "") or "").strip()
                        ai_desc = str(row.get("Agenda item description", "") or row.get("AI Description", "") or "").strip()

                        if not self.topic.matches_tdoc(tid, wg_name, ai=ai_val, title=title, related_wis=wi_val, ai_description=ai_desc):
                            continue

                        # Check decision classification
                        dec_type = classify_tdoc_decision_type(t_type, title, status)
                        if dec_type:
                            all_decisions.append(TopicDecision(
                                cohort_key=m_round.key,
                                wg=wg_name,
                                meeting_number=mtg_num,
                                date=mtg_date,
                                decision_type=dec_type,
                                title_or_summary=title,
                                tdoc_id=tid,
                                source=source,
                                target_spec=str(row.get("Spec", "") or row.get("spec", "")).strip(),
                                ai=ai_val,
                                wi=wi_val,
                                status=status,
                                source_url=resolve_tdoc_download_url(tid, folder_name),
                            ))

                        is_ls = "ls" in t_type.lower() or "liaison" in t_type.lower()
                        if is_ls:
                            ls_tracker.ingest_ls_tdoc(
                                tdoc_id=tid,
                                wg=wg_name,
                                meeting_number=mtg_num,
                                cycle=m_round.key,
                                tdoc_type=t_type,
                                title=title,
                                source=source,
                                status=status,
                                reply_to=row.get("Reply to", ""),
                                reply_in=row.get("Reply in", ""),
                            )

                        clean_src = CompanySanitizer.get_confirmed_source_text(source) or source
                        all_contributions.append({
                            "tdoc_id": tid,
                            "wg": wg_name,
                            "meeting_number": mtg_num,
                            "folder_name": folder_name,
                            "round_key": m_round.key,
                            "round_label": m_round.label,
                            "date": mtg_date,
                            "title": title,
                            "ai": ai_val,
                            "ai_desc": ai_desc,
                            "wi": wi_val,
                            "type": t_type,
                            "status": status,
                            "source": source,
                            "clean_source": clean_src,
                            "spec": str(row.get("Spec", "") or row.get("spec", "")).strip(),
                            "is_decision": bool(dec_type),
                            "decision_type": dec_type or "",
                            "source_url": resolve_tdoc_download_url(tid, folder_name),
                            "raw_row": row,
                        })

            if not self._is_cancelled:
                self.collection_finished.emit(all_decisions, all_contributions, ls_tracker)
        except Exception as e:
            logger.exception(f"[InputsCollectorThread] Collection failed: {e}")
            if not self._is_cancelled:
                self.collection_failed.emit(str(e))


class TopicInputsViewerDialog(QDialog):
    """Interactive dialog visualizing considered inputs, agreements, pCRs, CRs, and LSs."""
    export_requested = pyqtSignal()

    def __init__(
        self,
        meetings_db: MeetingsDatabase,
        topic: MetaTopicDefinition,
        selected_rounds: List[MeetingCohort],
        scope_mode: str = "full_dossier",
        timeframe_label: str = "",
        parent=None,
    ):
        super().__init__(None, Qt.Window | Qt.WindowMinMaxButtonsHint | Qt.WindowCloseButtonHint)
        self.setModal(False)
        self.setWindowModality(Qt.NonModal)
        self.setAttribute(Qt.WA_DeleteOnClose, True)
        self.setWindowTitle(f"⚖️ Considered Inputs & Decisions — {topic.name}")
        self.resize(1140, 720)

        self.meetings_db = meetings_db
        self.topic = topic
        self.rounds = list(selected_rounds)
        self.scope_mode = scope_mode
        self.timeframe_label = timeframe_label or f"{len(self.rounds)} Synchronized Meeting Rounds"

        self.all_decisions: List[TopicDecision] = []
        self.all_contributions: List[Dict[str, Any]] = []
        self.ls_tracker = CrossWgLiaisonTracker()

        self._displayed_decisions: List[TopicDecision] = []
        self._displayed_contribs: List[Dict[str, Any]] = []
        self._displayed_ls: List[Dict[str, Any]] = []
        self._current_detail_text: str = ""

        self._collector_worker: Optional[InputsCollectorThread] = None

        self._init_ui()
        self._start_collection()

    def _init_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(10)
        layout.setContentsMargins(14, 14, 14, 14)

        # 1. Header Card with topic overview and timeframe
        header_card = QFrame()
        header_card.setStyleSheet("background-color: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px;")
        h_layout = QVBoxLayout(header_card)
        h_layout.setContentsMargins(12, 10, 12, 10)
        h_layout.setSpacing(6)

        top_row = QHBoxLayout()
        lbl_title = QLabel(f"🌐 <b>Topic:</b> {self.topic.name}")
        lbl_title.setStyleSheet("font-size: 15px; font-weight: bold; color: #0F172A;")
        top_row.addWidget(lbl_title)

        top_row.addStretch()

        lbl_timeframe = QLabel(f"🗓️ <b>Scope:</b> {self.timeframe_label}")
        lbl_timeframe.setStyleSheet("color: #475569; font-size: 11px; background-color: #E2E8F0; border-radius: 4px; padding: 3px 8px;")
        top_row.addWidget(lbl_timeframe)
        h_layout.addLayout(top_row)

        desc_text = self.topic.description or "Cross-WG topic correlation"
        wgs_str = ", ".join(sorted(self.topic.target_wgs))
        lbl_desc = QLabel(f"<b>Target WGs:</b> <span style='color: #0369A1;'>{wgs_str}</span> | {desc_text}")
        lbl_desc.setStyleSheet("color: #64748B; font-size: 11px;")
        lbl_desc.setWordWrap(True)
        h_layout.addWidget(lbl_desc)

        layout.addWidget(header_card)

        # 2. Metric KPI Cards Strip
        kpi_frame = QFrame()
        kpi_frame.setStyleSheet("background-color: #FFFFFF; border: 1px solid #CBD5E1; border-radius: 6px;")
        kpi_layout = QHBoxLayout(kpi_frame)
        kpi_layout.setContentsMargins(10, 8, 10, 8)
        kpi_layout.setSpacing(16)

        self.kpi_total = QLabel("📊 Total Inputs: <b>0</b>")
        self.kpi_total.setStyleSheet("color: #1E293B; font-size: 12px;")
        kpi_layout.addWidget(self.kpi_total)

        self.kpi_decisions = QLabel("🟢 Decisions: <b>0</b>")
        self.kpi_decisions.setStyleSheet("color: #047857; font-size: 12px;")
        kpi_layout.addWidget(self.kpi_decisions)

        self.kpi_ls = QLabel("📬 Liaisons: <b>0</b>")
        self.kpi_ls.setStyleSheet("color: #6B21A8; font-size: 12px;")
        kpi_layout.addWidget(self.kpi_ls)

        self.kpi_disc = QLabel("📄 Discussions: <b>0</b>")
        self.kpi_disc.setStyleSheet("color: #475569; font-size: 12px;")
        kpi_layout.addWidget(self.kpi_disc)

        self.kpi_wgs = QLabel("👥 WGs: <b>0</b>")
        self.kpi_wgs.setStyleSheet("color: #0284C7; font-size: 12px;")
        kpi_layout.addWidget(self.kpi_wgs)

        kpi_layout.addStretch()

        self.btn_refresh = QPushButton("🔄 Re-Scan")
        self.btn_refresh.setStyleSheet(STYLE_SECONDARY_BTN)
        self.btn_refresh.clicked.connect(self._start_collection)
        kpi_layout.addWidget(self.btn_refresh)

        layout.addWidget(kpi_frame)

        # Progress bar strip (indeterminate/step during collection)
        self.prog_frame = QFrame()
        self.prog_frame.setStyleSheet("background-color: #EFF6FF; border: 1px solid #BFDBFE; border-radius: 6px;")
        p_layout = QHBoxLayout(self.prog_frame)
        p_layout.setContentsMargins(10, 6, 10, 6)
        self.lbl_prog_status = QLabel("Scanning meeting records...")
        self.lbl_prog_status.setStyleSheet("color: #1E40AF; font-size: 11px; font-weight: bold;")
        p_layout.addWidget(self.lbl_prog_status, 1)
        self.prog_bar = QProgressBar()
        self.prog_bar.setFixedHeight(12)
        self.prog_bar.setRange(0, 100)
        self.prog_bar.setTextVisible(False)
        self.prog_bar.setStyleSheet("""
            QProgressBar { border: 1px solid #93C5FD; border-radius: 3px; background-color: #DBEAFE; }
            QProgressBar::chunk { background-color: #3B82F6; border-radius: 2px; }
        """)
        p_layout.addWidget(self.prog_bar, 1)
        layout.addWidget(self.prog_frame)

        # 3. Interactive Filter Strip & Quick-Filter Pill Bars
        filter_frame = QFrame()
        filter_frame.setStyleSheet("background-color: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px;")
        filter_layout = QVBoxLayout(filter_frame)
        filter_layout.setContentsMargins(10, 8, 10, 8)
        filter_layout.setSpacing(8)

        # Row 1: Search + Meeting Round Selector + Reset
        search_row = QHBoxLayout()
        search_row.setSpacing(10)

        search_row.addWidget(QLabel("<b>🔍 Search:</b>"))
        self.edit_filter = QLineEdit()
        self.edit_filter.setPlaceholderText("Filter by TDoc ID, Title, Company, Agenda Item, or Decision text...")
        self.edit_filter.textChanged.connect(self._apply_filters)
        search_row.addWidget(self.edit_filter, 3)

        search_row.addWidget(QLabel("<b>Round:</b>"))
        self.combo_round_filter = QComboBox()
        self.combo_round_filter.addItem("All Meeting Rounds")
        for r in self.rounds:
            self.combo_round_filter.addItem(r.label)
        self.combo_round_filter.currentIndexChanged.connect(self._apply_filters)
        search_row.addWidget(self.combo_round_filter, 2)

        self.btn_reset_filters = QPushButton("↺ Reset")
        self.btn_reset_filters.setToolTip("Reset all search and filter pills")
        self.btn_reset_filters.setStyleSheet(STYLE_SECONDARY_BTN)
        self.btn_reset_filters.clicked.connect(self._reset_filters)
        search_row.addWidget(self.btn_reset_filters)
        filter_layout.addLayout(search_row)

        # Dropdown combos maintained for programmatic filtering & backward compatibility
        self.combo_wg_filter = QComboBox()
        self.combo_wg_filter.addItem("All WGs")
        for wg in sorted(self.topic.target_wgs):
            self.combo_wg_filter.addItem(wg)
        self.combo_wg_filter.setVisible(False)

        self.combo_type_filter = QComboBox()
        self.combo_type_filter.addItems([
            "All Inputs",
            "🟢 Decisions Only (Agreements, pCRs, CRs)",
            "Text Agreements Only",
            "Approved pCRs Only",
            "Approved CRs Only",
            "📬 Liaison Statements Only",
            "📄 Discussion Papers Only",
        ])
        self.combo_type_filter.setVisible(False)

        # Row 2: Working Group Quick-Filter Pills
        wg_row = QHBoxLayout()
        wg_row.setSpacing(6)
        lbl_wg_title = QLabel("<b>Working Group:</b>")
        lbl_wg_title.setStyleSheet("font-size: 11px; color: #475569;")
        lbl_wg_title.setFixedWidth(90)
        wg_row.addWidget(lbl_wg_title)

        self.wg_pills: Dict[str, FilterPillButton] = {}
        self.wg_pill_group = QButtonGroup(self)
        self.wg_pill_group.setExclusive(True)

        pill_all_wg = FilterPillButton("All WGs", "All WGs", count=None)
        pill_all_wg.setChecked(True)
        pill_all_wg.toggled.connect(lambda chk: self._on_wg_pill_toggled("All WGs", chk))
        self.wg_pills["All WGs"] = pill_all_wg
        self.wg_pill_group.addButton(pill_all_wg)
        wg_row.addWidget(pill_all_wg)

        for wg in sorted(self.topic.target_wgs):
            colors = get_wg_colors(wg)
            pill = FilterPillButton(wg, wg, count=0, colors=colors)
            pill.toggled.connect(lambda chk, w=wg: self._on_wg_pill_toggled(w, chk))
            self.wg_pills[wg] = pill
            self.wg_pill_group.addButton(pill)
            wg_row.addWidget(pill)
        wg_row.addStretch()
        filter_layout.addLayout(wg_row)

        # Row 3: Decision & Input Category Quick-Filter Pills
        cat_row = QHBoxLayout()
        cat_row.setSpacing(6)
        lbl_cat_title = QLabel("<b>Category:</b>")
        lbl_cat_title.setStyleSheet("font-size: 11px; color: #475569;")
        lbl_cat_title.setFixedWidth(90)
        cat_row.addWidget(lbl_cat_title)

        self.cat_pills: Dict[str, FilterPillButton] = {}
        self.cat_pill_group = QButtonGroup(self)
        self.cat_pill_group.setExclusive(True)

        cat_defs = [
            ("All Inputs", "All Inputs", None, None),
            ("🟢 Decisions Only (Agreements, pCRs, CRs)", "🟢 Decisions", 0, {"bg_active": "#10B981", "fg_active": "#FFFFFF", "border_active": "#059669", "bg_inactive": "#FFFFFF", "fg_inactive": "#047857", "border_inactive": "#A7F3D0", "hover": "#ECFDF5"}),
            ("Text Agreements Only", "📝 Agreements", 0, {"bg_active": "#10B981", "fg_active": "#FFFFFF", "border_active": "#059669", "bg_inactive": "#FFFFFF", "fg_inactive": "#065F46", "border_inactive": "#A7F3D0", "hover": "#ECFDF5"}),
            ("Approved pCRs Only", "📑 Approved pCRs", 0, {"bg_active": "#059669", "fg_active": "#FFFFFF", "border_active": "#047857", "bg_inactive": "#FFFFFF", "fg_inactive": "#065F46", "border_inactive": "#6EE7B7", "hover": "#ECFDF5"}),
            ("Approved CRs Only", "✅ Approved CRs", 0, {"bg_active": "#047857", "fg_active": "#FFFFFF", "border_active": "#065F46", "bg_inactive": "#FFFFFF", "fg_inactive": "#047857", "border_inactive": "#6EE7B7", "hover": "#ECFDF5"}),
            ("📬 Liaison Statements Only", "📬 Liaisons", 0, {"bg_active": "#7C3AED", "fg_active": "#FFFFFF", "border_active": "#6D28D9", "bg_inactive": "#FFFFFF", "fg_inactive": "#6B21A8", "border_inactive": "#DDD6FE", "hover": "#F5F3FF"}),
            ("📄 Discussion Papers Only", "📄 Discussions", 0, {"bg_active": "#64748B", "fg_active": "#FFFFFF", "border_active": "#475569", "bg_inactive": "#FFFFFF", "fg_inactive": "#334155", "border_inactive": "#CBD5E1", "hover": "#F8FAFC"}),
        ]

        for cat_key, cat_label, init_cnt, custom_col in cat_defs:
            pill = FilterPillButton(cat_key, cat_label, count=init_cnt, colors=custom_col)
            if cat_key == "All Inputs":
                pill.setChecked(True)
            pill.toggled.connect(lambda chk, k=cat_key: self._on_cat_pill_toggled(k, chk))
            self.cat_pills[cat_key] = pill
            self.cat_pill_group.addButton(pill)
            cat_row.addWidget(pill)
        cat_row.addStretch()
        filter_layout.addLayout(cat_row)

        layout.addWidget(filter_frame)

        # 4. Tabbed Visualizations
        self.tabs = QTabWidget()
        self.tabs.setStyleSheet("""
            QTabWidget::pane { border: 1px solid #CBD5E1; border-radius: 6px; background-color: #FFFFFF; }
            QTabBar::tab { padding: 6px 14px; font-weight: bold; font-size: 11px; }
            QTabBar::tab:selected { background-color: #FFFFFF; color: #10B981; border-bottom: 2px solid #10B981; }
        """)

        # Tab 1: Decisions Table
        self.table_decisions = QTableWidget()
        self.table_decisions.setStyleSheet(STYLE_TABLE)
        self.table_decisions.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table_decisions.verticalHeader().setVisible(False)
        self.table_decisions.itemDoubleClicked.connect(self._on_table_item_double_clicked)
        self.tabs.addTab(self.table_decisions, "⚖️ Technical Decisions && Agreements (0)")

        # Tab 2: All Contributions Table
        self.table_contributions = QTableWidget()
        self.table_contributions.setStyleSheet(STYLE_TABLE)
        self.table_contributions.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table_contributions.verticalHeader().setVisible(False)
        self.table_contributions.itemDoubleClicked.connect(self._on_table_item_double_clicked)
        self.tabs.addTab(self.table_contributions, "📄 All Considered Contributions (0)")

        # Tab 3: Liaison Statements Table
        self.table_ls = QTableWidget()
        self.table_ls.setStyleSheet(STYLE_TABLE)
        self.table_ls.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table_ls.verticalHeader().setVisible(False)
        self.table_ls.itemDoubleClicked.connect(self._on_table_item_double_clicked)
        self.tabs.addTab(self.table_ls, "📬 Inter-WG Liaison Statements (0)")

        # Tab 4: WG Distribution Matrix
        self.table_dist = QTableWidget()
        self.table_dist.setStyleSheet(STYLE_TABLE)
        self.table_dist.verticalHeader().setVisible(False)
        self.tabs.addTab(self.table_dist, "📊 WG && Round Distribution")

        # 5. Collapsible Agreement & Decision Detail Reader Drawer
        self.drawer_widget = QFrame()
        self.drawer_widget.setStyleSheet("background-color: #F8FAFC; border: 1px solid #CBD5E1; border-radius: 8px;")
        drawer_layout = QVBoxLayout(self.drawer_widget)
        drawer_layout.setContentsMargins(10, 8, 10, 8)
        drawer_layout.setSpacing(6)

        hdr_layout = QHBoxLayout()
        self.lbl_drawer_title = QLabel("📖 Decision & Agreement Detail Reader")
        self.lbl_drawer_title.setStyleSheet("font-weight: bold; font-size: 12px; color: #1E293B;")

        self.lbl_drawer_provenance = QLabel("")
        self.lbl_drawer_provenance.setStyleSheet("background-color: #DCFCE7; color: #166534; font-weight: bold; font-size: 11px; padding: 2px 8px; border-radius: 4px; border: 1px solid #BBF7D0;")
        self.lbl_drawer_provenance.setVisible(False)

        self.btn_copy_agreement_text = QPushButton("📋 Copy Verbatim Text")
        self.btn_copy_agreement_text.setStyleSheet(STYLE_SECONDARY_BTN)
        self.btn_copy_agreement_text.setToolTip("Copy the full verbatim agreement text or contribution summary to clipboard")
        self.btn_copy_agreement_text.clicked.connect(self._on_copy_detail_text)

        self.btn_toggle_drawer = QPushButton("▲ Collapse")
        self.btn_toggle_drawer.setStyleSheet(STYLE_SECONDARY_BTN)
        self.btn_toggle_drawer.clicked.connect(self._toggle_drawer)

        hdr_layout.addWidget(self.lbl_drawer_title)
        hdr_layout.addWidget(self.lbl_drawer_provenance)
        hdr_layout.addStretch()
        hdr_layout.addWidget(self.btn_copy_agreement_text)
        hdr_layout.addWidget(self.btn_toggle_drawer)
        drawer_layout.addLayout(hdr_layout)

        self.detail_text_edit = QTextEdit()
        self.detail_text_edit.setReadOnly(True)
        self.detail_text_edit.setMinimumHeight(90)
        self.detail_text_edit.setMaximumHeight(220)
        self.detail_text_edit.setStyleSheet("""
            QTextEdit {
                background-color: #FFFFFF;
                border: 1px solid #CBD5E1;
                border-radius: 6px;
                padding: 8px;
                font-family: 'Segoe UI', Arial, sans-serif;
                font-size: 11px;
                color: #1E293B;
            }
        """)
        self.detail_text_edit.setPlaceholderText("Select a decision, contribution, or liaison statement from the table above to read the full text, verbatim agreement wording, and technical provenance.")
        drawer_layout.addWidget(self.detail_text_edit)

        # Connect table selection change events
        self.table_decisions.itemSelectionChanged.connect(self._on_decision_selection_changed)
        self.table_contributions.itemSelectionChanged.connect(self._on_contribution_selection_changed)
        self.table_ls.itemSelectionChanged.connect(self._on_ls_selection_changed)
        self.tabs.currentChanged.connect(self._on_tab_changed)

        # Vertical Splitter
        self.splitter = QSplitter(Qt.Vertical)
        self.splitter.setChildrenCollapsible(False)
        self.splitter.addWidget(self.tabs)
        self.splitter.addWidget(self.drawer_widget)
        self.splitter.setStretchFactor(0, 3)
        self.splitter.setStretchFactor(1, 1)

        layout.addWidget(self.splitter, 1)

        # 6. Bottom Action Ribbon
        action_layout = QHBoxLayout()

        self.btn_open_doc = QPushButton("📄 Open Document")
        self.btn_open_doc.setStyleSheet(STYLE_SECONDARY_BTN)
        self.btn_open_doc.clicked.connect(self._on_open_selected_document)
        action_layout.addWidget(self.btn_open_doc)

        self.btn_open_portal = QPushButton("🌐 Open on 3GPP Portal")
        self.btn_open_portal.setStyleSheet(STYLE_SECONDARY_BTN)
        self.btn_open_portal.clicked.connect(self._on_open_selected_portal)
        action_layout.addWidget(self.btn_open_portal)

        self.btn_copy_md = QPushButton("📋 Copy Markdown Table")
        self.btn_copy_md.setStyleSheet(STYLE_SECONDARY_BTN)
        self.btn_copy_md.clicked.connect(self._on_copy_markdown_table)
        action_layout.addWidget(self.btn_copy_md)

        self.btn_export_csv = QPushButton("💾 Export to CSV...")
        self.btn_export_csv.setStyleSheet(STYLE_SECONDARY_BTN)
        self.btn_export_csv.clicked.connect(self._on_export_csv)
        action_layout.addWidget(self.btn_export_csv)

        action_layout.addStretch()

        self.btn_export_dossier = QPushButton("🚀 Proceed to Full Dossier Export...")
        self.btn_export_dossier.setStyleSheet(STYLE_PRIMARY_BTN)
        self.btn_export_dossier.clicked.connect(self._on_proceed_to_export)
        action_layout.addWidget(self.btn_export_dossier)

        btn_close = QPushButton("Close")
        btn_close.setStyleSheet(STYLE_SECONDARY_BTN)
        btn_close.clicked.connect(self.close)
        action_layout.addWidget(btn_close)

        layout.addLayout(action_layout)

    def _start_collection(self):
        if self._collector_worker and self._collector_worker.isRunning():
            self._collector_worker.cancel()
            self._collector_worker.wait(500)

        self.prog_frame.setVisible(True)
        self.prog_bar.setValue(0)
        self.lbl_prog_status.setText(f"Scanning meeting records across {len(self.rounds)} rounds...")
        self.btn_refresh.setEnabled(False)

        self._collector_worker = InputsCollectorThread(
            meetings_db=self.meetings_db,
            topic=self.topic,
            rounds=self.rounds,
            scope_mode=self.scope_mode,
            parent=None,
        )
        self._collector_worker.progress_status.connect(self._on_scan_progress)
        self._collector_worker.collection_finished.connect(self._on_scan_finished)
        self._collector_worker.collection_failed.connect(self._on_scan_failed)
        self._collector_worker.start()

    def _on_scan_progress(self, msg: str, cur_idx: int, total_count: int):
        self.lbl_prog_status.setText(msg)
        pct = int((cur_idx / max(1, total_count)) * 100)
        self.prog_bar.setValue(pct)

    def _on_scan_finished(self, decisions: List[TopicDecision], contributions: List[Dict[str, Any]], ls_tracker: CrossWgLiaisonTracker):
        self.all_decisions = decisions
        self.all_contributions = contributions
        self.ls_tracker = ls_tracker

        self.prog_frame.setVisible(False)
        self.btn_refresh.setEnabled(True)

        self._update_kpi_strip()
        self._apply_filters()
        self._populate_distribution_table()

    def _on_scan_failed(self, error_msg: str):
        self.prog_frame.setVisible(False)
        self.btn_refresh.setEnabled(True)
        QMessageBox.warning(self, "Collection Warning", f"Could not complete scan:\n{error_msg}")

    def _on_wg_pill_toggled(self, wg: str, checked: bool):
        if checked:
            idx = self.combo_wg_filter.findText(wg)
            if idx >= 0:
                self.combo_wg_filter.blockSignals(True)
                self.combo_wg_filter.setCurrentIndex(idx)
                self.combo_wg_filter.blockSignals(False)
            self._apply_filters()

    def _on_cat_pill_toggled(self, cat_key: str, checked: bool):
        if checked:
            idx = self.combo_type_filter.findText(cat_key)
            if idx >= 0:
                self.combo_type_filter.blockSignals(True)
                self.combo_type_filter.setCurrentIndex(idx)
                self.combo_type_filter.blockSignals(False)

            # Auto-switch to matching view for fast navigation
            if cat_key in ("🟢 Decisions Only (Agreements, pCRs, CRs)", "Text Agreements Only", "Approved pCRs Only", "Approved CRs Only"):
                if self.tabs.currentIndex() != 0:
                    self.tabs.setCurrentIndex(0)
            elif cat_key == "📬 Liaison Statements Only":
                if self.tabs.currentIndex() != 2:
                    self.tabs.setCurrentIndex(2)
            elif cat_key == "📄 Discussion Papers Only":
                if self.tabs.currentIndex() != 1:
                    self.tabs.setCurrentIndex(1)

            self._apply_filters()

    def _reset_filters(self):
        self.edit_filter.clear()
        self.combo_round_filter.setCurrentIndex(0)
        if hasattr(self, "wg_pills") and "All WGs" in self.wg_pills:
            self.wg_pills["All WGs"].setChecked(True)
        if hasattr(self, "cat_pills") and "All Inputs" in self.cat_pills:
            self.cat_pills["All Inputs"].setChecked(True)
        self._apply_filters()

    def _update_kpi_strip(self):
        tot_cnt = len(self.all_contributions)
        dec_cnt = len(self.all_decisions)
        ls_cnt = sum(1 for c in self.all_contributions if "ls" in c.get("type", "").lower())
        disc_cnt = sum(1 for c in self.all_contributions if "discussion" in c.get("type", "").lower())
        wgs_seen = set(c.get("wg") for c in self.all_contributions if c.get("wg"))

        self.kpi_total.setText(f"📊 Total Inputs: <b>{tot_cnt}</b>")
        self.kpi_decisions.setText(f"🟢 Decisions: <b>{dec_cnt}</b>")
        self.kpi_ls.setText(f"📬 Liaisons: <b>{ls_cnt}</b>")
        self.kpi_disc.setText(f"📄 Discussions: <b>{disc_cnt}</b>")
        self.kpi_wgs.setText(f"👥 WGs: <b>{len(wgs_seen)}</b>")

        self.tabs.setTabText(0, f"⚖️ Technical Decisions && Agreements ({dec_cnt})")
        self.tabs.setTabText(1, f"📄 All Considered Contributions ({tot_cnt})")
        self.tabs.setTabText(2, f"📬 Inter-WG Liaison Statements ({ls_cnt})")

        # Update WG pill counts
        if hasattr(self, "wg_pills"):
            if "All WGs" in self.wg_pills:
                self.wg_pills["All WGs"].set_count(tot_cnt)
            for wg, pill in self.wg_pills.items():
                if wg == "All WGs":
                    continue
                cnt = sum(1 for c in self.all_contributions if str(c.get("wg", "")).upper() == wg.upper())
                pill.set_count(cnt)

        # Update Category pill counts
        if hasattr(self, "cat_pills"):
            if "All Inputs" in self.cat_pills:
                self.cat_pills["All Inputs"].set_count(tot_cnt)
            if "🟢 Decisions Only (Agreements, pCRs, CRs)" in self.cat_pills:
                self.cat_pills["🟢 Decisions Only (Agreements, pCRs, CRs)"].set_count(dec_cnt)
            if "Text Agreements Only" in self.cat_pills:
                self.cat_pills["Text Agreements Only"].set_count(sum(1 for d in self.all_decisions if d.decision_type == "Text Agreement"))
            if "Approved pCRs Only" in self.cat_pills:
                self.cat_pills["Approved pCRs Only"].set_count(sum(1 for d in self.all_decisions if d.decision_type == "Approved pCR"))
            if "Approved CRs Only" in self.cat_pills:
                self.cat_pills["Approved CRs Only"].set_count(sum(1 for d in self.all_decisions if d.decision_type == "Approved CR"))
            if "📬 Liaison Statements Only" in self.cat_pills:
                self.cat_pills["📬 Liaison Statements Only"].set_count(ls_cnt)
            if "📄 Discussion Papers Only" in self.cat_pills:
                self.cat_pills["📄 Discussion Papers Only"].set_count(disc_cnt)

    def _apply_filters(self):
        query = self.edit_filter.text().strip().lower()
        selected_wg = self.combo_wg_filter.currentText()
        selected_cat = self.combo_type_filter.currentText()
        selected_round = self.combo_round_filter.currentText()

        # 1. Filter Decisions
        filtered_decs = []
        for d in self.all_decisions:
            if selected_wg != "All WGs" and d.wg.upper() != selected_wg.upper():
                continue
            if selected_round != "All Meeting Rounds":
                # Check cohort match
                matching_round = next((r for r in self.rounds if r.key == d.cohort_key), None)
                if not matching_round or matching_round.label != selected_round:
                    continue
            if selected_cat == "Text Agreements Only" and d.decision_type != "Text Agreement":
                continue
            if selected_cat == "Approved pCRs Only" and d.decision_type != "Approved pCR":
                continue
            if selected_cat == "Approved CRs Only" and d.decision_type != "Approved CR":
                continue
            if selected_cat == "📬 Liaison Statements Only" and d.decision_type != "Agreed LS":
                continue
            if selected_cat == "📄 Discussion Papers Only":
                continue

            if query:
                haystack = f"{d.tdoc_id or ''} {d.title_or_summary} {d.source} {d.ai} {d.target_spec}".lower()
                if query not in haystack:
                    continue
            filtered_decs.append(d)

        self._populate_decisions_table(filtered_decs)

        # 2. Filter Contributions
        filtered_contribs = []
        for c in self.all_contributions:
            if selected_wg != "All WGs" and c.get("wg", "").upper() != selected_wg.upper():
                continue
            if selected_round != "All Meeting Rounds" and c.get("round_label") != selected_round:
                continue
            if selected_cat == "🟢 Decisions Only (Agreements, pCRs, CRs)" and not c.get("is_decision"):
                continue
            if selected_cat == "Approved pCRs Only" and c.get("decision_type") != "Approved pCR":
                continue
            if selected_cat == "Approved CRs Only" and c.get("decision_type") != "Approved CR":
                continue
            if selected_cat == "Text Agreements Only" and c.get("decision_type") != "Text Agreement":
                continue
            if selected_cat == "📬 Liaison Statements Only" and "ls" not in c.get("type", "").lower():
                continue
            if selected_cat == "📄 Discussion Papers Only" and "discussion" not in c.get("type", "").lower():
                continue

            if query:
                haystack = f"{c.get('tdoc_id', '')} {c.get('title', '')} {c.get('source', '')} {c.get('ai', '')} {c.get('ai_desc', '')}".lower()
                if query not in haystack:
                    continue
            filtered_contribs.append(c)

        self._populate_contributions_table(filtered_contribs)

        # 3. Filter Liaison Statements
        filtered_ls = [c for c in filtered_contribs if "ls" in c.get("type", "").lower()]
        self._populate_ls_table(filtered_ls)

    def _populate_decisions_table(self, decisions: List[TopicDecision]):
        self.table_decisions.setUpdatesEnabled(False)
        self.table_decisions.blockSignals(True)
        try:
            self._displayed_decisions = list(decisions)
            self.table_decisions.setRowCount(len(decisions))
            self.table_decisions.setColumnCount(8)
            self.table_decisions.setHorizontalHeaderLabels([
                "Decision Type", "WG", "Meeting", "Date", "TDoc ID", "Decision / Substantive Scope", "Target Spec", "Agenda Item"
            ])

            header = self.table_decisions.horizontalHeader()
            for i in range(8):
                header.setSectionResizeMode(i, QHeaderView.Interactive)

            for row, d in enumerate(decisions):
                # 0: Decision Type Badge & Provenance
                if d.decision_type in ("Text Agreement", "Working Assumption"):
                    badge_text = f"🟢 {d.decision_type} (Notes)"
                    tip_text = (
                        f"<b>Binding Consensus Text</b> (Chairman's Notes AI {d.ai or 'N/A'})<br/>"
                        f"<b>Meeting:</b> {d.wg} #{d.meeting_number} ({d.date or 'N/A'})<br/>"
                        f"<b>Background Proposal:</b> {d.tdoc_id or 'None'} (Noted/Discussed)<br/>"
                        f"<i>In RAN WGs, only the verbatim text in Chairman's Notes is binding consensus.</i>"
                    )
                elif d.decision_type in ("Approved pCR", "Approved CR"):
                    badge_text = f"🟢 {d.decision_type} (Doc)"
                    tip_text = (
                        f"<b>Approved Specification Change</b> ({d.tdoc_id or 'N/A'})<br/>"
                        f"<b>Meeting:</b> {d.wg} #{d.meeting_number} ({d.date or 'N/A'})<br/>"
                        f"<b>Target Specification:</b> {d.target_spec or 'N/A'}<br/>"
                        f"<i>Full document attachment approved as baseline change.</i>"
                    )
                else:
                    badge_text = f"{d.badge_icon} {d.decision_type}"
                    tip_text = f"<b>{d.decision_type}</b> ({d.wg} #{d.meeting_number})<br/><b>AI:</b> {d.ai or 'N/A'}"

                badge_item = QTableWidgetItem(badge_text)
                badge_item.setTextAlignment(Qt.AlignCenter)
                badge_item.setToolTip(tip_text)
                bg_color, fg_color = BADGE_COLORS.get(d.decision_type, ("#F1F5F9", "#1E293B"))
                badge_item.setBackground(QColor(bg_color))
                badge_item.setForeground(QColor(fg_color))
                font = badge_item.font()
                font.setBold(True)
                badge_item.setFont(font)
                self.table_decisions.setItem(row, 0, badge_item)

                # 1: WG
                wg_item = QTableWidgetItem(d.wg)
                wg_item.setTextAlignment(Qt.AlignCenter)
                self.table_decisions.setItem(row, 1, wg_item)

                # 2: Meeting #
                mtg_item = QTableWidgetItem(f"#{d.meeting_number}")
                mtg_item.setTextAlignment(Qt.AlignCenter)
                self.table_decisions.setItem(row, 2, mtg_item)

                # 3: Date
                date_item = QTableWidgetItem(d.date)
                date_item.setTextAlignment(Qt.AlignCenter)
                self.table_decisions.setItem(row, 3, date_item)

                # 4: TDoc ID
                tid_str = d.tdoc_id or "—"
                tid_item = QTableWidgetItem(tid_str)
                tid_item.setTextAlignment(Qt.AlignCenter)
                if d.tdoc_id:
                    tid_item.setForeground(QColor("#0369A1"))
                    f = tid_item.font()
                    f.setBold(True)
                    tid_item.setFont(f)
                self.table_decisions.setItem(row, 4, tid_item)

                # 5: Decision / Substantive Scope
                scope_item = QTableWidgetItem(d.title_or_summary)
                scope_item.setToolTip(f"<b>{d.decision_type} ({d.wg} #{d.meeting_number})</b><br/>{d.title_or_summary}")
                self.table_decisions.setItem(row, 5, scope_item)

                # 6: Target Spec
                spec_item = QTableWidgetItem(d.target_spec or "—")
                spec_item.setTextAlignment(Qt.AlignCenter)
                self.table_decisions.setItem(row, 6, spec_item)

                # 7: Agenda Item
                ai_item = QTableWidgetItem(d.ai or "—")
                ai_item.setTextAlignment(Qt.AlignCenter)
                self.table_decisions.setItem(row, 7, ai_item)

            self.table_decisions.resizeColumnsToContents()
            header.setSectionResizeMode(5, QHeaderView.Stretch)
        finally:
            self.table_decisions.blockSignals(False)
            self.table_decisions.setUpdatesEnabled(True)

        if decisions:
            self.table_decisions.selectRow(0)
        else:
            if self.tabs.currentIndex() == 0:
                self.detail_text_edit.clear()
                self.lbl_drawer_provenance.setVisible(False)
                self._current_detail_text = ""

    def _populate_contributions_table(self, contribs: List[Dict[str, Any]]):
        self.table_contributions.setUpdatesEnabled(False)
        self.table_contributions.blockSignals(True)
        try:
            self._displayed_contribs = list(contribs)
            self.table_contributions.setRowCount(len(contribs))
            self.table_contributions.setColumnCount(9)
            self.table_contributions.setHorizontalHeaderLabels([
                "TDoc ID", "WG", "Meeting", "Date", "Agenda Item", "Type", "Status", "Title", "Source Company"
            ])

            header = self.table_contributions.horizontalHeader()
            for i in range(9):
                header.setSectionResizeMode(i, QHeaderView.Interactive)

            for row, c in enumerate(contribs):
                # 0: TDoc ID
                tid_item = QTableWidgetItem(c.get("tdoc_id", ""))
                tid_item.setTextAlignment(Qt.AlignCenter)
                tid_item.setForeground(QColor("#0369A1"))
                f = tid_item.font()
                f.setBold(True)
                tid_item.setFont(f)
                self.table_contributions.setItem(row, 0, tid_item)

                # 1: WG
                self.table_contributions.setItem(row, 1, QTableWidgetItem(c.get("wg", "")))
                # 2: Meeting #
                self.table_contributions.setItem(row, 2, QTableWidgetItem(f"#{c.get('meeting_number', '')}"))
                # 3: Date
                self.table_contributions.setItem(row, 3, QTableWidgetItem(c.get("date", "")))

                # 4: AI
                ai_item = QTableWidgetItem(c.get("ai", ""))
                ai_item.setToolTip(c.get("ai_desc", ""))
                self.table_contributions.setItem(row, 4, ai_item)

                # 5: Type
                t_type = c.get("type", "")
                t_item = QTableWidgetItem(t_type)
                t_item.setTextAlignment(Qt.AlignCenter)
                bg, fg = BADGE_COLORS.get(t_type, ("#FFFFFF", "#334155"))
                t_item.setBackground(QColor(bg))
                t_item.setForeground(QColor(fg))
                self.table_contributions.setItem(row, 5, t_item)

                # 6: Status
                status_item = QTableWidgetItem(c.get("status", ""))
                status_item.setTextAlignment(Qt.AlignCenter)
                if any(s in c.get("status", "").lower() for s in ("approved", "agreed", "endorsed")):
                    status_item.setForeground(QColor("#15803D"))
                    f_s = status_item.font()
                    f_s.setBold(True)
                    status_item.setFont(f_s)
                self.table_contributions.setItem(row, 6, status_item)

                # 7: Title
                title_item = QTableWidgetItem(c.get("title", ""))
                title_item.setToolTip(c.get("title", ""))
                self.table_contributions.setItem(row, 7, title_item)

                # 8: Source Company
                clean_co = c.get("clean_source") or CompanySanitizer.get_confirmed_source_text(c.get("source", "")) or c.get("source", "")
                co_item = QTableWidgetItem(clean_co)
                co_item.setToolTip(c.get("source", ""))
                self.table_contributions.setItem(row, 8, co_item)

            self.table_contributions.resizeColumnsToContents()
            header.setSectionResizeMode(7, QHeaderView.Stretch)
        finally:
            self.table_contributions.blockSignals(False)
            self.table_contributions.setUpdatesEnabled(True)

        if contribs and self.tabs.currentIndex() == 1:
            self.table_contributions.selectRow(0)

    def _populate_ls_table(self, ls_contribs: List[Dict[str, Any]]):
        self.table_ls.setUpdatesEnabled(False)
        self.table_ls.blockSignals(True)
        try:
            self._displayed_ls = list(ls_contribs)
            self.table_ls.setRowCount(len(ls_contribs))
            self.table_ls.setColumnCount(8)
            self.table_ls.setHorizontalHeaderLabels([
                "Direction", "TDoc ID", "WG", "Meeting", "Title", "Status", "Reply To", "Reply In"
            ])

            header = self.table_ls.horizontalHeader()
            for i in range(8):
                header.setSectionResizeMode(i, QHeaderView.Interactive)

            for row, c in enumerate(ls_contribs):
                raw = c.get("raw_row", {})
                t_type = c.get("type", "")
                is_in = "in" in t_type.lower()
                dir_text = "📥 LS IN" if is_in else "📤 LS OUT"
                dir_item = QTableWidgetItem(dir_text)
                dir_item.setTextAlignment(Qt.AlignCenter)
                bg, fg = BADGE_COLORS.get("LS in" if is_in else "LS out", ("#F3E8FF", "#6B21A8"))
                dir_item.setBackground(QColor(bg))
                dir_item.setForeground(QColor(fg))
                f = dir_item.font()
                f.setBold(True)
                dir_item.setFont(f)
                self.table_ls.setItem(row, 0, dir_item)

                # TDoc
                tid_item = QTableWidgetItem(c.get("tdoc_id", ""))
                tid_item.setTextAlignment(Qt.AlignCenter)
                tid_item.setForeground(QColor("#0369A1"))
                self.table_ls.setItem(row, 1, tid_item)

                self.table_ls.setItem(row, 2, QTableWidgetItem(c.get("wg", "")))
                self.table_ls.setItem(row, 3, QTableWidgetItem(f"#{c.get('meeting_number', '')}"))
                self.table_ls.setItem(row, 4, QTableWidgetItem(c.get("title", "")))
                self.table_ls.setItem(row, 5, QTableWidgetItem(c.get("status", "")))
                self.table_ls.setItem(row, 6, QTableWidgetItem(raw.get("Reply to", "") or raw.get("reply_to", "") or "—"))
                self.table_ls.setItem(row, 7, QTableWidgetItem(raw.get("Reply in", "") or raw.get("reply_in", "") or "—"))

            self.table_ls.resizeColumnsToContents()
            header.setSectionResizeMode(4, QHeaderView.Stretch)
        finally:
            self.table_ls.blockSignals(False)
            self.table_ls.setUpdatesEnabled(True)

        if ls_contribs and self.tabs.currentIndex() == 2:
            self.table_ls.selectRow(0)

    def _toggle_drawer(self):
        is_hidden = self.detail_text_edit.isHidden()
        self.detail_text_edit.setVisible(is_hidden)
        self.btn_toggle_drawer.setText("▲ Collapse" if is_hidden else "▼ Expand")

    def _on_copy_detail_text(self):
        text = self._current_detail_text or self.detail_text_edit.toPlainText().strip()
        if text:
            app = QApplication.instance()
            if app:
                app.clipboard().setText(text)
            self.lbl_drawer_title.setText("📖 Decision & Agreement Detail Reader — 📋 Copied!")
            QTimer.singleShot(2000, lambda: self.lbl_drawer_title.setText("📖 Decision & Agreement Detail Reader"))

    def _on_decision_selection_changed(self):
        row = self.table_decisions.currentRow()
        if row < 0 or row >= len(self._displayed_decisions):
            return
        d = self._displayed_decisions[row]
        self._current_detail_text = d.title_or_summary

        if d.decision_type in ("Text Agreement", "Working Assumption"):
            prov_text = f"🟢 {d.decision_type} (Chairman's Notes)"
            prov_style = "background-color: #DCFCE7; color: #166534; font-weight: bold; font-size: 11px; padding: 2px 8px; border-radius: 4px; border: 1px solid #BBF7D0;"
        elif d.decision_type in ("Approved pCR", "Approved CR"):
            prov_text = f"🟢 {d.decision_type} (Approved Document)"
            prov_style = "background-color: #ECFDF5; color: #065F46; font-weight: bold; font-size: 11px; padding: 2px 8px; border-radius: 4px; border: 1px solid #A7F3D0;"
        else:
            prov_text = f"{d.badge_icon} {d.decision_type}"
            prov_style = "background-color: #F1F5F9; color: #334155; font-weight: bold; font-size: 11px; padding: 2px 8px; border-radius: 4px; border: 1px solid #CBD5E1;"

        self.lbl_drawer_provenance.setText(prov_text)
        self.lbl_drawer_provenance.setStyleSheet(prov_style)
        self.lbl_drawer_provenance.setVisible(True)

        meta_parts = [
            f"<b>WG:</b> {d.wg}",
            f"<b>Meeting:</b> #{d.meeting_number} ({d.date or 'N/A'})",
            f"<b>AI:</b> <code>{d.ai or 'N/A'}</code>",
        ]
        if d.target_spec:
            meta_parts.append(f"<b>Spec:</b> <code>{d.target_spec}</code>")
        if d.wi:
            meta_parts.append(f"<b>WI:</b> <code>{d.wi}</code>")
        if d.tdoc_id:
            meta_parts.append(f"<b>TDoc / Background Proposal:</b> <b>{d.tdoc_id}</b> ({d.status})")

        html_lines = [
            f"<p style='margin:0 0 6px 0; font-size:11px;'>{' &nbsp;|&nbsp; '.join(meta_parts)}</p>"
        ]

        if d.decision_type in ("Text Agreement", "Working Assumption"):
            html_lines.append(
                "<div style='background-color:#F0FDF4; border-left:3px solid #16A34A; padding:4px 8px; margin-bottom:8px; font-size:11px; color:#166534;'>"
                "<b>Binding Consensus:</b> In 3GPP RAN Working Groups, only the verbatim text below recorded in the Chairman's Notes represents formal consensus. Cited background proposal was Noted.</div>"
            )
        elif d.decision_type in ("Approved pCR", "Approved CR"):
            html_lines.append(
                f"<div style='background-color:#ECFDF5; border-left:3px solid #059669; padding:4px 8px; margin-bottom:8px; font-size:11px; color:#065F46;'>"
                f"<b>Approved Specification Change:</b> The full document attachment of <b>{d.tdoc_id}</b> was formally approved as a baseline change for <code>{d.target_spec}</code>.</div>"
            )

        verbatim_html = d.title_or_summary.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace("\n", "<br/>")
        html_lines.append(f"<div style='background-color:#FFFFFF; border:1px solid #CBD5E1; border-radius:4px; padding:8px; font-family:Consolas, monospace; font-size:11px; line-height:1.4;'>{verbatim_html}</div>")

        self.detail_text_edit.setHtml("".join(html_lines))

    def _on_contribution_selection_changed(self):
        row = self.table_contributions.currentRow()
        if row < 0 or row >= len(self._displayed_contribs):
            return
        c = self._displayed_contribs[row]
        tid = c.get("tdoc_id", "")
        raw = c.get("raw_row", {})
        abstract = raw.get("Abstract", "") or raw.get("abstract", "") or raw.get("Secretary Remarks", "")

        text_parts = [f"{tid}: {c.get('title', '')}", f"Source: {c.get('source', '')}", f"Type: {c.get('type', '')} | Status: {c.get('status', '')}"]
        if abstract:
            text_parts.append(f"\nAbstract / Remarks:\n{abstract}")
        self._current_detail_text = "\n".join(text_parts)

        self.lbl_drawer_provenance.setText(f"📄 {c.get('type', 'TDoc')} ({c.get('status', 'N/A')})")
        self.lbl_drawer_provenance.setStyleSheet("background-color: #F1F5F9; color: #334155; font-weight: bold; font-size: 11px; padding: 2px 8px; border-radius: 4px; border: 1px solid #CBD5E1;")
        self.lbl_drawer_provenance.setVisible(True)

        meta_parts = [
            f"<b>TDoc:</b> <b>{tid}</b>",
            f"<b>WG:</b> {c.get('wg')} #{c.get('meeting_number')}",
            f"<b>AI:</b> <code>{c.get('ai')}</code>",
            f"<b>Source:</b> {c.get('source')}",
            f"<b>Status:</b> {c.get('status')}",
        ]
        if c.get("wi"):
            meta_parts.append(f"<b>WI:</b> <code>{c.get('wi')}</code>")
        if c.get("spec"):
            meta_parts.append(f"<b>Spec:</b> <code>{c.get('spec')}</code>")

        html_lines = [
            f"<p style='margin:0 0 4px 0; font-size:11px;'>{' &nbsp;|&nbsp; '.join(meta_parts)}</p>",
            f"<p style='margin:0 0 6px 0; font-size:12px; font-weight:bold; color:#1E293B;'>{c.get('title')}</p>",
        ]
        if abstract:
            abs_html = abstract.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace("\n", "<br/>")
            html_lines.append(f"<div style='background-color:#FFFFFF; border:1px solid #CBD5E1; border-radius:4px; padding:8px; font-size:11px;'><b>Abstract / Remarks:</b><br/>{abs_html}</div>")

        self.detail_text_edit.setHtml("".join(html_lines))

    def _on_ls_selection_changed(self):
        row = self.table_ls.currentRow()
        if row < 0 or row >= len(self._displayed_ls):
            return
        c = self._displayed_ls[row]
        tid = c.get("tdoc_id", "")
        raw = c.get("raw_row", {})
        reply_to = raw.get("Reply to", "") or raw.get("reply_to", "") or "—"
        reply_in = raw.get("Reply in", "") or raw.get("reply_in", "") or "—"

        self._current_detail_text = (
            f"Liaison Statement: {tid}\nTitle: {c.get('title')}\nSource: {c.get('source')}\n"
            f"Reply To: {reply_to}\nReply In: {reply_in}\nStatus: {c.get('status')}"
        )

        t_type = c.get("type", "")
        is_in = "in" in t_type.lower()
        dir_text = "📥 LS IN" if is_in else "📤 LS OUT"

        self.lbl_drawer_provenance.setText(f"{dir_text} ({c.get('status', 'N/A')})")
        self.lbl_drawer_provenance.setStyleSheet("background-color: #EDE9FE; color: #5B21B6; font-weight: bold; font-size: 11px; padding: 2px 8px; border-radius: 4px; border: 1px solid #DDD6FE;")
        self.lbl_drawer_provenance.setVisible(True)

        meta_parts = [
            f"<b>Liaison Statement:</b> <b>{tid}</b> ({dir_text})",
            f"<b>WG:</b> {c.get('wg')} #{c.get('meeting_number')}",
            f"<b>Source:</b> {c.get('source')}",
            f"<b>Status:</b> {c.get('status')}",
        ]
        html_lines = [
            f"<p style='margin:0 0 4px 0; font-size:11px;'>{' &nbsp;|&nbsp; '.join(meta_parts)}</p>",
            f"<p style='margin:0 0 6px 0; font-size:12px; font-weight:bold; color:#1E293B;'>{c.get('title')}</p>",
            f"<p style='margin:0; font-size:11px; color:#475569;'><b>Reply To:</b> <code>{reply_to}</code> &nbsp;|&nbsp; <b>Reply In:</b> <code>{reply_in}</code></p>",
        ]
        self.detail_text_edit.setHtml("".join(html_lines))

    def _on_tab_changed(self, idx: int):
        if idx == 0:
            self._on_decision_selection_changed()
        elif idx == 1:
            self._on_contribution_selection_changed()
        elif idx == 2:
            self._on_ls_selection_changed()
        else:
            self.detail_text_edit.clear()
            self.lbl_drawer_provenance.setVisible(False)
            self._current_detail_text = ""

    def _populate_distribution_table(self):
        wgs = sorted(list(self.topic.target_wgs))
        rounds = self.rounds

        self.table_dist.setRowCount(len(wgs) + 1)
        self.table_dist.setColumnCount(len(rounds) + 2)

        col_labels = ["Working Group"] + [r.label for r in rounds] + ["Total Across Scope"]
        self.table_dist.setHorizontalHeaderLabels(col_labels)

        header = self.table_dist.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        for i in range(1, len(rounds) + 1):
            header.setSectionResizeMode(i, QHeaderView.Stretch)
        header.setSectionResizeMode(len(rounds) + 1, QHeaderView.ResizeToContents)

        round_totals = {r.key: 0 for r in rounds}
        grand_total = 0

        for r_idx, wg in enumerate(wgs):
            wg_item = QTableWidgetItem(f"👥 {wg}")
            f = wg_item.font()
            f.setBold(True)
            wg_item.setFont(f)
            self.table_dist.setItem(r_idx, 0, wg_item)

            wg_row_total = 0
            for c_idx, m_round in enumerate(rounds, start=1):
                count = sum(1 for c in self.all_contributions if c.get("wg") == wg and c.get("round_key") == m_round.key)
                dec_cnt = sum(1 for d in self.all_decisions if d.wg == wg and d.cohort_key == m_round.key)
                text = f"{count} docs ({dec_cnt} dec)" if count or dec_cnt else "—"
                item = QTableWidgetItem(text)
                item.setTextAlignment(Qt.AlignCenter)
                self.table_dist.setItem(r_idx, c_idx, item)
                wg_row_total += count
                round_totals[m_round.key] += count

            tot_item = QTableWidgetItem(f"<b>{wg_row_total}</b>")
            tot_item.setTextAlignment(Qt.AlignCenter)
            self.table_dist.setItem(r_idx, len(rounds) + 1, tot_item)
            grand_total += wg_row_total

        # Total Summary Row
        last_row = len(wgs)
        tot_label = QTableWidgetItem("Total Documents")
        f_t = tot_label.font()
        f_t.setBold(True)
        tot_label.setFont(f_t)
        self.table_dist.setItem(last_row, 0, tot_label)

        for c_idx, m_round in enumerate(rounds, start=1):
            r_tot = QTableWidgetItem(str(round_totals[m_round.key]))
            r_tot.setTextAlignment(Qt.AlignCenter)
            f_r = r_tot.font()
            f_r.setBold(True)
            r_tot.setFont(f_r)
            self.table_dist.setItem(last_row, c_idx, r_tot)

        grand_item = QTableWidgetItem(str(grand_total))
        grand_item.setTextAlignment(Qt.AlignCenter)
        f_g = grand_item.font()
        f_g.setBold(True)
        grand_item.setFont(f_g)
        self.table_dist.setItem(last_row, len(rounds) + 1, grand_item)

    def _get_active_table(self) -> Optional[QTableWidget]:
        idx = self.tabs.currentIndex()
        if idx == 0:
            return self.table_decisions
        elif idx == 1:
            return self.table_contributions
        elif idx == 2:
            return self.table_ls
        return None

    def _get_selected_tdoc_id(self) -> Optional[str]:
        table = self._get_active_table()
        if not table:
            return None
        row = table.currentRow()
        if row < 0:
            return None
        if table == self.table_decisions:
            val = table.item(row, 4)
            return val.text().strip() if val and val.text() != "—" else None
        elif table == self.table_contributions:
            val = table.item(row, 0)
            return val.text().strip() if val else None
        elif table == self.table_ls:
            val = table.item(row, 1)
            return val.text().strip() if val else None
        return None

    def _on_table_item_double_clicked(self, item: QTableWidgetItem):
        self._on_open_selected_document()

    def _on_open_selected_document(self):
        tid = self._get_selected_tdoc_id()
        if not tid:
            QMessageBox.information(self, "No Document Selected", "Please select a row with a valid TDoc ID.")
            return

        # Find in candidate contributions
        contrib = next((c for c in self.all_contributions if c.get("tdoc_id") == tid), None)
        if not contrib:
            QMessageBox.information(self, "TDoc Info", f"TDoc ID: {tid}")
            return

        from modules.meetings.core.settings import MeetingsSettings
        cache_dir = MeetingsSettings().cache_dir
        folder_name = contrib.get("folder_name", "")
        mtg_dir = Path(cache_dir) / folder_name if cache_dir and folder_name else None

        # Check local files
        if mtg_dir and mtg_dir.is_dir():
            for p in mtg_dir.rglob(f"{tid}*"):
                if p.is_file() and not p.name.endswith(".json"):
                    try:
                        os.startfile(str(p))
                        return
                    except Exception as e:
                        logger.debug(f"Failed to open {p}: {e}")

        # If not cached locally, offer portal open
        ret = QMessageBox.question(
            self,
            "Document Not Cached",
            f"TDoc {tid} is not yet downloaded locally in meeting folder '{folder_name}'.\n\nWould you like to open it on the 3GPP Portal?",
            QMessageBox.Yes | QMessageBox.No,
        )
        if ret == QMessageBox.Yes:
            self._on_open_selected_portal()

    def _on_open_selected_portal(self):
        tid = self._get_selected_tdoc_id()
        if not tid:
            QMessageBox.information(self, "No Document Selected", "Please select a row with a valid TDoc ID.")
            return
        contrib = next((c for c in self.all_contributions if c.get("tdoc_id") == tid), None)
        folder_name = contrib.get("folder_name", "") if contrib else ""
        url = resolve_tdoc_download_url(tid, folder_name)
        if url:
            import webbrowser
            webbrowser.open(url)
        else:
            QMessageBox.warning(self, "URL Error", f"Could not construct 3GPP Portal URL for {tid}.")

    def _on_copy_markdown_table(self):
        table = self._get_active_table()
        if not table or table.rowCount() == 0:
            QMessageBox.information(self, "Empty Table", "No rows available to copy.")
            return

        cols = table.columnCount()
        headers = [table.horizontalHeaderItem(c).text() for c in range(cols)]
        lines = []
        lines.append("| " + " | ".join(headers) + " |")
        lines.append("| " + " | ".join([":---"] * cols) + " |")

        for r in range(table.rowCount()):
            row_vals = []
            for c in range(cols):
                it = table.item(r, c)
                val = it.text().replace("\n", " ").replace("|", "\\|").strip() if it else ""
                row_vals.append(val)
            lines.append("| " + " | ".join(row_vals) + " |")

        md_text = "\n".join(lines)
        QApplication.clipboard().setText(md_text)
        QMessageBox.information(self, "Copied", f"Copied {table.rowCount()} rows formatted as Markdown to clipboard.")

    def _on_export_csv(self):
        table = self._get_active_table()
        if not table or table.rowCount() == 0:
            QMessageBox.information(self, "Empty Table", "No rows available to export.")
            return

        path, _ = QFileDialog.getSaveFileName(
            self,
            "Export Table to CSV",
            f"Topic_Inputs_{self.topic.id}.csv",
            "CSV Files (*.csv);;All Files (*.*)",
        )
        if not path:
            return

        try:
            with open(path, "w", newline="", encoding="utf-8-sig") as f:
                writer = csv.writer(f)
                headers = [table.horizontalHeaderItem(c).text() for c in range(table.columnCount())]
                writer.writerow(headers)
                for r in range(table.rowCount()):
                    row_vals = [table.item(r, c).text().strip() if table.item(r, c) else "" for c in range(table.columnCount())]
                    writer.writerow(row_vals)
            QMessageBox.information(self, "Export Successful", f"Saved {table.rowCount()} rows to:\n{path}")
        except Exception as e:
            QMessageBox.critical(self, "Export Failed", f"Could not write CSV:\n{e}")

    def _on_proceed_to_export(self):
        self.export_requested.emit()
        self.close()

    def closeEvent(self, event):
        if self._collector_worker and self._collector_worker.isRunning():
            self._collector_worker.cancel()
            self._collector_worker.wait(3000)
        super().closeEvent(event)
