# --- File: src/modules/meetings/ui/session_plan_window.py ---
"""
Primary Desktop GUI for 3GPP SA2 Session Plan Management.
Provides Overview Grid with Time Labels, Agenda Mode (Gap Analysis with 1-Click Track),
Drafting Radar, My Topics, Timezone Synchronization, Diffing, and Exporting.
Includes direct drag & drop quick import on the GUI and rich tooltips.
"""
import logging
import re
import webbrowser
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Dict, List, Optional, Set, Tuple

from PyQt5.QtCore import QPoint, Qt, QTimer, pyqtSignal
from PyQt5.QtGui import QColor, QCursor, QFont, QIcon, QKeySequence
from PyQt5.QtWidgets import (
    QAction,
    QApplication,
    QButtonGroup,
    QComboBox,
    QDialog,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMenu,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QSplitter,
    QStackedWidget,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from core.utils.folder_opener import open_folder
from modules.meetings.core.agenda_manager import AgendaManager
from modules.meetings.core.session_plan_db import SessionPlanDatabase
from modules.meetings.core.session_plan_diff import SessionPlanDiffEngine
from modules.meetings.core.session_plan_exporter import SessionPlanExporter
from modules.meetings.core.session_plan_models import (
    AgendaSlotItem,
    MeetingStreamInfo,
    SessionPlanSnapshot,
    SessionSlot,
    UserSessionAnnotation,
)
from modules.meetings.core.session_plan_parser import SessionPlanParser
from modules.meetings.core.session_plan_timezone import SessionPlanTimezone
from modules.meetings.ui.session_plan_dialogs import (
    SessionPlanDiffDialog,
    SessionPlanImportDialog,
    SessionPlanNoteDialog,
    SessionPlanTimezoneDialog,
    SessionPlanTrackedTopicsDialog,
)
from modules.meetings.ui.tdocs_drop_overlay import DropOverlayWidget

logger = logging.getLogger(__name__)

DAY_NAMES = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
QUARTER_ORDER = ["Q0", "Q1", "Q2", "Lunch", "Q3", "Q4", "Q5"]

QUARTER_DISPLAY_LABELS = {
    "Q0": "Q0\n08:00 – 08:50",
    "Q1": "Q1\n09:00 – 10:30",
    "Q2": "Q2\n11:00 – 12:30",
    "Lunch": "Lunch\n13:00 – 14:00",
    "Q3": "Q3\n14:00 – 15:30",
    "Q4": "Q4\n16:00 – 17:30",
    "Q5": "Q5\n18:00 – 19:30",
}


class SessionPlanWindow(QWidget):
    """Primary window for exploring, comparing, and managing meeting session plans."""

    def __init__(
        self,
        meeting_dir: Path,
        mtg_info: dict,
        agenda_map: Optional[dict] = None,
        tracked_tokens: Optional[Set[str]] = None,
        parent=None,
    ):
        super().__init__(parent)
        self.meeting_dir = Path(meeting_dir)
        self.mtg_info = mtg_info or {}
        self.meeting_number = str(mtg_info.get("meeting_number", "177"))
        self.location = str(mtg_info.get("location", ""))
        self.start_date_str = str(mtg_info.get("start_date", "2026-10-12"))
        self.agenda_map = agenda_map or AgendaManager.load_local_agenda(self.meeting_dir / "Agenda")

        # Initialize SQLite database under <meeting_dir>/Agenda/session_plans.db
        self.db = SessionPlanDatabase(self.meeting_dir / "Agenda" / "session_plans.db")

        # Load tracked tokens (merge Work Items watchlist with session_plans.db persistent tracked topics)
        self.external_tracked_tokens = {t.lower() for t in (tracked_tokens or set())}
        self.db_tracked = self.db.get_tracked_topics(self.meeting_number)
        self.tracked_tokens = set(self.external_tracked_tokens) | set(self.db_tracked.keys())

        # Timezone state
        self.tz_offset_min = 120  # Default +2 hours (Europe summer/Prague)
        self.tz_name = "Europe/Prague (UTC+02:00)"
        self._init_timezone()

        self.active_snapshot: Optional[SessionPlanSnapshot] = None
        self.annotations: Dict[str, UserSessionAnnotation] = {}

        self.setWindowTitle(f"🗓️ Session Plan — SA2 #{self.meeting_number} ({self.location})")
        self.resize(1380, 840)
        self.setAcceptDrops(True)

        self._setup_ui()
        self.drop_overlay = DropOverlayWidget(self)

        self._load_data()

        # Real-time clock timer
        self.clock_timer = QTimer(self)
        self.clock_timer.timeout.connect(self._update_live_clock)
        self.clock_timer.start(1000)
        self._update_live_clock()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if hasattr(self, "drop_overlay") and self.drop_overlay:
            self.drop_overlay.setGeometry(self.rect())

    # --- Drag & Drop Direct Ingestion ---

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            for url in event.mimeData().urls():
                ext = Path(url.toLocalFile()).suffix.lower()
                if ext in [".docx", ".doc", ".zip", ".htm", ".html", ".csv", ".txt"]:
                    if hasattr(self, "drop_overlay"):
                        self.drop_overlay.setGeometry(self.rect())
                        self.drop_overlay.show()
                        self.drop_overlay.raise_()
                    event.acceptProposedAction()
                    return
        event.ignore()

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()

    def dragLeaveEvent(self, event):
        if hasattr(self, "drop_overlay"):
            self.drop_overlay.hide()
        event.accept()

    def dropEvent(self, event):
        if hasattr(self, "drop_overlay"):
            self.drop_overlay.hide()

        valid_files = []
        for url in event.mimeData().urls():
            p = Path(url.toLocalFile())
            if p.suffix.lower() in [".docx", ".doc", ".zip", ".htm", ".html", ".csv", ".txt"]:
                valid_files.append(p)

        if valid_files:
            event.acceptProposedAction()
            for vf in valid_files:
                self._handle_manual_import("file", str(vf))
        else:
            event.ignore()

    # --- Timezone Initialization ---

    def _init_timezone(self):
        """Initializes venue timezone from DB metadata, Agenda deadlines, or city lookup."""
        meta = self.db.get_meeting_meta(self.meeting_number)
        if meta and meta.get("tz_offset_minutes") is not None:
            self.tz_offset_min = meta["tz_offset_minutes"]
            self.tz_name = meta.get("tz_name") or f"UTC{self.tz_offset_min//60:+03d}:00"
            return

        res = SessionPlanTimezone.lookup_city_timezone(self.location)
        if res:
            self.tz_offset_min, self.tz_name = res
            self.db.set_meeting_meta(
                meeting_number=self.meeting_number,
                location=self.location,
                tz_offset_minutes=self.tz_offset_min,
                tz_name=self.tz_name,
                tz_confirmed=0,
            )

    # --- UI Layout ---

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 10, 12, 10)
        layout.setSpacing(8)

        # 1. Top Ribbon
        self._setup_top_ribbon(layout)

        # 2. Sub-Ribbon: Navigation & Filters
        self._setup_navigation_ribbon(layout)

        # 3. Stacked View Pages
        self.view_stack = QStackedWidget()
        self._setup_overview_view()
        self._setup_agenda_view()
        self._setup_drafting_view()
        self._setup_tracked_view()

        layout.addWidget(self.view_stack, 1)

    def _setup_top_ribbon(self, parent_layout: QVBoxLayout):
        ribbon = QHBoxLayout()
        ribbon.setSpacing(8)

        # Title
        title_lbl = QLabel(f"<b>SA2 #{self.meeting_number}</b>")
        title_lbl.setStyleSheet("font-size: 16px; color: #0F172A;")
        title_lbl.setToolTip(f"SA WG2 Meeting #{self.meeting_number} ({self.location})")
        ribbon.addWidget(title_lbl)

        # Snapshot Combobox
        self.snap_combo = QComboBox()
        self.snap_combo.setMinimumWidth(230)
        self.snap_combo.setStyleSheet("padding: 4px 8px; font-weight: 600; border-radius: 6px; border: 1px solid #CBD5E1;")
        self.snap_combo.setToolTip("Select or switch between imported session plan snapshots and revisions.")
        self.snap_combo.currentIndexChanged.connect(self._on_snapshot_changed)
        ribbon.addWidget(self.snap_combo)

        # Timezone Pill
        self.tz_btn = QPushButton(f"🌍 {self.tz_name}")
        self.tz_btn.setStyleSheet("""
            QPushButton {
                background: #F1F5F9; color: #334155; font-size: 11px; font-weight: 600;
                padding: 4px 10px; border: 1px solid #CBD5E1; border-radius: 999px;
            }
            QPushButton:hover { background: #E2E8F0; }
        """)
        self.tz_btn.setToolTip("Meeting Venue Timezone. Click to change or verify against your Austrian corporate laptop clock.")
        self.tz_btn.clicked.connect(self._open_timezone_dialog)
        ribbon.addWidget(self.tz_btn)

        # Real-time Clock Strip
        self.clock_lbl = QLabel()
        self.clock_lbl.setStyleSheet("font-size: 11px; color: #64748B; font-weight: 500; margin-left: 8px;")
        self.clock_lbl.setToolTip("Current physical time at the meeting venue and on your Austrian corporate laptop.")
        ribbon.addWidget(self.clock_lbl)

        ribbon.addStretch()

        # Action Buttons
        self.import_btn = QPushButton("📥 Import ▾")
        self.import_btn.setStyleSheet("padding: 5px 12px; font-weight: 600; border-radius: 6px; background: #2563EB; color: #FFFFFF;")
        self.import_btn.setToolTip("Import a session plan document from URL, TDoc#, local file, or clipboard paste, or drag & drop files directly onto this window.")
        import_menu = QMenu(self)
        import_menu.addAction("🔗 Enter URL or TDoc#...", self._import_via_url).setToolTip("Fetch via 3GPP URL or type a TDoc number (e.g. S2-2610195).")
        import_menu.addAction("📂 Browse Local Document...", self._import_via_file).setToolTip("Select a .docx, .doc, .zip, or HTML table from your disk.")
        import_menu.addAction("📋 Paste Markdown/Text Table...", self._import_via_paste).setToolTip("Paste table text directly from an email, TOHRU chat, or notes.")
        self.import_btn.setMenu(import_menu)
        ribbon.addWidget(self.import_btn)

        self.diff_btn = QPushButton("⇄ Compare")
        self.diff_btn.setStyleSheet("padding: 5px 10px; font-weight: 500; border-radius: 6px; background: #F1F5F9; border: 1px solid #CBD5E1;")
        self.diff_btn.setToolTip("Compare the active session plan against an earlier revision to see added, removed, rescheduled, or volume-changed sessions.")
        self.diff_btn.clicked.connect(self._open_diff_dialog)
        ribbon.addWidget(self.diff_btn)

        self.export_btn = QPushButton("📤 Export ▾")
        self.export_btn.setStyleSheet("padding: 5px 10px; font-weight: 500; border-radius: 6px; background: #F1F5F9; border: 1px solid #CBD5E1;")
        self.export_btn.setToolTip("Export to responsive offline Mobile HTML (no PDF table separator issues!), strict UTC .ics Calendar, or Outlook Email summary.")
        export_menu = QMenu(self)
        export_menu.addAction("📱 Export Mobile HTML (Offline)...", self._export_mobile_html).setToolTip("Generate a responsive, single-file HTML view optimized for iPhone/Android Safari/Chrome.")
        export_menu.addAction("🌐 Open Mobile View in Default Browser", self._open_mobile_html_in_browser).setToolTip("Export and instantly open the mobile-optimized session plan in your browser.")
        export_menu.addAction("📅 Export iCalendar (.ics)...", self._export_ics).setToolTip("Export calendar events with strict UTC timestamps for Outlook, Apple, or Google Calendar.")
        export_menu.addAction("✉️ Copy Email Summary for Colleagues", self._copy_email_summary).setToolTip("Copy rich HTML email summary to clipboard ready to paste into Microsoft Outlook.")
        export_menu.addSeparator()
        export_menu.addAction("📂 Open Session Plans Folder", self._open_plans_folder).setToolTip("Open <meeting_dir>/Agenda/SessionPlans/ in Windows Explorer.")
        self.export_btn.setMenu(export_menu)
        ribbon.addWidget(self.export_btn)

        parent_layout.addLayout(ribbon)

    def _setup_navigation_ribbon(self, parent_layout: QVBoxLayout):
        sub_layout = QHBoxLayout()
        sub_layout.setSpacing(10)

        # Segmented view mode buttons
        self.mode_group = QButtonGroup(self)
        self.mode_group.setExclusive(True)

        self.btn_grid = QPushButton("🗓️ Overview Grid")
        self.btn_grid.setToolTip("Full 5-day spatial overview showing parallel Streams 1, 2, and 3 across all Quarters and Time Units.")
        self.btn_agenda = QPushButton("📋 Agenda Mode (Gap Check)")
        self.btn_agenda.setToolTip("Cross-references the official meeting agenda against scheduled sessions. Highlights which Agenda Items are scheduled vs unscheduled (missing from plan).")
        self.btn_drafting = QPushButton("🛠️ Drafting Radar (0)")
        self.btn_drafting.setToolTip("Chronological tracking of early morning (08:00), lunch (13:00), and evening (18:00) informal and formal drafting sessions.")
        self.btn_tracked = QPushButton("🎯 My Topics (0)")
        self.btn_tracked.setToolTip("Shows only sessions matching your followed Work Items and tracked Agenda Items, highlighting schedule conflicts.")

        for idx, btn in enumerate([self.btn_grid, self.btn_agenda, self.btn_drafting, self.btn_tracked]):
            btn.setCheckable(True)
            btn.setStyleSheet("""
                QPushButton {
                    padding: 5px 12px; font-size: 12px; font-weight: 600; border: 1px solid #CBD5E1;
                    background: #FFFFFF; color: #475569; border-radius: 6px;
                }
                QPushButton:checked {
                    background: #3B82F6; color: #FFFFFF; border-color: #3B82F6;
                }
            """)
            self.mode_group.addButton(btn, idx)
            sub_layout.addWidget(btn)

        self.btn_grid.setChecked(True)
        self.mode_group.buttonClicked[int].connect(self._on_view_mode_changed)

        # Quick Track Topics manager
        self.btn_manage_topics = QPushButton("⭐ Track Topics...")
        self.btn_manage_topics.setStyleSheet("""
            QPushButton {
                padding: 5px 10px; font-size: 12px; font-weight: 600; border: 1px solid #DDD6FE;
                background: #F5F3FF; color: #6D28D9; border-radius: 6px;
            }
            QPushButton:hover { background: #EDE9FE; }
        """)
        self.btn_manage_topics.setToolTip("Select which Agenda Items or Work Items to follow in 'My Topics' and highlight in purple.")
        self.btn_manage_topics.clicked.connect(self._open_manage_topics_dialog)
        sub_layout.addWidget(self.btn_manage_topics)

        sub_layout.addStretch()

        # Search Bar
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("🔍 Filter AI, topic, convener, room...")
        self.search_input.setToolTip("Filter sessions and agenda items by AI number (e.g. 20.6.18), topic keyword, room, or convener.")
        self.search_input.setClearButtonEnabled(True)
        self.search_input.setMinimumWidth(260)
        self.search_input.setStyleSheet("padding: 5px 10px; border: 1px solid #CBD5E1; border-radius: 6px;")
        self.search_input.textChanged.connect(self._apply_search_filter)
        sub_layout.addWidget(self.search_input)

        parent_layout.addLayout(sub_layout)

    # --- View 1: Overview Grid ---

    def _setup_overview_view(self):
        page = QWidget()
        page_layout = QVBoxLayout(page)
        page_layout.setContentsMargins(0, 0, 0, 0)
        page_layout.setSpacing(6)

        # Room & Convener Guide Strip
        self.room_strip = QFrame()
        self.room_strip.setStyleSheet("background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; padding: 6px 10px;")
        self.room_strip_layout = QHBoxLayout(self.room_strip)
        self.room_strip_layout.setContentsMargins(0, 0, 0, 0)
        self.room_strip_layout.setSpacing(12)

        self.r1_lbl = QLabel("🏢 <b>Stream 1:</b> Main Room (B2, Fl. 2) [300] · MCC / Chair · SA2_Main")
        self.r1_lbl.setStyleSheet("color: #4338CA; font-size: 11px;")
        self.r1_lbl.setToolTip("Stream 1 (Main Room): Chaired by SA2 Chair, notes taken by MCC. Hosts Plenary and Single Stream approvals.")
        self.room_strip_layout.addWidget(self.r1_lbl)

        self.r2_lbl = QLabel("🏢 <b>Stream 2:</b> Breakout 1 (B3, Fl. 2) [100] · Mike · SA2_Brk1")
        self.r2_lbl.setStyleSheet("color: #047857; font-size: 11px;")
        self.r2_lbl.setToolTip("Stream 2 (Breakout 1): Chaired by Convener 1 (Mike). TOHRU channel: SA2_Brk1.")
        self.room_strip_layout.addWidget(self.r2_lbl)

        self.r3_lbl = QLabel("🏢 <b>Stream 3:</b> Breakout 2 (D11/D12, Fl. 3) [120] · Wanqiang · SA2_Brk2")
        self.r3_lbl.setStyleSheet("color: #B45309; font-size: 11px;")
        self.r3_lbl.setToolTip("Stream 3 (Breakout 2): Chaired by Convener 2 (Wanqiang). TOHRU channel: SA2_Brk2.")
        self.room_strip_layout.addWidget(self.r3_lbl)

        self.room_strip_layout.addStretch()
        page_layout.addWidget(self.room_strip)

        # Overview Table
        self.grid_table = QTableWidget()
        self.grid_table.setColumnCount(5)
        self.grid_table.setHorizontalHeaderLabels(DAY_NAMES)
        self.grid_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.grid_table.verticalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.grid_table.verticalHeader().setDefaultSectionSize(95)
        self.grid_table.setAlternatingRowColors(True)
        self.grid_table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.grid_table.customContextMenuRequested.connect(self._on_grid_context_menu)
        self.grid_table.cellDoubleClicked.connect(self._on_grid_cell_double_clicked)
        page_layout.addWidget(self.grid_table)

        page_layout.addWidget(QLabel("💡 <i>Double-click any cell to add personal notes / priority. Right-click to track topic or view details.</i>"))

        self.view_stack.addWidget(page)

    # --- View 2: Agenda Mode ---

    def _setup_agenda_view(self):
        page = QWidget()
        page_layout = QVBoxLayout(page)
        page_layout.setContentsMargins(0, 0, 0, 0)

        self.agenda_table = QTableWidget()
        self.agenda_table.setColumnCount(6)
        self.agenda_table.setHorizontalHeaderLabels([
            "⭐", "AI #", "Topic Description", "Status", "Allocated Quarters & Rooms", "Actions"
        ])
        self.agenda_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.agenda_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.agenda_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self.agenda_table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeToContents)
        self.agenda_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.agenda_table.setAlternatingRowColors(True)
        page_layout.addWidget(self.agenda_table)

        self.view_stack.addWidget(page)

    # --- View 3: Drafting Radar ---

    def _setup_drafting_view(self):
        page = QWidget()
        page_layout = QVBoxLayout(page)
        page_layout.setContentsMargins(0, 0, 0, 0)
        page_layout.setSpacing(6)

        self.drafting_table = QTableWidget()
        self.drafting_table.setColumnCount(6)
        self.drafting_table.setHorizontalHeaderLabels([
            "Day", "Quarter", "Time (Venue / Vienna)", "Stream / Room", "Convener", "Drafting Topic & AIs"
        ])
        self.drafting_table.horizontalHeader().setSectionResizeMode(5, QHeaderView.Stretch)
        self.drafting_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.drafting_table.setAlternatingRowColors(True)
        self.drafting_table.cellDoubleClicked.connect(self._on_drafting_cell_double_clicked)
        page_layout.addWidget(self.drafting_table)
        page_layout.addWidget(QLabel("💡 <i>Double-click any drafting session to view and highlight in the 5-day Overview Grid.</i>"))

        self.view_stack.addWidget(page)

    # --- View 4: My Topics ---

    def _setup_tracked_view(self):
        page = QWidget()
        page_layout = QVBoxLayout(page)
        page_layout.setContentsMargins(0, 0, 0, 0)
        page_layout.setSpacing(6)

        self.tracked_table = QTableWidget()
        self.tracked_table.setColumnCount(6)
        self.tracked_table.setHorizontalHeaderLabels([
            "Day", "Quarter", "Time (Venue / Vienna)", "Stream / Room", "Matched Topic", "My Notes / Priority / Conflict"
        ])
        self.tracked_table.horizontalHeader().setSectionResizeMode(4, QHeaderView.Stretch)
        self.tracked_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.tracked_table.setAlternatingRowColors(True)
        self.tracked_table.cellDoubleClicked.connect(self._on_tracked_cell_double_clicked)
        self.tracked_table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.tracked_table.customContextMenuRequested.connect(self._on_tracked_context_menu)
        page_layout.addWidget(self.tracked_table)
        page_layout.addWidget(QLabel("💡 <i>Double-click any session to jump to the Overview Grid. Right-click to edit notes or copy details.</i>"))

        self.view_stack.addWidget(page)

    # --- Data Loading & Rendering ---

    def _load_data(self):
        """Loads snapshots, persistent tracked topics, and annotations from the SQLite database."""
        self.annotations = self.db.get_annotations(self.meeting_number)
        self.db_tracked = self.db.get_tracked_topics(self.meeting_number)
        self.tracked_tokens = set(self.external_tracked_tokens) | set(self.db_tracked.keys())

        snapshots = self.db.list_snapshots(self.meeting_number)

        self.snap_combo.blockSignals(True)
        self.snap_combo.clear()
        for snap_row in snapshots:
            label = f"{snap_row['snapshot_label']} ({snap_row['imported_at'][:16]})"
            self.snap_combo.addItem(label, snap_row["id"])
        self.snap_combo.blockSignals(False)

        if snapshots:
            active_snap = self.db.get_active_snapshot(self.meeting_number)
            self.active_snapshot = active_snap
            idx = self.snap_combo.findData(active_snap.id if active_snap else snapshots[0]["id"])
            if idx >= 0:
                self.snap_combo.setCurrentIndex(idx)
        else:
            self.active_snapshot = None

        self._render_current_views()

    def _render_current_views(self):
        """Refreshes all 4 views with the currently active snapshot."""
        if not self.active_snapshot:
            self._render_empty_views()
            return

        self._update_room_guide_strip()
        self._render_overview_grid()
        self._render_agenda_mode()
        self._render_drafting_radar()
        self._render_tracked_topics()

    def _update_room_guide_strip(self):
        """Updates room guide labels with active snapshot metadata."""
        if not self.active_snapshot:
            return
        streams = self.active_snapshot.streams
        if 1 in streams:
            s1 = streams[1]
            self.r1_lbl.setText(f"🏢 <b>Stream 1:</b> {s1.room_name} [{s1.capacity}] · {s1.convener or 'MCC / Chair'} · {s1.tohru_room}")
        if 2 in streams:
            s2 = streams[2]
            self.r2_lbl.setText(f"🏢 <b>Stream 2:</b> {s2.room_name} [{s2.capacity}] · {s2.convener or 'Convener 1'} · {s2.tohru_room}")
        if 3 in streams:
            s3 = streams[3]
            self.r3_lbl.setText(f"🏢 <b>Stream 3:</b> {s3.room_name} [{s3.capacity}] · {s3.convener or 'Convener 2'} · {s3.tohru_room}")

    def _render_empty_views(self):
        self.grid_table.setRowCount(0)
        self.agenda_table.setRowCount(0)
        self.drafting_table.setRowCount(0)
        self.tracked_table.setRowCount(0)

    def _is_slot_tracked(self, s: SessionSlot) -> bool:
        """Determines whether a slot matches any user-tracked topic, WI token, or agenda item."""
        if not self.tracked_tokens:
            return False
        raw_lower = s.raw_text.lower()
        if any(tk in raw_lower for tk in self.tracked_tokens if tk):
            return True
        for it in s.items:
            if it.ai_number.lower() in self.tracked_tokens or any(tk in it.topic_label.lower() for tk in self.tracked_tokens if tk):
                return True
        return False

    def _render_overview_grid(self):
        """Populates the 5-day Overview Grid with full Quarter Times and Sub-Slot timings."""
        self.grid_table.setRowCount(len(QUARTER_ORDER))

        # 1. Vertical Headers with Full Quarter Times
        v_headers = [QUARTER_DISPLAY_LABELS.get(q, q) for q in QUARTER_ORDER]
        self.grid_table.setVerticalHeaderLabels(v_headers)

        # 2. Horizontal Headers with Day and Date
        h_headers = []
        try:
            base_date = datetime.strptime(self.start_date_str, "%Y-%m-%d").date()
        except Exception:
            base_date = None

        venue_today = SessionPlanTimezone.get_venue_current_datetime(self.tz_offset_min).date()

        for idx, d_name in enumerate(DAY_NAMES):
            if base_date:
                cur_date = base_date + timedelta(days=idx)
                d_str = f"{d_name} ({cur_date.strftime('%d %b')})"
                if cur_date == venue_today:
                    d_str += " 🟢 Today"
                h_headers.append(d_str)
            else:
                h_headers.append(d_name)

        self.grid_table.setHorizontalHeaderLabels(h_headers)

        # Group slots by (day_index, quarter)
        slot_map: Dict[Tuple[int, str], List[SessionSlot]] = {}
        for s in self.active_snapshot.slots:
            slot_map.setdefault((s.day_index, s.quarter), []).append(s)

        for r_idx, q in enumerate(QUARTER_ORDER):
            for c_idx in range(5):
                slots_in_cell = slot_map.get((c_idx, q), [])
                if not slots_in_cell:
                    item = QTableWidgetItem("—")
                    item.setTextAlignment(Qt.AlignCenter)
                    item.setForeground(QColor("#94A3B8"))
                    self.grid_table.setItem(r_idx, c_idx, item)
                    continue

                # Build multi-stream cell text with exact sub-slot times
                cell_lines = []
                is_cell_drafting = False
                is_cell_tracked = False

                for s in sorted(slots_in_cell, key=lambda x: (x.sub_slot, x.stream_number)):
                    if s.sub_slot > 0:
                        time_tag = f"({s.start_time_local}–{s.end_time_local})"
                        cell_lines.append(f"<b>Str {s.stream_number} [TU {s.sub_slot} {time_tag}]:</b> {s.raw_text}")
                    else:
                        cell_lines.append(f"<b>Str {s.stream_number}:</b> {s.raw_text}")

                    if s.is_drafting:
                        is_cell_drafting = True
                    if self._is_slot_tracked(s):
                        is_cell_tracked = True

                cell_widget = QLabel("<br>".join(cell_lines))
                cell_widget.setWordWrap(True)
                cell_widget.setContentsMargins(6, 6, 6, 6)

                if is_cell_drafting:
                    cell_widget.setStyleSheet("background: #FEF2F2; color: #991B1B; border-radius: 4px; padding: 4px;")
                elif is_cell_tracked:
                    cell_widget.setStyleSheet("background: #F5F3FF; color: #5B21B6; border-radius: 4px; padding: 4px;")
                else:
                    cell_widget.setStyleSheet("padding: 4px;")

                self.grid_table.setCellWidget(r_idx, c_idx, cell_widget)

    def _render_agenda_mode(self):
        """Cross-references Agenda document items against allocated slots with 1-click ⭐ Track."""
        scheduled_ais: Dict[str, List[SessionSlot]] = {}
        for s in self.active_snapshot.slots:
            for it in s.items:
                if it.ai_number:
                    scheduled_ais.setdefault(it.ai_number, []).append(s)

        if self.agenda_map:
            all_ais = sorted(self.agenda_map.keys(), key=self._ai_sort_key)
        else:
            all_ais = sorted(scheduled_ais.keys(), key=self._ai_sort_key)

        num_sched = sum(1 for ai in all_ais if scheduled_ais.get(ai))
        num_unsched = len(all_ais) - num_sched
        self.btn_agenda.setText(f"📋 Agenda Mode ({num_sched} Sched · {num_unsched} Unscheduled)")

        self.agenda_table.setRowCount(len(all_ais))

        for row, ai_num in enumerate(all_ais):
            if ai_num in self.agenda_map:
                ai_item = self.agenda_map[ai_num]
                desc = ai_item.description
                acronym = getattr(ai_item, "acronym", None)
            else:
                matched_s = scheduled_ais.get(ai_num, [])
                desc = matched_s[0].raw_text if matched_s else f"Agenda Item {ai_num}"
                acronym = None

            is_tracked = (ai_num.lower() in self.tracked_tokens) or (acronym and acronym.lower() in self.tracked_tokens)

            # Col 0: Star toggle button
            star_btn = QPushButton("⭐" if is_tracked else "☆")
            star_btn.setFixedWidth(28)
            star_btn.setStyleSheet("border: none; font-size: 14px; background: transparent;")
            star_btn.setToolTip("Click to track/untrack this topic in 'My Topics'")
            star_btn.clicked.connect(lambda _, a=ai_num, l=desc: self._toggle_track_ai(a, l))
            self.agenda_table.setCellWidget(row, 0, star_btn)

            item_num = QTableWidgetItem(ai_num)
            item_desc = QTableWidgetItem(desc)

            matched_slots = scheduled_ais.get(ai_num, [])
            if matched_slots:
                status_item = QTableWidgetItem(f"🟢 Scheduled ({len(matched_slots)})")
                status_item.setForeground(QColor("#166534"))
                slot_descs = [f"{s.day_name[:3]} {s.quarter} (Str {s.stream_number})" for s in matched_slots]
                alloc_item = QTableWidgetItem(", ".join(slot_descs))
            else:
                status_item = QTableWidgetItem("🔴 Unscheduled")
                status_item.setForeground(QColor("#991B1B"))
                alloc_item = QTableWidgetItem("—")

            self.agenda_table.setItem(row, 1, item_num)
            self.agenda_table.setItem(row, 2, item_desc)
            self.agenda_table.setItem(row, 3, status_item)
            self.agenda_table.setItem(row, 4, alloc_item)

            action_btn = QPushButton("👁️ Show")
            action_btn.setStyleSheet("padding: 2px 6px; font-size: 11px;")
            action_btn.setToolTip("Filter to this Agenda Item across the Overview Grid")
            action_btn.clicked.connect(lambda _, a=ai_num: self._filter_by_ai(a))
            self.agenda_table.setCellWidget(row, 5, action_btn)

    def _render_drafting_radar(self):
        """Populates the Drafting Radar view."""
        drafting_slots = [s for s in self.active_snapshot.slots if s.is_drafting]
        self.btn_drafting.setText(f"🛠️ Drafting Radar ({len(drafting_slots)})")
        self.drafting_table.setRowCount(len(drafting_slots))

        streams = self.active_snapshot.streams
        for row, s in enumerate(drafting_slots):
            s_info = streams.get(s.stream_number)
            r_name = s_info.room_name if s_info else f"Stream {s.stream_number}"
            convener = s_info.convener if s_info else ""

            time_str = SessionPlanTimezone.format_dual_time_range(
                s.start_time_local, s.end_time_local, self.tz_offset_min
            )

            self.drafting_table.setItem(row, 0, QTableWidgetItem(s.day_name))
            self.drafting_table.setItem(row, 1, QTableWidgetItem(s.quarter))
            self.drafting_table.setItem(row, 2, QTableWidgetItem(time_str))
            self.drafting_table.setItem(row, 3, QTableWidgetItem(f"Str {s.stream_number} ({r_name})"))
            self.drafting_table.setItem(row, 4, QTableWidgetItem(convener))
            self.drafting_table.setItem(row, 5, QTableWidgetItem(s.raw_text))

    def _render_tracked_topics(self):
        """Populates the My Topics view with parallel room conflict detection."""
        tracked_slots = [s for s in self.active_snapshot.slots if self._is_slot_tracked(s)]

        # Conflict detection: same day and quarter, multiple streams
        time_slot_map: Dict[Tuple[int, str], List[SessionSlot]] = {}
        for s in tracked_slots:
            time_slot_map.setdefault((s.day_index, s.quarter), []).append(s)

        conflicts = {k: slots for k, slots in time_slot_map.items() if len(slots) > 1}

        if conflicts:
            self.btn_tracked.setText(f"🎯 My Topics ({len(tracked_slots)}) ⚠️ {len(conflicts)} Conflict(s)")
        else:
            self.btn_tracked.setText(f"🎯 My Topics ({len(tracked_slots)})")

        self.btn_manage_topics.setText(f"⭐ Track Topics ({len(self.tracked_tokens)})")
        self.tracked_table.setRowCount(len(tracked_slots))

        streams = self.active_snapshot.streams
        for row, s in enumerate(tracked_slots):
            s_info = streams.get(s.stream_number)
            r_name = s_info.room_name if s_info else f"Stream {s.stream_number}"

            time_str = SessionPlanTimezone.format_dual_time_range(
                s.start_time_local, s.end_time_local, self.tz_offset_min
            )

            ann = self.annotations.get(s.slot_key)
            note_str = ann.comment if ann else ""

            is_conflict = (s.day_index, s.quarter) in conflicts
            note_display = note_str
            if is_conflict:
                other_streams = [f"Str {o.stream_number}" for o in conflicts[(s.day_index, s.quarter)] if o != s]
                conflict_tag = f"⚠️ CONFLICT with {', '.join(other_streams)}"
                note_display = f"{conflict_tag} | {note_str}" if note_str else conflict_tag

            item_day = QTableWidgetItem(s.day_name)
            item_q = QTableWidgetItem(s.quarter)
            item_time = QTableWidgetItem(time_str)
            item_str = QTableWidgetItem(f"Str {s.stream_number} ({r_name})")
            item_raw = QTableWidgetItem(s.raw_text)
            item_note = QTableWidgetItem(note_display or "—")

            if is_conflict:
                for itm in [item_day, item_q, item_time, item_str, item_raw, item_note]:
                    itm.setBackground(QColor("#FEF3C7"))
                item_note.setForeground(QColor("#B45309"))

            self.tracked_table.setItem(row, 0, item_day)
            self.tracked_table.setItem(row, 1, item_q)
            self.tracked_table.setItem(row, 2, item_time)
            self.tracked_table.setItem(row, 3, item_str)
            self.tracked_table.setItem(row, 4, item_raw)
            self.tracked_table.setItem(row, 5, item_note)

    # --- Live Clock & Timezone ---

    def _update_live_clock(self):
        """Updates real-time Venue Time vs Austrian Laptop Time in the header."""
        venue_dt = SessionPlanTimezone.get_venue_current_datetime(self.tz_offset_min)
        laptop_offset = SessionPlanTimezone.get_laptop_tz_offset_minutes()

        venue_str = venue_dt.strftime("%a %H:%M")
        now_utc = datetime.now(timezone.utc)
        laptop_dt = now_utc + timedelta(minutes=laptop_offset)
        laptop_str = laptop_dt.strftime("%H:%M")

        if self.tz_offset_min == laptop_offset:
            self.clock_lbl.setText(f"🕒 <b>{venue_str}</b> (Venue & Austrian Time)")
        else:
            self.clock_lbl.setText(f"🏛️ Venue: <b>{venue_str}</b>  |  💻 Vienna: <b>{laptop_str}</b>")

    def _open_timezone_dialog(self):
        dlg = SessionPlanTimezoneDialog(
            meeting_number=self.meeting_number,
            location=self.location,
            initial_offset_min=self.tz_offset_min,
            initial_tz_name=self.tz_name,
            parent=self,
        )
        if dlg.exec_() == QDialog.Accepted:
            self.tz_offset_min = dlg.get_selected_offset()
            self.tz_name = dlg.get_selected_tz_name()
            self.tz_btn.setText(f"🌍 {self.tz_name}")
            self.db.set_meeting_meta(
                meeting_number=self.meeting_number,
                location=self.location,
                tz_offset_minutes=self.tz_offset_min,
                tz_name=self.tz_name,
                tz_confirmed=1,
            )
            self._render_current_views()

    # --- Tracking & Topics Management ---

    def _toggle_track_ai(self, ai_num: str, label: str = ""):
        """Toggles tracking for an Agenda Item."""
        ai_clean = ai_num.lower().strip()
        if ai_clean in self.tracked_tokens:
            self.tracked_tokens.discard(ai_clean)
            self.db.remove_tracked_topic(self.meeting_number, ai_clean)
        else:
            self.tracked_tokens.add(ai_clean)
            self.db.add_tracked_topic(self.meeting_number, ai_clean, label or f"AI {ai_num}")

        self._render_current_views()

    def _open_manage_topics_dialog(self):
        dlg = SessionPlanTrackedTopicsDialog(
            agenda_map=self.agenda_map,
            currently_tracked={k: k for k in self.tracked_tokens},
            followed_wi_tokens=self.external_tracked_tokens,
            parent=self,
        )
        if dlg.exec_() == QDialog.Accepted:
            selected = dlg.get_selected_topics()
            self.db.set_tracked_topics(self.meeting_number, selected)
            self.db_tracked = selected
            self.tracked_tokens = set(self.external_tracked_tokens) | set(selected.keys())
            self._render_current_views()
            QMessageBox.information(self, "Topics Updated", f"Now tracking {len(self.tracked_tokens)} topic(s). Overview Grid and My Topics have been refreshed!")

    # --- Navigation & Filters ---

    def _on_view_mode_changed(self, idx: int):
        self.view_stack.setCurrentIndex(idx)

    def _apply_search_filter(self, text: str):
        q = text.lower().strip()
        for r in range(self.agenda_table.rowCount()):
            ai_item = self.agenda_table.item(r, 1)
            desc_item = self.agenda_table.item(r, 2)
            txt = f"{ai_item.text()} {desc_item.text()}".lower() if ai_item and desc_item else ""
            self.agenda_table.setRowHidden(r, bool(q and q not in txt))

    def _filter_by_ai(self, ai_num: str):
        self.btn_grid.setChecked(True)
        self.view_stack.setCurrentIndex(0)
        self.search_input.setText(ai_num)

    def _on_grid_cell_double_clicked(self, row: int, col: int):
        q = QUARTER_ORDER[row]
        day = DAY_NAMES[col]
        slot_key = f"{day}|{q}|0|1"
        ann = self.annotations.get(slot_key)
        dlg = SessionPlanNoteDialog(f"{day} {q}", ann, self)
        if dlg.exec_() == QDialog.Accepted:
            badge, colleague, comment = dlg.get_data()
            self.db.set_annotation(
                meeting_number=self.meeting_number,
                target_type="slot",
                target_key=slot_key,
                status_badge=badge,
                assigned_colleague=colleague,
                comment=comment,
            )
            self.annotations = self.db.get_annotations(self.meeting_number)
            self._render_current_views()

    def _on_drafting_cell_double_clicked(self, row: int, col: int):
        d_item = self.drafting_table.item(row, 0)
        q_item = self.drafting_table.item(row, 1)
        if not d_item or not q_item:
            return
        d_name = d_item.text().strip()
        q_name = q_item.text().strip()
        if d_name in DAY_NAMES and q_name in QUARTER_ORDER:
            c = DAY_NAMES.index(d_name)
            r = QUARTER_ORDER.index(q_name)
            self.btn_grid.setChecked(True)
            self.view_stack.setCurrentIndex(0)
            self.grid_table.setCurrentCell(r, c)

    def _on_tracked_cell_double_clicked(self, row: int, col: int):
        d_item = self.tracked_table.item(row, 0)
        q_item = self.tracked_table.item(row, 1)
        if not d_item or not q_item:
            return
        d_name = d_item.text().strip()
        q_name = q_item.text().strip()
        if d_name in DAY_NAMES and q_name in QUARTER_ORDER:
            c = DAY_NAMES.index(d_name)
            r = QUARTER_ORDER.index(q_name)
            self.btn_grid.setChecked(True)
            self.view_stack.setCurrentIndex(0)
            self.grid_table.setCurrentCell(r, c)

    def _on_tracked_context_menu(self, pos: QPoint):
        r = self.tracked_table.rowAt(pos.y())
        if r < 0:
            return
        d_item = self.tracked_table.item(r, 0)
        q_item = self.tracked_table.item(r, 1)
        topic_item = self.tracked_table.item(r, 4)
        if not d_item or not q_item or not topic_item:
            return

        d_name = d_item.text().strip()
        q_name = q_item.text().strip()
        raw_text = topic_item.text().strip()

        menu = QMenu(self)
        menu.setToolTipsVisible(True)

        act_jump = menu.addAction(f"🏢 Show in Overview Grid ({d_name} {q_name})")
        act_jump.triggered.connect(lambda: self._on_tracked_cell_double_clicked(r, 0))

        if d_name in DAY_NAMES and q_name in QUARTER_ORDER:
            c = DAY_NAMES.index(d_name)
            grid_r = QUARTER_ORDER.index(q_name)
            act_note = menu.addAction(f"📝 Edit Note for {d_name} {q_name}...")
            act_note.triggered.connect(lambda: self._on_grid_cell_double_clicked(grid_r, c))

        menu.addSeparator()
        act_copy = menu.addAction("📋 Copy Slot Details")
        act_copy.triggered.connect(lambda: self._copy_tracked_row(r))

        menu.exec_(self.tracked_table.viewport().mapToGlobal(pos))

    def _copy_tracked_row(self, row: int):
        parts = [
            self.tracked_table.item(row, i).text()
            for i in range(self.tracked_table.columnCount())
            if self.tracked_table.item(row, i)
        ]
        from PyQt5.QtWidgets import QApplication
        QApplication.clipboard().setText(" | ".join(parts))

    def _on_grid_context_menu(self, pos: QPoint):
        item = self.grid_table.itemAt(pos)
        r = self.grid_table.rowAt(pos.y())
        c = self.grid_table.columnAt(pos.x())
        if r < 0 or c < 0:
            return

        q = QUARTER_ORDER[r]
        day = DAY_NAMES[c]

        # Find slots in this cell
        slots_here = [s for s in self.active_snapshot.slots if s.day_index == c and s.quarter == q] if self.active_snapshot else []

        menu = QMenu(self)
        menu.setToolTipsVisible(True)

        act_note = menu.addAction(f"📝 Edit Note / Priority for {day} {q}...")
        act_note.triggered.connect(lambda: self._on_grid_cell_double_clicked(r, c))

        if slots_here:
            menu.addSeparator()
            # Add track options for extracted topics
            for s in slots_here:
                for it in s.items:
                    label = it.topic_label or f"AI {it.ai_number}"
                    token = it.ai_number or it.topic_label
                    if not token:
                        continue
                    is_tr = token.lower() in self.tracked_tokens
                    act_tr = menu.addAction(f"{'⭐ Untrack' if is_tr else '⭐ Track'}: {label[:40]}")
                    act_tr.triggered.connect(lambda _, tok=token, lbl=label: self._toggle_track_ai(tok, lbl))

        menu.exec_(self.grid_table.viewport().mapToGlobal(pos))

    def _on_snapshot_changed(self, idx: int):
        snap_id = self.snap_combo.itemData(idx)
        if snap_id:
            self.db.set_active_snapshot(snap_id)
            self.active_snapshot = self.db.get_snapshot_by_id(snap_id)
            self._render_current_views()

    # --- Ingestion Entry Points ---

    def _import_via_url(self):
        dlg = SessionPlanImportDialog(self.meeting_number, self)
        dlg.import_requested.connect(self._handle_manual_import)
        dlg.exec_()

    def _import_via_file(self):
        f_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Session Plan Document",
            str(self.meeting_dir / "Agenda" / "SessionPlans"),
            "Session Plan Files (*.docx *.doc *.zip *.htm *.html *.txt);;All Files (*.*)",
        )
        if f_path:
            self._handle_manual_import("file", f_path)

    def _import_via_paste(self):
        dlg = SessionPlanImportDialog(self.meeting_number, self)
        dlg.import_requested.connect(self._handle_manual_import)
        dlg.exec_()

    def _handle_manual_import(self, mode: str, payload: str):
        try:
            if mode == "file":
                p = Path(payload)
                ext = p.suffix.lower()

                # Check if dropped file is an agenda.csv
                if p.name.lower() in ("agenda.csv", "agenda.txt"):
                    target_csv = self.meeting_dir / "Agenda" / "agenda.csv"
                    import shutil
                    shutil.copy2(p, target_csv)
                    self.agenda_map = AgendaManager.load_local_agenda(self.meeting_dir / "Agenda")
                    self._render_current_views()
                    QMessageBox.information(self, "Agenda Updated", f"Successfully loaded {len(self.agenda_map)} Agenda Items from agenda.csv!")
                    return

                staged = SessionPlanParser.stage_session_plan_file(p, self.meeting_dir)
                if ext == ".docx":
                    snap = SessionPlanParser.parse_docx(staged, meeting_number=self.meeting_number)
                else:
                    snap = SessionPlanParser.parse_text_or_markdown(
                        staged.read_text(encoding="utf-8", errors="replace"),
                        meeting_number=self.meeting_number,
                        cached_file_path=str(staged),
                    )
            elif mode == "paste":
                snap = SessionPlanParser.parse_text_or_markdown(payload, meeting_number=self.meeting_number)
            else:
                # URL or TDoc#
                from modules.meetings.core.tdoc_file_handler import TDocFileHandler
                from modules.meetings.core.url_router import URLRouter
                from core.network.network_state import NetworkState

                tdoc_candidate = payload.strip().upper()
                wg_name = self.mtg_info.get("wg_name", "").upper()
                candidate_urls = []
                if NetworkState.get_instance().is_local_active():
                    candidate_urls.append(URLRouter._get_local_server_base(wg_name) + "/Docs")
                docs_url = self.mtg_info.get("docs_folder_url", "")
                if docs_url:
                    if not docs_url.startswith("http"):
                        docs_url = "https://www.3gpp.org/ftp/" + docs_url.lstrip("/")
                    candidate_urls.append(docs_url)

                tdoc_dir = self.meeting_dir / tdoc_candidate
                extracted = []
                for u in candidate_urls:
                    extracted = TDocFileHandler.download_and_extract_tdoc(tdoc_candidate, u, tdoc_dir, timeout=10)
                    if extracted:
                        break

                if not extracted and tdoc_dir.exists():
                    extracted = list(tdoc_dir.glob("*.docx")) + list(tdoc_dir.glob("*.doc"))

                if not extracted:
                    QMessageBox.warning(self, "Download Failed", f"Could not acquire document for {payload}.")
                    return

                staged = SessionPlanParser.stage_session_plan_file(extracted[0], self.meeting_dir, tdoc_candidate)
                if staged.suffix.lower() == ".docx":
                    snap = SessionPlanParser.parse_docx(staged, meeting_number=self.meeting_number)
                else:
                    snap = SessionPlanParser.parse_text_or_markdown(
                        staged.read_text(encoding="utf-8", errors="replace"),
                        meeting_number=self.meeting_number,
                        cached_file_path=str(staged),
                    )

            snap_id = self.db.save_snapshot(snap)
            self._load_data()
            QMessageBox.information(self, "Import Successful", f"Successfully loaded session plan:\n{snap.snapshot_label}")

        except Exception as e:
            logger.error(f"[SessionPlanWindow] Import failed: {e}", exc_info=True)
            QMessageBox.critical(self, "Import Error", f"Failed to parse session plan:\n{e}")

    # --- Diffing & Exporting ---

    def _open_diff_dialog(self):
        snaps = self.db.list_snapshots(self.meeting_number)
        if len(snaps) < 2:
            QMessageBox.information(
                self,
                "Comparison Requires 2 Snapshots",
                "You need at least two session plan snapshots to run a comparison.\nImport an earlier or revised session plan first.",
            )
            return

        snap_a = self.db.get_snapshot_by_id(snaps[1]["id"])
        snap_b = self.db.get_snapshot_by_id(snaps[0]["id"])
        diff_res = SessionPlanDiffEngine.compare_snapshots(snap_a, snap_b)
        dlg = SessionPlanDiffDialog(diff_res, self)
        dlg.exec_()

    def _export_mobile_html(self) -> Optional[Path]:
        if not self.active_snapshot:
            return None
        out_dir = self.meeting_dir / "Export"
        out_path = out_dir / f"SA2_{self.meeting_number}_Session_Plan_Mobile.html"
        SessionPlanExporter.export_mobile_html(
            snapshot=self.active_snapshot,
            output_path=out_path,
            meeting_tz_offset_min=self.tz_offset_min,
            tz_name=self.tz_name,
            annotations=self.annotations,
            tracked_tokens=self.tracked_tokens,
        )
        QMessageBox.information(self, "Mobile HTML Exported", f"Saved responsive mobile view to:\n{out_path}\n\nYou can copy this file to your phone or open it directly in Safari/Chrome offline.")
        return out_path

    def _open_mobile_html_in_browser(self):
        if not self.active_snapshot:
            return
        out_path = self._export_mobile_html()
        if out_path and out_path.exists():
            webbrowser.open(str(out_path))

    def _export_ics(self):
        if not self.active_snapshot:
            return
        out_dir = self.meeting_dir / "Export"
        out_path = out_dir / f"SA2_{self.meeting_number}_Schedule.ics"
        SessionPlanExporter.export_ics(
            snapshot=self.active_snapshot,
            output_path=out_path,
            meeting_start_date_str=self.start_date_str,
            meeting_tz_offset_min=self.tz_offset_min,
            only_tracked=False,
            tracked_tokens=self.tracked_tokens,
            annotations=self.annotations,
        )
        QMessageBox.information(self, "Calendar Exported", f"Generated .ics calendar file:\n{out_path}\n\nStrictly anchored in UTC so your Outlook and smartphone trigger reminders at the exact venue time.")

    def _copy_email_summary(self):
        if not self.active_snapshot:
            return
        subj, body = SessionPlanExporter.generate_email_summary(
            snapshot=self.active_snapshot,
            meeting_tz_offset_min=self.tz_offset_min,
            tz_name=self.tz_name,
            tracked_tokens=self.tracked_tokens,
            annotations=self.annotations,
        )
        clipboard = QApplication.clipboard()
        clipboard.setText(body)
        QMessageBox.information(self, "Email Summary Copied", f"Subject: {subj}\n\nRich HTML email summary copied to your clipboard!\nYou can paste it directly into Microsoft Outlook or your mail client.")

    def _open_plans_folder(self):
        plans_dir = self.meeting_dir / "Agenda" / "SessionPlans"
        plans_dir.mkdir(parents=True, exist_ok=True)
        open_folder(plans_dir)

    @staticmethod
    def _ai_sort_key(ai_str: str) -> tuple:
        parts = []
        for p in re.split(r"(\d+)", ai_str):
            if p.isdigit():
                parts.append(int(p))
            elif p:
                parts.append(p)
        return tuple(parts)
