# --- File: src/modules/meetings/ui/session_plan_window.py ---
"""
Primary Desktop GUI for 3GPP SA2 Session Plan Management.
Provides Overview Grid with Time Labels, Agenda Mode (Gap Analysis with 1-Click Track),
Drafting Radar, My Topics, Timezone Synchronization, Diffing, and Exporting.
Includes direct drag & drop quick import on the GUI and rich tooltips.
"""
import html
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
    QCheckBox,
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

from dataclasses import dataclass
from modules.meetings.core.session_plan_parser import (
    PERIOD_ORDER,
    PERIOD_TIMES,
    SessionPlanParser,
)

DAY_NAMES = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
QUARTER_ORDER = PERIOD_ORDER
QUARTER_TIMES = PERIOD_TIMES


@dataclass
class GridRowDescriptor:
    period: str                  # "Q0", "Break 1", "Q1", ...
    stream_number: int           # 1, 2, 3 (or 0 for break)
    row_type: str                # "stream" or "break"
    display_label: str           # Vertical header string
    time_range: Tuple[str, str]  # ("09:00", "10:30")


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
        self.grid_rows: List[GridRowDescriptor] = []
        self.empty_break_periods: Set[str] = set()

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

    def showEvent(self, event):
        super().showEvent(event)
        if hasattr(self, "clock_timer") and self.clock_timer is not None and not self.clock_timer.isActive():
            self.clock_timer.start(1000)
            self._update_live_clock()

    def closeEvent(self, event):
        if hasattr(self, "clock_timer") and self.clock_timer is not None:
            self.clock_timer.stop()
        for dlg in self.findChildren(QDialog):
            try:
                dlg.close()
            except Exception:
                pass
        super().closeEvent(event)
        event.accept()

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

        # Room & Convener Guide & Stream Toggle Strip
        self.room_strip = QFrame()
        self.room_strip.setStyleSheet("background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; padding: 4px 10px;")
        self.room_strip_layout = QHBoxLayout(self.room_strip)
        self.room_strip_layout.setContentsMargins(4, 2, 4, 2)
        self.room_strip_layout.setSpacing(12)

        self.btn_toggle_s1 = QCheckBox("🏢 Stream 1: Main Room (B2) [300]")
        self.btn_toggle_s1.setChecked(True)
        self.btn_toggle_s1.setStyleSheet("QCheckBox { color: #4338CA; font-weight: 600; font-size: 11px; padding: 2px 4px; }")
        self.btn_toggle_s1.setToolTip("Toggle visibility of Stream 1 (Main Room / Plenary / Single Stream).")
        self.btn_toggle_s1.toggled.connect(self._update_row_visibility)
        self.room_strip_layout.addWidget(self.btn_toggle_s1)

        self.btn_toggle_s2 = QCheckBox("🏢 Stream 2: Breakout 1 (B3) [100]")
        self.btn_toggle_s2.setChecked(True)
        self.btn_toggle_s2.setStyleSheet("QCheckBox { color: #047857; font-weight: 600; font-size: 11px; padding: 2px 4px; }")
        self.btn_toggle_s2.setToolTip("Toggle visibility of Stream 2 (Breakout 1 / Mike).")
        self.btn_toggle_s2.toggled.connect(self._update_row_visibility)
        self.room_strip_layout.addWidget(self.btn_toggle_s2)

        self.btn_toggle_s3 = QCheckBox("🏢 Stream 3: Breakout 2 (D11/D12) [120]")
        self.btn_toggle_s3.setChecked(True)
        self.btn_toggle_s3.setStyleSheet("QCheckBox { color: #B45309; font-weight: 600; font-size: 11px; padding: 2px 4px; }")
        self.btn_toggle_s3.setToolTip("Toggle visibility of Stream 3 (Breakout 2 / Wanqiang).")
        self.btn_toggle_s3.toggled.connect(self._update_row_visibility)
        self.room_strip_layout.addWidget(self.btn_toggle_s3)

        sep = QFrame()
        sep.setFrameShape(QFrame.VLine)
        sep.setStyleSheet("color: #CBD5E1;")
        self.room_strip_layout.addWidget(sep)

        self.chk_breaks = QCheckBox("☕ Show Breaks")
        self.chk_breaks.setChecked(True)
        self.chk_breaks.setStyleSheet("QCheckBox { color: #475569; font-weight: 600; font-size: 11px; padding: 2px 4px; }")
        self.chk_breaks.setToolTip("Toggle visibility of Coffee Breaks, Lunch Breaks, and Short Breaks.")
        self.chk_breaks.toggled.connect(self._update_row_visibility)
        self.room_strip_layout.addWidget(self.chk_breaks)

        self.chk_hide_empty_breaks = QCheckBox("⚡ Hide Empty Breaks")
        self.chk_hide_empty_breaks.setChecked(True)
        self.chk_hide_empty_breaks.setStyleSheet("QCheckBox { color: #64748B; font-size: 11px; padding: 2px 4px; }")
        self.chk_hide_empty_breaks.setToolTip("When checked, automatically hides break rows that have no scheduled drafting sessions or notes.")
        self.chk_hide_empty_breaks.toggled.connect(self._update_row_visibility)
        self.room_strip_layout.addWidget(self.chk_hide_empty_breaks)

        self.room_strip_layout.addStretch()
        page_layout.addWidget(self.room_strip)

        # Overview Table (10 columns: 2 per day)
        self.grid_table = QTableWidget()
        self.grid_table.setColumnCount(10)
        self.grid_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.grid_table.verticalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.grid_table.verticalHeader().setDefaultSectionSize(60)
        self.grid_table.setAlternatingRowColors(False)
        self.grid_table.setStyleSheet("QTableWidget { gridline-color: #E2E8F0; selection-background-color: #E0E7FF; }")
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
        """Updates room guide labels and toggle checkboxes with active snapshot metadata."""
        if not self.active_snapshot:
            return
        streams = self.active_snapshot.streams
        if 1 in streams:
            s1 = streams[1]
            self.btn_toggle_s1.setText(f"🏢 Stream 1: {s1.room_name} [{s1.capacity}] · {s1.convener or 'MCC / Chair'} · {s1.tohru_room}")
        if 2 in streams:
            s2 = streams[2]
            self.btn_toggle_s2.setText(f"🏢 Stream 2: {s2.room_name} [{s2.capacity}] · {s2.convener or 'Convener 1'} · {s2.tohru_room}")
        if 3 in streams:
            s3 = streams[3]
            self.btn_toggle_s3.setText(f"🏢 Stream 3: {s3.room_name} [{s3.capacity}] · {s3.convener or 'Convener 2'} · {s3.tohru_room}")

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

    def _build_grid_rows(self) -> List[GridRowDescriptor]:
        """Constructs the 23-row spatial grid descriptors covering quarters and breaks."""
        rows = []
        streams = self.active_snapshot.streams if self.active_snapshot else {}
        for p in PERIOD_ORDER:
            times = PERIOD_TIMES[p]
            if any(b in p.lower() for b in ["break", "coffee", "lunch"]):
                clean_name = (
                    p.replace(" 1", "")
                    .replace(" 2", "")
                    .replace("Coffee", "☕ Coffee Break")
                    .replace("Break", "☕ Short Break")
                    .replace("Lunch", "🍱 Lunch Break")
                )
                lbl = f"{clean_name}\n{times[0]} – {times[1]}"
                rows.append(GridRowDescriptor(p, 0, "break", lbl, (times[0], times[1])))
            else:
                for s_num in [1, 2, 3]:
                    s_info = streams.get(s_num)
                    r_name = s_info.room_name if s_info else ("Main" if s_num == 1 else f"Brk {s_num-1}")
                    short_r = r_name.split("(")[0].strip()
                    lbl = f"{p} · Str {s_num} ({short_r})\n{times[0]} – {times[1]}"
                    rows.append(GridRowDescriptor(p, s_num, "stream", lbl, (times[0], times[1])))
        return rows

    def _update_row_visibility(self):
        """Hides/shows rows according to interactive stream and break toggle checkboxes."""
        show_s1 = self.btn_toggle_s1.isChecked()
        show_s2 = self.btn_toggle_s2.isChecked()
        show_s3 = self.btn_toggle_s3.isChecked()
        show_breaks = self.chk_breaks.isChecked()
        hide_empty = self.chk_hide_empty_breaks.isChecked()

        for row_idx, r_desc in enumerate(self.grid_rows):
            if r_desc.row_type == "stream":
                if r_desc.stream_number == 1:
                    self.grid_table.setRowHidden(row_idx, not show_s1)
                elif r_desc.stream_number == 2:
                    self.grid_table.setRowHidden(row_idx, not show_s2)
                elif r_desc.stream_number == 3:
                    self.grid_table.setRowHidden(row_idx, not show_s3)
            elif r_desc.row_type == "break":
                if not show_breaks:
                    self.grid_table.setRowHidden(row_idx, True)
                elif hide_empty and r_desc.period in self.empty_break_periods:
                    self.grid_table.setRowHidden(row_idx, True)
                else:
                    self.grid_table.setRowHidden(row_idx, False)

    def _build_rich_slot_tooltip(
        self, slot: SessionSlot, day_idx: int, period: str, stream_num: int, tu: int
    ) -> str:
        """Builds rich formatted HTML tooltip with timing, stream, AI details, notes, and conflicts."""
        day_name = DAY_NAMES[day_idx]
        date_str = ""
        try:
            base_date = datetime.strptime(self.start_date_str, "%Y-%m-%d").date()
            cur_date = base_date + timedelta(days=day_idx)
            date_str = cur_date.strftime("%d %b")
        except Exception:
            pass

        times = PERIOD_TIMES.get(period, ("09:00", "10:30"))
        if tu == 1:
            s_t, e_t = times[2], times[3]
            half_lbl = "1st Half (TU 1)"
        elif tu == 2:
            s_t, e_t = times[4], times[5]
            half_lbl = "2nd Half (TU 2)"
        else:
            s_t, e_t = times[0], times[1]
            half_lbl = "Full Quarter"

        time_dual = SessionPlanTimezone.format_dual_time_range(s_t, e_t, self.tz_offset_min)

        s_info = self.active_snapshot.streams.get(stream_num) if self.active_snapshot else None
        r_name = s_info.room_name if s_info else f"Stream {stream_num}"
        cap = f" [{s_info.capacity}]" if s_info and s_info.capacity else ""
        convener = s_info.convener if s_info else ("SA2 Chair / MCC" if stream_num == 1 else "")
        tohru = s_info.tohru_room if s_info else ""

        ai_bullets = []
        for it in slot.items:
            if it.ai_number:
                ai_info = self.agenda_map.get(it.ai_number)
                desc = ai_info.description if ai_info else it.topic_label
                ai_bullets.append(f"• <b>AI {it.ai_number}:</b> {html.escape(desc)}{' [' + it.tdoc_count_str + ']' if it.tdoc_count_str else ''}")

        ann = self.annotations.get(slot.slot_key)

        tt = f"""<div style='font-family:Segoe UI, sans-serif; font-size:12px; min-width:240px; padding:4px;'>
            <div style='font-weight:bold; font-size:13px; color:#1E293B; margin-bottom:4px;'>
                📅 {day_name} ({date_str}) · {period} · {half_lbl}
            </div>
            <div style='color:#475569; margin-bottom:6px;'>
                🕒 <b>Time:</b> {time_dual}
            </div>
            <div style='color:#475569; margin-bottom:6px;'>
                🏢 <b>Stream {stream_num}:</b> {r_name}{cap} &nbsp;|&nbsp; 👤 <b>Chair:</b> {convener} &nbsp;|&nbsp; 📡 <b>TOHRU:</b> {tohru}
            </div>
            <hr style='border:none; border-top:1px solid #CBD5E1; margin:6px 0;'>
            <div style='font-weight:bold; color:#0F172A; margin-bottom:4px;'>
                📋 {html.escape(slot.raw_text)}
            </div>"""

        if ai_bullets:
            tt += f"<div style='color:#64748B; font-size:11px; margin-left:4px; margin-top:4px;'>{'<br>'.join(ai_bullets)}</div>"

        if self._is_slot_tracked(slot):
            tt += "<div style='margin-top:6px; padding:3px 6px; background:#F5F3FF; color:#5B21B6; border-radius:4px; font-weight:600;'>🎯 Tracked in My Topics</div>"

        if ann and (ann.comment or ann.assigned_colleague):
            c_str = f" ({ann.assigned_colleague})" if ann.assigned_colleague else ""
            tt += f"<div style='margin-top:4px; padding:3px 6px; background:#F1F5F9; color:#334155; border-radius:4px;'>📝 <b>Note:</b> {html.escape(ann.comment)}{c_str}</div>"

        tt += "</div>"
        return tt

    def _create_slot_cell_widget(
        self, slot: SessionSlot, day_idx: int, period: str, stream_num: int, is_half: bool = False, tu: int = 0
    ) -> QWidget:
        """Creates clean visual card for a session slot with zero inline timestamps."""
        w = QWidget()
        lay = QVBoxLayout(w)
        lay.setContentsMargins(4, 3, 4, 3)
        lay.setSpacing(2)

        lbl = QLabel(slot.raw_text)
        lbl.setWordWrap(True)
        lbl.setAlignment(Qt.AlignTop | Qt.AlignLeft)

        is_tracked = self._is_slot_tracked(slot)

        if slot.is_drafting:
            lbl.setStyleSheet("background: #FEF2F2; color: #991B1B; border: 1px solid #FECACA; border-radius: 4px; padding: 4px; font-weight: 500;")
        elif is_tracked:
            lbl.setStyleSheet("background: #F5F3FF; color: #5B21B6; border: 1px solid #DDD6FE; border-radius: 4px; padding: 4px; font-weight: 500;")
        else:
            border_c = "#4338CA" if stream_num == 1 else ("#047857" if stream_num == 2 else "#B45309")
            lbl.setStyleSheet(f"background: #FFFFFF; color: #1E293B; border: 1px solid #E2E8F0; border-left: 3px solid {border_c}; border-radius: 4px; padding: 4px;")

        lay.addWidget(lbl)

        tooltip_html = self._build_rich_slot_tooltip(slot, day_idx, period, stream_num, tu)
        w.setToolTip(tooltip_html)
        lbl.setToolTip(tooltip_html)
        return w

    def _create_break_cell_widget(self, slot: SessionSlot, day_idx: int, period: str) -> QWidget:
        """Creates cell widget for break drafting sessions or boundary schedules."""
        w = QWidget()
        lay = QVBoxLayout(w)
        lay.setContentsMargins(4, 2, 4, 2)

        lbl = QLabel(slot.raw_text)
        lbl.setWordWrap(True)
        lbl.setAlignment(Qt.AlignTop | Qt.AlignLeft)

        if slot.is_drafting:
            lbl.setStyleSheet("background: #FEF2F2; color: #991B1B; border: 1px solid #FECACA; border-radius: 4px; padding: 4px; font-weight: 600;")
        elif slot.is_plenary:
            lbl.setStyleSheet("background: #EFF6FF; color: #1E40AF; border: 1px solid #BFDBFE; border-radius: 4px; padding: 4px; font-weight: 600;")
        else:
            lbl.setStyleSheet("background: #F8FAFC; color: #334155; border: 1px solid #E2E8F0; border-radius: 4px; padding: 4px;")

        lay.addWidget(lbl)

        day_name = DAY_NAMES[day_idx]
        times = PERIOD_TIMES.get(period, ("12:30", "14:00"))
        time_dual = SessionPlanTimezone.format_dual_time_range(times[0], times[1], self.tz_offset_min)
        tt = f"""<div style='font-family:Segoe UI, sans-serif; font-size:12px; min-width:220px; padding:4px;'>
            <div style='font-weight:bold; font-size:13px; color:#1E293B; margin-bottom:4px;'>
                ☕ {day_name} · {period}
            </div>
            <div style='color:#475569; margin-bottom:6px;'>
                🕒 <b>Time:</b> {time_dual}
            </div>
            <hr style='border:none; border-top:1px solid #CBD5E1; margin:6px 0;'>
            <div style='font-weight:bold; color:#0F172A;'>
                {html.escape(slot.raw_text)}
            </div>
        </div>"""
        w.setToolTip(tt)
        lbl.setToolTip(tt)
        return w

    def _render_overview_grid(self):
        """Populates the 10-column spatial Overview Grid with separated stream rows, graphical half-quarters, and breaks."""
        self.grid_table.clearSpans()
        self.grid_rows = self._build_grid_rows()
        self.grid_table.setRowCount(len(self.grid_rows))

        # 1. Vertical Headers with Period and Stream/Break
        v_headers = [r.display_label for r in self.grid_rows]
        self.grid_table.setVerticalHeaderLabels(v_headers)

        # 2. Horizontal Headers with Day and Half-Quarter Columns (10 columns)
        h_headers = []
        try:
            base_date = datetime.strptime(self.start_date_str, "%Y-%m-%d").date()
        except Exception:
            base_date = None

        venue_today = SessionPlanTimezone.get_venue_current_datetime(self.tz_offset_min).date()

        for idx, d_name in enumerate(DAY_NAMES):
            if base_date:
                cur_date = base_date + timedelta(days=idx)
                d_str = f"{d_name[:3]} ({cur_date.strftime('%d %b')})"
                if cur_date == venue_today:
                    d_str += " 🟢 Today"
            else:
                d_str = d_name

            h_headers.append(f"{d_str} · 1st Half")
            h_headers.append(f"{d_str} · 2nd Half")

        self.grid_table.setHorizontalHeaderLabels(h_headers)

        # Determine which break periods have content
        active_break_periods = set()
        for s in self.active_snapshot.slots:
            if any(b in s.quarter.lower() for b in ["break", "coffee", "lunch"]):
                if s.raw_text.strip() and s.raw_text.strip() != "-":
                    active_break_periods.add(s.quarter)
        self.empty_break_periods = {r.period for r in self.grid_rows if r.row_type == "break" and r.period not in active_break_periods}

        # Group slots
        slot_map: Dict[Tuple[int, str, int], List[SessionSlot]] = {}
        for s in self.active_snapshot.slots:
            slot_map.setdefault((s.day_index, s.quarter, s.stream_number), []).append(s)

        break_slot_map: Dict[Tuple[int, str], List[SessionSlot]] = {}
        for s in self.active_snapshot.slots:
            if any(b in s.quarter.lower() for b in ["break", "coffee", "lunch"]):
                break_slot_map.setdefault((s.day_index, s.quarter), []).append(s)

        for row_idx, r_desc in enumerate(self.grid_rows):
            if r_desc.row_type == "stream":
                for day_idx in range(5):
                    tu1_col = day_idx * 2
                    tu2_col = day_idx * 2 + 1
                    matching = slot_map.get((day_idx, r_desc.period, r_desc.stream_number), [])

                    full_slot = next((s for s in matching if s.sub_slot == 0), None)
                    tu1_slot = next((s for s in matching if s.sub_slot == 1), None)
                    tu2_slot = next((s for s in matching if s.sub_slot == 2), None)

                    if full_slot:
                        self.grid_table.setSpan(row_idx, tu1_col, 1, 2)
                        w = self._create_slot_cell_widget(full_slot, day_idx, r_desc.period, r_desc.stream_number, is_half=False, tu=0)
                        self.grid_table.setCellWidget(row_idx, tu1_col, w)
                    elif tu1_slot or tu2_slot:
                        if tu1_slot:
                            w1 = self._create_slot_cell_widget(tu1_slot, day_idx, r_desc.period, r_desc.stream_number, is_half=True, tu=1)
                            self.grid_table.setCellWidget(row_idx, tu1_col, w1)
                        else:
                            item_e = QTableWidgetItem("—")
                            item_e.setTextAlignment(Qt.AlignCenter)
                            item_e.setForeground(QColor("#CBD5E1"))
                            self.grid_table.setItem(row_idx, tu1_col, item_e)

                        if tu2_slot:
                            w2 = self._create_slot_cell_widget(tu2_slot, day_idx, r_desc.period, r_desc.stream_number, is_half=True, tu=2)
                            self.grid_table.setCellWidget(row_idx, tu2_col, w2)
                        else:
                            item_e = QTableWidgetItem("—")
                            item_e.setTextAlignment(Qt.AlignCenter)
                            item_e.setForeground(QColor("#CBD5E1"))
                            self.grid_table.setItem(row_idx, tu2_col, item_e)
                    else:
                        self.grid_table.setSpan(row_idx, tu1_col, 1, 2)
                        item_e = QTableWidgetItem("—")
                        item_e.setTextAlignment(Qt.AlignCenter)
                        item_e.setForeground(QColor("#E2E8F0"))
                        self.grid_table.setItem(row_idx, tu1_col, item_e)

            elif r_desc.row_type == "break":
                for day_idx in range(5):
                    tu1_col = day_idx * 2
                    self.grid_table.setSpan(row_idx, tu1_col, 1, 2)
                    matching_b = break_slot_map.get((day_idx, r_desc.period), [])
                    valid_b = [s for s in matching_b if s.raw_text.strip() and s.raw_text.strip() != "-"]
                    if valid_b:
                        w = self._create_break_cell_widget(valid_b[0], day_idx, r_desc.period)
                        self.grid_table.setCellWidget(row_idx, tu1_col, w)
                    else:
                        item_b = QTableWidgetItem("—")
                        item_b.setTextAlignment(Qt.AlignCenter)
                        item_b.setForeground(QColor("#CBD5E1"))
                        item_b.setBackground(QColor("#F8FAFC"))
                        times = PERIOD_TIMES[r_desc.period]
                        item_b.setToolTip(f"{r_desc.period}: {times[0]} – {times[1]} (Break)")
                        self.grid_table.setItem(row_idx, tu1_col, item_b)

        self._update_row_visibility()

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
        if row < 0 or row >= len(self.grid_rows):
            return
        r_desc = self.grid_rows[row]
        day_idx = col // 2
        if day_idx < 0 or day_idx >= len(DAY_NAMES):
            return
        day = DAY_NAMES[day_idx]
        span = self.grid_table.columnSpan(row, col)
        tu = 1 if col % 2 == 0 else 2
        sub_slot = 0 if span == 2 else tu

        slot_key = f"{day}|{r_desc.period}|{sub_slot}|{r_desc.stream_number}"
        ann = self.annotations.get(slot_key)
        label = f"{day} {r_desc.period} (Str {r_desc.stream_number})" if r_desc.stream_number > 0 else f"{day} {r_desc.period}"
        dlg = SessionPlanNoteDialog(label, ann, self)
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
        s_item = self.drafting_table.item(row, 3)
        if not d_item or not q_item:
            return
        d_name = d_item.text().strip()
        q_name = q_item.text().strip()
        m_s = re.search(r"Str\s*(\d+)", s_item.text() if s_item else "")
        stream_num = int(m_s.group(1)) if m_s else 1

        if d_name in DAY_NAMES:
            day_idx = DAY_NAMES.index(d_name)
            target_row = None
            for r_i, r_d in enumerate(self.grid_rows):
                if r_d.period == q_name and (r_d.stream_number == stream_num or r_d.row_type == "break"):
                    target_row = r_i
                    break

            if target_row is not None:
                self.btn_grid.setChecked(True)
                self.view_stack.setCurrentIndex(0)
                self.grid_table.setCurrentCell(target_row, day_idx * 2)

    def _on_tracked_cell_double_clicked(self, row: int, col: int):
        d_item = self.tracked_table.item(row, 0)
        q_item = self.tracked_table.item(row, 1)
        s_item = self.tracked_table.item(row, 3)
        if not d_item or not q_item:
            return
        d_name = d_item.text().strip()
        q_name = q_item.text().strip()
        m_s = re.search(r"Str\s*(\d+)", s_item.text() if s_item else "")
        stream_num = int(m_s.group(1)) if m_s else 1

        if d_name in DAY_NAMES:
            day_idx = DAY_NAMES.index(d_name)
            target_row = None
            for r_i, r_d in enumerate(self.grid_rows):
                if r_d.period == q_name and (r_d.stream_number == stream_num or r_d.row_type == "break"):
                    target_row = r_i
                    break

            if target_row is not None:
                self.btn_grid.setChecked(True)
                self.view_stack.setCurrentIndex(0)
                self.grid_table.setCurrentCell(target_row, day_idx * 2)

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
        r = self.grid_table.rowAt(pos.y())
        c = self.grid_table.columnAt(pos.x())
        if r < 0 or c < 0 or r >= len(self.grid_rows):
            return

        r_desc = self.grid_rows[r]
        day_idx = c // 2
        if day_idx < 0 or day_idx >= len(DAY_NAMES):
            return
        day = DAY_NAMES[day_idx]
        span = self.grid_table.columnSpan(r, c)
        tu = 1 if c % 2 == 0 else 2
        sub_slot = 0 if span == 2 else tu

        # Find slots in this cell
        if r_desc.row_type == "stream":
            slots_here = [
                s for s in (self.active_snapshot.slots if self.active_snapshot else [])
                if s.day_index == day_idx and s.quarter == r_desc.period and s.stream_number == r_desc.stream_number and (s.sub_slot == sub_slot or s.sub_slot == 0)
            ]
        else:
            slots_here = [
                s for s in (self.active_snapshot.slots if self.active_snapshot else [])
                if s.day_index == day_idx and s.quarter == r_desc.period
            ]

        menu = QMenu(self)
        menu.setToolTipsVisible(True)

        lbl_target = f"{day} {r_desc.period} (Str {r_desc.stream_number})" if r_desc.stream_number > 0 else f"{day} {r_desc.period}"
        act_note = menu.addAction(f"📝 Edit Note / Priority for {lbl_target}...")
        act_note.triggered.connect(lambda: self._on_grid_cell_double_clicked(r, c))

        if slots_here:
            menu.addSeparator()
            for s in slots_here:
                for it in s.items:
                    label = it.topic_label or f"AI {it.ai_number}"
                    token = it.ai_number or it.topic_label
                    if not token:
                        continue
                    is_tr = token.lower() in self.tracked_tokens
                    act_tr = menu.addAction(f"{'⭐ Untrack' if is_tr else '⭐ Track'}: {label[:40]}")
                    act_tr.triggered.connect(lambda _, tok=token, lbl=label: self._toggle_track_ai(tok, lbl))

            menu.addSeparator()
            act_copy = menu.addAction("📋 Copy Slot Details")
            act_copy.triggered.connect(lambda: self._copy_slot_text(slots_here[0].raw_text))

        menu.exec_(self.grid_table.viewport().mapToGlobal(pos))

    def _copy_slot_text(self, text: str):
        from PyQt5.QtWidgets import QApplication
        QApplication.clipboard().setText(text)

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
