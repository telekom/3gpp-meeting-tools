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
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

from core.utils.folder_opener import open_folder
from modules.meetings.core.agenda_manager import AgendaManager
from modules.meetings.core.session_plan_db import SessionPlanDatabase
from modules.meetings.core.session_plan_diff import SessionPlanDiffEngine
from modules.meetings.core.session_plan_exporter import SessionPlanExporter
from modules.meetings.core.session_plan_models import (
    DELEGATE_PALETTE,
    AgendaSlotItem,
    DelegationMember,
    DelegationPlan,
    MeetingStreamInfo,
    SessionAppointment,
    SessionPlanSnapshot,
    SessionSlot,
    TopicOwnership,
    UserSessionAnnotation,
    compute_unique_initials,
    resolve_slot_effective_delegates,
)
from modules.meetings.core.session_plan_parser import SessionPlanParser
from modules.meetings.core.session_plan_timezone import SessionPlanTimezone
from modules.meetings.ui.delegate_owner_picker import DelegateOwnerPicker
from modules.meetings.ui.session_plan_dialogs import (
    DelegationImportDialog,
    DelegationMemberDialog,
    MultiDelegateSelectDialog,
    SessionDispatchDialog,
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

        # Delegation state
        self.delegation_plan: DelegationPlan = self.db.get_delegation_plan(self.meeting_number)
        # Multi-select Delegate Focus: empty set == whole team (no dimming); otherwise roster names.
        self.grid_focus_delegates: Set[str] = set()
        self.grid_highlight_delegate: str = "ALL"  # Backward-compat mirror: "ALL", "SELF", member name, or "MULTI"
        self.dynamic_delegate_chips: List[QCheckBox] = []
        # Per-render caches (collisions are expensive: O(slots x owners) and were recomputed per cell)
        self._grid_collisions_cache: Optional[Dict[Tuple[str, int, str], List[SessionSlot]]] = None
        self._initials_cache: Tuple[Tuple[str, ...], Dict[str, str]] = ((), {})

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
        self.drop_overlay = DropOverlayWidget(
            self,
            title="Drop Session Plan file(s) to import",
            supported_exts=[".docx", ".doc", ".zip", ".htm", ".csv"],
        )

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
        self._setup_team_view()

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
        export_menu.addSeparator()
        export_menu.addAction("📅 Export Full Meeting .ics Calendar...", self._export_ics).setToolTip("Export calendar events with strict UTC timestamps for Outlook, Apple, or Google Calendar.")
        export_menu.addAction("📅 Export Delegate .ics (Personal / Colleague)...", self._export_delegate_ics).setToolTip("Export personal calendar events scoped to a specific delegate with assigned topics.")
        export_menu.addSeparator()
        export_menu.addAction("✉️ Copy Personal/Tracked Email Summary", self._copy_email_summary).setToolTip("Copy rich HTML email summary of personal tracked topics to clipboard.")
        export_menu.addAction("📧 Copy Complete Delegation Email Summary...", self._on_export_delegation_email).setToolTip("Copy rich HTML delegation plan and machine-readable interchange comment for team email sharing.")
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
        self.btn_tracked = QPushButton("🎯 Delegate Agenda (0)")
        self.btn_tracked.setToolTip("View personal schedules by delegate (or your own agenda) and the entire delegation timeline.")
        self.btn_team = QPushButton("👥 Team Allocations (Clean)")
        self.btn_team.setToolTip("Distribute Agenda Items across team members, detect double-booking clashes, and appoint available delegates on a session-by-session basis.")

        for idx, btn in enumerate([self.btn_grid, self.btn_agenda, self.btn_drafting, self.btn_tracked, self.btn_team]):
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

        # Delegate Focus / Heatmap Highlight Strip
        self.delegate_strip = QFrame()
        self.delegate_strip.setStyleSheet("background: #FFFFFF; border: 1px solid #E2E8F0; border-radius: 8px; padding: 4px 10px;")
        self.delegate_strip_layout = QHBoxLayout(self.delegate_strip)
        self.delegate_strip_layout.setContentsMargins(4, 2, 4, 2)
        self.delegate_strip_layout.setSpacing(6)

        lbl_f = QLabel("<b>Delegate Focus:</b>")
        lbl_f.setStyleSheet("font-size: 11px; color: #475569;")
        lbl_f.setToolTip("Tick one or more delegates to highlight their sessions; all other sessions are dimmed.")
        self.delegate_strip_layout.addWidget(lbl_f)

        self.btn_chip_all = QPushButton("🌈 All Team")
        self.btn_chip_all.setToolTip("Clear the focus selection and show every session at full contrast.")
        self.btn_chip_all.setCursor(Qt.PointingHandCursor)
        self.btn_chip_all.clicked.connect(lambda: self._on_delegate_highlight_clicked("ALL"))
        self.delegate_strip_layout.addWidget(self.btn_chip_all)

        # Dynamic "Only Me" + per-delegate focus checkboxes live in their own sub-layout
        self.delegate_chips_layout = QHBoxLayout()
        self.delegate_chips_layout.setContentsMargins(0, 0, 0, 0)
        self.delegate_chips_layout.setSpacing(6)
        self.delegate_strip_layout.addLayout(self.delegate_chips_layout)

        self.btn_chip_me: Optional[QPushButton] = None
        self.dynamic_delegate_chips: List[QCheckBox] = []
        self.delegate_strip_layout.addStretch()

        self.lbl_badge_legend = QLabel(
            "<span style='color:#64748B;'>Badges: <b>●</b> owner · <b>◌</b> dashed = ad-hoc cover · "
            "<b style='color:#DC2626;'>◯</b> red ring = clash</span>"
        )
        self.lbl_badge_legend.setStyleSheet("font-size: 10px;")
        self.delegate_strip_layout.addWidget(self.lbl_badge_legend)
        page_layout.addWidget(self.delegate_strip)
        self._style_focus_buttons()

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

    # --- View 4: Delegate Agenda (formerly My Topics) ---

    def _setup_tracked_view(self):
        page = QWidget()
        page_layout = QVBoxLayout(page)
        page_layout.setContentsMargins(0, 0, 0, 0)
        page_layout.setSpacing(6)

        # Delegate Switcher Bar
        top_bar = QHBoxLayout()
        top_bar.setSpacing(8)
        top_bar.addWidget(QLabel("<b>Delegate Agenda:</b>"))
        self.delegate_agenda_combo = QComboBox()
        self.delegate_agenda_combo.setMinimumWidth(230)
        self.delegate_agenda_combo.setStyleSheet("padding: 3px 8px; font-weight: 600; border: 1px solid #CBD5E1; border-radius: 6px;")
        self.delegate_agenda_combo.currentIndexChanged.connect(self._on_delegate_agenda_changed)
        top_bar.addWidget(self.delegate_agenda_combo)

        top_bar.addStretch()
        page_layout.addLayout(top_bar)

        self.tracked_table = QTableWidget()
        self.tracked_table.setColumnCount(7)
        self.tracked_table.setHorizontalHeaderLabels([
            "Day", "Quarter", "Time (Venue / Vienna)", "Stream / Room", "Delegate", "Matched Topic", "My Notes / Priority / Conflict"
        ])
        self.tracked_table.horizontalHeader().setSectionResizeMode(5, QHeaderView.Stretch)
        self.tracked_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.tracked_table.setAlternatingRowColors(True)
        self.tracked_table.cellDoubleClicked.connect(self._on_tracked_cell_double_clicked)
        self.tracked_table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.tracked_table.customContextMenuRequested.connect(self._on_tracked_context_menu)
        page_layout.addWidget(self.tracked_table)
        page_layout.addWidget(QLabel("💡 <i>Double-click any session to jump to the Overview Grid. Right-click to edit notes or copy details.</i>"))

        self.view_stack.addWidget(page)

    # --- View 5: Team Allocations (HoD Dispatcher & Topic Ownership) ---

    def _setup_team_view(self):
        page = QWidget()
        page_layout = QVBoxLayout(page)
        page_layout.setContentsMargins(0, 0, 0, 0)
        page_layout.setSpacing(8)

        # 1. Roster Card
        self.team_roster_card = QFrame()
        self.team_roster_card.setStyleSheet("background: #FFFFFF; border: 1px solid #E2E8F0; border-radius: 8px; padding: 6px 12px;")
        rc_lay = QHBoxLayout(self.team_roster_card)
        rc_lay.setContentsMargins(4, 4, 4, 4)
        rc_lay.setSpacing(10)

        rc_lay.addWidget(QLabel("<b>👥 Delegation Roster:</b>"))

        self.team_chips_container = QWidget()
        self.team_member_chips_layout = QHBoxLayout(self.team_chips_container)
        self.team_member_chips_layout.setContentsMargins(0, 0, 0, 0)
        self.team_member_chips_layout.setSpacing(6)
        rc_lay.addWidget(self.team_chips_container)

        self.btn_add_member = QPushButton("➕ Add Member...")
        self.btn_add_member.setStyleSheet("padding: 3px 10px; font-size: 11px; font-weight: 600; border: 1px dashed #3B82F6; color: #2563EB; border-radius: 6px; background: #EFF6FF;")
        self.btn_add_member.clicked.connect(self._on_add_delegation_member)
        rc_lay.addWidget(self.btn_add_member)

        rc_lay.addStretch()

        # Action Buttons
        self.btn_export_team_email = QPushButton("📧 Export Email Summary...")
        self.btn_export_team_email.setStyleSheet("padding: 5px 12px; font-weight: 600; font-size: 11px; border-radius: 6px; background: #2563EB; color: #FFFFFF;")
        self.btn_export_team_email.setToolTip("Generate an executive delegation summary email with color codes, topic distribution, and personal agendas.")
        self.btn_export_team_email.clicked.connect(self._on_export_delegation_email)
        rc_lay.addWidget(self.btn_export_team_email)

        self.btn_copy_team_json = QPushButton("📋 Copy Plan JSON")
        self.btn_copy_team_json.setStyleSheet("padding: 5px 10px; font-size: 11px; border-radius: 6px; border: 1px solid #CBD5E1; background: #F8FAFC;")
        self.btn_copy_team_json.setToolTip("Copy machine-readable JSON representation of this team plan to clipboard.")
        self.btn_copy_team_json.clicked.connect(self._on_copy_delegation_json)
        rc_lay.addWidget(self.btn_copy_team_json)

        self.btn_import_team_plan = QPushButton("📥 Import Team Plan...")
        self.btn_import_team_plan.setStyleSheet("padding: 5px 10px; font-size: 11px; font-weight: 600; border-radius: 6px; border: 1px solid #CBD5E1; background: #F1F5F9;")
        self.btn_import_team_plan.setToolTip("Import delegation plan from pasted email or exported .json file.")
        self.btn_import_team_plan.clicked.connect(self._on_import_delegation_plan)
        rc_lay.addWidget(self.btn_import_team_plan)

        page_layout.addWidget(self.team_roster_card)

        # 2. Tabs: Topic Ownership vs Clash & Dispatch Board
        self.team_tabs = QTabWidget()

        # --- Tab 1: Topic Ownership Matrix ---
        tab_topics = QWidget()
        tt_lay = QVBoxLayout(tab_topics)
        tt_lay.setContentsMargins(4, 6, 4, 4)
        tt_lay.setSpacing(6)

        tt_filter_box = QHBoxLayout()
        self.team_ai_search = QLineEdit()
        self.team_ai_search.setPlaceholderText("🔍 Filter AI number, topic, or owner...")
        self.team_ai_search.setClearButtonEnabled(True)
        self.team_ai_search.textChanged.connect(self._filter_topic_ownership_table)
        tt_filter_box.addWidget(self.team_ai_search)

        self.btn_auto_assign_wis = QPushButton("🎯 Auto-Assign Followed WIs")
        self.btn_auto_assign_wis.setToolTip("Assign topics matching your followed Work Items watchlist to yourself.")
        self.btn_auto_assign_wis.clicked.connect(self._auto_assign_followed_wis_to_self)
        tt_filter_box.addWidget(self.btn_auto_assign_wis)
        tt_lay.addLayout(tt_filter_box)

        self.team_topics_table = QTableWidget()
        self.team_topics_table.setColumnCount(5)
        self.team_topics_table.setHorizontalHeaderLabels([
            "AI #", "Topic Description", "Scheduled Slots in Plan", "Lead Topic Owner", "Coverage Status"
        ])
        self.team_topics_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.team_topics_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.team_topics_table.setAlternatingRowColors(True)
        tt_lay.addWidget(self.team_topics_table)
        self.team_tabs.addTab(tab_topics, "📋 Topic Ownership Matrix")

        # --- Tab 2: Clash & Dispatch Board ---
        tab_clashes = QWidget()
        tc_lay = QVBoxLayout(tab_clashes)
        tc_lay.setContentsMargins(4, 6, 4, 4)
        tc_lay.setSpacing(6)

        self.clash_kpi_lbl = QLabel()
        self.clash_kpi_lbl.setStyleSheet("font-size: 12px; font-weight: 600; padding: 6px 10px; border-radius: 6px;")
        tc_lay.addWidget(self.clash_kpi_lbl)

        self.team_clashes_table = QTableWidget()
        self.team_clashes_table.setColumnCount(6)
        self.team_clashes_table.setHorizontalHeaderLabels([
            "Slot (Day & Quarter)", "Time", "Colliding Delegate", "Stream A", "Stream B", "Action / Appoint Coverage"
        ])
        self.team_clashes_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.Stretch)
        self.team_clashes_table.horizontalHeader().setSectionResizeMode(4, QHeaderView.Stretch)
        self.team_clashes_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.team_clashes_table.setAlternatingRowColors(True)
        tc_lay.addWidget(self.team_clashes_table)

        tc_lay.addWidget(QLabel("💡 <i>Click 'Appoint Available Delegate' to inspect who is free in that time slot and appoint coverage on a session-by-session basis.</i>"))
        self.team_tabs.addTab(tab_clashes, "⚡ Clash & Dispatch Board")

        page_layout.addWidget(self.team_tabs, 1)
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
        """Refreshes all 5 views with the currently active snapshot."""
        if not self.active_snapshot:
            self._render_empty_views()
            return

        self._update_room_guide_strip()
        self._update_overview_delegate_strip()
        self._render_overview_grid()
        self._render_agenda_mode()
        self._render_drafting_radar()
        self._render_tracked_topics()
        self._render_team_view()

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
        self.team_topics_table.setRowCount(0)
        self.team_clashes_table.setRowCount(0)

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
        if " (" in time_dual:
            parts = time_dual.split(" (", 1)
            time_display = f"{parts[0]} Venue Time ({parts[1]}"
        else:
            time_display = f"{time_dual} (Venue Time)"

        s_info = self.active_snapshot.streams.get(stream_num) if self.active_snapshot else None
        r_name = s_info.room_name if s_info else f"Stream {stream_num}"
        cap = f" ({s_info.capacity} seats)" if s_info and s_info.capacity else ""
        convener = s_info.convener if s_info else ("SA2 Chair / MCC" if stream_num == 1 else "")
        tohru = s_info.tohru_room if s_info else ""

        meta_parts = [f"🏢 <b>Stream {stream_num}:</b> {html.escape(r_name)}{cap}"]
        if convener:
            meta_parts.append(f"👤 <b>Chair:</b> {html.escape(convener)}")
        if tohru:
            meta_parts.append(f"📡 <b>TOHRU:</b> {html.escape(tohru)}")
        room_line = " &nbsp;|&nbsp; ".join(meta_parts)

        ai_bullets = []
        for it in slot.items:
            if it.ai_number:
                ai_info = self.agenda_map.get(it.ai_number)
                desc = ai_info.description if ai_info else it.topic_label
                tdoc_str = f" [{it.tdoc_count_str}]" if it.tdoc_count_str else ""
                ai_bullets.append(f"• <b>AI {it.ai_number}:</b> {html.escape(desc)}{tdoc_str}")

        tt = f"""<div style='font-family:Segoe UI, sans-serif; font-size:12px; min-width:280px; padding:4px;'>
            <div style='font-weight:700; font-size:13px; color:#0F172A; margin-bottom:4px;'>
                📅 {day_name} ({date_str}) · {period} · {half_lbl}
            </div>
            <div style='color:#475569; font-size:11.5px; margin-bottom:3px;'>
                🕒 <b>Time:</b> {time_display}
            </div>
            <div style='color:#475569; font-size:11.5px; margin-bottom:6px;'>
                {room_line}
            </div>
            <hr style='border:none; border-top:1px solid #E2E8F0; margin:6px 0;'>"""

        raw_note = slot.raw_text.strip() if slot.raw_text else ""
        if raw_note:
            tt += f"""<div style='font-size:10.5px; font-weight:700; color:#475569; text-transform:uppercase; letter-spacing:0.5px; margin-bottom:2px;'>
                📋 Chair's Slot Note:
            </div>
            <div style='font-weight:600; font-size:12px; color:#0F172A; margin-bottom:6px;'>
                {html.escape(raw_note)}
            </div>"""

        if ai_bullets:
            count_str = f" ({len(ai_bullets)})" if len(ai_bullets) > 1 else ""
            tt += f"""<div style='font-size:10.5px; font-weight:700; color:#475569; text-transform:uppercase; letter-spacing:0.5px; margin-top:6px; margin-bottom:3px;'>
                📑 Agenda Items Covered{count_str}:
            </div>
            <div style='color:#334155; font-size:11px; margin-left:4px; line-height:1.45;'>
                {'<br>'.join(ai_bullets)}
            </div>"""

        if self._is_slot_tracked(slot):
            tt += "<div style='margin-top:6px; padding:3px 8px; background:#F5F3FF; border:1px solid #DDD6FE; color:#5B21B6; border-radius:4px; font-weight:600; font-size:11.5px;'>🎯 Tracked in My Topics</div>"

        del_names, is_cov, is_all = self._get_slot_effective_delegates(slot)
        if is_all:
            cov_lbl = " <span style='color:#6366F1; font-weight:normal;'>(Ad-Hoc Coverage)</span>" if is_cov else " <span style='color:#6366F1; font-weight:normal;'>(General / All Team Topic)</span>"
            tt += f"<div style='margin-top:6px; padding:4px 8px; background:#EEF2FF; border:1px solid #C7D2FE; color:#3730A3; border-radius:4px; font-size:11.5px;'>👥 <b>Delegation:</b> <b>All Team</b>{cov_lbl}</div>"
        elif del_names:
            role_lbl = " <span style='color:#64748B; font-weight:normal;'>(Ad-Hoc Coverage)</span>" if is_cov else " <span style='color:#64748B; font-weight:normal;'>(Topic Owner)</span>"
            badges_html = []
            for dn in del_names:
                m = self.delegation_plan.members.get(dn) if self.delegation_plan else None
                col = m.color if m else "#2563EB"
                bg_t = m.bg_tint if m else DelegationMember.compute_tint(col, 0.12)
                bdr_t = m.border_tint if m else DelegationMember.compute_tint(col, 0.35)
                badges_html.append(f"<span style='background:{bg_t}; border:1px solid {bdr_t}; color:{col}; padding:1px 6px; border-radius:3px; font-weight:700;'>● {html.escape(dn)}</span>")
            tt += f"<div style='margin-top:6px; padding:4px 8px; background:#F8FAFC; border:1px solid #E2E8F0; border-radius:4px; font-size:11.5px;'>👤 <b>Delegation:</b> {' '.join(badges_html)}{role_lbl}</div>"

        ann = self.annotations.get(slot.slot_key)
        if ann and (ann.comment or ann.assigned_colleague):
            note_items = []
            if ann.comment:
                note_items.append(f"“{html.escape(ann.comment)}”")
            if ann.assigned_colleague:
                note_items.append(f"(Assigned: <b>{html.escape(ann.assigned_colleague)}</b>)")
            tt += f"<div style='margin-top:6px; padding:4px 8px; background:#FEF3C7; border:1px solid #FDE68A; color:#92400E; border-radius:4px; font-size:11.5px;'>📝 <b>Personal Note:</b> {' '.join(note_items)}</div>"

        tt += "</div>"
        return tt

    def _make_delegate_badge(
        self,
        member: DelegationMember,
        initials_text: str,
        is_cov: bool = False,
        is_clash: bool = False,
        clash_slots: Optional[List[SessionSlot]] = None,
    ) -> QLabel:
        """Creates a compact circular/pill badge showing 2-letter delegate initials."""
        lbl = QLabel(initials_text)
        lbl.setAlignment(Qt.AlignCenter)

        role_desc = "Ad-hoc coverage override" if is_cov else "Topic owner"
        tt_lines = [f"<b>{html.escape(member.name)}</b> ({html.escape(initials_text)}) — {role_desc}"]
        if is_clash and clash_slots:
            for cs in clash_slots:
                tt_lines.append(f"⚠️ Clashes with Stream {cs.stream_number}: {html.escape(cs.raw_text[:48])}")
        elif is_clash:
            tt_lines.append("⚠️ Parallel-stream clash detected!")
        lbl.setToolTip("<br>".join(tt_lines))

        border_style = "dashed" if is_cov else "solid"
        if is_clash:
            bdr = f"2px {border_style} #DC2626"
        else:
            bdr = f"1px {border_style} {member.border_tint}"

        lbl.setStyleSheet(
            f"background: {member.bg_tint}; color: {member.color}; font-weight: 700; font-size: 9px; "
            f"border: {bdr}; border-radius: 8px; padding: 1px 5px; min-width: 16px;"
        )
        return lbl

    def _create_slot_cell_widget(
        self, slot: SessionSlot, day_idx: int, period: str, stream_num: int, is_half: bool = False, tu: int = 0
    ) -> QWidget:
        """Creates clean visual card for a session slot with delegate coloring and conflict tags."""
        w = QWidget()
        lay = QVBoxLayout(w)
        lay.setContentsMargins(4, 3, 4, 3)
        lay.setSpacing(2)

        del_names, is_cov, is_all = self._get_slot_effective_delegates(slot)
        assigned_members = [
            self.delegation_plan.members[name]
            for name in del_names
            if self.delegation_plan and name in self.delegation_plan.members
        ]

        collisions = (
            self._grid_collisions_cache
            if self._grid_collisions_cache is not None
            else self._detect_delegation_collisions()
        )
        colliding_members = []
        clash_partners: Dict[str, List[SessionSlot]] = {}
        for d_name in del_names:
            quarter_clashes = collisions.get((d_name, slot.day_index, slot.quarter), [])
            if any(slot.id == cs.id or slot.slot_key == cs.slot_key for cs in quarter_clashes):
                colliding_members.append(d_name)
                clash_partners[d_name] = [
                    cs for cs in quarter_clashes
                    if cs.slot_key != slot.slot_key and self._slots_time_overlap(cs, slot)
                ]
        is_collision = len(colliding_members) > 0

        # Compact badge strip: colored initials dots instead of full names (full names in tooltips)
        if is_all or assigned_members:
            pill_row = QHBoxLayout()
            pill_row.setContentsMargins(0, 0, 0, 0)
            pill_row.setSpacing(3)

            if is_all:
                all_pill = QLabel("👥 ALL")
                all_pill.setToolTip("Whole delegation" + (" (ad-hoc coverage)" if is_cov else " (All Team topic)"))
                border_style = "dashed" if is_cov else "solid"
                all_pill.setStyleSheet(
                    "background: #EEF2FF; color: #4338CA; font-weight: 700; font-size: 9px; "
                    f"border: 1px {border_style} #6366F1; border-radius: 8px; padding: 1px 5px;"
                )
                pill_row.addWidget(all_pill)
            else:
                initials = self._delegate_initials_map()
                max_badges = 4
                shown = assigned_members if len(assigned_members) <= max_badges else assigned_members[: max_badges - 1]
                for m in shown:
                    pill_row.addWidget(
                        self._make_delegate_badge(
                            m,
                            initials.get(m.name, m.initials),
                            is_cov=is_cov,
                            is_clash=m.name in colliding_members,
                            clash_slots=clash_partners.get(m.name, []),
                        )
                    )
                if len(assigned_members) > len(shown):
                    rest = assigned_members[len(shown):]
                    more = QLabel(f"+{len(rest)}")
                    more.setToolTip(", ".join(m.name for m in rest))
                    more.setStyleSheet(
                        "background: #F1F5F9; color: #334155; font-weight: 700; font-size: 9px; "
                        "border: 1px solid #CBD5E1; border-radius: 9px; padding: 1px 4px;"
                    )
                    pill_row.addWidget(more)

            if is_collision:
                warn_pill = QLabel("⚠")
                who = ", ".join(colliding_members)
                warn_pill.setToolTip(f"Parallel-stream clash for: {who}")
                warn_pill.setStyleSheet(
                    "background: #FEF3C7; color: #B45309; font-weight: bold; font-size: 10px; "
                    "border: 1px solid #F59E0B; border-radius: 8px; padding: 0px 4px;"
                )
                pill_row.addWidget(warn_pill)

            pill_row.addStretch()
            lay.addLayout(pill_row)

        lbl = QLabel(slot.raw_text)
        lbl.setWordWrap(True)
        lbl.setAlignment(Qt.AlignTop | Qt.AlignLeft)

        is_tracked = self._is_slot_tracked(slot)

        # Dimming check if a (multi-select) delegate focus is active on the Overview Grid
        is_dimmed = False
        is_focused = False
        if self.grid_focus_delegates:
            focus_lower = {n.lower() for n in self.grid_focus_delegates}
            if is_all or any(dn.lower() in focus_lower for dn in del_names):
                is_focused = True
            else:
                is_dimmed = True

        if is_dimmed:
            lbl.setStyleSheet("background: #F8FAFC; color: #94A3B8; border: 1px solid #E2E8F0; border-radius: 4px; padding: 3px;")
        elif is_collision:
            lbl.setStyleSheet("background: #FFFBEB; color: #92400E; border: 1px solid #F59E0B; border-left: 4px solid #D97706; border-radius: 4px; padding: 4px; font-weight: 500;")
        elif is_all:
            if is_focused:
                lbl.setStyleSheet("background: #EEF2FF; color: #0F172A; border: 1.5px solid #4338CA; border-left: 5px solid #4338CA; border-radius: 4px; padding: 4px; font-weight: 600;")
            else:
                lbl.setStyleSheet("background: #FFFFFF; color: #0F172A; border: 1px solid #CBD5E1; border-left: 4px solid #4338CA; border-radius: 4px; padding: 4px; font-weight: 500;")
        elif assigned_members:
            primary_m = assigned_members[0]
            if is_focused:
                lbl.setStyleSheet(f"background: {primary_m.bg_tint}; color: #0F172A; border: 1.5px solid {primary_m.color}; border-left: 5px solid {primary_m.color}; border-radius: 4px; padding: 4px; font-weight: 600;")
            else:
                lbl.setStyleSheet(f"background: #FFFFFF; color: #0F172A; border: 1px solid #CBD5E1; border-left: 4px solid {primary_m.color}; border-radius: 4px; padding: 4px; font-weight: 500;")
        elif slot.is_drafting:
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
        self._grid_collisions_cache = self._detect_delegation_collisions()
        try:
            self._render_overview_grid_impl()
        finally:
            self._grid_collisions_cache = None

    def _render_overview_grid_impl(self):
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
        """Populates the Delegate Agenda view with delegate filtering and conflict detection."""
        if not self.active_snapshot:
            return

        # 1. Update Delegate Agenda combo items if needed
        self.delegate_agenda_combo.blockSignals(True)
        cur_selection = self.delegate_agenda_combo.currentData()
        self.delegate_agenda_combo.clear()

        self_member = next((m for m in self.delegation_plan.members.values() if m.is_self), None) if self.delegation_plan else None
        if self_member:
            self.delegate_agenda_combo.addItem(f"👤 Myself ({self_member.name})", "SELF")
        else:
            self.delegate_agenda_combo.addItem("👤 My Tracked Topics (Personal)", "SELF")

        if self.delegation_plan:
            for name, m in self.delegation_plan.members.items():
                self.delegate_agenda_combo.addItem(f"● {name}", name)

        self.delegate_agenda_combo.addItem("👥 Whole Delegation (Master Schedule)", "ALL")

        # Restore selection
        idx = self.delegate_agenda_combo.findData(cur_selection)
        if idx >= 0:
            self.delegate_agenda_combo.setCurrentIndex(idx)
        else:
            self.delegate_agenda_combo.setCurrentIndex(0)
        self.delegate_agenda_combo.blockSignals(False)

        selected_scope = self.delegate_agenda_combo.currentData() or "SELF"

        # 2. Filter target slots based on scope
        target_slots: List[SessionSlot] = []
        for s in self.active_snapshot.slots:
            del_names, is_cov, is_all = self._get_slot_effective_delegates(s)
            if selected_scope == "ALL":
                if del_names or is_all or self._is_slot_tracked(s):
                    target_slots.append(s)
            elif selected_scope == "SELF":
                if self_member:
                    if is_all or any(dn.lower() == self_member.name.lower() for dn in del_names):
                        target_slots.append(s)
                    elif not del_names and self._is_slot_tracked(s):
                        target_slots.append(s)
                else:
                    if self._is_slot_tracked(s):
                        target_slots.append(s)
            else:
                if is_all or any(dn.lower() == selected_scope.lower() for dn in del_names):
                    target_slots.append(s)

        # 3. Conflict detection: same day and quarter, multiple streams
        time_slot_map: Dict[Tuple[int, str], List[SessionSlot]] = {}
        for s in target_slots:
            time_slot_map.setdefault((s.day_index, s.quarter), []).append(s)

        conflicts = {k: slots for k, slots in time_slot_map.items() if len(slots) > 1}

        if conflicts:
            self.btn_tracked.setText(f"🎯 Delegate Agenda ({len(target_slots)}) ⚠️ {len(conflicts)} Clash(es)")
        else:
            self.btn_tracked.setText(f"🎯 Delegate Agenda ({len(target_slots)})")

        self.btn_manage_topics.setText(f"⭐ Track Topics ({len(self.tracked_tokens)})")
        self.tracked_table.setRowCount(len(target_slots))

        streams = self.active_snapshot.streams
        for row, s in enumerate(target_slots):
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

            del_names, is_cov, is_all = self._get_slot_effective_delegates(s)
            cov_lbl = " (cov)" if is_cov else ""
            if is_all:
                del_txt = f"👥 All Team{cov_lbl}"
            elif del_names:
                del_txt = f"{', '.join(del_names)}{cov_lbl}"
            else:
                del_txt = "—"

            item_day = QTableWidgetItem(s.day_name)
            item_day.setData(Qt.UserRole, s)
            item_q = QTableWidgetItem(s.quarter)
            item_time = QTableWidgetItem(time_str)
            item_str = QTableWidgetItem(f"Str {s.stream_number} ({r_name})")
            item_del = QTableWidgetItem(del_txt)

            if is_all:
                item_del.setForeground(QColor("#4338CA"))
                f = item_del.font()
                f.setBold(True)
                item_del.setFont(f)
            elif len(del_names) == 1 and self.delegation_plan:
                m = self.delegation_plan.members.get(del_names[0])
                if m:
                    item_del.setForeground(QColor(m.color))
                    f = item_del.font()
                    f.setBold(True)
                    item_del.setFont(f)
            elif len(del_names) > 1:
                item_del.setForeground(QColor("#0F172A"))
                f = item_del.font()
                f.setBold(True)
                item_del.setFont(f)

            item_raw = QTableWidgetItem(s.raw_text)
            item_note = QTableWidgetItem(note_display or "—")

            if is_conflict:
                for itm in [item_day, item_q, item_time, item_str, item_del, item_raw, item_note]:
                    itm.setBackground(QColor("#FEF3C7"))
                item_note.setForeground(QColor("#B45309"))

            self.tracked_table.setItem(row, 0, item_day)
            self.tracked_table.setItem(row, 1, item_q)
            self.tracked_table.setItem(row, 2, item_time)
            self.tracked_table.setItem(row, 3, item_str)
            self.tracked_table.setItem(row, 4, item_del)
            self.tracked_table.setItem(row, 5, item_raw)
            self.tracked_table.setItem(row, 6, item_note)

    # --- Delegation Engine & Team View ---

    def _get_slot_effective_delegates(self, slot: SessionSlot) -> Tuple[List[str], bool, bool]:
        """
        Returns (delegates_list, is_appointment_override, is_all).
        Priority:
        1. Explicit session appointment (slot override).
        2. Topic ownership from extracted slot items.
        3. Boundary-aware raw text match against assigned AI tokens
           (``20.6.2`` must never match ``20.6.24``).
        Delegates to the shared ``resolve_slot_effective_delegates`` so the grid,
        clash board, and exporters always agree.
        """
        return resolve_slot_effective_delegates(slot, self.delegation_plan)

    def _get_slot_effective_delegate(self, slot: SessionSlot) -> Tuple[Optional[str], bool]:
        """Backward-compatible single-string delegate helper."""
        dels, is_cov, is_all = self._get_slot_effective_delegates(slot)
        if not dels:
            return None, False
        if is_all:
            return "ALL", is_cov
        return ", ".join(dels), is_cov

    def _slots_time_overlap(self, s1: SessionSlot, s2: SessionSlot) -> bool:
        """
        Determines whether two slots physically overlap in meeting time.
        - Must be the same day and quarter.
        - If either slot spans the full quarter (sub_slot == 0), they overlap.
        - If both are half-quarters (sub_slot 1 or 2), they overlap iff sub_slot == sub_slot.
        """
        if s1.day_index != s2.day_index or s1.quarter != s2.quarter:
            return False
        if s1.sub_slot == 0 or s2.sub_slot == 0:
            return True
        return s1.sub_slot == s2.sub_slot

    def _detect_delegation_collisions(self) -> Dict[Tuple[str, int, str], List[SessionSlot]]:
        """
        Detects parallel stream collisions for delegates taking half-quarter TU1/TU2 into account.
        Key: (delegate_name, day_index, quarter) -> list of colliding SessionSlots.
        Only slots that physically overlap in time across different active streams (>0) are flagged.
        """
        if not self.active_snapshot or not self.delegation_plan:
            return {}

        slot_map: Dict[Tuple[str, int, str], List[SessionSlot]] = {}
        for s in self.active_snapshot.slots:
            del_names, _, _ = self._get_slot_effective_delegates(s)
            for del_name in del_names:
                key = (del_name, s.day_index, s.quarter)
                slot_map.setdefault(key, []).append(s)

        collisions: Dict[Tuple[str, int, str], List[SessionSlot]] = {}
        for (del_name, d_idx, q_name), slots in slot_map.items():
            colliding_slots: List[SessionSlot] = []
            for i, s1 in enumerate(slots):
                for j, s2 in enumerate(slots):
                    if i != j and s1.stream_number != s2.stream_number and s1.stream_number > 0 and s2.stream_number > 0:
                        if self._slots_time_overlap(s1, s2):
                            if s1 not in colliding_slots:
                                colliding_slots.append(s1)
                            if s2 not in colliding_slots:
                                colliding_slots.append(s2)
            if colliding_slots:
                collisions[(del_name, d_idx, q_name)] = colliding_slots

        return collisions

    def _compute_delegate_availability(self, target_slot: SessionSlot) -> Dict[str, dict]:
        """
        For each member in the roster, determines availability for target_slot:
        returns: member_name -> {'status': 'free'|'busy', 'reason': str}
        """
        res = {}
        if not self.active_snapshot or not self.delegation_plan:
            return res

        for name, m in self.delegation_plan.members.items():
            busy_slots = []
            for s in self.active_snapshot.slots:
                if (
                    s.stream_number != target_slot.stream_number
                    and s.stream_number > 0
                    and self._slots_time_overlap(s, target_slot)
                ):
                    del_names, _, _ = self._get_slot_effective_delegates(s)
                    if any(dn.lower() == name.lower() for dn in del_names):
                        busy_slots.append(s)
                        busy_slots.append(s)

            if busy_slots:
                streams_str = ", ".join(f"Str {bs.stream_number}" for bs in busy_slots)
                res[name] = {"status": "busy", "reason": f"Busy in {streams_str} ({busy_slots[0].raw_text[:28]}...)"}
            else:
                res[name] = {"status": "free", "reason": "Available (no parallel session)"}

        return res

    def _delegate_initials_map(self) -> Dict[str, str]:
        """Returns roster-unique badge initials (cached until the roster changes)."""
        names = tuple(self.delegation_plan.members.keys()) if self.delegation_plan else ()
        if self._initials_cache[0] != names:
            self._initials_cache = (names, compute_unique_initials(names))
        return self._initials_cache[1]

    def _self_member(self) -> Optional[DelegationMember]:
        if not self.delegation_plan:
            return None
        return next((m for m in self.delegation_plan.members.values() if m.is_self), None)

    def _update_overview_delegate_strip(self):
        """Rebuilds the multi-select Delegate Focus checkboxes above the Overview Grid."""
        if not hasattr(self, "delegate_chips_layout"):
            return

        while self.delegate_chips_layout.count():
            item = self.delegate_chips_layout.takeAt(0)
            w = item.widget()
            if w:
                w.deleteLater()
        self.dynamic_delegate_chips.clear()
        self.btn_chip_me = None

        roster = self.delegation_plan.members if self.delegation_plan else {}
        # Drop focus entries for delegates that were removed/renamed
        self.grid_focus_delegates = {n for n in self.grid_focus_delegates if n in roster}

        self_member = self._self_member()
        if self_member:
            btn_me = QPushButton(f"👤 Only Me ({self_member.name})")
            btn_me.setToolTip("Focus on your own sessions only (shortcut for ticking just yourself).")
            btn_me.setCursor(Qt.PointingHandCursor)
            btn_me.clicked.connect(lambda: self._on_delegate_highlight_clicked("SELF"))
            self.delegate_chips_layout.addWidget(btn_me)
            self.btn_chip_me = btn_me

        initials = self._delegate_initials_map()
        for name, m in roster.items():
            chk = QCheckBox(f"{initials.get(name, m.initials)} · {name}")
            chk.setProperty("delegate_name", name)
            chk.setCursor(Qt.PointingHandCursor)
            chk.setToolTip(f"Highlight sessions covered by {name}. Tick several delegates to combine.")
            chk.setChecked(name in self.grid_focus_delegates)
            chk.toggled.connect(lambda checked, n=name: self._on_focus_delegate_toggled(n, checked))
            self.delegate_chips_layout.addWidget(chk)
            self.dynamic_delegate_chips.append(chk)

        self._style_focus_buttons()

    def _style_focus_buttons(self):
        """Applies active/inactive styling to the focus strip according to the focus set."""
        base = "padding: 2px 8px; font-size: 11px; font-weight: 600; border-radius: 12px;"
        active = "border: 1px solid #3B82F6; background: #3B82F6; color: #FFFFFF;"
        idle = "border: 1px solid #CBD5E1; background: #FFFFFF; color: #334155;"
        self.btn_chip_all.setStyleSheet(f"QPushButton {{ {base} {active if not self.grid_focus_delegates else idle} }}")

        self_member = self._self_member()
        if self.btn_chip_me is not None:
            only_me = bool(self_member) and self.grid_focus_delegates == {self_member.name}
            col = self_member.color if self_member else "#0F172A"
            style = f"border: 1px solid {col}; background: {col}; color: #FFFFFF;" if only_me else idle
            self.btn_chip_me.setStyleSheet(f"QPushButton {{ {base} {style} }}")

        for chk in self.dynamic_delegate_chips:
            name = chk.property("delegate_name")
            m = self.delegation_plan.members.get(name) if self.delegation_plan else None
            col = m.color if m else "#334155"
            bg = m.bg_tint if m else "#FFFFFF"
            bdr = m.border_tint if m else "#CBD5E1"
            if chk.isChecked():
                bg, bdr = (m.compute_tint(col, 0.22) if m else "#E2E8F0"), col
            chk.setStyleSheet(
                f"QCheckBox {{ {base} padding: 2px 8px 2px 6px; color: {col}; background: {bg}; border: 1px solid {bdr}; }}"
                f"QCheckBox::indicator {{ width: 12px; height: 12px; }}"
            )

    def _sync_focus_mirror(self):
        """Keeps the legacy single-value ``grid_highlight_delegate`` attribute coherent."""
        self_member = self._self_member()
        if not self.grid_focus_delegates:
            self.grid_highlight_delegate = "ALL"
        elif self_member and self.grid_focus_delegates == {self_member.name}:
            self.grid_highlight_delegate = "SELF"
        elif len(self.grid_focus_delegates) == 1:
            self.grid_highlight_delegate = next(iter(self.grid_focus_delegates))
        else:
            self.grid_highlight_delegate = "MULTI"

    def _set_focus_delegates(self, names: Set[str]):
        """Sets the focus set, syncs checkboxes without signal storms, and re-renders the grid once."""
        self.grid_focus_delegates = set(names)
        for chk in self.dynamic_delegate_chips:
            chk.blockSignals(True)
            chk.setChecked(chk.property("delegate_name") in self.grid_focus_delegates)
            chk.blockSignals(False)
        self._sync_focus_mirror()
        self._style_focus_buttons()
        self._render_overview_grid()

    def _on_focus_delegate_toggled(self, name: str, checked: bool):
        names = set(self.grid_focus_delegates)
        if checked:
            names.add(name)
        else:
            names.discard(name)
        self._set_focus_delegates(names)

    def _on_delegate_highlight_clicked(self, key: str):
        """
        Backward-compatible focus entry point:
        "ALL" clears the focus, "SELF" focuses only the local user, a roster name toggles that delegate.
        """
        if key == "ALL":
            self._set_focus_delegates(set())
        elif key == "SELF":
            me = self._self_member()
            self._set_focus_delegates({me.name} if me else set())
        else:
            names = set(self.grid_focus_delegates)
            names.symmetric_difference_update({key})
            self._set_focus_delegates(names)

    def _on_delegate_agenda_changed(self):
        self._render_tracked_topics()

    def _get_scheduled_ais_map(self) -> Dict[str, List[SessionSlot]]:
        """Maps AI numbers to the list of SessionSlots in which they are scheduled."""
        scheduled_ais: Dict[str, List[SessionSlot]] = {}
        if self.active_snapshot:
            for s in self.active_snapshot.slots:
                for it in s.items:
                    if it.ai_number:
                        scheduled_ais.setdefault(it.ai_number, []).append(s)
        return scheduled_ais

    def _render_team_roster_chips(self):
        """Refreshes the horizontal roster chips strip at the top of the Team view."""
        if not self.delegation_plan:
            return
        while self.team_member_chips_layout.count():
            item = self.team_member_chips_layout.takeAt(0)
            w = item.widget()
            if w:
                w.deleteLater()

        for name, m in self.delegation_plan.members.items():
            chip = QPushButton(f"● {name}" + (" (Me)" if m.is_self else ""))
            chip.setStyleSheet(
                f"padding: 3px 10px; font-size: 11px; font-weight: 600; border-radius: 12px; "
                f"border: 1px solid {m.border_tint}; background: {m.bg_tint}; color: {m.color};"
            )
            chip.setCursor(Qt.PointingHandCursor)
            chip.clicked.connect(lambda _, mem=m: self._on_edit_delegation_member(mem))
            chip.setContextMenuPolicy(Qt.CustomContextMenu)
            chip.customContextMenuRequested.connect(
                lambda pos, n=name, mem=m, btn=chip: self._show_member_context_menu(pos, n, mem, btn)
            )
            self.team_member_chips_layout.addWidget(chip)

    def _render_team_topics(self):
        """Populates the Topic Ownership Matrix table with comboboxes and status tags."""
        if not self.delegation_plan:
            return

        scheduled_ais = self._get_scheduled_ais_map()
        if self.agenda_map:
            all_ais = sorted(self.agenda_map.keys(), key=self._ai_sort_key)
        else:
            all_ais = sorted(scheduled_ais.keys(), key=self._ai_sort_key)

        self.team_topics_table.setRowCount(len(all_ais))
        collisions = self._detect_delegation_collisions()
        members_list = list(self.delegation_plan.members.values())
        initials_map = self._delegate_initials_map()

        for row, ai_num in enumerate(all_ais):
            if ai_num in self.agenda_map:
                ai_item = self.agenda_map[ai_num]
                desc = ai_item.description
            else:
                matched_s = scheduled_ais.get(ai_num, [])
                desc = matched_s[0].raw_text if matched_s else f"Agenda Item {ai_num}"

            slots_for_ai = scheduled_ais.get(ai_num, [])
            slots_summary = ", ".join(f"{s.day_name[:3]} {s.quarter} (Str {s.stream_number})" for s in slots_for_ai[:3])
            if len(slots_for_ai) > 3:
                slots_summary += f" +{len(slots_for_ai)-3} more"
            if not slots_summary:
                slots_summary = "— (Not scheduled)"

            to = self.delegation_plan.topic_owners.get(ai_num)
            owner_name = to.owner_name if to else ""
            owners_list = to.get_owners([m.name for m in members_list]) if to else []

            has_clash = False
            if owners_list and slots_for_ai:
                for s in slots_for_ai:
                    for o_name in owners_list:
                        quarter_clashes = collisions.get((o_name, s.day_index, s.quarter), [])
                        if any(s.id == cs.id or s.slot_key == cs.slot_key for cs in quarter_clashes):
                            has_clash = True
                            break
                    if has_clash:
                        break

            item_ai = QTableWidgetItem(ai_num)
            item_desc = QTableWidgetItem(desc)
            item_slots = QTableWidgetItem(slots_summary)

            # Checkbox-based multi-select owner picker (replaces combo + custom dialog round-trip)
            picker = DelegateOwnerPicker(members_list, owner_name, initials=initials_map)
            picker.ownerChanged.connect(lambda new_owner, a=ai_num: self._on_topic_owner_changed(a, new_owner))

            if has_clash:
                status_txt = "🔴 Collision in parallel stream"
                status_col = "#991B1B"
                status_bg = "#FEF2F2"
            elif owner_name:
                cov_lbl = "All Team" if to and to.is_all() else (f"{len(owners_list)} delegates" if len(owners_list) > 1 else owner_name)
                status_txt = f"🟢 Covered ({cov_lbl} · {len(slots_for_ai)} slot(s))" if slots_for_ai else f"⚪ Assigned ({cov_lbl})"
                status_col = "#065F46"
                status_bg = "#ECFDF5"
            else:
                status_txt = "🟡 Unassigned"
                status_col = "#92400E"
                status_bg = "#FFFBEB"

            item_status = QTableWidgetItem(status_txt)
            item_status.setForeground(QColor(status_col))
            item_status.setBackground(QColor(status_bg))

            self.team_topics_table.setItem(row, 0, item_ai)
            self.team_topics_table.setItem(row, 1, item_desc)
            self.team_topics_table.setItem(row, 2, item_slots)
            self.team_topics_table.setCellWidget(row, 3, picker)
            self.team_topics_table.setItem(row, 4, item_status)

    def _update_team_topic_status_cells(self):
        """Fast status column refresh for the Topic Ownership Matrix without recreating controls."""
        if not self.delegation_plan:
            return
        collisions = self._detect_delegation_collisions()
        scheduled_ais = self._get_scheduled_ais_map()
        members_list = list(self.delegation_plan.members.values())
        for row in range(self.team_topics_table.rowCount()):
            ai_item = self.team_topics_table.item(row, 0)
            if not ai_item:
                continue
            ai_num = ai_item.text().strip()
            to = self.delegation_plan.topic_owners.get(ai_num)
            owner_name = to.owner_name if to else ""
            owners_list = to.get_owners([m.name for m in members_list]) if to else []
            slots_for_ai = scheduled_ais.get(ai_num, [])

            has_clash = False
            if owners_list and slots_for_ai:
                for s in slots_for_ai:
                    for o_name in owners_list:
                        quarter_clashes = collisions.get((o_name, s.day_index, s.quarter), [])
                        if any(s.id == cs.id or s.slot_key == cs.slot_key for cs in quarter_clashes):
                            has_clash = True
                            break
                    if has_clash:
                        break

            if has_clash:
                status_txt = "🔴 Collision in parallel stream"
                status_col = "#991B1B"
                status_bg = "#FEF2F2"
            elif owner_name:
                cov_lbl = "All Team" if to and to.is_all() else (f"{len(owners_list)} delegates" if len(owners_list) > 1 else owner_name)
                status_txt = f"🟢 Covered ({cov_lbl} · {len(slots_for_ai)} slot(s))" if slots_for_ai else f"⚪ Assigned ({cov_lbl})"
                status_col = "#065F46"
                status_bg = "#ECFDF5"
            else:
                status_txt = "🟡 Unassigned"
                status_col = "#92400E"
                status_bg = "#FFFBEB"

            item_status = self.team_topics_table.item(row, 4)
            if not item_status:
                item_status = QTableWidgetItem(status_txt)
                self.team_topics_table.setItem(row, 4, item_status)
            else:
                item_status.setText(status_txt)
            item_status.setForeground(QColor(status_col))
            item_status.setBackground(QColor(status_bg))

    def _render_team_clashes(self):
        """Renders the Clash & Dispatch Board with multi-stream choices and actions."""
        if not self.delegation_plan:
            return

        collisions = self._detect_delegation_collisions()
        clash_count = len(collisions)
        self.btn_team.setText(f"👥 Team Allocations ({clash_count} ⚠️)" if clash_count else "👥 Team Allocations (Clean)")
        self.team_tabs.setTabText(1, f"⚡ Clash & Dispatch Board ({clash_count})")

        if clash_count == 0:
            self.clash_kpi_lbl.setText("✅ No collisions detected! All parallel streams have distinct delegates or single coverage.")
            self.clash_kpi_lbl.setStyleSheet(
                "background: #ECFDF5; color: #065F46; border: 1px solid #A7F3D0; font-size: 12px; font-weight: 600; padding: 6px 10px; border-radius: 6px;"
            )
            self.team_clashes_table.setRowCount(0)
        else:
            self.clash_kpi_lbl.setText(
                f"⚠️ {clash_count} Double-Booking Collision(s) Detected! Appoint an available delegate to cover parallel rooms."
            )
            self.clash_kpi_lbl.setStyleSheet(
                "background: #FFFBEB; color: #92400E; border: 1px solid #FDE68A; font-size: 12px; font-weight: 600; padding: 6px 10px; border-radius: 6px;"
            )
            self.team_clashes_table.setRowCount(clash_count)

            for row, ((del_name, d_idx, q_name), slots) in enumerate(collisions.items()):
                s_a = slots[0]
                time_str = SessionPlanTimezone.format_dual_time_range(
                    s_a.start_time_local, s_a.end_time_local, self.tz_offset_min
                )

                item_slot = QTableWidgetItem(f"{s_a.day_name} {s_a.quarter}")
                item_time = QTableWidgetItem(time_str)

                item_del = QTableWidgetItem(del_name)
                m = self.delegation_plan.members.get(del_name)
                if m:
                    item_del.setForeground(QColor(m.color))
                f = item_del.font()
                f.setBold(True)
                item_del.setFont(f)

                s_info_a = self.active_snapshot.streams.get(s_a.stream_number) if self.active_snapshot else None
                r_a = s_info_a.room_name if s_info_a else f"Stream {s_a.stream_number}"
                item_stra = QTableWidgetItem(f"Str {s_a.stream_number} ({r_a}): {s_a.raw_text}")

                other_desc = []
                for other_s in slots[1:]:
                    s_inf = self.active_snapshot.streams.get(other_s.stream_number) if self.active_snapshot else None
                    r_oth = s_inf.room_name if s_inf else f"Stream {other_s.stream_number}"
                    other_desc.append(f"Str {other_s.stream_number} ({r_oth}): {other_s.raw_text}")
                item_strb = QTableWidgetItem("\n".join(other_desc) if other_desc else f"Str {slots[1].stream_number}: {slots[1].raw_text}")

                btn_appoint = QPushButton("⚡ Appoint Coverage ▾")
                btn_appoint.setStyleSheet(
                    "padding: 3px 8px; font-size: 11px; font-weight: 600; border-radius: 6px; background: #059669; color: #FFFFFF;"
                )
                menu = QMenu(btn_appoint)
                for s_opt in slots:
                    s_inf = self.active_snapshot.streams.get(s_opt.stream_number) if self.active_snapshot else None
                    rm = s_inf.room_name if s_inf else f"Stream {s_opt.stream_number}"
                    act = menu.addAction(f"Cover Str {s_opt.stream_number} ({rm}): {s_opt.raw_text[:30]}...")
                    act.triggered.connect(lambda _, slot=s_opt, orig=del_name: self._on_dispatch_slot(slot, orig))
                btn_appoint.setMenu(menu)

                self.team_clashes_table.setItem(row, 0, item_slot)
                self.team_clashes_table.setItem(row, 1, item_time)
                self.team_clashes_table.setItem(row, 2, item_del)
                self.team_clashes_table.setItem(row, 3, item_stra)
                self.team_clashes_table.setItem(row, 4, item_strb)
                self.team_clashes_table.setCellWidget(row, 5, btn_appoint)

    def _render_team_view(self):
        """Refreshes the Team Allocations view: Roster chips, Topic Ownership Matrix, and Clash Board."""
        if not self.delegation_plan:
            return
        self._render_team_roster_chips()
        self._render_team_topics()
        self._render_team_clashes()

    def _show_member_context_menu(self, pos: QPoint, name: str, member: DelegationMember, btn: QPushButton):
        menu = QMenu(self)
        menu.addAction("✏️ Edit Member...", lambda: self._on_edit_delegation_member(member))
        if not member.is_self:
            menu.addAction("👤 Set as Myself", lambda: self._on_set_self_delegate(name))
        menu.addSeparator()
        menu.addAction("🗑️ Remove Member", lambda: self._on_delete_delegation_member(name))
        menu.exec_(btn.mapToGlobal(pos))

    def _on_add_delegation_member(self):
        existing = list(self.delegation_plan.members.keys()) if self.delegation_plan else []
        dlg = DelegationMemberDialog(parent=self, member=None, existing_names=existing)
        if dlg.exec_() == QDialog.Accepted:
            new_mem = dlg.get_member()
            self.db.save_delegation_member(self.meeting_number, new_mem)
            self.delegation_plan = self.db.get_delegation_plan(self.meeting_number)
            self._render_current_views()

    def _on_edit_delegation_member(self, member: DelegationMember):
        existing = list(self.delegation_plan.members.keys()) if self.delegation_plan else []
        dlg = DelegationMemberDialog(parent=self, member=member, existing_names=existing)
        if dlg.exec_() == QDialog.Accepted:
            updated_mem = dlg.get_member()
            self.db.save_delegation_member(self.meeting_number, updated_mem)
            self.delegation_plan = self.db.get_delegation_plan(self.meeting_number)
            self._render_current_views()

    def _on_delete_delegation_member(self, name: str):
        ans = QMessageBox.question(
            self,
            "Confirm Delete",
            f"Are you sure you want to remove '{name}' from the delegation roster?\nTheir assigned topics and appointments will be cleared.",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if ans == QMessageBox.Yes:
            self.db.delete_delegation_member(self.meeting_number, name)
            self.delegation_plan = self.db.get_delegation_plan(self.meeting_number)
            self._render_current_views()

    def _on_set_self_delegate(self, name: str):
        self.db.set_self_delegate(self.meeting_number, name)
        self.delegation_plan = self.db.get_delegation_plan(self.meeting_number)
        self._render_current_views()

    def _on_topic_owner_changed(self, ai_number: str, owner_name: str):
        self.db.set_topic_owner(self.meeting_number, ai_number, owner_name)
        self.delegation_plan = self.db.get_delegation_plan(self.meeting_number)
        self._render_overview_grid()
        self._render_tracked_topics()
        self._render_team_clashes()
        self._update_team_topic_status_cells()

    def _on_topic_owner_combo_activated(self, ai_num: str, desc: str, combo: QComboBox):
        choice = combo.currentData()
        if choice == "__CUSTOM__":
            members = list(self.delegation_plan.members.values()) if self.delegation_plan else []
            to = self.delegation_plan.topic_owners.get(ai_num) if self.delegation_plan else None
            cur = to.owner_name if to else ""
            dlg = MultiDelegateSelectDialog(f"AI {ai_num}: {desc}", members, cur, self)
            if dlg.exec_() == QDialog.Accepted:
                new_owner = dlg.get_owner_string()
                self._on_topic_owner_changed(ai_num, new_owner)
                self._render_team_topics()
            else:
                to = self.delegation_plan.topic_owners.get(ai_num) if self.delegation_plan else None
                cur = to.owner_name if to else ""
                idx = combo.findData(cur)
                if idx >= 0:
                    combo.setCurrentIndex(idx)
        else:
            self._on_topic_owner_changed(ai_num, choice)

    def _on_dispatch_slot(self, slot: SessionSlot, original_owner: str):
        members = list(self.delegation_plan.members.values()) if self.delegation_plan else []
        if not members:
            QMessageBox.information(self, "No Roster", "Please add delegation members before appointing coverage.")
            return

        avail = self._compute_delegate_availability(slot)
        dlg = SessionDispatchDialog(slot, original_owner, members, avail, self)
        if dlg.exec_() == QDialog.Accepted:
            appointed, reason = dlg.get_result()
            if appointed:
                self.db.set_session_appointment(
                    meeting_number=self.meeting_number,
                    slot_key=slot.slot_key,
                    appointed_delegate=appointed,
                    original_owner=original_owner,
                    reason=reason,
                )
            else:
                self.db.delete_session_appointment(self.meeting_number, slot.slot_key)

            self.delegation_plan = self.db.get_delegation_plan(self.meeting_number)
            self._render_current_views()

    def _on_revert_dispatch(self, slot: SessionSlot):
        """Removes ad-hoc appointment override, restoring ownership to original topic owner."""
        self.db.delete_session_appointment(self.meeting_number, slot.slot_key)
        self.delegation_plan = self.db.get_delegation_plan(self.meeting_number)
        self._render_current_views()

    def _filter_topic_ownership_table(self, query: str):
        q = query.lower().strip()
        for r in range(self.team_topics_table.rowCount()):
            ai_item = self.team_topics_table.item(r, 0)
            desc_item = self.team_topics_table.item(r, 1)
            combo = self.team_topics_table.cellWidget(r, 3)
            owner_txt = combo.currentText() if combo else ""
            txt = f"{ai_item.text()} {desc_item.text()} {owner_txt}".lower() if ai_item and desc_item else ""
            self.team_topics_table.setRowHidden(r, bool(q and q not in txt))

    def _auto_assign_followed_wis_to_self(self):
        self_member = next((m for m in self.delegation_plan.members.values() if m.is_self), None) if self.delegation_plan else None
        if not self_member:
            QMessageBox.warning(self, "No Self Identity", "Please set one of the delegation members as 'Myself' first.")
            return

        count = 0
        if self.agenda_map:
            for ai_num, item in self.agenda_map.items():
                acronym = getattr(item, "acronym", "") or ""
                desc = getattr(item, "description", "") or ""
                matches = (
                    ai_num.lower() in self.tracked_tokens
                    or (acronym and acronym.lower() in self.tracked_tokens)
                    or any(t in desc.lower() for t in self.tracked_tokens if t)
                )
                if matches:
                    self.db.set_topic_owner(self.meeting_number, ai_num, self_member.name)
                    count += 1

        self.delegation_plan = self.db.get_delegation_plan(self.meeting_number)
        self._render_current_views()
        QMessageBox.information(self, "Auto-Assignment Complete", f"Assigned {count} topic(s) to {self_member.name}.")

    def _on_export_delegation_email(self):
        if not self.active_snapshot or not self.delegation_plan:
            return
        subject, html_body, plain_text = SessionPlanExporter.generate_delegation_email_summary(
            snapshot=self.active_snapshot,
            plan=self.delegation_plan,
            meeting_tz_offset_min=self.tz_offset_min,
            tz_name=self.tz_name,
            agenda_map=self.agenda_map,
        )
        QApplication.clipboard().setText(html_body)
        QMessageBox.information(
            self,
            "Delegation Summary Copied",
            f"Subject: {subject}\n\nPre-formatted executive rich HTML email summary copied to clipboard!\nYou can paste it directly into Microsoft Outlook or Teams.",
        )

    def _on_copy_delegation_json(self):
        if not self.delegation_plan:
            return
        payload = SessionPlanExporter.serialize_delegation_plan_json(self.delegation_plan)
        QApplication.clipboard().setText(payload)
        QMessageBox.information(
            self,
            "Plan JSON Copied",
            "Delegation Plan JSON copied to clipboard.\nColleagues can import this using 'Import Team Plan'.",
        )

    def _on_import_delegation_plan(self):
        dlg = DelegationImportDialog(
            current_roster=list(self.delegation_plan.members.keys()) if self.delegation_plan else [],
            parent=self,
        )
        if dlg.exec_() == QDialog.Accepted:
            data, is_merge, claimed_self = dlg.get_data()
            if data:
                count = self.db.import_delegation_plan_dict(self.meeting_number, data, merge=is_merge)
                if claimed_self:
                    self.db.set_self_delegate(self.meeting_number, claimed_self)
                self.delegation_plan = self.db.get_delegation_plan(self.meeting_number)
                self._render_current_views()
                QMessageBox.information(
                    self,
                    "Import Complete",
                    f"Successfully imported delegation plan ({count} items updated).\nYour identity is set to: {claimed_self or 'Default'}.",
                )

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
        topic_item = self.tracked_table.item(r, 5)
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

        s = d_item.data(Qt.UserRole)
        if self.delegation_plan and self.delegation_plan.members and s:
            menu.addSeparator()
            del_name, is_cov = self._get_slot_effective_delegate(s)
            disp_name = "All Team" if del_name == "ALL" else (del_name or "")
            app = self.delegation_plan.appointments.get(s.slot_key)
            orig_owner = app.original_owner if (app and app.original_owner) else (del_name or "")
            if is_cov:
                act_disp = menu.addAction(f"⚡ Change Coverage for this Session (Covered by {disp_name})...")
                act_disp.triggered.connect(lambda _, slot=s, orig=orig_owner: self._on_dispatch_slot(slot, orig))
                act_rev = menu.addAction(f"↺ Revert Coverage to {orig_owner or 'Topic Owner'}")
                act_rev.triggered.connect(lambda _, slot=s: self._on_revert_dispatch(slot))
            else:
                curr_str = f"({disp_name})" if disp_name else "(Unassigned)"
                act_disp = menu.addAction(f"⚡ Appoint Coverage for this Session {curr_str}...")
                act_disp.triggered.connect(lambda _, slot=s, orig=orig_owner: self._on_dispatch_slot(slot, orig))

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

            if self.delegation_plan and self.delegation_plan.members:
                menu.addSeparator()
                for s in slots_here:
                    del_name, is_cov = self._get_slot_effective_delegate(s)
                    disp_name = "All Team" if del_name == "ALL" else (del_name or "")
                    app = self.delegation_plan.appointments.get(s.slot_key)
                    orig_owner = app.original_owner if (app and app.original_owner) else (del_name or "")
                    str_lbl = f"Str {s.stream_number}: " if len(slots_here) > 1 else ""

                    if is_cov:
                        act_disp = menu.addAction(f"⚡ Change Coverage for {str_lbl}this Session (Covered by {disp_name})...")
                        act_disp.triggered.connect(lambda _, slot=s, orig=orig_owner: self._on_dispatch_slot(slot, orig))
                        act_rev = menu.addAction(f"↺ Revert Coverage for {str_lbl}this Session to {orig_owner or 'Topic Owner'}")
                        act_rev.triggered.connect(lambda _, slot=s: self._on_revert_dispatch(slot))
                    else:
                        curr_str = f"({disp_name})" if disp_name else "(Unassigned)"
                        act_disp = menu.addAction(f"⚡ Appoint Coverage for {str_lbl}this Session {curr_str}...")
                        act_disp.triggered.connect(lambda _, slot=s, orig=orig_owner: self._on_dispatch_slot(slot, orig))

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
            delegation_plan=self.delegation_plan,
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
            delegation_plan=self.delegation_plan,
        )
        QMessageBox.information(
            self,
            "Calendar Exported",
            f"Generated .ics calendar file:\n{out_path}\n\nStrictly anchored in UTC so your Outlook and smartphone trigger reminders at the exact venue time.",
        )

    def _export_delegate_ics(self):
        """Exports an .ics calendar file scoped to a specific delegation member."""
        if not self.active_snapshot:
            return
        if not self.delegation_plan or not self.delegation_plan.members:
            QMessageBox.information(self, "No Roster", "Please add delegation members before exporting a delegate calendar.")
            return

        default_name = ""
        if hasattr(self, "tracked_delegate_filter") and self.tracked_delegate_filter.currentIndex() > 0:
            default_name = self.tracked_delegate_filter.currentData()
        if not default_name:
            for name, m in self.delegation_plan.members.items():
                if m.is_self:
                    default_name = name
                    break
        if not default_name:
            default_name = list(self.delegation_plan.members.keys())[0]

        from PyQt5.QtWidgets import QInputDialog
        names = list(self.delegation_plan.members.keys())
        default_idx = names.index(default_name) if default_name in names else 0
        target, ok = QInputDialog.getItem(
            self,
            "Export Delegate Calendar",
            "Select delegate to export .ics calendar for:",
            names,
            default_idx,
            False,
        )
        if not ok or not target:
            return

        out_dir = self.meeting_dir / "Export"
        safe_name = re.sub(r"[^\w\-]", "_", target)
        out_path = out_dir / f"SA2_{self.meeting_number}_{safe_name}.ics"

        SessionPlanExporter.export_ics(
            snapshot=self.active_snapshot,
            output_path=out_path,
            meeting_start_date_str=self.start_date_str,
            meeting_tz_offset_min=self.tz_offset_min,
            only_tracked=False,
            tracked_tokens=self.tracked_tokens,
            annotations=self.annotations,
            delegation_plan=self.delegation_plan,
            target_delegate=target,
        )
        QMessageBox.information(
            self,
            "Delegate Calendar Exported",
            f"Generated personalized .ics calendar file for {target}:\n{out_path}\n\nStrictly anchored in UTC with assigned session topics.",
        )

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
