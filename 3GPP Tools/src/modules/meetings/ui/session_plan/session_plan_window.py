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

from PyQt5.QtCore import QPoint, QSettings, Qt, QTimer
from PyQt5.QtGui import QColor
from PyQt5.QtWidgets import (
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
    QSizePolicy,
    QStackedWidget,
    QTableWidget,
    QTableWidgetItem,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

from core.utils.folder_opener import open_folder
from modules.meetings.core.agenda_manager import AgendaManager
from modules.meetings.core.session_plan.session_plan_db import SessionPlanDatabase
from modules.meetings.core.session_plan.session_plan_diff import SessionPlanDiffEngine
from modules.meetings.core.session_plan.session_plan_exporter import SessionPlanExporter
from modules.meetings.core.session_plan.session_plan_models import (
    DelegationMember,
    DelegationPlan,
    SessionPlanSnapshot,
    SessionSlot,
    UserSessionAnnotation,
    ai_numbers_match,
    build_slot_ai_bullets_html,
    compute_unique_initials,
    expand_ai_number,
    format_slot_display_text,
    get_slot_unassigned_ais,
    is_slot_not_followed,
    resolve_slot_effective_delegates,
)
from modules.meetings.core.session_plan.session_plan_timezone import SessionPlanTimezone
from modules.meetings.ui.delegate_owner_picker import DelegateOwnerPicker
from modules.meetings.ui.session_plan.session_plan_dialogs import (
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
from core.ui.theme import C

logger = logging.getLogger(__name__)

from dataclasses import dataclass
from modules.meetings.core.session_plan.session_plan_parser import (
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
        super().__init__(None)  # Pass None so Win32 does not set parent as GW_OWNER
        self._parent_ref = parent
        self.setWindowFlags(Qt.Window)
        if parent:
            try:
                parent.destroyed.connect(self.close)
            except Exception:
                pass
        self.meeting_dir = Path(meeting_dir)
        self.mtg_info = mtg_info or {}
        self.meeting_number = str(mtg_info.get("meeting_number", "177"))
        self.location = str(mtg_info.get("location", ""))
        self.start_date_str = str(mtg_info.get("start_date", "2026-10-12"))

        # Load local agenda.csv with multiple candidate fallbacks
        self.agenda_map = agenda_map
        if not self.agenda_map:
            candidates = [
                self.meeting_dir / "Agenda",
                self.meeting_dir,
                self.meeting_dir.parent / "Agenda",
            ]
            for cand in candidates:
                if (cand / "agenda.csv").is_file() or (cand / "Agenda.csv").is_file():
                    self.agenda_map = AgendaManager.load_local_agenda(cand)
                    if self.agenda_map:
                        break
        if not self.agenda_map:
            self.agenda_map = AgendaManager.load_local_agenda(self.meeting_dir / "Agenda")

        # Initialize SQLite database under <agenda_dir>/session_plans.db
        agenda_dir = self.meeting_dir if self.meeting_dir.name.lower() == "agenda" else self.meeting_dir / "Agenda"
        self.db = SessionPlanDatabase(agenda_dir / "session_plans.db")

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
            title="Drop Session Plan file or email attachment to import",
            supported_exts=[".docx", ".doc", ".zip", ".htm", ".txt", ".csv"],
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
        if hasattr(self, "grid_table") and hasattr(self, "view_stack") and self.view_stack.currentIndex() == 0:
            QTimer.singleShot(0, self.grid_table.resizeRowsToContents)

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

    def _is_valid_drop_data(self, mime_data) -> bool:
        """Determines if the drag payload contains a session plan file, Outlook attachment, or table text."""
        if not mime_data:
            return False

        # 1. Local URLs (Files dragged from Explorer, Desktop, or temp folder)
        if mime_data.hasUrls():
            for url in mime_data.urls():
                p = Path(url.toLocalFile())
                ext = p.suffix.lower()
                if ext in [".docx", ".doc", ".docm", ".zip", ".htm", ".html", ".csv", ".txt"]:
                    return True
            return False

        # 2. Virtual Outlook attachments (FileGroupDescriptorW / FileContents)
        formats = mime_data.formats()
        for f in formats:
            if "FileGroupDescriptor" in f:
                return True

        # 3. Text or HTML table dragged from email body
        if mime_data.hasText():
            txt = mime_data.text()
            if any(k in txt.lower() for k in ("stream 1", "stream 2", "main room", "breakout 1", "breakout 2", "q1", "q2", "monday")):
                return True

        return False

    def _extract_outlook_attachments(self, mime_data) -> List[Path]:
        """Extracts virtual file attachments dragged directly from Microsoft Outlook."""
        import struct
        formats = mime_data.formats()
        fgd_format = None
        is_unicode = True
        for f in formats:
            if "FileGroupDescriptorW" in f:
                fgd_format = f
                is_unicode = True
                break
            elif "FileGroupDescriptor" in f:
                fgd_format = f
                is_unicode = False
                break

        if not fgd_format:
            return []

        fc_format = None
        for f in formats:
            if "FileContents" in f:
                fc_format = f
                break

        if not fc_format:
            return []

        descriptor_bytes = bytes(mime_data.data(fgd_format))
        if len(descriptor_bytes) < 4:
            return []

        num_files = struct.unpack("<I", descriptor_bytes[:4])[0]
        extracted = []
        contents_bytes = bytes(mime_data.data(fc_format))
        rec_size = 592 if is_unicode else 332
        name_offset = 72

        temp_dir = self.meeting_dir / "Agenda" / "SessionPlans" / "Dropped"
        temp_dir.mkdir(parents=True, exist_ok=True)

        for i in range(min(num_files, 5)):
            rec_start = 4 + i * rec_size
            if rec_start + rec_size > len(descriptor_bytes):
                break

            raw_name = descriptor_bytes[rec_start + name_offset : rec_start + rec_size]
            if is_unicode:
                filename = raw_name.decode("utf-16le", errors="ignore").split("\0")[0].strip()
            else:
                filename = raw_name.decode("mbcs", errors="ignore").split("\0")[0].strip()

            ext = Path(filename).suffix.lower()
            if ext in [".docx", ".doc", ".docm", ".zip", ".htm", ".html", ".csv", ".txt"]:
                target_path = temp_dir / filename
                target_path.write_bytes(contents_bytes)
                extracted.append(target_path)
                break  # Typically 1 attachment dropped at a time

        return extracted

    def _extract_drop_payloads(self, mime_data) -> List[Tuple[str, str]]:
        """
        Extracts imported items as a list of (mode, payload) tuples:
        - ('file', file_path_str) for local files or extracted Outlook attachments
        - ('paste', text_payload) for dragged table text
        """
        results = []

        # 1. Local files
        if mime_data.hasUrls():
            for url in mime_data.urls():
                p = Path(url.toLocalFile())
                if p.suffix.lower() in [".docx", ".doc", ".docm", ".zip", ".htm", ".html", ".csv", ".txt"] and p.is_file():
                    results.append(("file", str(p)))

        # 2. Outlook virtual attachments
        if not results:
            outlook_files = self._extract_outlook_attachments(mime_data)
            for of in outlook_files:
                results.append(("file", str(of)))

        # 3. Dragged email text or HTML table
        if not results and mime_data.hasText():
            txt = mime_data.text().strip()
            if any(k in txt.lower() for k in ("stream 1", "stream 2", "main room", "breakout 1", "breakout 2", "q1", "q2", "monday")):
                results.append(("paste", txt))

        return results

    def dragEnterEvent(self, event):
        if self._is_valid_drop_data(event.mimeData()):
            if hasattr(self, "drop_overlay") and self.drop_overlay:
                self.drop_overlay.setGeometry(self.rect())
                self.drop_overlay.show()
                self.drop_overlay.raise_()
            event.acceptProposedAction()
            return
        event.ignore()

    def dragMoveEvent(self, event):
        if self._is_valid_drop_data(event.mimeData()):
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

        payloads = self._extract_drop_payloads(event.mimeData())
        if payloads:
            event.acceptProposedAction()
            for mode, payload in payloads:
                self._handle_manual_import(mode, payload)
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
        title_lbl.setStyleSheet(f"font-size: 16px; color: {C.TEXT_PRIMARY};")
        title_lbl.setToolTip(f"SA WG2 Meeting #{self.meeting_number} ({self.location})")
        ribbon.addWidget(title_lbl)

        # Snapshot Combobox
        self.snap_combo = QComboBox()
        self.snap_combo.setMinimumWidth(230)
        self.snap_combo.setStyleSheet(f"padding: 4px 8px; font-weight: 600; border-radius: 6px; border: 1px solid {C.BORDER_STRONG};")
        self.snap_combo.setToolTip("Select or switch between imported session plan snapshots and revisions.")
        self.snap_combo.currentIndexChanged.connect(self._on_snapshot_changed)
        ribbon.addWidget(self.snap_combo)

        # Timezone Pill
        self.tz_btn = QPushButton(f"🌍 {self.tz_name}")
        self.tz_btn.setStyleSheet(f"""
            QPushButton {{
                background: {C.BG_CANVAS}; color: {C.TEXT_BODY}; font-size: 11px; font-weight: 600;
                padding: 4px 10px; border: 1px solid {C.BORDER_STRONG}; border-radius: 999px;
            }}
            QPushButton:hover {{ background: {C.BORDER_SUBTLE}; }}
        """)
        self.tz_btn.setToolTip("Meeting Venue Timezone. Click to change or verify against your Austrian corporate laptop clock.")
        self.tz_btn.clicked.connect(self._open_timezone_dialog)
        ribbon.addWidget(self.tz_btn)

        # Real-time Clock Strip
        self.clock_lbl = QLabel()
        self.clock_lbl.setStyleSheet(f"font-size: 11px; color: {C.TEXT_MUTED}; font-weight: 500; margin-left: 8px;")
        self.clock_lbl.setToolTip("Current physical time at the meeting venue and on your Austrian corporate laptop.")
        ribbon.addWidget(self.clock_lbl)

        ribbon.addStretch()

        # Action Buttons
        self.import_btn = QPushButton("📥 Import ▾")
        self.import_btn.setStyleSheet(f"padding: 5px 12px; font-weight: 600; border-radius: 6px; background: {C.ACCENT_ALT}; color: {C.BG_CARD};")
        self.import_btn.setToolTip("Import a session plan document from URL, TDoc#, local file, or clipboard paste, or drag & drop files directly onto this window.")
        import_menu = QMenu(self)
        import_menu.addAction("🔗 Enter URL or TDoc#...", self._import_via_url).setToolTip("Fetch via 3GPP URL or type a TDoc number (e.g. S2-2610195).")
        import_menu.addAction("📂 Browse Local Document...", self._import_via_file).setToolTip("Select a .docx, .doc, .zip, or HTML table from your disk.")
        import_menu.addAction("📋 Paste Markdown/Text Table...", self._import_via_paste).setToolTip("Paste table text directly from an email, TOHRU chat, or notes.")
        self.import_btn.setMenu(import_menu)
        ribbon.addWidget(self.import_btn)

        self.diff_btn = QPushButton("⇄ Compare")
        self.diff_btn.setStyleSheet(f"padding: 5px 10px; font-weight: 500; border-radius: 6px; background: {C.BG_CANVAS}; border: 1px solid {C.BORDER_STRONG};")
        self.diff_btn.setToolTip("Compare the active session plan against an earlier revision to see added, removed, rescheduled, or volume-changed sessions.")
        self.diff_btn.clicked.connect(self._open_diff_dialog)
        ribbon.addWidget(self.diff_btn)

        self.export_btn = QPushButton("📤 Export ▾")
        self.export_btn.setStyleSheet(f"padding: 5px 10px; font-weight: 500; border-radius: 6px; background: {C.BG_CANVAS}; border: 1px solid {C.BORDER_STRONG};")
        self.export_btn.setToolTip("Export to responsive offline Mobile HTML (no PDF table separator issues!), strict UTC .ics Calendar, or Outlook Email summary.")
        export_menu = QMenu(self)
        export_menu.addAction("🌐 Export Interactive HTML (Overview && Mobile)...", self._export_mobile_html).setToolTip("Generate an interactive, responsive standalone HTML file featuring the 5-Day Overview Grid and Mobile Day Cards with delegate toggling.")
        export_menu.addAction("🌐 Open Interactive HTML in Default Browser", self._open_mobile_html_in_browser).setToolTip("Export and instantly open the interactive session plan (Overview Grid & Mobile) in your browser.")
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
            btn.setStyleSheet(f"""
                QPushButton {{
                    padding: 5px 12px; font-size: 12px; font-weight: 600; border: 1px solid {C.BORDER_STRONG};
                    background: {C.BG_CARD}; color: {C.TEXT_SECONDARY}; border-radius: 6px;
                }}
                QPushButton:checked {{
                    background: {C.INFO_ACCENT}; color: {C.BG_CARD}; border-color: {C.INFO_ACCENT};
                }}
            """)
            self.mode_group.addButton(btn, idx)
            sub_layout.addWidget(btn)

        self.btn_grid.setChecked(True)
        self.mode_group.buttonClicked[int].connect(self._on_view_mode_changed)

        # Quick-Jump Conflicts Button
        self.btn_conflicts_quick = QPushButton("⚠️ Conflicts (0)")
        self.btn_conflicts_quick.setCursor(Qt.PointingHandCursor)
        self.btn_conflicts_quick.setStyleSheet(f"""
            QPushButton {{
                padding: 5px 10px; font-size: 12px; font-weight: 600; border: 1px solid {C.DANGER_BORDER};
                background: {C.DANGER_SOFT}; color: {C.DANGER}; border-radius: 6px;
            }}
            QPushButton:hover {{ background: {C.DANGER_SOFT_HOVER}; }}
        """)
        self.btn_conflicts_quick.setToolTip("Jump directly to the Parallel Stream Conflicts & Coverage Dispatch Board.")
        self.btn_conflicts_quick.clicked.connect(self._jump_to_conflicts_board)
        sub_layout.addWidget(self.btn_conflicts_quick)

        # Quick Track Topics manager
        self.btn_manage_topics = QPushButton("⭐ Track Topics...")
        self.btn_manage_topics.setStyleSheet(f"""
            QPushButton {{
                padding: 5px 10px; font-size: 12px; font-weight: 600; border: 1px solid {C.VIOLET_BORDER};
                background: {C.VIOLET_SURFACE}; color: {C.VIOLET_PRIMARY}; border-radius: 6px;
            }}
            QPushButton:hover {{ background: {C.VIOLET_SOFT}; }}
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
        self.search_input.setStyleSheet(f"padding: 5px 10px; border: 1px solid {C.BORDER_STRONG}; border-radius: 6px;")
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
        self.room_strip.setStyleSheet(f"background: {C.BG_SURFACE}; border: 1px solid {C.BORDER_SUBTLE}; border-radius: 6px; padding: 2px 8px;")
        self.room_strip_layout = QHBoxLayout(self.room_strip)
        self.room_strip_layout.setContentsMargins(2, 1, 2, 1)
        self.room_strip_layout.setSpacing(10)

        self.btn_toggle_s1 = QCheckBox("🏢 Str 1: Main Room")
        self.btn_toggle_s1.setChecked(True)
        self.btn_toggle_s1.setStyleSheet(f"QCheckBox {{ color: {C.INDIGO_DARK}; font-weight: 600; font-size: 11px; padding: 1px 4px; }}")
        self.btn_toggle_s1.setToolTip("Toggle visibility of Stream 1 (Main Room / Plenary / Single Stream).")
        self.btn_toggle_s1.toggled.connect(self._update_row_visibility)
        self.room_strip_layout.addWidget(self.btn_toggle_s1)

        self.btn_toggle_s2 = QCheckBox("🏢 Str 2: Breakout 1")
        self.btn_toggle_s2.setChecked(True)
        self.btn_toggle_s2.setStyleSheet(f"QCheckBox {{ color: {C.EMERALD_DARK}; font-weight: 600; font-size: 11px; padding: 1px 4px; }}")
        self.btn_toggle_s2.setToolTip("Toggle visibility of Stream 2 (Breakout 1 / Mike).")
        self.btn_toggle_s2.toggled.connect(self._update_row_visibility)
        self.room_strip_layout.addWidget(self.btn_toggle_s2)

        self.btn_toggle_s3 = QCheckBox("🏢 Str 3: Breakout 2")
        self.btn_toggle_s3.setChecked(True)
        self.btn_toggle_s3.setStyleSheet(f"QCheckBox {{ color: {C.WARNING}; font-weight: 600; font-size: 11px; padding: 1px 4px; }}")
        self.btn_toggle_s3.setToolTip("Toggle visibility of Stream 3 (Breakout 2 / Wanqiang).")
        self.btn_toggle_s3.toggled.connect(self._update_row_visibility)
        self.room_strip_layout.addWidget(self.btn_toggle_s3)

        sep = QFrame()
        sep.setFrameShape(QFrame.VLine)
        sep.setStyleSheet(f"color: {C.BORDER_STRONG};")
        self.room_strip_layout.addWidget(sep)

        self.chk_breaks = QCheckBox("☕ Show Breaks")
        self.chk_breaks.setChecked(True)
        self.chk_breaks.setStyleSheet(f"QCheckBox {{ color: {C.TEXT_SECONDARY}; font-weight: 600; font-size: 11px; padding: 2px 4px; }}")
        self.chk_breaks.setToolTip("Toggle visibility of Coffee Breaks, Lunch Breaks, and Short Breaks.")
        self.chk_breaks.toggled.connect(self._update_row_visibility)
        self.room_strip_layout.addWidget(self.chk_breaks)

        self.chk_hide_empty_breaks = QCheckBox("⚡ Hide Empty Breaks")
        self.chk_hide_empty_breaks.setChecked(True)
        self.chk_hide_empty_breaks.setStyleSheet(f"QCheckBox {{ color: {C.TEXT_MUTED}; font-size: 11px; padding: 2px 4px; }}")
        self.chk_hide_empty_breaks.setToolTip("When checked, automatically hides break rows that have no scheduled drafting sessions or notes.")
        self.chk_hide_empty_breaks.toggled.connect(self._update_row_visibility)
        self.room_strip_layout.addWidget(self.chk_hide_empty_breaks)

        self.room_strip_layout.addStretch()
        page_layout.addWidget(self.room_strip)

        # Delegate Focus / Heatmap Highlight Strip
        self.delegate_strip = QFrame()
        self.delegate_strip.setStyleSheet(f"background: {C.BG_CARD}; border: 1px solid {C.BORDER_SUBTLE}; border-radius: 6px; padding: 2px 8px;")
        self.delegate_strip_layout = QHBoxLayout(self.delegate_strip)
        self.delegate_strip_layout.setContentsMargins(2, 1, 2, 1)
        self.delegate_strip_layout.setSpacing(6)

        lbl_f = QLabel("<b>Delegate Focus:</b>")
        lbl_f.setStyleSheet(f"font-size: 11px; color: {C.TEXT_SECONDARY};")
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

        self.lbl_badge_legend = QLabel(f"<span style='color:{C.TEXT_MUTED};'>ℹ️ Legend</span>")
        self.lbl_badge_legend.setStyleSheet(
            f"font-size: 10.5px; font-weight: 600; background: {C.BG_SURFACE}; "
            f"border: 1px solid {C.BORDER_SUBTLE}; border-radius: 9px; padding: 1px 7px;"
        )
        self.lbl_badge_legend.setToolTip(
            "<b>Cell &amp; Badge Legend</b><br>"
            "• <b>Solid pill</b> = Topic owner<br>"
            "• <b>Dashed pill</b> = Ad-hoc session cover<br>"
            "• <b style='color:#DC2626;'>Red ring / ⚠</b> = Parallel-stream clash<br>"
            "• <b>❓</b> = Unassigned Agenda Item<br>"
            "• <b>⛔</b> = Explicitly Not Followed (Out of scope)"
        )
        self.delegate_strip_layout.addWidget(self.lbl_badge_legend)
        page_layout.addWidget(self.delegate_strip)
        self._style_focus_buttons()

        # Overview Table (10 columns: 2 per day)
        self.grid_table = QTableWidget()
        self.grid_table.setColumnCount(10)
        self.grid_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.grid_table.horizontalHeader().setDefaultAlignment(Qt.AlignCenter)
        self.grid_table.verticalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.grid_table.verticalHeader().setDefaultSectionSize(44)
        self.grid_table.setAlternatingRowColors(False)
        self.grid_table.setStyleSheet(f"QTableWidget {{ gridline-color: {C.BORDER_SUBTLE}; selection-background-color: {C.INDIGO_SOFT}; }}")
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
        self.delegate_agenda_combo.setStyleSheet(f"padding: 3px 8px; font-weight: 600; border: 1px solid {C.BORDER_STRONG}; border-radius: 6px;")
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
        self.team_roster_card.setStyleSheet(f"background: {C.BG_CARD}; border: 1px solid {C.BORDER_SUBTLE}; border-radius: 8px; padding: 6px 12px;")
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
        self.btn_add_member.setStyleSheet(f"padding: 3px 10px; font-size: 11px; font-weight: 600; border: 1px dashed {C.INFO_ACCENT}; color: {C.ACCENT_ALT}; border-radius: 6px; background: {C.INFO_SOFT};")
        self.btn_add_member.clicked.connect(self._on_add_delegation_member)
        rc_lay.addWidget(self.btn_add_member)

        rc_lay.addStretch()

        # Action Buttons
        self.btn_export_team_email = QPushButton("📧 Export Email Summary...")
        self.btn_export_team_email.setStyleSheet(f"padding: 5px 12px; font-weight: 600; font-size: 11px; border-radius: 6px; background: {C.ACCENT_ALT}; color: {C.BG_CARD};")
        self.btn_export_team_email.setToolTip("Generate an executive delegation summary email with color codes, topic distribution, and personal agendas.")
        self.btn_export_team_email.clicked.connect(self._on_export_delegation_email)
        rc_lay.addWidget(self.btn_export_team_email)

        self.btn_copy_team_json = QPushButton("📋 Copy Plan JSON")
        self.btn_copy_team_json.setStyleSheet(f"padding: 5px 10px; font-size: 11px; border-radius: 6px; border: 1px solid {C.BORDER_STRONG}; background: {C.BG_SURFACE};")
        self.btn_copy_team_json.setToolTip("Copy machine-readable JSON representation of this team plan to clipboard.")
        self.btn_copy_team_json.clicked.connect(self._on_copy_delegation_json)
        rc_lay.addWidget(self.btn_copy_team_json)

        self.btn_import_team_plan = QPushButton("📥 Import Team Plan...")
        self.btn_import_team_plan.setStyleSheet(f"padding: 5px 10px; font-size: 11px; font-weight: 600; border-radius: 6px; border: 1px solid {C.BORDER_STRONG}; background: {C.BG_CANVAS};")
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
        tt_filter_box.setSpacing(8)

        self.team_ai_search = QLineEdit()
        self.team_ai_search.setPlaceholderText("🔍 Filter AI number, topic, or owner...")
        self.team_ai_search.setClearButtonEnabled(True)
        self.team_ai_search.setMinimumWidth(260)
        self.team_ai_search.setStyleSheet(f"padding: 5px 10px; border: 1px solid {C.BORDER_STRONG}; border-radius: 6px; font-size: 12px;")
        self.team_ai_search.textChanged.connect(self._filter_topic_ownership_table)
        tt_filter_box.addWidget(self.team_ai_search)

        settings = QSettings("3GPP_Tools", "SessionPlanWindow")
        saved_scheduled = settings.value("team_scheduled_only", False, type=bool)

        self.btn_toggle_scheduled_only = QPushButton("📅 Scheduled Only")
        self.btn_toggle_scheduled_only.setCheckable(True)
        self.btn_toggle_scheduled_only.setChecked(saved_scheduled)
        self.btn_toggle_scheduled_only.setCursor(Qt.PointingHandCursor)
        self.btn_toggle_scheduled_only.setToolTip(
            "Toggle visibility of unscheduled Agenda Items (hide rows with no allocated sessions in the session plan)."
        )
        self.btn_toggle_scheduled_only.toggled.connect(self._on_scheduled_only_toggled)
        self._style_scheduled_toggle(saved_scheduled)
        tt_filter_box.addWidget(self.btn_toggle_scheduled_only)
        self.chk_scheduled_only = self.btn_toggle_scheduled_only

        self.btn_auto_assign_wis = QPushButton("🎯 Auto-Assign Followed WIs...")
        self.btn_auto_assign_wis.setCursor(Qt.PointingHandCursor)
        self.btn_auto_assign_wis.setStyleSheet(
            f"QPushButton {{ padding: 5px 12px; font-size: 11.5px; font-weight: 600; border-radius: 6px; border: 1px solid {C.BORDER_STRONG}; background: {C.BG_SURFACE}; color: {C.TEXT_STRONG}; }}"
            f"QPushButton:hover {{ background: {C.INFO_SOFT}; border-color: {C.INFO_ACCENT}; color: {C.INFO}; }}"
        )
        self.btn_auto_assign_wis.setToolTip("Assign topics matching your followed Work Items watchlist to a delegate.")
        self.btn_auto_assign_wis.clicked.connect(self._auto_assign_followed_wis_to_self)
        tt_filter_box.addWidget(self.btn_auto_assign_wis)

        self.btn_reset_session_overrides = QPushButton("↺ Reset Session Overrides")
        self.btn_reset_session_overrides.setCursor(Qt.PointingHandCursor)
        self.btn_reset_session_overrides.setStyleSheet(
            f"QPushButton {{ padding: 5px 11px; font-size: 11.5px; font-weight: 600; border-radius: 6px; border: 1px solid {C.WARNING_BORDER}; background: {C.WARNING_SOFT}; color: {C.WARNING_DARK}; }}"
            f"QPushButton:hover {{ background: {C.WARNING_SOFT_ALT}; border-color: {C.WARNING_ACCENT}; }}"
        )
        self.btn_reset_session_overrides.setToolTip(
            "Clear all session-level coverage overrides for this meeting while keeping persistent Default AI Assignments intact."
        )
        self.btn_reset_session_overrides.clicked.connect(self._on_reset_all_session_overrides)
        tt_filter_box.addWidget(self.btn_reset_session_overrides)

        tt_filter_box.addStretch()
        lbl_persist_hint = QLabel("📌 <i>Default AI Assignments persist across meetings · Session overrides apply to individual slots</i>")
        lbl_persist_hint.setStyleSheet(f"font-size: 11px; color: {C.TEXT_MUTED};")
        tt_filter_box.addWidget(lbl_persist_hint)
        tt_lay.addLayout(tt_filter_box)

        self.team_topics_table = QTableWidget()
        self.team_topics_table.setColumnCount(5)
        self.team_topics_table.setHorizontalHeaderLabels([
            "AI #", "Topic Description", "📅 Scheduled Slots", "👤 Topic Owner", "🛡️ Coverage Status"
        ])
        h_header = self.team_topics_table.horizontalHeader()
        h_header.setSectionResizeMode(0, QHeaderView.Interactive)
        h_header.setSectionResizeMode(1, QHeaderView.Stretch)
        h_header.setSectionResizeMode(2, QHeaderView.Interactive)
        h_header.setSectionResizeMode(3, QHeaderView.Interactive)
        h_header.setSectionResizeMode(4, QHeaderView.Interactive)

        self.team_topics_table.setColumnWidth(0, 65)    # AI #
        self.team_topics_table.setColumnWidth(2, 220)   # 📅 Scheduled Slots
        self.team_topics_table.setColumnWidth(3, 195)   # 👤 Topic Owner
        self.team_topics_table.setColumnWidth(4, 185)   # 🛡️ Coverage Status
        if self.team_topics_table.horizontalHeaderItem(3):
            self.team_topics_table.horizontalHeaderItem(3).setToolTip(
                "Default AI Assignment — persists across meetings automatically and can be updated between meetings."
            )
        if self.team_topics_table.horizontalHeaderItem(4):
            self.team_topics_table.horizontalHeaderItem(4).setToolTip(
                "Coverage status across scheduled sessions, including any session-level collision overrides."
            )

        self.team_topics_table.verticalHeader().setDefaultSectionSize(36)
        self.team_topics_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.team_topics_table.setAlternatingRowColors(True)
        self.team_topics_table.setStyleSheet(f"""
            QHeaderView::section {{
                background-color: {C.BG_SURFACE};
                color: {C.TEXT_BODY};
                font-weight: 700;
                font-size: 11.5px;
                padding: 6px 8px;
                border: 1px solid {C.BORDER_SUBTLE};
                border-top: none;
                border-left: none;
            }}
            QTableWidget {{
                gridline-color: {C.BG_CANVAS};
                selection-background-color: {C.INFO_SOFT};
                selection-color: {C.TEXT_PRIMARY};
                font-size: 12px;
            }}
            QTableWidget::item {{
                padding: 4px 8px;
            }}
        """)
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
        self.team_clashes_table.setColumnCount(7)
        self.team_clashes_table.setHorizontalHeaderLabels([
            "📅 Slot & TU", "🕒 Time", "👤 Colliding Delegate", "Stream A (Assigned)", "Parallel Stream B", "🟢 Free Colleagues", "🛡️ Action / Coverage"
        ])
        h_clashes = self.team_clashes_table.horizontalHeader()
        h_clashes.setSectionResizeMode(0, QHeaderView.Interactive)
        h_clashes.setSectionResizeMode(1, QHeaderView.Interactive)
        h_clashes.setSectionResizeMode(2, QHeaderView.Interactive)
        h_clashes.setSectionResizeMode(3, QHeaderView.Stretch)
        h_clashes.setSectionResizeMode(4, QHeaderView.Stretch)
        h_clashes.setSectionResizeMode(5, QHeaderView.Interactive)
        h_clashes.setSectionResizeMode(6, QHeaderView.Interactive)

        self.team_clashes_table.setColumnWidth(0, 125)
        self.team_clashes_table.setColumnWidth(1, 145)
        self.team_clashes_table.setColumnWidth(2, 140)
        self.team_clashes_table.setColumnWidth(5, 165)
        self.team_clashes_table.setColumnWidth(6, 185)

        self.team_clashes_table.verticalHeader().setDefaultSectionSize(42)
        self.team_clashes_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.team_clashes_table.setAlternatingRowColors(True)
        self.team_clashes_table.cellDoubleClicked.connect(self._on_clash_cell_double_clicked)
        self.team_clashes_table.setStyleSheet(f"""
            QHeaderView::section {{
                background-color: {C.BG_SURFACE};
                color: {C.TEXT_BODY};
                font-weight: 700;
                font-size: 11.5px;
                padding: 6px 8px;
                border: 1px solid {C.BORDER_SUBTLE};
                border-top: none;
                border-left: none;
            }}
            QTableWidget {{
                gridline-color: {C.BG_CANVAS};
                selection-background-color: {C.INFO_SOFT};
                selection-color: {C.TEXT_PRIMARY};
                font-size: 12px;
            }}
            QTableWidget::item {{
                padding: 4px 8px;
            }}
        """)
        tc_lay.addWidget(self.team_clashes_table)

        tc_lay.addWidget(QLabel("💡 <i>Double-click any row to jump to that slot in the Overview Grid. Click '⚡ Appoint Coverage ▾' to 1-click assign a free colleague or open the dispatch dialog.</i>"))
        self.team_tabs.addTab(tab_clashes, "⚡ Clash && Dispatch Board")

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
        for s_num, chk, fallback_chair in (
            (1, self.btn_toggle_s1, "MCC / Chair"),
            (2, self.btn_toggle_s2, "Convener 1"),
            (3, self.btn_toggle_s3, "Convener 2"),
        ):
            if s_num in streams:
                s_info = streams[s_num]
                short_room = s_info.room_name.split("(")[0].strip() or s_info.room_name
                chk.setText(f"🏢 Str {s_num}: {short_room}")
                chk.setToolTip(
                    f"<b>Stream {s_num}:</b> {html.escape(s_info.room_name)} [{s_info.capacity} seats]<br>"
                    f"<b>Chair:</b> {html.escape(s_info.convener or fallback_chair)}<br>"
                    f"<b>TOHRU:</b> {html.escape(s_info.tohru_room)}"
                )

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
            if it.ai_number:
                for sub_ai in expand_ai_number(it.ai_number, self.agenda_map):
                    if sub_ai.lower() in self.tracked_tokens:
                        return True
        return False

    def _get_period_effective_times(self, period: str) -> Tuple[str, str, str, str, str, str]:
        """Returns (full_s, full_e, tu1_s, tu1_e, tu2_s, tu2_e) from active snapshot slots or fallback PERIOD_TIMES."""
        default_t = PERIOD_TIMES.get(period, ("09:00", "10:30", "09:00", "09:45", "09:45", "10:30"))
        if not self.active_snapshot or not self.active_snapshot.slots:
            return default_t
        p_slots = [s for s in self.active_snapshot.slots if s.quarter == period]
        if not p_slots:
            return default_t
        full_slots = [s for s in p_slots if s.sub_slot == 0 and s.start_time_local and s.end_time_local]
        if full_slots:
            s_str, e_str = full_slots[0].start_time_local, full_slots[0].end_time_local
            return SessionPlanParser._resolve_quarter_times(period, (s_str, e_str))
        tu1_starts = [s.start_time_local for s in p_slots if s.sub_slot == 1 and s.start_time_local]
        tu2_ends = [s.end_time_local for s in p_slots if s.sub_slot == 2 and s.end_time_local]
        s_str = min(tu1_starts) if tu1_starts else default_t[0]
        e_str = max(tu2_ends) if tu2_ends else default_t[1]
        return SessionPlanParser._resolve_quarter_times(period, (s_str, e_str))

    def _build_grid_rows(self) -> List[GridRowDescriptor]:
        """Constructs the 23-row spatial grid descriptors covering quarters and breaks."""
        rows = []
        streams = self.active_snapshot.streams if self.active_snapshot else {}
        for p in PERIOD_ORDER:
            times = self._get_period_effective_times(p)
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

        times = self._get_period_effective_times(period)
        if tu == 1:
            s_t = slot.start_time_local or times[2]
            e_t = slot.end_time_local or times[3]
            half_lbl = "1st Half (TU 1)"
        elif tu == 2:
            s_t = slot.start_time_local or times[4]
            e_t = slot.end_time_local or times[5]
            half_lbl = "2nd Half (TU 2)"
        else:
            s_t = slot.start_time_local or times[0]
            e_t = slot.end_time_local or times[1]
            half_lbl = "Full Quarter"

        time_dual = SessionPlanTimezone.format_dual_time_range(s_t, e_t, self.tz_offset_min)
        if " (" in time_dual:
            parts = time_dual.split(" (", 1)
            time_display = f"{parts[0]} Venue Time ({parts[1]}"
        else:
            time_display = f"{time_dual} (Venue Time)"

        s_info = self.active_snapshot.streams.get(stream_num) if self.active_snapshot else None
        r_name = getattr(slot, "room_override", "") or (s_info.room_name if s_info else f"Stream {stream_num}")
        cap = f" ({s_info.capacity} seats)" if s_info and s_info.capacity else ""
        convener = s_info.convener if s_info else ("SA2 Chair / MCC" if stream_num == 1 else "")
        tohru = s_info.tohru_room if s_info else ""

        meta_parts = [f"🏢 <b>Stream {stream_num}:</b> {html.escape(r_name)}{cap}"]
        if convener:
            meta_parts.append(f"👤 <b>Chair:</b> {html.escape(convener)}")
        if tohru:
            meta_parts.append(f"📡 <b>TOHRU:</b> {html.escape(tohru)}")
        room_line = " &nbsp;|&nbsp; ".join(meta_parts)

        ai_bullets = build_slot_ai_bullets_html(
            slot, agenda_map=self.agenda_map, plan=self.delegation_plan
        )

        tt = f"""<div style='font-family:Segoe UI, sans-serif; font-size:12px; min-width:280px; padding:4px;'>
            <div style='font-weight:700; font-size:13px; color:{C.TEXT_PRIMARY}; margin-bottom:4px;'>
                📅 {day_name} ({date_str}) · {period} · {half_lbl}
            </div>
            <div style='color:{C.TEXT_SECONDARY}; font-size:11.5px; margin-bottom:3px;'>
                🕒 <b>Time:</b> {time_display}
            </div>
            <div style='color:{C.TEXT_SECONDARY}; font-size:11.5px; margin-bottom:6px;'>
                {room_line}
            </div>
            <hr style='border:none; border-top:1px solid {C.BORDER_SUBTLE}; margin:6px 0;'>"""

        raw_note = slot.raw_text.strip() if slot.raw_text else ""
        if raw_note:
            tt += f"""<div style='font-size:10.5px; font-weight:700; color:{C.TEXT_SECONDARY}; text-transform:uppercase; letter-spacing:0.5px; margin-bottom:2px;'>
                📋 Chair's Slot Note:
            </div>
            <div style='font-weight:600; font-size:12px; color:{C.TEXT_PRIMARY}; margin-bottom:6px;'>
                {html.escape(raw_note)}
            </div>"""

        if ai_bullets:
            count_str = f" ({len(ai_bullets)})" if len(ai_bullets) > 1 else ""
            tt += f"""<div style='font-size:10.5px; font-weight:700; color:{C.TEXT_SECONDARY}; text-transform:uppercase; letter-spacing:0.5px; margin-top:6px; margin-bottom:3px;'>
                📑 Agenda Items Covered{count_str}:
            </div>
            <div style='color:{C.TEXT_BODY}; font-size:11px; margin-left:4px; line-height:1.45;'>
                {'<br>'.join(ai_bullets)}
            </div>"""

        if self._is_slot_tracked(slot):
            tt += f"<div style='margin-top:6px; padding:3px 8px; background:{C.VIOLET_SURFACE}; border:1px solid {C.VIOLET_BORDER}; color:{C.VIOLET_DARK}; border-radius:4px; font-weight:600; font-size:11.5px;'>🎯 Tracked in My Topics</div>"

        del_names, is_cov, is_all = self._get_slot_effective_delegates(slot)
        if is_all:
            cov_lbl = " <span style='color:#6366F1; font-weight:normal;'>(Ad-Hoc Coverage)</span>" if is_cov else " <span style='color:#6366F1; font-weight:normal;'>(General / All Team Topic)</span>"
            tt += f"<div style='margin-top:6px; padding:4px 8px; background:{C.WG_SA_BG}; border:1px solid {C.WG_SA_BORDER}; color:{C.WG_SA_TEXT}; border-radius:4px; font-size:11.5px;'>👥 <b>Delegation:</b> <b>All Team</b>{cov_lbl}</div>"
        elif del_names:
            role_lbl = f" <span style='color:{C.TEXT_MUTED}; font-weight:normal;'>(Ad-Hoc Coverage)</span>" if is_cov else f" <span style='color:{C.TEXT_MUTED}; font-weight:normal;'>(Topic Owner)</span>"
            badges_html = []
            for dn in del_names:
                m = self.delegation_plan.members.get(dn) if self.delegation_plan else None
                col = m.color if m else C.ACCENT_ALT
                bg_t = m.bg_tint if m else DelegationMember.compute_tint(col, 0.12)
                bdr_t = m.border_tint if m else DelegationMember.compute_tint(col, 0.35)
                badges_html.append(f"<span style='background:{bg_t}; border:1px solid {bdr_t}; color:{col}; padding:1px 6px; border-radius:3px; font-weight:700;'>● {html.escape(dn)}</span>")
            tt += f"<div style='margin-top:6px; padding:4px 8px; background:{C.BG_SURFACE}; border:1px solid {C.BORDER_SUBTLE}; border-radius:4px; font-size:11.5px;'>👤 <b>Delegation:</b> {' '.join(badges_html)}{role_lbl}</div>"
        elif not del_names and is_slot_not_followed(slot, self.delegation_plan, agenda_map=self.agenda_map):
            tt += f"<div style='margin-top:6px; padding:4px 8px; background:{C.BG_CANVAS}; border:1px solid {C.BORDER_STRONG}; color:{C.TEXT_MUTED}; border-radius:4px; font-size:11.5px;'>⛔ <b>Delegation:</b> <i>Not Followed (Out of Scope)</i></div>"

        ann = self.annotations.get(slot.slot_key)
        if ann and (ann.comment or ann.assigned_colleague):
            note_items = []
            if ann.comment:
                note_items.append(f"“{html.escape(ann.comment)}”")
            if ann.assigned_colleague:
                note_items.append(f"(Assigned: <b>{html.escape(ann.assigned_colleague)}</b>)")
            tt += f"<div style='margin-top:6px; padding:4px 8px; background:{C.WARNING_SOFT_ALT}; border:1px solid {C.WARNING_BORDER}; color:{C.WARNING_DARK}; border-radius:4px; font-size:11.5px;'>📝 <b>Personal Note:</b> {' '.join(note_items)}</div>"

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
            bdr = f"2px {border_style} {C.DANGER_ALT}"
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
        """Creates a compact unified card for a session slot with right-aligned pills and neutral unassigned styling."""
        w = QWidget()
        outer_lay = QVBoxLayout(w)
        outer_lay.setContentsMargins(2, 2, 2, 2)
        outer_lay.setSpacing(0)

        card = QFrame()
        card.setObjectName("slotCard")
        card_lay = QHBoxLayout(card)
        card_lay.setContentsMargins(6, 4, 6, 4)
        card_lay.setSpacing(4)

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
        is_tracked = self._is_slot_tracked(slot)
        is_not_followed_slot = is_slot_not_followed(slot, self.delegation_plan, agenda_map=self.agenda_map)
        room_ov = getattr(slot, "room_override", "") or ""
        clean_text, is_empty_drafting = format_slot_display_text(slot.raw_text, room_ov)
        unassigned_ais = get_slot_unassigned_ais(slot, self.delegation_plan, agenda_map=self.agenda_map)
        has_unassigned_ai = bool(not is_all and not is_not_followed_slot and unassigned_ais)

        # Dimming check if a (multi-select) delegate focus is active on the Overview Grid
        is_dimmed = False
        is_focused = False
        if self.grid_focus_delegates:
            focus_lower = {n.lower() for n in self.grid_focus_delegates}
            if is_all or any(dn.lower() in focus_lower for dn in del_names):
                is_focused = True
            else:
                is_dimmed = True

        # Resolve unified card palette
        font_weight = 400
        if is_dimmed:
            card_bg = C.BG_SURFACE
            card_bdr = f"1px solid {C.BORDER_SUBTLE}"
            text_color = C.TEXT_FAINT
        elif is_collision:
            card_bg = C.WARNING_SOFT
            card_bdr = f"1px solid {C.WARNING_ACCENT}; border-left: 4px solid {C.WARNING_AMBER}"
            text_color = C.WARNING_DARK
            font_weight = 500
        elif is_all:
            text_color = C.TEXT_PRIMARY
            if is_focused:
                card_bg = C.WG_SA_BG
                card_bdr = f"1.5px solid {C.INDIGO_DARK}; border-left: 5px solid {C.INDIGO_DARK}"
                font_weight = 600
            else:
                card_bg = DelegationMember.compute_tint(C.INDIGO_DARK, 0.06)
                card_bdr = f"1px solid {C.BORDER_STRONG}; border-left: 4px solid {C.INDIGO_DARK}"
                font_weight = 500
        elif assigned_members:
            primary_m = assigned_members[0]
            text_color = C.TEXT_PRIMARY
            if is_focused:
                card_bg = primary_m.bg_tint
                card_bdr = f"1.5px solid {primary_m.color}; border-left: 5px solid {primary_m.color}"
                font_weight = 600
            else:
                card_bg = DelegationMember.compute_tint(primary_m.color, 0.07)
                card_bdr = f"1px solid {primary_m.border_tint}; border-left: 4px solid {primary_m.color}"
                font_weight = 500
        elif is_empty_drafting:
            card_bg = C.BG_SURFACE
            card_bdr = f"1px solid {C.BORDER_SUBTLE}"
            text_color = C.TEXT_MUTED
        elif slot.is_drafting:
            card_bg = C.DANGER_SOFT
            card_bdr = f"1px solid {C.DANGER_BORDER}; border-left: 4px solid {C.DANGER_ALT}"
            text_color = C.DANGER_DARK
            font_weight = 500
        elif is_not_followed_slot:
            card_bg = C.BG_SURFACE
            card_bdr = f"1px dashed {C.BORDER_STRONG}"
            text_color = C.TEXT_MUTED
        elif is_tracked:
            card_bg = C.VIOLET_SURFACE
            card_bdr = f"1px solid {C.VIOLET_BORDER}"
            text_color = C.VIOLET_DARK
            font_weight = 500
        else:
            card_bg = C.BG_CARD
            card_bdr = f"1px solid {C.BORDER_SUBTLE}; border-left: 3px solid {C.BORDER_STRONG}"
            text_color = C.TEXT_STRONG

        card.setStyleSheet(f"QFrame#slotCard {{ background: {card_bg}; border: {card_bdr}; border-radius: 5px; }}")

        # Left column: optional room override chip + wrapped topic text
        body_lay = QVBoxLayout()
        body_lay.setContentsMargins(0, 0, 0, 0)
        body_lay.setSpacing(2)

        if room_ov:
            room_chip = QLabel(f"📍 {room_ov}")
            room_chip.setStyleSheet(
                f"background: {C.BG_SURFACE}; color: {C.TEXT_SECONDARY}; font-size: 9.5px; "
                f"font-weight: 700; border: 1px solid {C.BORDER_SUBTLE}; border-radius: 3px; padding: 0px 4px;"
            )
            room_chip.setSizePolicy(QSizePolicy.Maximum, QSizePolicy.Fixed)
            body_lay.addWidget(room_chip, 0, Qt.AlignLeft)

        if clean_text:
            lbl = QLabel(clean_text)
            lbl.setWordWrap(True)
            lbl.setAlignment(Qt.AlignTop | Qt.AlignLeft)
            lbl.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.MinimumExpanding)
            lbl.setStyleSheet(f"background: transparent; border: none; color: {text_color}; font-weight: {font_weight};")
            body_lay.addWidget(lbl)

        body_lay.addStretch(1)
        card_lay.addLayout(body_lay, 1)

        # Right column: compact top-aligned delegate / status pills
        pill_widgets: List[QWidget] = []
        if is_all:
            all_pill = QLabel("👥 ALL")
            all_pill.setToolTip("Whole delegation" + (" (ad-hoc coverage)" if is_cov else " (All Team topic)"))
            border_style = "dashed" if is_cov else "solid"
            all_pill.setStyleSheet(
                f"background: {C.WG_SA_BG}; color: {C.INDIGO_DARK}; font-weight: 700; font-size: 8.5px; "
                f"border: 1px {border_style} #6366F1; border-radius: 7px; padding: 1px 4px;"
            )
            pill_widgets.append(all_pill)
        elif assigned_members:
            initials = self._delegate_initials_map()
            max_badges = 3
            shown = assigned_members if len(assigned_members) <= max_badges else assigned_members[: max_badges - 1]
            for m in shown:
                pill_widgets.append(
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
                    f"background: {C.BG_CANVAS}; color: {C.TEXT_BODY}; font-weight: 700; font-size: 8.5px; "
                    f"border: 1px solid {C.BORDER_STRONG}; border-radius: 7px; padding: 1px 4px;"
                )
                pill_widgets.append(more)
        elif is_not_followed_slot:
            nf_pill = QLabel("⛔")
            nf_pill.setToolTip("Explicitly Not Followed (Out of scope)")
            nf_pill.setStyleSheet(
                f"background: {C.BG_CANVAS}; color: {C.TEXT_MUTED}; font-size: 9px; "
                f"border: 1px solid {C.BORDER_STRONG}; border-radius: 7px; padding: 0px 4px;"
            )
            pill_widgets.append(nf_pill)
        if has_unassigned_ai:
            un_pill = QLabel("❓")
            un_lbl = ", ".join(unassigned_ais)
            un_pill.setToolTip(f"Unassigned Agenda Item(s): {un_lbl} — right-click to assign")
            un_pill.setStyleSheet(
                "background: #FEF3C7; color: #B45309; font-weight: 700; font-size: 9px; "
                "border: 1px dashed #F59E0B; border-radius: 7px; padding: 0px 4px;"
            )
            pill_widgets.append(un_pill)

        if is_collision:
            warn_pill = QLabel("⚠")
            who = ", ".join(colliding_members)
            warn_pill.setToolTip(f"Parallel-stream clash for: {who}")
            warn_pill.setStyleSheet(
                f"background: {C.WARNING_SOFT_ALT}; color: {C.WARNING}; font-weight: bold; font-size: 9.5px; "
                f"border: 1px solid {C.WARNING_ACCENT}; border-radius: 7px; padding: 0px 4px;"
            )
            pill_widgets.append(warn_pill)

        if pill_widgets:
            pill_col = QVBoxLayout()
            pill_col.setContentsMargins(0, 0, 0, 0)
            pill_col.setSpacing(2)
            for pw in pill_widgets:
                pill_col.addWidget(pw, 0, Qt.AlignRight)
            pill_col.addStretch(1)
            card_lay.addLayout(pill_col, 0)

        outer_lay.addWidget(card)

        tooltip_html = self._build_rich_slot_tooltip(slot, day_idx, period, stream_num, tu)
        w.setToolTip(tooltip_html)
        card.setToolTip(tooltip_html)
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
            lbl.setStyleSheet(f"background: {C.DANGER_SOFT}; color: {C.DANGER_DARK}; border: 1px solid {C.DANGER_BORDER}; border-radius: 4px; padding: 4px; font-weight: 600;")
        elif slot.is_plenary:
            lbl.setStyleSheet(f"background: {C.INFO_SOFT}; color: {C.INFO_DARK}; border: 1px solid {C.INFO_BORDER}; border-radius: 4px; padding: 4px; font-weight: 600;")
        else:
            lbl.setStyleSheet(f"background: {C.BG_SURFACE}; color: {C.TEXT_BODY}; border: 1px solid {C.BORDER_SUBTLE}; border-radius: 4px; padding: 4px;")

        lay.addWidget(lbl)

        day_name = DAY_NAMES[day_idx]
        times = PERIOD_TIMES.get(period, ("12:30", "14:00"))
        time_dual = SessionPlanTimezone.format_dual_time_range(times[0], times[1], self.tz_offset_min)
        tt = f"""<div style='font-family:Segoe UI, sans-serif; font-size:12px; min-width:220px; padding:4px;'>
            <div style='font-weight:bold; font-size:13px; color:{C.TEXT_STRONG}; margin-bottom:4px;'>
                ☕ {day_name} · {period}
            </div>
            <div style='color:{C.TEXT_SECONDARY}; margin-bottom:6px;'>
                🕒 <b>Time:</b> {time_dual}
            </div>
            <hr style='border:none; border-top:1px solid {C.BORDER_STRONG}; margin:6px 0;'>
            <div style='font-weight:bold; color:{C.TEXT_PRIMARY};'>
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
        self.grid_table.setRowCount(0)
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
                d_str = f"{d_name[:3]} {cur_date.strftime('%d %b')}"
                if cur_date == venue_today:
                    d_str += " 🟢"
            else:
                d_str = d_name

            h_headers.append(f"{d_str}\n1st Half")
            h_headers.append(f"{d_str}\n2nd Half")

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
                            item_e.setForeground(QColor(C.BORDER_STRONG))
                            self.grid_table.setItem(row_idx, tu1_col, item_e)

                        if tu2_slot:
                            w2 = self._create_slot_cell_widget(tu2_slot, day_idx, r_desc.period, r_desc.stream_number, is_half=True, tu=2)
                            self.grid_table.setCellWidget(row_idx, tu2_col, w2)
                        else:
                            item_e = QTableWidgetItem("—")
                            item_e.setTextAlignment(Qt.AlignCenter)
                            item_e.setForeground(QColor(C.BORDER_STRONG))
                            self.grid_table.setItem(row_idx, tu2_col, item_e)
                    else:
                        self.grid_table.setSpan(row_idx, tu1_col, 1, 2)
                        item_e = QTableWidgetItem("—")
                        item_e.setTextAlignment(Qt.AlignCenter)
                        item_e.setForeground(QColor(C.BORDER_SUBTLE))
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
                        item_b.setForeground(QColor(C.BORDER_STRONG))
                        item_b.setBackground(QColor(C.BG_SURFACE))
                        times = PERIOD_TIMES[r_desc.period]
                        item_b.setToolTip(f"{r_desc.period}: {times[0]} – {times[1]} (Break)")
                        self.grid_table.setItem(row_idx, tu1_col, item_b)

        self._update_row_visibility()
        QTimer.singleShot(0, self.grid_table.resizeRowsToContents)

    @staticmethod
    def _slots_for_ai(ai_num: str, scheduled_ais: Dict[str, List[SessionSlot]]) -> List[SessionSlot]:
        """Returns scheduled slots matching ``ai_num`` directly or via ``.x`` wildcard equivalence."""
        if ai_num in scheduled_ais:
            return scheduled_ais[ai_num]
        matched: List[SessionSlot] = []
        seen_keys: Set[str] = set()
        for sched_ai, slots_list in scheduled_ais.items():
            if ai_numbers_match(ai_num, sched_ai):
                for s in slots_list:
                    if s.slot_key not in seen_keys:
                        seen_keys.add(s.slot_key)
                        matched.append(s)
        return matched

    def _combined_ai_keys(self, scheduled_ais: Dict[str, List[SessionSlot]]) -> List[str]:
        """Combines agenda_map keys with unrepresented scheduled AI numbers (e.g. ``30.x``)."""
        if not self.agenda_map:
            return sorted(scheduled_ais.keys(), key=self._ai_sort_key)
        keys = set(self.agenda_map.keys())
        for sched_ai in scheduled_ais.keys():
            if sched_ai not in keys and not any(ai_numbers_match(sched_ai, ak) for ak in self.agenda_map):
                keys.add(sched_ai)
        return sorted(keys, key=self._ai_sort_key)

    def _resolve_ai_description(self, ai_num: str, scheduled_ais: Dict[str, List[SessionSlot]]) -> Tuple[str, Optional[str]]:
        """Resolves human-readable description and optional acronym for an agenda item number."""
        if self.agenda_map and ai_num in self.agenda_map:
            ai_item = self.agenda_map[ai_num]
            return ai_item.description, getattr(ai_item, "acronym", None)

        # Look for the specific AgendaSlotItem matching this ai_number (or .x wildcard)
        matched_slots = self._slots_for_ai(ai_num, scheduled_ais)
        for s in matched_slots:
            for it in s.items:
                if it.ai_number == ai_num or ai_numbers_match(ai_num, it.ai_number):
                    label = it.topic_label or it.display_expr
                    if label:
                        return label, None

        if matched_slots:
            first_s = matched_slots[0]
            if first_s.items:
                for it in first_s.items:
                    lbl = it.topic_label or it.display_expr
                    if lbl:
                        return lbl, None
            for l in first_s.raw_text.splitlines():
                l_s = l.strip()
                if l_s and not re.search(r"^\s*(?:in\s+)?(main room|breakout\s*\d)\b", l_s, re.I):
                    return l_s, None
            return first_s.raw_text.splitlines()[0].strip(), None

        return f"Agenda Item {ai_num}", None

    def _render_agenda_mode(self):
        """Cross-references Agenda document items against allocated slots with 1-click ⭐ Track."""
        scheduled_ais = self._get_scheduled_ais_map()
        all_ais = self._combined_ai_keys(scheduled_ais)

        num_sched = sum(1 for ai in all_ais if self._slots_for_ai(ai, scheduled_ais))
        num_unsched = len(all_ais) - num_sched
        self.btn_agenda.setText(f"📋 Agenda Mode ({num_sched} Sched · {num_unsched} Unscheduled)")

        self.agenda_table.setRowCount(len(all_ais))

        for row, ai_num in enumerate(all_ais):
            desc, acronym = self._resolve_ai_description(ai_num, scheduled_ais)

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

            matched_slots = self._slots_for_ai(ai_num, scheduled_ais)
            if matched_slots:
                status_item = QTableWidgetItem(f"🟢 Scheduled ({len(matched_slots)})")
                status_item.setForeground(QColor(C.SUCCESS_DARK))
                slot_descs = [f"{s.day_name[:3]} {s.quarter} (Str {s.stream_number})" for s in matched_slots]
                alloc_item = QTableWidgetItem(", ".join(slot_descs))
            else:
                status_item = QTableWidgetItem("🔴 Unscheduled")
                status_item.setForeground(QColor(C.DANGER_DARK))
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
            r_name = getattr(s, "room_override", "") or (s_info.room_name if s_info else f"Stream {s.stream_number}")
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

        # Add Whole Delegation (Master Schedule) first
        self.delegate_agenda_combo.addItem("👥 Whole Delegation (Master Schedule)", "ALL")

        # Add individual delegates
        if self.delegation_plan:
            for name, m in self.delegation_plan.members.items():
                self.delegate_agenda_combo.addItem(f"● {name}", name)

        # Add personal / unassigned starred topics option
        self.delegate_agenda_combo.addItem("⭐ Starred Topics (Personal Watchlist)", "TRACKED")

        # Restore selection
        idx = self.delegate_agenda_combo.findData(cur_selection)
        if idx >= 0:
            self.delegate_agenda_combo.setCurrentIndex(idx)
        else:
            self.delegate_agenda_combo.setCurrentIndex(0)
        self.delegate_agenda_combo.blockSignals(False)

        selected_scope = self.delegate_agenda_combo.currentData() or "ALL"

        # 2. Filter target slots based on scope
        target_slots: List[SessionSlot] = []
        for s in self.active_snapshot.slots:
            del_names, is_cov, is_all = self._get_slot_effective_delegates(s)
            if selected_scope == "ALL":
                if del_names or is_all or self._is_slot_tracked(s):
                    target_slots.append(s)
            elif selected_scope in ("TRACKED", "SELF"):
                if self._is_slot_tracked(s):
                    target_slots.append(s)
            else:
                if is_all or any(dn.lower() == selected_scope.lower() for dn in del_names):
                    target_slots.append(s)

        # 3. Conflict detection: accurate per-delegate & half-quarter (TU1/TU2) time overlap across distinct streams
        slot_conflicts_map: Dict[str, List[Tuple[SessionSlot, List[str]]]] = {}
        for i, s1 in enumerate(target_slots):
            if s1.stream_number <= 0:
                continue
            dels1, _, is_all1 = self._get_slot_effective_delegates(s1)
            set1 = set(self.delegation_plan.members.keys()) if (is_all1 and self.delegation_plan) else set(dels1)

            for j, s2 in enumerate(target_slots):
                if i == j or s2.stream_number <= 0 or s1.stream_number == s2.stream_number:
                    continue
                if not self._slots_time_overlap(s1, s2):
                    continue

                if selected_scope in ("TRACKED", "SELF"):
                    shared_dels = ["Tracked"]
                elif selected_scope != "ALL":
                    shared_dels = [selected_scope]
                else:
                    dels2, _, is_all2 = self._get_slot_effective_delegates(s2)
                    set2 = set(self.delegation_plan.members.keys()) if (is_all2 and self.delegation_plan) else set(dels2)
                    shared_dels = sorted(set1 & set2)

                if shared_dels:
                    slot_conflicts_map.setdefault(s1.slot_key, []).append((s2, shared_dels))

        clash_slot_count = len(slot_conflicts_map)
        if clash_slot_count:
            self.btn_tracked.setText(f"🎯 Delegate Agenda ({len(target_slots)}) ⚠️ {clash_slot_count} Clash(es)")
        else:
            self.btn_tracked.setText(f"🎯 Delegate Agenda ({len(target_slots)})")

        self.btn_manage_topics.setText(f"⭐ Track Topics ({len(self.tracked_tokens)})")
        self.tracked_table.setRowCount(len(target_slots))

        streams = self.active_snapshot.streams
        for row, s in enumerate(target_slots):
            s_info = streams.get(s.stream_number)
            r_name = getattr(s, "room_override", "") or (s_info.room_name if s_info else f"Stream {s.stream_number}")

            time_str = SessionPlanTimezone.format_dual_time_range(
                s.start_time_local, s.end_time_local, self.tz_offset_min
            )

            ann = self.annotations.get(s.slot_key)
            note_str = ann.comment if ann else ""

            conflict_entries = slot_conflicts_map.get(s.slot_key, [])
            is_conflict = bool(conflict_entries)
            note_display = note_str
            if is_conflict:
                other_parts = []
                for o_slot, sh_dels in conflict_entries:
                    tu_lbl = f" TU{o_slot.sub_slot}" if o_slot.sub_slot in (1, 2) else ""
                    if selected_scope == "ALL" and sh_dels and sh_dels != ["Tracked"]:
                        other_parts.append(f"Str {o_slot.stream_number}{tu_lbl} ({', '.join(sh_dels)})")
                    else:
                        other_parts.append(f"Str {o_slot.stream_number}{tu_lbl}")
                conflict_tag = f"⚠️ CONFLICT with {', '.join(other_parts)}"
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
                item_del.setForeground(QColor(C.INDIGO_DARK))
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
                item_del.setForeground(QColor(C.TEXT_PRIMARY))
                f = item_del.font()
                f.setBold(True)
                item_del.setFont(f)

            item_raw = QTableWidgetItem(s.raw_text)
            item_note = QTableWidgetItem(note_display or "—")

            if is_conflict:
                for itm in [item_day, item_q, item_time, item_str, item_del, item_raw, item_note]:
                    itm.setBackground(QColor(C.WARNING_SOFT_ALT))
                item_note.setForeground(QColor(C.WARNING))

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
        return resolve_slot_effective_delegates(slot, self.delegation_plan, agenda_map=self.agenda_map)

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
        active = f"border: 1px solid {C.INFO_ACCENT}; background: {C.INFO_ACCENT}; color: {C.BG_CARD};"
        idle = f"border: 1px solid {C.BORDER_STRONG}; background: {C.BG_CARD}; color: {C.TEXT_BODY};"
        self.btn_chip_all.setStyleSheet(f"QPushButton {{ {base} {active if not self.grid_focus_delegates else idle} }}")

        self_member = self._self_member()
        if self.btn_chip_me is not None:
            only_me = bool(self_member) and self.grid_focus_delegates == {self_member.name}
            col = self_member.color if self_member else C.TEXT_PRIMARY
            style = f"border: 1px solid {col}; background: {col}; color: {C.BG_CARD};" if only_me else idle
            self.btn_chip_me.setStyleSheet(f"QPushButton {{ {base} {style} }}")

        for chk in self.dynamic_delegate_chips:
            name = chk.property("delegate_name")
            m = self.delegation_plan.members.get(name) if self.delegation_plan else None
            col = m.color if m else C.TEXT_BODY
            bg = m.bg_tint if m else C.BG_CARD
            bdr = m.border_tint if m else C.BORDER_STRONG
            if chk.isChecked():
                bg, bdr = (m.compute_tint(col, 0.22) if m else C.BORDER_SUBTLE), col
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
            chip = QPushButton(f"● {name}")
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
        all_ais = self._combined_ai_keys(scheduled_ais)

        self.team_topics_table.setRowCount(len(all_ais))
        collisions = self._detect_delegation_collisions()
        members_list = list(self.delegation_plan.members.values())
        initials_map = self._delegate_initials_map()

        for row, ai_num in enumerate(all_ais):
            desc, _ = self._resolve_ai_description(ai_num, scheduled_ais)

            slots_for_ai = self._slots_for_ai(ai_num, scheduled_ais)
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
            item_ai.setTextAlignment(Qt.AlignCenter)
            item_ai.setToolTip(f"Agenda Item {ai_num}")

            item_desc = QTableWidgetItem(desc)
            item_desc.setTextAlignment(Qt.AlignVCenter | Qt.AlignLeft)
            item_desc.setToolTip(desc)

            item_slots = QTableWidgetItem(slots_summary)
            item_slots.setTextAlignment(Qt.AlignVCenter | Qt.AlignLeft)
            item_slots.setData(Qt.UserRole, len(slots_for_ai))
            item_slots.setToolTip(f"Scheduled slots for AI {ai_num}:\n{slots_summary}")
            if not slots_for_ai:
                item_slots.setForeground(QColor(C.TEXT_FAINT))  # Muted slate for unscheduled
            else:
                item_slots.setForeground(QColor(C.TEXT_PRIMARY))
                font_slots = item_slots.font()
                font_slots.setBold(True)
                item_slots.setFont(font_slots)

            # Checkbox-based multi-select owner picker (replaces combo + custom dialog round-trip)
            picker = DelegateOwnerPicker(members_list, owner_name, initials=initials_map)
            picker.ownerChanged.connect(lambda new_owner, a=ai_num: self._on_topic_owner_changed(a, new_owner))

            override_slots = [
                s for s in slots_for_ai
                if self.delegation_plan and s.slot_key in self.delegation_plan.appointments
            ]
            override_suffix = f" · ⚡ {len(override_slots)} session override(s)" if override_slots else ""

            if has_clash:
                status_txt = f"🔴 Collision in parallel stream{override_suffix}"
                status_col = C.DANGER_DARK
                status_bg = C.DANGER_SOFT
            elif to and to.is_not_followed():
                status_txt = f"⚪ Not Followed{override_suffix}"
                status_col = C.TEXT_MUTED
                status_bg = C.BG_CANVAS
            elif owner_name:
                cov_lbl = "All Team" if to and to.is_all() else (f"{len(owners_list)} delegates" if len(owners_list) > 1 else owner_name)
                status_txt = (
                    f"🟢 Covered ({cov_lbl} · {len(slots_for_ai)} slot(s){override_suffix})"
                    if slots_for_ai
                    else f"⚪ Assigned (Default: {cov_lbl})"
                )
                status_col = C.WG_RAN_TEXT
                status_bg = C.WG_RAN_BG
            elif override_slots:
                status_txt = f"⚡ Covered via {len(override_slots)} session override(s) (No default owner)"
                status_col = C.INFO
                status_bg = C.INFO_SOFT
            else:
                status_txt = "🟡 Unassigned"
                status_col = C.WARNING_DARK
                status_bg = C.WARNING_SOFT

            item_status = QTableWidgetItem(status_txt)
            item_status.setTextAlignment(Qt.AlignVCenter | Qt.AlignLeft)
            item_status.setForeground(QColor(status_col))
            item_status.setBackground(QColor(status_bg))
            item_status.setToolTip(status_txt)
            font_status = item_status.font()
            font_status.setBold(True)
            item_status.setFont(font_status)

            self.team_topics_table.setItem(row, 0, item_ai)
            self.team_topics_table.setItem(row, 1, item_desc)
            self.team_topics_table.setItem(row, 2, item_slots)
            self.team_topics_table.setCellWidget(row, 3, picker)
            self.team_topics_table.setItem(row, 4, item_status)

        # Update button text with dynamic counts
        num_sched = sum(1 for ai in all_ais if scheduled_ais.get(ai))
        total_ais = len(all_ais)
        if hasattr(self, "btn_toggle_scheduled_only"):
            is_chk = self.btn_toggle_scheduled_only.isChecked()
            lbl = f"✓ 📅 Scheduled Only ({num_sched}/{total_ais})" if is_chk else f"📅 Scheduled Only ({num_sched}/{total_ais})"
            self.btn_toggle_scheduled_only.setText(lbl)
            self._style_scheduled_toggle(is_chk)

        if hasattr(self, "btn_reset_session_overrides"):
            app_count = len(self.delegation_plan.appointments) if self.delegation_plan else 0
            self.btn_reset_session_overrides.setText(f"↺ Reset Session Overrides ({app_count})" if app_count else "↺ Reset Session Overrides")
            self.btn_reset_session_overrides.setEnabled(app_count > 0)

        # Apply current filter
        self._filter_topic_ownership_table()

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
            slots_for_ai = self._slots_for_ai(ai_num, scheduled_ais)

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

            override_slots = [
                s for s in slots_for_ai
                if self.delegation_plan and s.slot_key in self.delegation_plan.appointments
            ]
            override_suffix = f" · ⚡ {len(override_slots)} session override(s)" if override_slots else ""

            if has_clash:
                status_txt = f"🔴 Collision in parallel stream{override_suffix}"
                status_col = C.DANGER_DARK
                status_bg = C.DANGER_SOFT
            elif to and to.is_not_followed():
                status_txt = f"⚪ Not Followed{override_suffix}"
                status_col = C.TEXT_MUTED
                status_bg = C.BG_CANVAS
            elif owner_name:
                cov_lbl = "All Team" if to and to.is_all() else (f"{len(owners_list)} delegates" if len(owners_list) > 1 else owner_name)
                status_txt = (
                    f"🟢 Covered ({cov_lbl} · {len(slots_for_ai)} slot(s){override_suffix})"
                    if slots_for_ai
                    else f"⚪ Assigned (Default: {cov_lbl})"
                )
                status_col = C.WG_RAN_TEXT
                status_bg = C.WG_RAN_BG
            elif override_slots:
                status_txt = f"⚡ Covered via {len(override_slots)} session override(s) (No default owner)"
                status_col = C.INFO
                status_bg = C.INFO_SOFT
            else:
                status_txt = "🟡 Unassigned"
                status_col = C.WARNING_DARK
                status_bg = C.WARNING_SOFT

            item_status = self.team_topics_table.item(row, 4)
            if not item_status:
                item_status = QTableWidgetItem(status_txt)
                self.team_topics_table.setItem(row, 4, item_status)
            else:
                item_status.setText(status_txt)
            item_status.setForeground(QColor(status_col))
            item_status.setBackground(QColor(status_bg))

        if hasattr(self, "btn_reset_session_overrides"):
            app_count = len(self.delegation_plan.appointments) if self.delegation_plan else 0
            self.btn_reset_session_overrides.setText(f"↺ Reset Session Overrides ({app_count})" if app_count else "↺ Reset Session Overrides")
            self.btn_reset_session_overrides.setEnabled(app_count > 0)

    def _jump_to_conflicts_board(self):
        """Switches to the Team Allocations view and selects the Clash & Dispatch Board tab."""
        self.btn_team.setChecked(True)
        self.view_stack.setCurrentIndex(4)
        if hasattr(self, "team_tabs"):
            self.team_tabs.setCurrentIndex(1)

    def _on_clash_cell_double_clicked(self, row: int, col: int):
        """Jumps from a conflict row in the Clash & Dispatch Board directly to that slot in the Overview Grid."""
        slot_item = self.team_clashes_table.item(row, 0)
        if not slot_item:
            return
        slot_obj = slot_item.data(Qt.UserRole)
        if not slot_obj:
            return
        day_idx = slot_obj.day_index
        q_name = slot_obj.quarter
        stream_num = slot_obj.stream_number or 1
        target_row = None
        for r_i, r_d in enumerate(self.grid_rows):
            if r_d.period == q_name and (r_d.stream_number == stream_num or r_d.row_type == "break"):
                target_row = r_i
                break
        if target_row is not None:
            self.btn_grid.setChecked(True)
            self.view_stack.setCurrentIndex(0)
            col_idx = day_idx * 2 + (1 if slot_obj.sub_slot == 2 else 0)
            self.grid_table.setCurrentCell(target_row, col_idx)

    def _quick_appoint_colleague(self, slot: SessionSlot, original_owner: str, appointed_delegate: str):
        """1-click direct appointment of an available colleague to resolve a parallel stream collision."""
        self.db.set_session_appointment(
            meeting_number=self.meeting_number,
            slot_key=slot.slot_key,
            appointed_delegate=appointed_delegate,
            original_owner=original_owner,
            reason=f"Covering parallel stream collision for {original_owner}",
        )
        self.delegation_plan = self.db.get_delegation_plan(self.meeting_number)
        self._render_current_views()

    def _on_reset_all_session_overrides(self):
        """Clears all session-level coverage overrides for this meeting while keeping persistent Default AI Assignments."""
        if not self.delegation_plan or not self.delegation_plan.appointments:
            return
        ans = QMessageBox.question(
            self,
            "Reset Session Overrides",
            f"Clear all {len(self.delegation_plan.appointments)} session-level override(s) for SA2 #{self.meeting_number}?\n\nPersistent Default AI Assignments across meetings will remain untouched.",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if ans == QMessageBox.Yes:
            self.db.clear_session_appointments(self.meeting_number)
            self.delegation_plan = self.db.get_delegation_plan(self.meeting_number)
            self._render_current_views()

    def _render_team_clashes(self):
        """Renders the Clash & Dispatch Board with multi-stream choices, free colleague detection, and 1-click actions."""
        if not self.delegation_plan:
            return

        collisions = self._detect_delegation_collisions()
        clash_count = len(collisions)
        self.btn_team.setText(f"👥 Team Allocations ({clash_count} ⚠️)" if clash_count else "👥 Team Allocations (Clean)")
        self.team_tabs.setTabText(1, f"⚡ Clash && Dispatch Board ({clash_count})")
        if hasattr(self, "btn_conflicts_quick"):
            self.btn_conflicts_quick.setText(f"⚠️ Conflicts ({clash_count})")
            if clash_count > 0:
                self.btn_conflicts_quick.setStyleSheet(
                    f"QPushButton {{ padding: 5px 10px; font-size: 12px; font-weight: 700; border: 1px solid {C.DANGER_BORDER_ALT}; background: {C.DANGER_SOFT}; color: {C.DANGER}; border-radius: 6px; }}"
                    f"QPushButton:hover {{ background: {C.DANGER_SOFT_HOVER}; }}"
                )
            else:
                self.btn_conflicts_quick.setStyleSheet(
                    f"QPushButton {{ padding: 5px 10px; font-size: 12px; font-weight: 600; border: 1px solid {C.WG_RAN_BORDER}; background: {C.WG_RAN_BG}; color: {C.WG_RAN_TEXT}; border-radius: 6px; }}"
                    "QPushButton:hover { background: #D1FAE5; }"
                )

        if clash_count == 0:
            self.clash_kpi_lbl.setText("✅ No collisions detected! All parallel streams have distinct delegates or single coverage.")
            self.clash_kpi_lbl.setStyleSheet(
                f"background: {C.WG_RAN_BG}; color: {C.WG_RAN_TEXT}; border: 1px solid {C.WG_RAN_BORDER}; font-size: 12px; font-weight: 600; padding: 6px 10px; border-radius: 6px;"
            )
            self.team_clashes_table.setRowCount(0)
        else:
            self.clash_kpi_lbl.setText(
                f"⚠️ {clash_count} Double-Booking Collision(s) Detected! Appoint an available colleague to cover one of the parallel rooms."
            )
            self.clash_kpi_lbl.setStyleSheet(
                f"background: {C.WARNING_SOFT}; color: {C.WARNING_DARK}; border: 1px solid {C.WARNING_BORDER}; font-size: 12px; font-weight: 600; padding: 6px 10px; border-radius: 6px;"
            )
            self.team_clashes_table.setRowCount(clash_count)

            for row, ((del_name, d_idx, q_name), slots) in enumerate(collisions.items()):
                s_a = slots[0]
                time_str = SessionPlanTimezone.format_dual_time_range(
                    s_a.start_time_local, s_a.end_time_local, self.tz_offset_min
                )

                # Determine overlapping TUs
                has_tu1 = all(s.sub_slot in (0, 1) for s in slots) or any(
                    s1.sub_slot in (0, 1) and s2.sub_slot in (0, 1) and s1.stream_number != s2.stream_number
                    for s1 in slots for s2 in slots
                )
                has_tu2 = all(s.sub_slot in (0, 2) for s in slots) or any(
                    s1.sub_slot in (0, 2) and s2.sub_slot in (0, 2) and s1.stream_number != s2.stream_number
                    for s1 in slots for s2 in slots
                )
                tu_tag = " (TU 1&2)" if (has_tu1 and has_tu2) else (" (TU 1)" if has_tu1 else (" (TU 2)" if has_tu2 else ""))

                item_slot = QTableWidgetItem(f"{s_a.day_name} {s_a.quarter}{tu_tag}")
                item_slot.setData(Qt.UserRole, s_a)
                item_slot.setToolTip("Double-click to jump to this slot in the Overview Grid.")
                item_time = QTableWidgetItem(time_str)

                item_del = QTableWidgetItem(f"● {del_name}")
                m = self.delegation_plan.members.get(del_name)
                if m:
                    item_del.setForeground(QColor(m.color))
                f = item_del.font()
                f.setBold(True)
                item_del.setFont(f)

                def _format_clash_slot_desc(slot_obj: SessionSlot) -> str:
                    s_inf = self.active_snapshot.streams.get(slot_obj.stream_number) if self.active_snapshot else None
                    rm = getattr(slot_obj, "room_override", "") or (s_inf.room_name if s_inf else f"Stream {slot_obj.stream_number}")
                    _, is_cov_flag, is_all_flag = self._get_slot_effective_delegates(slot_obj)
                    role_lbl = "[All Team]" if is_all_flag else ("[Override]" if is_cov_flag else "[Owner]")
                    tu_sub = f" [TU {slot_obj.sub_slot}]" if slot_obj.sub_slot in (1, 2) else ""
                    return f"Str {slot_obj.stream_number}{tu_sub} ({rm}) {role_lbl}: {slot_obj.raw_text}"

                item_stra = QTableWidgetItem(_format_clash_slot_desc(s_a))
                item_stra.setToolTip(_format_clash_slot_desc(s_a))

                other_desc = [_format_clash_slot_desc(other_s) for other_s in slots[1:]]
                strb_txt = "\n".join(other_desc) if other_desc else _format_clash_slot_desc(slots[-1])
                item_strb = QTableWidgetItem(strb_txt)
                item_strb.setToolTip(strb_txt)

                # Compute free colleagues for this collision
                avail_map = self._compute_delegate_availability(s_a)
                free_colleagues = [
                    c_name for c_name, c_info in avail_map.items()
                    if c_info.get("status") == "free" and c_name.lower() != del_name.lower()
                ]
                if free_colleagues:
                    free_txt = "🟢 " + ", ".join(free_colleagues)
                    free_col = C.WG_RAN_TEXT
                else:
                    free_txt = "⚠️ All busy"
                    free_col = C.DANGER_DARK
                item_free = QTableWidgetItem(free_txt)
                item_free.setForeground(QColor(free_col))
                item_free.setToolTip(
                    f"Available colleagues during {s_a.day_name} {s_a.quarter}: {', '.join(free_colleagues)}"
                    if free_colleagues
                    else "All delegation members are currently assigned to parallel streams in this slot."
                )

                btn_appoint = QPushButton("⚡ Appoint Coverage ▾")
                btn_appoint.setStyleSheet(
                    f"padding: 3px 8px; font-size: 11px; font-weight: 600; border-radius: 6px; background: {C.EMERALD_PRIMARY}; color: {C.BG_CARD};"
                )
                menu = QMenu(btn_appoint)
                if free_colleagues:
                    for fc in free_colleagues:
                        for s_opt in slots:
                            act_q = menu.addAction(f"⚡ Quick-Appoint {fc} (🟢 Free) → Cover Str {s_opt.stream_number}")
                            act_q.triggered.connect(
                                lambda _, slot=s_opt, orig=del_name, target=fc: self._quick_appoint_colleague(slot, orig, target)
                            )
                    menu.addSeparator()

                for s_opt in slots:
                    s_inf = self.active_snapshot.streams.get(s_opt.stream_number) if self.active_snapshot else None
                    rm = getattr(s_opt, "room_override", "") or (s_inf.room_name if s_inf else f"Stream {s_opt.stream_number}")
                    act = menu.addAction(f"🔧 Dispatch Dialog: Cover Str {s_opt.stream_number} ({rm}): {s_opt.raw_text[:28]}...")
                    act.triggered.connect(lambda _, slot=s_opt, orig=del_name: self._on_dispatch_slot(slot, orig))

                menu.addSeparator()
                act_jump = menu.addAction(f"🔍 Show {s_a.day_name} {s_a.quarter} in Overview Grid")
                act_jump.triggered.connect(lambda _, r=row: self._on_clash_cell_double_clicked(r, 0))
                btn_appoint.setMenu(menu)

                self.team_clashes_table.setItem(row, 0, item_slot)
                self.team_clashes_table.setItem(row, 1, item_time)
                self.team_clashes_table.setItem(row, 2, item_del)
                self.team_clashes_table.setItem(row, 3, item_stra)
                self.team_clashes_table.setItem(row, 4, item_strb)
                self.team_clashes_table.setItem(row, 5, item_free)
                self.team_clashes_table.setCellWidget(row, 6, btn_appoint)

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

    def _style_scheduled_toggle(self, checked: bool):
        base = "padding: 5px 12px; font-size: 11.5px; font-weight: 600; border-radius: 6px;"
        if checked:
            self.btn_toggle_scheduled_only.setStyleSheet(
                f"QPushButton {{ {base} background: {C.INFO_SOFT}; border: 1.5px solid {C.ACCENT_ALT}; color: {C.INFO}; }}"
                f"QPushButton:hover {{ background: {C.INFO_SOFT_STRONG}; }}"
            )
        else:
            self.btn_toggle_scheduled_only.setStyleSheet(
                f"QPushButton {{ {base} background: {C.BG_SURFACE}; border: 1px solid {C.BORDER_STRONG}; color: {C.TEXT_SECONDARY}; }}"
                f"QPushButton:hover {{ background: {C.BG_CANVAS}; color: {C.TEXT_STRONG}; border-color: {C.TEXT_FAINT}; }}"
            )

    def _on_scheduled_only_toggled(self, checked: bool):
        settings = QSettings("3GPP_Tools", "SessionPlanWindow")
        settings.setValue("team_scheduled_only", checked)
        if hasattr(self, "btn_toggle_scheduled_only"):
            total_ais = self.team_topics_table.rowCount()
            num_sched = sum(
                1 for r in range(total_ais)
                if (self.team_topics_table.item(r, 2) and (self.team_topics_table.item(r, 2).data(Qt.UserRole) or 0) > 0)
            )
            lbl = f"✓ 📅 Scheduled Only ({num_sched}/{total_ais})" if checked else f"📅 Scheduled Only ({num_sched}/{total_ais})"
            self.btn_toggle_scheduled_only.setText(lbl)
            self._style_scheduled_toggle(checked)
        self._filter_topic_ownership_table()

    def _filter_topic_ownership_table(self, query: Optional[str] = None):
        q = (self.team_ai_search.text() if hasattr(self, "team_ai_search") else (query or "")).lower().strip()
        scheduled_only = self.btn_toggle_scheduled_only.isChecked() if hasattr(self, "btn_toggle_scheduled_only") else False

        for r in range(self.team_topics_table.rowCount()):
            ai_item = self.team_topics_table.item(r, 0)
            desc_item = self.team_topics_table.item(r, 1)
            slots_item = self.team_topics_table.item(r, 2)
            picker = self.team_topics_table.cellWidget(r, 3)

            slot_count = slots_item.data(Qt.UserRole) if slots_item else 0
            if slot_count is None:
                slot_count = 0 if (not slots_item or "(not scheduled)" in slots_item.text().lower() or "—" in slots_item.text()) else 1

            if scheduled_only and slot_count == 0:
                self.team_topics_table.setRowHidden(r, True)
                continue

            owner_txt = ""
            if picker:
                if hasattr(picker, "text"):
                    owner_txt = picker.text()
                elif hasattr(picker, "currentText"):
                    owner_txt = picker.currentText()

            txt = f"{ai_item.text()} {desc_item.text()} {owner_txt}".lower() if ai_item and desc_item else ""
            self.team_topics_table.setRowHidden(r, bool(q and q not in txt))

    def _auto_assign_followed_wis_to_self(self):
        if not self.delegation_plan or not self.delegation_plan.members:
            QMessageBox.warning(self, "No Roster", "Please add delegation members before assigning topics.")
            return

        names = list(self.delegation_plan.members.keys())
        if len(names) == 1:
            target_name = names[0]
        else:
            from PyQt5.QtWidgets import QInputDialog
            target_name, ok = QInputDialog.getItem(
                self,
                "Auto-Assign Followed WIs",
                "Select delegate to assign followed Work Item topics to:",
                names,
                0,
                False,
            )
            if not ok or not target_name:
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
                    self.db.set_topic_owner(self.meeting_number, ai_num, target_name)
                    count += 1

        self.delegation_plan = self.db.get_delegation_plan(self.meeting_number)
        self._render_current_views()
        QMessageBox.information(self, "Auto-Assignment Complete", f"Assigned {count} topic(s) to {target_name}.")

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
            self.clock_lbl.setText(f"🕒 <b>{venue_str}</b> (Venue &amp; Austrian Time)")
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
                    if it.ai_number:
                        expanded_ais = expand_ai_number(it.ai_number, self.agenda_map)
                        for target_ai in expanded_ais[:12]:
                            ai_obj = self.agenda_map.get(target_ai) if self.agenda_map else None
                            raw_lbl = getattr(ai_obj, "description", "") if ai_obj else (it.topic_label or f"AI {target_ai}")
                            label = f"AI {target_ai}: {raw_lbl}" if (len(expanded_ais) > 1 or target_ai != it.ai_number) else (it.topic_label or f"AI {target_ai}")
                            is_tr = target_ai.lower() in self.tracked_tokens
                            act_tr = menu.addAction(f"{'⭐ Untrack' if is_tr else '⭐ Track'}: {label[:44]}")
                            act_tr.triggered.connect(lambda _, tok=target_ai, lbl=raw_lbl: self._toggle_track_ai(tok, lbl))
                    else:
                        token = it.topic_label
                        if not token:
                            continue
                        is_tr = token.lower() in self.tracked_tokens
                        act_tr = menu.addAction(f"{'⭐ Untrack' if is_tr else '⭐ Track'}: {token[:40]}")
                        act_tr.triggered.connect(lambda _, tok=token, lbl=token: self._toggle_track_ai(tok, lbl))

            if self.delegation_plan and self.delegation_plan.members:
                menu.addSeparator()

                # Direct Topic Owner Assignment Submenu for each AI covered (expanding .x wildcards)
                for s in slots_here:
                    for it in s.items:
                        if not it.ai_number:
                            continue
                        expanded_ais = expand_ai_number(it.ai_number, self.agenda_map)
                        for target_ai in expanded_ais[:12]:
                            current_to = self.delegation_plan.topic_owners.get(target_ai)
                            curr_owner_str = current_to.owner_name if current_to else ""
                            curr_disp = " (Not Followed)" if (current_to and current_to.is_not_followed()) else (f" ({curr_owner_str})" if curr_owner_str else " (Unassigned)")
                            ai_menu = menu.addMenu(f"👤 Assign Topic Owner (AI {target_ai}){curr_disp}...")

                            act_all = ai_menu.addAction("👥 All Team")
                            act_all.triggered.connect(lambda _, ai=target_ai: self._set_topic_owner_direct(ai, "ALL"))

                            ai_menu.addSeparator()
                            for m_name in self.delegation_plan.members.keys():
                                is_curr = (curr_owner_str == m_name)
                                prefix = "✓ " if is_curr else "● "
                                act_m = ai_menu.addAction(f"{prefix}{m_name}")
                                act_m.triggered.connect(lambda _, ai=target_ai, n=m_name: self._set_topic_owner_direct(ai, n))

                            ai_menu.addSeparator()
                            act_multi = ai_menu.addAction("⚡ Custom Multi-Delegate...")
                            act_multi.triggered.connect(lambda _, ai=target_ai, cur=curr_owner_str: self._open_multi_delegate_dialog(ai, cur))

                            ai_menu.addSeparator()
                            act_not_followed = ai_menu.addAction("⛔ Mark as Not Followed")
                            act_not_followed.triggered.connect(lambda _, ai=target_ai: self._set_topic_owner_direct(ai, "NOT_FOLLOWED"))

                            if curr_owner_str:
                                act_clear = ai_menu.addAction("✕ Clear (Unassign)")
                                act_clear.triggered.connect(lambda _, ai=target_ai: self._set_topic_owner_direct(ai, ""))

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

    def _set_topic_owner_direct(self, ai_number: str, owner_name: str):
        """Directly updates topic owner from Overview Grid context menu and refreshes views."""
        self.db.set_topic_owner(self.meeting_number, ai_number, owner_name)
        self.delegation_plan = self.db.get_delegation_plan(self.meeting_number)
        self._render_current_views()

    def _open_multi_delegate_dialog(self, ai_number: str, current_owner: str):
        """Opens MultiDelegateSelectDialog for compound multi-delegate ownership from grid."""
        from modules.meetings.ui.session_plan.session_plan_dialogs import MultiDelegateSelectDialog
        ai_desc = ""
        if self.agenda_map and ai_number in self.agenda_map:
            ai_desc = getattr(self.agenda_map[ai_number], "description", "")
        title = f"AI {ai_number}: {ai_desc}" if ai_desc else f"AI {ai_number}"
        members_list = list(self.delegation_plan.members.values()) if self.delegation_plan else []
        dlg = MultiDelegateSelectDialog(title, members_list, current_owner, parent=self)
        if dlg.exec_() == QDialog.Accepted:
            new_owner = dlg.get_owner_string()
            self._set_topic_owner_direct(ai_number, new_owner)

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

                # If dropped or selected file is a zip archive, unpack it first
                if ext == ".zip":
                    import zipfile
                    extract_dir = self.meeting_dir / "Agenda" / "SessionPlans" / "Extracted" / p.stem
                    extract_dir.mkdir(parents=True, exist_ok=True)
                    with zipfile.ZipFile(p, "r") as zf:
                        zf.extractall(extract_dir)
                    doc_candidates = list(extract_dir.glob("*.docx")) or list(extract_dir.glob("*.doc*")) or list(extract_dir.glob("*.htm*"))
                    if doc_candidates:
                        p = doc_candidates[0]
                        ext = p.suffix.lower()

                staged = SessionPlanParser.stage_session_plan_file(p, self.meeting_dir)
                if ext in (".docx", ".docm"):
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
                from modules.meetings.core.tdocs.tdoc_file_handler import TDocFileHandler
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
            meeting_dates_str=self.start_date_str,
            annotations=self.annotations,
            tracked_tokens=self.tracked_tokens,
            delegation_plan=self.delegation_plan,
            agenda_map=self.agenda_map,
        )
        QMessageBox.information(
            self,
            "Interactive HTML Exported",
            f"Saved interactive Session Plan (5-Day Overview Grid + Mobile Day Cards) to:\n{out_path}\n\nFeatures interactive delegate toggling and spatial overview for easy sharing with colleagues.",
        )
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
            agenda_map=self.agenda_map,
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
            agenda_map=self.agenda_map,
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
