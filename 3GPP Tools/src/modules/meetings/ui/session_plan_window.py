# --- File: src/modules/meetings/ui/session_plan_window.py ---
"""
Primary Desktop GUI for 3GPP SA2 Session Plan Management.
Provides Overview Grid, Agenda Mode (Gap Analysis), Drafting Radar, My Topics,
Timezone Synchronization, Diffing, and Exporting.
"""
import logging
import re
from datetime import datetime, timezone
from pathlib import Path
from typing import Dict, List, Optional, Set

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
)

logger = logging.getLogger(__name__)

DAY_NAMES = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
QUARTER_ORDER = ["Q0", "Q1", "Q2", "Lunch", "Q3", "Q4", "Q5"]


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
        self.agenda_map = agenda_map or AgendaManager.load_local_agenda(self.meeting_dir / "Agenda")
        self.tracked_tokens = {t.lower() for t in (tracked_tokens or set())}

        # Initialize SQLite database under <meeting_dir>/Agenda/session_plans.db
        self.db = SessionPlanDatabase(self.meeting_dir / "Agenda" / "session_plans.db")

        # Timezone state
        self.tz_offset_min = 120  # Default +2 hours (Europe summer/Prague)
        self.tz_name = "Europe/Prague (UTC+02:00)"
        self._init_timezone()

        self.active_snapshot: Optional[SessionPlanSnapshot] = None
        self.annotations: Dict[str, UserSessionAnnotation] = {}

        self.setWindowTitle(f"🗓️ Session Plan — SA2 #{self.meeting_number} ({self.location})")
        self.resize(1360, 820)

        self._setup_ui()
        self._load_data()

        # Real-time clock timer
        self.clock_timer = QTimer(self)
        self.clock_timer.timeout.connect(self._update_live_clock)
        self.clock_timer.start(1000)
        self._update_live_clock()

    def _init_timezone(self):
        """Initializes venue timezone from DB metadata, Agenda deadlines, or city lookup."""
        meta = self.db.get_meeting_meta(self.meeting_number)
        if meta and meta.get("tz_offset_minutes") is not None:
            self.tz_offset_min = meta["tz_offset_minutes"]
            self.tz_name = meta.get("tz_name") or f"UTC{self.tz_offset_min//60:+03d}:00"
            return

        # Attempt auto-detection from city
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

        # Title / Meeting badge
        title_lbl = QLabel(f"<b>SA2 #{self.meeting_number}</b>")
        title_lbl.setStyleSheet("font-size: 16px; color: #0F172A;")
        ribbon.addWidget(title_lbl)

        # Snapshot Combobox
        self.snap_combo = QComboBox()
        self.snap_combo.setMinimumWidth(220)
        self.snap_combo.setStyleSheet("padding: 4px 8px; font-weight: 600; border-radius: 6px; border: 1px solid #CBD5E1;")
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
        self.tz_btn.setToolTip("Click to adjust Venue Timezone vs Austrian Corporate Laptop Clock")
        self.tz_btn.clicked.connect(self._open_timezone_dialog)
        ribbon.addWidget(self.tz_btn)

        # Real-time Clock Strip
        self.clock_lbl = QLabel()
        self.clock_lbl.setStyleSheet("font-size: 11px; color: #64748B; font-weight: 500; margin-left: 8px;")
        ribbon.addWidget(self.clock_lbl)

        ribbon.addStretch()

        # Action Buttons
        self.import_btn = QPushButton("📥 Import ▾")
        self.import_btn.setStyleSheet("padding: 5px 12px; font-weight: 600; border-radius: 6px; background: #2563EB; color: #FFFFFF;")
        import_menu = QMenu(self)
        import_menu.addAction("🔗 Enter URL / TDoc#...", self._import_via_url)
        import_menu.addAction("📂 Browse Local File...", self._import_via_file)
        import_menu.addAction("📋 Paste Markdown/Text Table...", self._import_via_paste)
        self.import_btn.setMenu(import_menu)
        ribbon.addWidget(self.import_btn)

        self.diff_btn = QPushButton("⇄ Compare")
        self.diff_btn.setStyleSheet("padding: 5px 10px; font-weight: 500; border-radius: 6px; background: #F1F5F9; border: 1px solid #CBD5E1;")
        self.diff_btn.clicked.connect(self._open_diff_dialog)
        ribbon.addWidget(self.diff_btn)

        self.export_btn = QPushButton("📤 Export ▾")
        self.export_btn.setStyleSheet("padding: 5px 10px; font-weight: 500; border-radius: 6px; background: #F1F5F9; border: 1px solid #CBD5E1;")
        export_menu = QMenu(self)
        export_menu.addAction("📱 Export Mobile HTML (Offline)...", self._export_mobile_html)
        export_menu.addAction("📅 Export iCalendar (.ics)...", self._export_ics)
        export_menu.addAction("✉️ Copy Email Summary for Colleagues", self._copy_email_summary)
        export_menu.addSeparator()
        export_menu.addAction("📂 Open Session Plans Folder", self._open_plans_folder)
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
        self.btn_agenda = QPushButton("📋 Agenda Mode (Gap Check)")
        self.btn_drafting = QPushButton("🛠️ Drafting Radar (0)")
        self.btn_tracked = QPushButton("🎯 My Topics (0)")

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

        sub_layout.addStretch()

        # Search Bar
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("🔍 Filter AI, topic, convener, room...")
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

        self.grid_table = QTableWidget()
        self.grid_table.setColumnCount(5)
        self.grid_table.setHorizontalHeaderLabels(DAY_NAMES)
        self.grid_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.grid_table.verticalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.grid_table.verticalHeader().setDefaultSectionSize(95)
        self.grid_table.setAlternatingRowColors(True)
        self.grid_table.cellDoubleClicked.connect(self._on_grid_cell_double_clicked)
        page_layout.addWidget(self.grid_table)

        self.view_stack.addWidget(page)

    # --- View 2: Agenda Mode ---

    def _setup_agenda_view(self):
        page = QWidget()
        page_layout = QVBoxLayout(page)
        page_layout.setContentsMargins(0, 0, 0, 0)

        self.agenda_table = QTableWidget()
        self.agenda_table.setColumnCount(5)
        self.agenda_table.setHorizontalHeaderLabels([
            "AI #", "Topic Description", "Status", "Allocated Quarters & Rooms", "Actions"
        ])
        self.agenda_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.agenda_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self.agenda_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.agenda_table.setAlternatingRowColors(True)
        page_layout.addWidget(self.agenda_table)

        self.view_stack.addWidget(page)

    # --- View 3: Drafting Radar ---

    def _setup_drafting_view(self):
        page = QWidget()
        page_layout = QVBoxLayout(page)
        page_layout.setContentsMargins(0, 0, 0, 0)

        self.drafting_table = QTableWidget()
        self.drafting_table.setColumnCount(6)
        self.drafting_table.setHorizontalHeaderLabels([
            "Day", "Quarter", "Time (Venue / Vienna)", "Stream / Room", "Convener", "Drafting Topic & AIs"
        ])
        self.drafting_table.horizontalHeader().setSectionResizeMode(5, QHeaderView.Stretch)
        self.drafting_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.drafting_table.setAlternatingRowColors(True)
        page_layout.addWidget(self.drafting_table)

        self.view_stack.addWidget(page)

    # --- View 4: My Topics ---

    def _setup_tracked_view(self):
        page = QWidget()
        page_layout = QVBoxLayout(page)
        page_layout.setContentsMargins(0, 0, 0, 0)

        self.tracked_table = QTableWidget()
        self.tracked_table.setColumnCount(6)
        self.tracked_table.setHorizontalHeaderLabels([
            "Day", "Quarter", "Time (Venue / Vienna)", "Stream / Room", "Matched Topic", "My Notes / Priority"
        ])
        self.tracked_table.horizontalHeader().setSectionResizeMode(4, QHeaderView.Stretch)
        self.tracked_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.tracked_table.setAlternatingRowColors(True)
        page_layout.addWidget(self.tracked_table)

        self.view_stack.addWidget(page)

    # --- Data Loading & Rendering ---

    def _load_data(self):
        """Loads snapshots and annotations from the SQLite database."""
        self.annotations = self.db.get_annotations(self.meeting_number)
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

        self._render_overview_grid()
        self._render_agenda_mode()
        self._render_drafting_radar()
        self._render_tracked_topics()

    def _render_empty_views(self):
        self.grid_table.setRowCount(0)
        self.agenda_table.setRowCount(0)
        self.drafting_table.setRowCount(0)
        self.tracked_table.setRowCount(0)

    def _render_overview_grid(self):
        """Populates the 5-day Overview Grid."""
        self.grid_table.setRowCount(len(QUARTER_ORDER))
        self.grid_table.setVerticalHeaderLabels(QUARTER_ORDER)

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

                # Build multi-stream cell text
                cell_lines = []
                is_cell_drafting = False
                is_cell_tracked = False

                for s in sorted(slots_in_cell, key=lambda x: (x.sub_slot, x.stream_number)):
                    sub_str = f" [TU {s.sub_slot}]" if s.sub_slot > 0 else ""
                    cell_lines.append(f"<b>Str {s.stream_number}{sub_str}:</b> {s.raw_text}")
                    if s.is_drafting:
                        is_cell_drafting = True
                    for it in s.items:
                        if it.ai_number.lower() in self.tracked_tokens or any(tk in it.topic_label.lower() for tk in self.tracked_tokens):
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
        """Cross-references Agenda document items against allocated slots."""
        # Map: ai_number -> list of slots
        scheduled_ais: Dict[str, List[SessionSlot]] = {}
        for s in self.active_snapshot.slots:
            for it in s.items:
                if it.ai_number:
                    scheduled_ais.setdefault(it.ai_number, []).append(s)

        all_ais = sorted(self.agenda_map.keys(), key=self._ai_sort_key)
        self.agenda_table.setRowCount(len(all_ais))

        for row, ai_num in enumerate(all_ais):
            ai_item = self.agenda_map[ai_num]
            desc = ai_item.description

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

            self.agenda_table.setItem(row, 0, item_num)
            self.agenda_table.setItem(row, 1, item_desc)
            self.agenda_table.setItem(row, 2, status_item)
            self.agenda_table.setItem(row, 3, alloc_item)

            action_btn = QPushButton("👁️ Show")
            action_btn.setStyleSheet("padding: 2px 6px; font-size: 11px;")
            action_btn.clicked.connect(lambda _, a=ai_num: self._filter_by_ai(a))
            self.agenda_table.setCellWidget(row, 4, action_btn)

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
        """Populates the My Topics view."""
        tracked_slots = []
        for s in self.active_snapshot.slots:
            for it in s.items:
                if it.ai_number.lower() in self.tracked_tokens or any(tk in it.topic_label.lower() for tk in self.tracked_tokens):
                    tracked_slots.append(s)
                    break

        self.btn_tracked.setText(f"🎯 My Topics ({len(tracked_slots)})")
        self.tracked_table.setRowCount(len(tracked_slots))

        streams = self.active_snapshot.streams
        for row, s in enumerate(tracked_slots):
            s_info = streams.get(s.stream_number)
            r_name = s_info.room_name if s_info else f"Stream {s.stream_number}"

            time_str = SessionPlanTimezone.format_dual_time_range(
                s.start_time_local, s.end_time_local, self.tz_offset_min
            )

            ann = self.annotations.get(s.slot_key)
            note_str = ann.comment if ann else "—"

            self.tracked_table.setItem(row, 0, QTableWidgetItem(s.day_name))
            self.tracked_table.setItem(row, 1, QTableWidgetItem(s.quarter))
            self.tracked_table.setItem(row, 2, QTableWidgetItem(time_str))
            self.tracked_table.setItem(row, 3, QTableWidgetItem(f"Str {s.stream_number} ({r_name})"))
            self.tracked_table.setItem(row, 4, QTableWidgetItem(s.raw_text))
            self.tracked_table.setItem(row, 5, QTableWidgetItem(note_str))

    # --- Live Clock & Timezone ---

    def _update_live_clock(self):
        """Updates real-time Venue Time vs Austrian Laptop Time in the header."""
        venue_dt = SessionPlanTimezone.get_venue_current_datetime(self.tz_offset_min)
        laptop_offset = SessionPlanTimezone.get_laptop_tz_offset_minutes()

        venue_str = venue_dt.strftime("%a %H:%M")
        now_utc = datetime.now(timezone.utc)
        from datetime import timedelta
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

    # --- Navigation & Filters ---

    def _on_view_mode_changed(self, idx: int):
        self.view_stack.setCurrentIndex(idx)

    def _apply_search_filter(self, text: str):
        q = text.lower().strip()
        # Filter agenda table
        for r in range(self.agenda_table.rowCount()):
            ai_item = self.agenda_table.item(r, 0)
            desc_item = self.agenda_table.item(r, 1)
            txt = f"{ai_item.text()} {desc_item.text()}".lower() if ai_item and desc_item else ""
            self.agenda_table.setRowHidden(r, bool(q and q not in txt))

    def _filter_by_ai(self, ai_num: str):
        self.btn_grid.setChecked(True)
        self.view_stack.setCurrentIndex(0)
        self.search_input.setText(ai_num)

    def _on_grid_cell_double_clicked(self, row: int, col: int):
        q = QUARTER_ORDER[row]
        day = DAY_NAMES[col]
        slot_key = f"{day}|{q}|0|1"  # Default
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
            "Session Plan Files (*.docx *.doc *.zip *.htm *.html);;All Files (*.*)",
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
                staged = SessionPlanParser.stage_session_plan_file(p, self.meeting_dir)
                if ext == ".docx":
                    snap = SessionPlanParser.parse_docx(staged, meeting_number=self.meeting_number)
                else:
                    # Try text/html parser
                    snap = SessionPlanParser.parse_text_or_markdown(
                        staged.read_text(encoding="utf-8", errors="replace"),
                        meeting_number=self.meeting_number,
                        cached_file_path=str(staged),
                    )
            elif mode == "paste":
                snap = SessionPlanParser.parse_text_or_markdown(payload, meeting_number=self.meeting_number)
            else:
                QMessageBox.information(self, "URL Fetch", f"Fetching session plan from {payload}...")
                return

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

        # Compare active against previous
        snap_a = self.db.get_snapshot_by_id(snaps[1]["id"])
        snap_b = self.db.get_snapshot_by_id(snaps[0]["id"])
        diff_res = SessionPlanDiffEngine.compare_snapshots(snap_a, snap_b)
        dlg = SessionPlanDiffDialog(diff_res, self)
        dlg.exec_()

    def _export_mobile_html(self):
        if not self.active_snapshot:
            return
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

    def _export_ics(self):
        if not self.active_snapshot:
            return
        out_dir = self.meeting_dir / "Export"
        out_path = out_dir / f"SA2_{self.meeting_number}_Schedule.ics"
        start_date = self.mtg_info.get("start_date", "2026-10-12")
        SessionPlanExporter.export_ics(
            snapshot=self.active_snapshot,
            output_path=out_path,
            meeting_start_date_str=start_date,
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
