# --- File: src/modules/meetings/ui/session_plan/window/main_window.py ---
"""
SessionPlanWindow shell: constructor, Qt lifecycle events, data loading and view dispatch.
Feature logic lives in the mixins/ package (see README.md for a 'where to edit what' map).
"""
import html
from pathlib import Path
from typing import (
    Dict,
    List,
    Optional,
    Set,
    Tuple,
)
from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtWidgets import (
    QCheckBox,
    QDialog,
    QLabel,
    QMessageBox,
    QStackedWidget,
    QVBoxLayout,
    QWidget,
)
from modules.meetings.core.agenda_manager import AgendaManager
from modules.meetings.core.session_plan.session_plan_db import SessionPlanDatabase
from modules.meetings.core.session_plan.session_plan_models import (
    DelegationPlan,
    SessionPlanSnapshot,
    SessionSlot,
    UserSessionAnnotation,
    collect_meeting_physical_rooms,
    lookup_room_location_note,
)
from modules.meetings.core.session_plan.session_plan_timezone import SessionPlanTimezone
from modules.meetings.ui.session_plan.session_plan_dialogs import (
    SessionPlanRoomLocationsDialog,
    SessionPlanTimezoneDialog,
    SessionPlanTrackedTopicsDialog,
)
from modules.meetings.ui.tdocs_drop_overlay import DropOverlayWidget
from core.ui.theme import C
from modules.meetings.ui.session_plan.window.constants import GridRowDescriptor, _format_compact_snapshot_label
from modules.meetings.ui.session_plan.window.mixins import DropImportMixin, RibbonSetupMixin, TeamSetupMixin, GridModelMixin, SlotCellsMixin, OverviewRenderMixin, DelegationMixin, TeamRenderMixin, NowContextMixin, NowCardsMixin, InteractionsMixin, ExportsMixin


class SessionPlanWindow(
    DropImportMixin,
    RibbonSetupMixin,
    TeamSetupMixin,
    GridModelMixin,
    SlotCellsMixin,
    OverviewRenderMixin,
    DelegationMixin,
    TeamRenderMixin,
    NowContextMixin,
    NowCardsMixin,
    InteractionsMixin,
    ExportsMixin,
    QWidget,
):
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
        self.room_locations: Dict[str, str] = {}
        self.slot_locations: Dict[str, str] = {}
        self.now_auto_follow: bool = True
        self.now_selected_unit_key: Optional[str] = None
        self._last_now_live_bucket: Optional[Tuple[int, int]] = None
        self.grid_rows: List[GridRowDescriptor] = []
        self.empty_break_periods: Set[str] = set()
        self.grid_day_filter_mode: str = "ALL"  # "ALL", "TODAY_PLUS", or "CUSTOM"
        self.grid_selected_days: Set[int] = set()

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

    # --- Timezone Initialization ---

    def _init_timezone(self):
        """Initializes venue timezone from meeting_timezone.json, DB metadata, or city lookup."""
        mtg_root = self.meeting_dir.parent if self.meeting_dir.name.lower() == "agenda" else self.meeting_dir
        folder_tz = SessionPlanTimezone.load_meeting_folder_timezone(mtg_root)
        meta = self.db.get_meeting_meta(self.meeting_number)

        if folder_tz and folder_tz.get("source") == "manual_override":
            self.tz_offset_min = int(folder_tz["tz_offset_minutes"])
            self.tz_name = folder_tz.get("tz_label") or f"UTC{self.tz_offset_min//60:+03d}:00"
            return

        if meta and meta.get("tz_offset_minutes") is not None:
            self.tz_offset_min = meta["tz_offset_minutes"]
            self.tz_name = meta.get("tz_name") or f"UTC{self.tz_offset_min//60:+03d}:00"
            return

        if folder_tz:
            self.tz_offset_min = int(folder_tz["tz_offset_minutes"])
            self.tz_name = folder_tz.get("tz_label") or f"UTC{self.tz_offset_min//60:+03d}:00"
            return

        res = SessionPlanTimezone.lookup_city_timezone(self.location, self.start_date_str)
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
        self._setup_now_view()
        self._setup_agenda_view()
        self._setup_drafting_view()
        self._setup_tracked_view()
        self._setup_team_view()

        layout.addWidget(self.view_stack, 1)

    # --- Data Loading & Rendering ---

    def _load_data(self):
        """Loads snapshots, persistent tracked topics, annotations, and room/slot locations from SQLite/folder."""
        self.annotations = self.db.get_annotations(self.meeting_number)
        self.room_locations = self.db.get_room_locations(self.meeting_number, meeting_dir=self.meeting_dir)
        self.slot_locations = self.db.get_slot_locations(self.meeting_number, meeting_dir=self.meeting_dir)
        self.db_tracked = self.db.get_tracked_topics(self.meeting_number)
        self.tracked_tokens = set(self.external_tracked_tokens) | set(self.db_tracked.keys())

        snapshots = self.db.list_snapshots(self.meeting_number)

        self.snap_combo.blockSignals(True)
        self.snap_combo.clear()
        for snap_row in snapshots:
            full_label = f"{snap_row['snapshot_label']} ({snap_row['imported_at'][:16]})"
            compact_label = _format_compact_snapshot_label(snap_row["snapshot_label"], snap_row["imported_at"])
            idx_row = self.snap_combo.count()
            self.snap_combo.addItem(compact_label, snap_row["id"])
            self.snap_combo.setItemData(idx_row, full_label, Qt.ToolTipRole)
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
        """Refreshes all 6 views with the currently active snapshot."""
        if not self.active_snapshot:
            self._render_empty_views()
            return

        self._update_room_guide_strip()
        self._update_overview_delegate_strip()
        self._render_overview_grid()
        self._render_now_view()
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
                loc_note = lookup_room_location_note(short_room, self.room_locations)
                loc_line = f"<br><b>Location Note:</b> {html.escape(loc_note)}" if loc_note else ""
                chk.setText(f"🏢 Str {s_num}: {short_room}")
                chk.setToolTip(
                    f"<b>Stream {s_num}:</b> {html.escape(s_info.room_name)} [{s_info.capacity} seats]{loc_line}<br>"
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
        if hasattr(self, "now_content_layout"):
            self._clear_now_layout()
            empty_lbl = QLabel("📭 Import a Session Plan snapshot to see what is happening now and upcoming.")
            empty_lbl.setStyleSheet(f"color: {C.TEXT_MUTED}; font-size: 13px; padding: 24px;")
            empty_lbl.setAlignment(Qt.AlignCenter)
            self.now_content_layout.addWidget(empty_lbl)
            self.now_content_layout.addStretch(1)

    def _open_room_locations_dialog(self, focus_room: str = ""):
        """Opens the dialog for editing hand-made floor/wing notes per physical room."""
        streams = self.active_snapshot.streams if self.active_snapshot else {}
        room_names = collect_meeting_physical_rooms(self.active_snapshot, self.room_locations)
        room_infos: List[Tuple[str, str]] = []
        for rn in room_names:
            sub = "Custom / Ad-hoc Room"
            for s_num, s_info in sorted(streams.items()):
                base_r = s_info.room_name.split("(")[0].strip() or s_info.room_name
                if base_r.lower() == rn.lower():
                    cap_s = f" · {s_info.capacity} seats" if s_info.capacity else ""
                    ch_s = f" · {s_info.convener}" if s_info.convener else ""
                    sub = f"Stream {s_num}{cap_s}{ch_s}"
                    break
            room_infos.append((rn, sub))

        current_prefill = {
            rn: lookup_room_location_note(rn, self.room_locations) or self.room_locations.get(rn, "")
            for rn in room_names
        }
        dlg = SessionPlanRoomLocationsDialog(
            meeting_number=self.meeting_number,
            room_infos=room_infos,
            current_locations=current_prefill,
            focus_room=focus_room,
            parent=self,
        )
        if dlg.exec_() == QDialog.Accepted:
            updated = dlg.get_locations()
            self.db.set_all_room_locations(
                self.meeting_number, updated, meeting_dir=self.meeting_dir
            )
            self.room_locations = self.db.get_room_locations(self.meeting_number, meeting_dir=self.meeting_dir)
            self._render_current_views()

    def _edit_slot_or_room_location(self, physical_room: str, slot: Optional[SessionSlot] = None):
        """Quickly edits the physical room's location note from a card in the Now tab."""
        self._open_room_locations_dialog(focus_room=physical_room)

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
            mtg_root = self.meeting_dir.parent if self.meeting_dir.name.lower() == "agenda" else self.meeting_dir
            SessionPlanTimezone.save_meeting_folder_timezone(
                meeting_dir=mtg_root,
                tz_offset_minutes=self.tz_offset_min,
                tz_label=self.tz_name,
                source="manual_override",
                location=self.location,
                wg=str(self.mtg_info.get("wg_name", "")),
                meeting_number=self.meeting_number,
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
        if hasattr(self, "now_page") and self.view_stack.currentWidget() == self.now_page and self.active_snapshot:
            self._render_now_view()

    def _apply_search_filter(self, text: str):
        q = text.lower().strip()
        for r in range(self.agenda_table.rowCount()):
            ai_item = self.agenda_table.item(r, 1)
            desc_item = self.agenda_table.item(r, 2)
            txt = f"{ai_item.text()} {desc_item.text()}".lower() if ai_item and desc_item else ""
            self.agenda_table.setRowHidden(r, bool(q and q not in txt))

    def _filter_by_ai(self, ai_num: str):
        self.btn_grid.setChecked(True)
        self.view_stack.setCurrentWidget(self.overview_page)
        self.search_input.setText(ai_num)

    def _on_snapshot_changed(self, idx: int):
        snap_id = self.snap_combo.itemData(idx)
        if snap_id:
            self.db.set_active_snapshot(snap_id)
            self.active_snapshot = self.db.get_snapshot_by_id(snap_id)
            self._render_current_views()
