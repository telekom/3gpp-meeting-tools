# --- File: src/modules/meetings/ui/session_plan/window/mixins/overview_render.py ---
"""
Rendering of the Overview grid, Agenda gap-analysis mode, Drafting radar and Tracked topics.

Mixin of SessionPlanWindow: relies on attributes set in SessionPlanWindow.__init__.
"""
import html
from typing import (
    Dict,
    List,
    Tuple,
)
from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtGui import QColor
from PyQt5.QtWidgets import QPushButton, QTableWidgetItem
from modules.meetings.core.session_plan.session_plan_models import (
    SessionSlot,
    format_slot_display_text,
    format_venue_time_range,
    resolve_effective_room_and_location,
)
from core.ui.theme import C
from modules.meetings.core.session_plan.session_plan_parser import PERIOD_TIMES
from modules.meetings.ui.session_plan.window.constants import DAY_NAMES


class OverviewRenderMixin:
    """Rendering of the Overview grid, Agenda gap-analysis mode, Drafting radar and Tracked topics."""

    def _render_overview_grid(self):
        """Populates the 10-column spatial Overview Grid with separated stream rows, graphical half-quarters, and breaks."""
        self._grid_collisions_cache = self._detect_delegation_collisions()
        try:
            self._render_overview_grid_impl()
        finally:
            self._grid_collisions_cache = None

    def _make_day_tint_item(self, day_bg_hex: str) -> QTableWidgetItem:
        """Creates a non-editable background-tinted item so cellWidget margins match day zebra/today columns."""
        bg_item = QTableWidgetItem("")
        bg_item.setFlags(Qt.ItemIsEnabled)
        bg_item.setBackground(QColor(day_bg_hex))
        return bg_item

    def _is_plain_empty_drafting(self, slot: SessionSlot) -> bool:
        """True when a slot is a bare `'Drafting:'` placeholder with no delegates, tracked notes, or room override."""
        room_ov = getattr(slot, "room_override", "") or ""
        _, is_empty_drafting = format_slot_display_text(slot.raw_text, room_ov)
        if not is_empty_drafting:
            return False
        ann = self.annotations.get(slot.slot_key)
        if ann and (ann.comment or ann.assigned_colleague):
            return False
        del_names, _, is_all = self._get_slot_effective_delegates(slot)
        if is_all or del_names:
            return False
        return True

    def _make_empty_drafting_item(self, slot: SessionSlot, day_idx: int, period: str, stream_num: int, tu: int, day_bg_hex: str) -> QTableWidgetItem:
        """Renders bare empty `'Drafting:'` placeholders as a quiet muted dash while preserving the rich tooltip."""
        item = QTableWidgetItem("—")
        item.setTextAlignment(Qt.AlignCenter)
        item.setForeground(QColor(C.TEXT_FAINT))
        item.setBackground(QColor(day_bg_hex))
        item.setToolTip(self._build_rich_slot_tooltip(slot, day_idx, period, stream_num, tu))
        return item

    def _render_overview_grid_impl(self):
        self.grid_table.clearContents()
        self.grid_table.clearSpans()
        self.grid_table.setRowCount(0)
        self.grid_rows = self._build_grid_rows()
        self.grid_table.setRowCount(len(self.grid_rows))

        # 1. Vertical Headers with Time-First Period and Stream/Break
        v_headers = [r.display_label for r in self.grid_rows]
        self.grid_table.setVerticalHeaderLabels(v_headers)
        for row_idx, r_desc in enumerate(self.grid_rows):
            vh_item = self.grid_table.verticalHeaderItem(row_idx)
            if vh_item is not None:
                t_s, t_e = r_desc.time_range
                tt_lines = [f"<b>{html.escape(r_desc.period)}</b> · {t_s} – {t_e} (Venue Time)"]
                if r_desc.tu_split_label:
                    tt_lines.append(f"🕒 {html.escape(r_desc.tu_split_label)}")
                if r_desc.row_type == "stream":
                    tt_lines.append(f"🏢 Stream {r_desc.stream_number} · {html.escape(r_desc.short_room)}")
                if r_desc.is_live_period:
                    tt_lines.append("<b>🟢 HAPPENING NOW</b>")
                vh_item.setToolTip("<br>".join(tt_lines))

        # 2. Horizontal Headers with Day and Half-Quarter Columns (10 columns)
        h_headers = []
        today_idx = self._get_today_day_index()
        for idx in range(len(DAY_NAMES)):
            d_str, is_today = self._get_day_header_info(idx)
            if is_today:
                d_str += " 🟢"
            h_headers.append(f"{d_str}\n1st Half")
            h_headers.append(f"{d_str}\n2nd Half")

        self.grid_table.setHorizontalHeaderLabels(h_headers)

        # Determine which break periods have content (kept for test/API compatibility)
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
                    day_bg = self._get_day_bg_hex(day_idx, today_idx=today_idx)
                    matching = slot_map.get((day_idx, r_desc.period, r_desc.stream_number), [])

                    full_slot = next((s for s in matching if s.sub_slot == 0), None)
                    tu1_slot = next((s for s in matching if s.sub_slot == 1), None)
                    tu2_slot = next((s for s in matching if s.sub_slot == 2), None)

                    if full_slot:
                        self.grid_table.setSpan(row_idx, tu1_col, 1, 2)
                        if self._is_plain_empty_drafting(full_slot):
                            self.grid_table.setItem(
                                row_idx, tu1_col,
                                self._make_empty_drafting_item(full_slot, day_idx, r_desc.period, r_desc.stream_number, 0, day_bg),
                            )
                        else:
                            self.grid_table.setItem(row_idx, tu1_col, self._make_day_tint_item(day_bg))
                            w = self._create_slot_cell_widget(full_slot, day_idx, r_desc.period, r_desc.stream_number, is_half=False, tu=0)
                            self.grid_table.setCellWidget(row_idx, tu1_col, w)
                    elif tu1_slot or tu2_slot:
                        if tu1_slot:
                            if self._is_plain_empty_drafting(tu1_slot):
                                self.grid_table.setItem(
                                    row_idx, tu1_col,
                                    self._make_empty_drafting_item(tu1_slot, day_idx, r_desc.period, r_desc.stream_number, 1, day_bg),
                                )
                            else:
                                self.grid_table.setItem(row_idx, tu1_col, self._make_day_tint_item(day_bg))
                                w1 = self._create_slot_cell_widget(tu1_slot, day_idx, r_desc.period, r_desc.stream_number, is_half=True, tu=1)
                                self.grid_table.setCellWidget(row_idx, tu1_col, w1)
                        else:
                            item_e = QTableWidgetItem("—")
                            item_e.setTextAlignment(Qt.AlignCenter)
                            item_e.setForeground(QColor(C.BORDER_STRONG))
                            item_e.setBackground(QColor(day_bg))
                            self.grid_table.setItem(row_idx, tu1_col, item_e)

                        if tu2_slot:
                            if self._is_plain_empty_drafting(tu2_slot):
                                self.grid_table.setItem(
                                    row_idx, tu2_col,
                                    self._make_empty_drafting_item(tu2_slot, day_idx, r_desc.period, r_desc.stream_number, 2, day_bg),
                                )
                            else:
                                self.grid_table.setItem(row_idx, tu2_col, self._make_day_tint_item(day_bg))
                                w2 = self._create_slot_cell_widget(tu2_slot, day_idx, r_desc.period, r_desc.stream_number, is_half=True, tu=2)
                                self.grid_table.setCellWidget(row_idx, tu2_col, w2)
                        else:
                            item_e = QTableWidgetItem("—")
                            item_e.setTextAlignment(Qt.AlignCenter)
                            item_e.setForeground(QColor(C.BORDER_STRONG))
                            item_e.setBackground(QColor(day_bg))
                            self.grid_table.setItem(row_idx, tu2_col, item_e)
                    else:
                        self.grid_table.setSpan(row_idx, tu1_col, 1, 2)
                        item_e = QTableWidgetItem("—")
                        item_e.setTextAlignment(Qt.AlignCenter)
                        item_e.setForeground(QColor(C.BORDER_STRONG))
                        item_e.setBackground(QColor(day_bg))
                        self.grid_table.setItem(row_idx, tu1_col, item_e)

            elif r_desc.row_type == "break":
                for day_idx in range(5):
                    tu1_col = day_idx * 2
                    day_bg = self._get_day_bg_hex(day_idx, today_idx=today_idx)
                    self.grid_table.setSpan(row_idx, tu1_col, 1, 2)
                    matching_b = break_slot_map.get((day_idx, r_desc.period), [])
                    valid_b = [s for s in matching_b if s.raw_text.strip() and s.raw_text.strip() != "-"]
                    if valid_b:
                        self.grid_table.setItem(row_idx, tu1_col, self._make_day_tint_item(day_bg))
                        w = self._create_break_cell_widget(valid_b[0], day_idx, r_desc.period)
                        self.grid_table.setCellWidget(row_idx, tu1_col, w)
                    else:
                        item_b = QTableWidgetItem("—")
                        item_b.setTextAlignment(Qt.AlignCenter)
                        item_b.setForeground(QColor(C.BORDER_STRONG))
                        item_b.setBackground(QColor(C.SUCCESS_SOFT if today_idx == day_idx else C.BG_SURFACE))
                        times = PERIOD_TIMES[r_desc.period]
                        item_b.setToolTip(f"{r_desc.period}: {times[0]} – {times[1]} (Break)")
                        self.grid_table.setItem(row_idx, tu1_col, item_b)

        self._update_row_visibility()
        self._update_column_visibility()
        QTimer.singleShot(0, self.grid_table.resizeRowsToContents)

    def _render_agenda_mode(self):
        """Cross-references Agenda document items against allocated slots with 1-click ⭐ Track."""
        scheduled_ais = self._get_scheduled_ais_map()
        all_ais = self._combined_ai_keys(scheduled_ais)

        slots_by_ai = {ai: self._slots_for_ai(ai, scheduled_ais) for ai in all_ais}
        num_sched = sum(bool(slots) for slots in slots_by_ai.values())
        num_unsched = len(all_ais) - num_sched
        self.btn_agenda.setText(f"📋 Agenda ({num_sched} Sched / {len(all_ais)})")
        self.btn_agenda.setToolTip(
            f"Agenda Mode: {num_sched} Scheduled · {num_unsched} Unscheduled ({len(all_ais)} total AIs).\n"
            "Cross-references the official meeting agenda against scheduled sessions."
        )

        self._agenda_tracking_buttons = {}
        self.agenda_table.setRowCount(len(all_ais))

        for row, ai_num in enumerate(all_ais):
            desc, acronym = self._resolve_ai_description(ai_num, scheduled_ais, slots_by_ai[ai_num])

            is_tracked = (ai_num.lower() in self.tracked_tokens) or (acronym and acronym.lower() in self.tracked_tokens)

            # Col 0: Star toggle button
            star_btn = QPushButton("⭐" if is_tracked else "☆")
            topic_keys = {ai_num.lower()}
            if acronym:
                topic_keys.add(acronym.lower())
            for key in topic_keys:
                self._agenda_tracking_buttons.setdefault(key, []).append((star_btn, topic_keys))
            star_btn.setFixedWidth(28)
            star_btn.setStyleSheet("QPushButton { border: none; font-size: 14px; background: transparent; }")
            star_btn.setToolTip("Click to track/untrack this topic in 'My Topics'")
            star_btn.clicked.connect(lambda _, a=ai_num, l=desc: self._toggle_track_ai(a, l))
            self.agenda_table.setCellWidget(row, 0, star_btn)

            item_num = QTableWidgetItem(ai_num)
            item_desc = QTableWidgetItem(desc)

            matched_slots = slots_by_ai[ai_num]
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
            r_name, loc_note, _ = resolve_effective_room_and_location(
                s, streams=streams, room_locations=self.room_locations, slot_locations=self.slot_locations
            )
            r_disp = f"{r_name} · {loc_note}" if loc_note else r_name
            convener = s_info.convener if s_info else ""

            time_str = format_venue_time_range(s.start_time_local, s.end_time_local)

            self.drafting_table.setItem(row, 0, QTableWidgetItem(s.day_name))
            self.drafting_table.setItem(row, 1, QTableWidgetItem(s.quarter))
            item_time = QTableWidgetItem(time_str)
            f_t = item_time.font()
            f_t.setBold(True)
            item_time.setFont(f_t)
            item_time.setForeground(QColor(C.TEXT_PRIMARY))
            self.drafting_table.setItem(row, 2, item_time)
            self.drafting_table.setItem(row, 3, QTableWidgetItem(f"Str {s.stream_number} ({r_disp})"))
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
            for name, _m in self.delegation_plan.members.items():
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
            self.btn_tracked.setText(f"🎯 Delegate Agenda ({len(target_slots)} · ⚠️ {clash_slot_count})")
            self.btn_tracked.setToolTip(
                f"Delegate Agenda: {len(target_slots)} sessions · {clash_slot_count} parallel-stream clash(es) detected."
            )
        else:
            self.btn_tracked.setText(f"🎯 Delegate Agenda ({len(target_slots)})")
            self.btn_tracked.setToolTip(
                "View personal schedules by delegate (or your own agenda) and the entire delegation timeline."
            )

        self.btn_manage_topics.setText(f"⭐ Track Topics ({len(self.tracked_tokens)})")
        self.tracked_table.setRowCount(len(target_slots))

        streams = self.active_snapshot.streams
        for row, s in enumerate(target_slots):
            r_name, loc_note, _ = resolve_effective_room_and_location(
                s, streams=streams, room_locations=self.room_locations, slot_locations=self.slot_locations
            )
            r_disp = f"{r_name} · {loc_note}" if loc_note else r_name

            time_str = format_venue_time_range(s.start_time_local, s.end_time_local)

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
            f_t = item_time.font()
            f_t.setBold(True)
            item_time.setFont(f_t)
            item_time.setForeground(QColor(C.TEXT_PRIMARY))
            item_str = QTableWidgetItem(f"Str {s.stream_number} ({r_disp})")
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
