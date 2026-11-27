# --- File: src/modules/meetings/ui/session_plan/window/mixins/grid_model.py ---
"""
Grid row model: period timings, live period, day filters, row/column visibility,
AI-number helpers and grid navigation (jump to slot / ensure day visible).

Mixin of SessionPlanWindow: relies on attributes set in SessionPlanWindow.__init__.
"""
import re
from datetime import datetime, timedelta
from typing import (
    Dict,
    List,
    Optional,
    Set,
    Tuple,
)
from PyQt5.QtCore import QTimer
from modules.meetings.core.session_plan.session_plan_models import (
    SessionSlot,
    ai_numbers_match,
    expand_ai_number,
    format_break_period_label,
)
from modules.meetings.core.session_plan.session_plan_timezone import SessionPlanTimezone
from core.ui.theme import C
from modules.meetings.core.session_plan.session_plan_parser import (
    PERIOD_ORDER,
    PERIOD_TIMES,
    SessionPlanParser,
)
from modules.meetings.ui.session_plan.window.constants import DAY_NAMES, GridRowDescriptor


class GridModelMixin:
    """Grid row model: period timings, live period, day filters, row/column visibility,"""

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

    def _get_live_period_name(self) -> Optional[str]:
        """Returns the period code (e.g. 'Q1', 'Lunch') active right now in Venue time, or None."""
        venue_dt = SessionPlanTimezone.get_venue_current_datetime(self.tz_offset_min)
        try:
            base_date = datetime.strptime(self.start_date_str, "%Y-%m-%d").date()
            day_diff = (venue_dt.date() - base_date).days
        except Exception:
            return None
        if day_diff < 0 or day_diff > 4:
            return None
        cur_hhmm = venue_dt.strftime("%H:%M")
        for p in PERIOD_ORDER:
            times = self._get_period_effective_times(p)
            if times[0] <= cur_hhmm < times[1]:
                return p
        return None

    def _build_grid_rows(self) -> List[GridRowDescriptor]:
        """Constructs the 23-row time-first spatial grid descriptors covering quarters and breaks."""
        rows = []
        streams = self.active_snapshot.streams if self.active_snapshot else {}
        live_period = self._get_live_period_name()
        for p in PERIOD_ORDER:
            times = self._get_period_effective_times(p)
            tu_split = f"1st Half: {times[2]}–{times[3]} · 2nd Half: {times[4]}–{times[5]}"
            is_live = (p == live_period)
            if any(b in p.lower() for b in ["break", "coffee", "lunch"]):
                clean_name = format_break_period_label(p)
                lbl = f"{clean_name}\n{times[0]} – {times[1]}"
                rows.append(
                    GridRowDescriptor(
                        p,
                        0,
                        "break",
                        lbl,
                        (times[0], times[1]),
                        short_room="",
                        local_time_range=("", ""),
                        tu_split_label="",
                        is_live_period=is_live,
                    )
                )
            else:
                for s_num in [1, 2, 3]:
                    s_info = streams.get(s_num)
                    r_name = s_info.room_name if s_info else ("Main" if s_num == 1 else f"Brk {s_num-1}")
                    short_r = r_name.split("(")[0].strip()
                    lbl = f"{p} · {times[0]} – {times[1]}\nStr {s_num} ({short_r})"
                    rows.append(
                        GridRowDescriptor(
                            p,
                            s_num,
                            "stream",
                            lbl,
                            (times[0], times[1]),
                            short_room=short_r,
                            local_time_range=("", ""),
                            tu_split_label=tu_split,
                            is_live_period=is_live,
                        )
                    )
        return rows

    def _get_today_day_index(self) -> Optional[int]:
        """Returns 0..4 when the current venue date falls within the Mon-Fri meeting week, else None."""
        try:
            base_date = datetime.strptime(self.start_date_str, "%Y-%m-%d").date()
        except Exception:
            return None
        venue_today = SessionPlanTimezone.get_venue_current_datetime(self.tz_offset_min).date()
        diff = (venue_today - base_date).days
        return diff if 0 <= diff <= 4 else None

    def _get_day_header_info(self, day_idx: int) -> Tuple[str, bool]:
        """Returns `(formatted_day_label, is_today)` for horizontal column header painting."""
        d_name = DAY_NAMES[day_idx] if 0 <= day_idx < len(DAY_NAMES) else f"Day {day_idx + 1}"
        try:
            base_date = datetime.strptime(self.start_date_str, "%Y-%m-%d").date()
            cur_date = base_date + timedelta(days=day_idx)
            venue_today = SessionPlanTimezone.get_venue_current_datetime(self.tz_offset_min).date()
            return f"{d_name[:3]} {cur_date.strftime('%d %b')}", (cur_date == venue_today)
        except Exception:
            return d_name, False

    def _get_day_bg_hex(self, day_idx: int, today_idx: Optional[int] = None) -> str:
        """Returns subtle column background tint for alternating days and Today highlight."""
        if today_idx is None:
            today_idx = self._get_today_day_index()
        if today_idx is not None and day_idx == today_idx:
            return C.SUCCESS_SOFT
        return C.BG_SURFACE if (day_idx % 2 == 1) else C.BG_CARD

    def _style_day_filter_buttons(self):
        """Styles the Overview Grid Day Filter pill buttons (`All 5`, `Today+`, `Mon`..`Fri`)."""
        if not hasattr(self, "btn_day_all"):
            return
        base = "padding: 2px 7px; font-size: 10.5px; font-weight: 600; border-radius: 5px;"
        active_all = f"background: {C.ACCENT_PRIMARY}; color: {C.BG_CARD}; border: 1px solid {C.ACCENT_HOVER};"
        active_today = f"background: {C.EMERALD_PRIMARY}; color: {C.BG_CARD}; border: 1px solid {C.EMERALD_DARK};"
        active_day = f"background: {C.INFO_ALT}; color: {C.BG_CARD}; border: 1px solid {C.ACCENT_HOVER};"
        idle = f"background: {C.BG_CARD}; color: {C.TEXT_SECONDARY}; border: 1px solid {C.BORDER_STRONG};"

        self.btn_day_all.blockSignals(True)
        self.btn_day_all.setChecked(self.grid_day_filter_mode == "ALL")
        self.btn_day_all.blockSignals(False)
        self.btn_day_all.setStyleSheet(
            f"QPushButton {{ {base} {active_all if self.grid_day_filter_mode == 'ALL' else idle} }}"
        )

        self.btn_day_today_plus.blockSignals(True)
        self.btn_day_today_plus.setChecked(self.grid_day_filter_mode == "TODAY_PLUS")
        self.btn_day_today_plus.blockSignals(False)
        self.btn_day_today_plus.setStyleSheet(
            f"QPushButton {{ {base} {active_today if self.grid_day_filter_mode == 'TODAY_PLUS' else idle} }}"
        )

        for idx, btn in enumerate(getattr(self, "day_filter_btns", [])):
            is_sel = (self.grid_day_filter_mode == "CUSTOM" and idx in self.grid_selected_days)
            btn.blockSignals(True)
            btn.setChecked(is_sel)
            btn.blockSignals(False)
            btn.setStyleSheet(f"QPushButton {{ {base} {active_day if is_sel else idle} }}")

    def _on_day_filter_clicked(self, target):
        """Handles clicks on `All 5`, `Today+`, or individual `Mon`..`Fri` Day Filter pills."""
        if target == "ALL":
            self.grid_day_filter_mode = "ALL"
            self.grid_selected_days.clear()
        elif target == "TODAY_PLUS":
            self.grid_day_filter_mode = "TODAY_PLUS"
            self.grid_selected_days.clear()
        elif isinstance(target, int):
            if self.grid_day_filter_mode != "CUSTOM":
                self.grid_day_filter_mode = "CUSTOM"
                self.grid_selected_days = {target}
            else:
                if target in self.grid_selected_days:
                    self.grid_selected_days.discard(target)
                else:
                    self.grid_selected_days.add(target)
                if not self.grid_selected_days or len(self.grid_selected_days) >= len(DAY_NAMES):
                    self.grid_day_filter_mode = "ALL"
                    self.grid_selected_days.clear()

        self._style_day_filter_buttons()
        self._update_column_visibility()

    def _update_column_visibility(self):
        """Shows/hides day column pairs in `grid_table` according to the active Day Filter."""
        if not hasattr(self, "grid_table"):
            return
        if self.grid_day_filter_mode == "TODAY_PLUS":
            today_idx = self._get_today_day_index()
            if today_idx is not None:
                visible_days = set(range(today_idx, len(DAY_NAMES)))
            else:
                visible_days = set(range(len(DAY_NAMES)))
        elif self.grid_day_filter_mode == "CUSTOM" and self.grid_selected_days:
            visible_days = set(self.grid_selected_days)
        else:
            visible_days = set(range(len(DAY_NAMES)))

        for d_idx in range(len(DAY_NAMES)):
            hide = d_idx not in visible_days
            self.grid_table.setColumnHidden(d_idx * 2, hide)
            self.grid_table.setColumnHidden(d_idx * 2 + 1, hide)

        if hasattr(self, "horizontal_day_header"):
            self.horizontal_day_header.viewport().update()
        QTimer.singleShot(0, self.grid_table.resizeRowsToContents)

    def _update_row_visibility(self):
        """Hides/shows rows according to interactive stream and break toggle checkboxes."""
        show_s1 = self.btn_toggle_s1.isChecked()
        show_s2 = self.btn_toggle_s2.isChecked()
        show_s3 = self.btn_toggle_s3.isChecked()
        show_breaks = self.chk_breaks.isChecked()

        for row_idx, r_desc in enumerate(self.grid_rows):
            if r_desc.row_type == "stream":
                if r_desc.stream_number == 1:
                    self.grid_table.setRowHidden(row_idx, not show_s1)
                elif r_desc.stream_number == 2:
                    self.grid_table.setRowHidden(row_idx, not show_s2)
                elif r_desc.stream_number == 3:
                    self.grid_table.setRowHidden(row_idx, not show_s3)
            elif r_desc.row_type == "break":
                self.grid_table.setRowHidden(row_idx, not show_breaks)
        if hasattr(self, "grid_table") and self.grid_table.verticalHeader():
            self.grid_table.verticalHeader().viewport().update()

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

    def _resolve_ai_description(self, ai_num: str, scheduled_ais: Dict[str, List[SessionSlot]], matched_slots=None) -> Tuple[str, Optional[str]]:
        """Resolves human-readable description and optional acronym for an agenda item number."""
        if self.agenda_map and ai_num in self.agenda_map:
            ai_item = self.agenda_map[ai_num]
            return ai_item.description, getattr(ai_item, "acronym", None)

        # Look for the specific AgendaSlotItem matching this ai_number (or .x wildcard)
        if matched_slots is None:
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

    def _ensure_day_visible(self, day_idx: int):
        """Ensures the target day columns are visible if currently hidden by Day Filter."""
        if not (0 <= day_idx < len(DAY_NAMES)) or not hasattr(self, "grid_table"):
            return
        if self.grid_table.isColumnHidden(day_idx * 2):
            self._on_day_filter_clicked("ALL")

    def _jump_to_slot_in_grid(self, day_idx: int, period: str, stream_num: int = 1, sub_slot: int = 0):
        target_row = None
        for r_i, r_d in enumerate(self.grid_rows):
            if r_d.period == period and (r_d.stream_number == stream_num or r_d.row_type == "break"):
                target_row = r_i
                break
        if target_row is not None:
            self._ensure_day_visible(day_idx)
            self.btn_grid.setChecked(True)
            self.view_stack.setCurrentWidget(self.overview_page)
            col_idx = day_idx * 2 + (1 if sub_slot == 2 else 0)
            self.grid_table.setCurrentCell(target_row, col_idx)

    @staticmethod
    def _ai_sort_key(ai_str: str) -> tuple:
        parts = []
        for p in re.split(r"(\d+)", ai_str):
            if p.isdigit():
                parts.append(int(p))
            elif p:
                parts.append(p)
        return tuple(parts)
