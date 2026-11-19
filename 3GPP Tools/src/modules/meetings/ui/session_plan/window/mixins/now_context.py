# --- File: src/modules/meetings/ui/session_plan/window/mixins/now_context.py ---
"""
NOW view state: live clock, timeline units, live/selected context resolution, navigation.

Mixin of SessionPlanWindow: relies on attributes set in SessionPlanWindow.__init__.
"""
from datetime import (
    datetime,
    timedelta,
    timezone,
)
from typing import (
    Dict,
    List,
    Optional,
    Set,
    Tuple,
)
from modules.meetings.core.session_plan.session_plan_models import SessionSlot, format_venue_time_range
from modules.meetings.core.session_plan.session_plan_timezone import SessionPlanTimezone
from core.ui.theme import C
from modules.meetings.core.session_plan.session_plan_parser import PERIOD_ORDER
from modules.meetings.ui.session_plan.window.constants import DAY_NAMES


class NowContextMixin:
    """NOW view state: live clock, timeline units, live/selected context resolution, navigation."""

    # --- Live Clock & Timezone ---

    def _update_live_clock(self):
        """Updates real-time Venue Time vs Laptop Time pill in the header and refreshes live highlights."""
        venue_dt = SessionPlanTimezone.get_venue_current_datetime(self.tz_offset_min)
        laptop_offset = SessionPlanTimezone.get_laptop_tz_offset_minutes()

        venue_str = venue_dt.strftime("%a %H:%M")
        now_utc = datetime.now(timezone.utc)
        laptop_dt = now_utc + timedelta(minutes=laptop_offset)
        laptop_str = laptop_dt.strftime("%H:%M")

        if self.tz_offset_min == laptop_offset:
            self.clock_lbl.setText(
                f"🕒 Venue &amp; Local: <b style='font-size:13px; color:{C.INFO_DARK};'>{venue_str}</b>"
            )
        else:
            self.clock_lbl.setText(
                f"🏛️ Venue: <b style='font-size:13px; color:{C.INFO_DARK};'>{venue_str}</b>"
                f" &nbsp;<span style='color:{C.BORDER_STRONG};'>|</span>&nbsp; "
                f"🏠 Local: <b style='font-size:13px; color:{C.TEXT_PRIMARY};'>{laptop_str}</b>"
            )

        bucket = (venue_dt.hour, venue_dt.minute)
        if self._last_now_live_bucket != bucket:
            self._last_now_live_bucket = bucket
            live_p = self._get_live_period_name()
            if hasattr(self, "grid_rows") and self.grid_rows:
                changed = False
                for r_desc in self.grid_rows:
                    new_live = (r_desc.period == live_p)
                    if r_desc.is_live_period != new_live:
                        r_desc.is_live_period = new_live
                        changed = True
                if changed and hasattr(self, "grid_table"):
                    self.grid_table.verticalHeader().viewport().update()
            if self.active_snapshot and hasattr(self, "now_page"):
                self._update_now_tab_badge()
                if self.now_auto_follow and self.view_stack.currentWidget() == self.now_page:
                    self._render_now_view()

    # --- Live "Now" & Upcoming Timeline Engine + Room Location Editors ---

    def _clear_now_layout(self):
        """Safely clears all widgets and sub-layouts from `self.now_content_layout`."""
        if not hasattr(self, "now_content_layout"):
            return
        while self.now_content_layout.count():
            item = self.now_content_layout.takeAt(0)
            w = item.widget()
            if w is not None:
                w.setParent(None)
                w.deleteLater()
            elif item.layout() is not None:
                self._clear_sub_layout(item.layout())

    @staticmethod
    def _clear_sub_layout(lay):
        if lay is None:
            return
        while lay.count():
            child = lay.takeAt(0)
            cw = child.widget()
            if cw is not None:
                cw.setParent(None)
                cw.deleteLater()
            else:
                sub_l = child.layout()
                if sub_l is not None:
                    NowContextMixin._clear_sub_layout(sub_l)

    def _build_now_timeline_units(self) -> List[dict]:
        """
        Builds ordered chronological schedule units (`day_idx|period|tu`) across Mon–Fri
        for active sessions and non-empty breaks, splitting into TU 1/TU 2 when half-quarters are used.
        """
        if not self.active_snapshot:
            return []

        units: List[dict] = []
        by_day_period: Dict[Tuple[int, str], List[SessionSlot]] = {}
        days_with_slots: Set[int] = set()

        for sl in self.active_snapshot.slots:
            if not sl.raw_text.strip() or sl.raw_text.strip() == "-":
                continue
            by_day_period.setdefault((sl.day_index, sl.quarter), []).append(sl)
            days_with_slots.add(sl.day_index)

        all_periods = list(PERIOD_ORDER)
        for sl in self.active_snapshot.slots:
            if sl.quarter and sl.quarter not in all_periods:
                all_periods.append(sl.quarter)

        for d_idx in sorted(days_with_slots):
            if d_idx < 0 or d_idx >= len(DAY_NAMES):
                continue
            d_name = DAY_NAMES[d_idx]
            for p in all_periods:
                p_slots = by_day_period.get((d_idx, p), [])
                if not p_slots:
                    continue
                times = self._get_period_effective_times(p)
                is_break = any(b in p.lower() for b in ("break", "coffee", "lunch"))
                has_halves = (not is_break) and any(s.sub_slot in (1, 2) for s in p_slots)

                if has_halves:
                    units.append({
                        "key": f"{d_idx}|{p}|1",
                        "day_idx": d_idx,
                        "day_name": d_name,
                        "period": p,
                        "tu": 1,
                        "is_break": False,
                        "start": times[2],
                        "end": times[3],
                        "half_label": "1st Half (TU 1)",
                    })
                    units.append({
                        "key": f"{d_idx}|{p}|2",
                        "day_idx": d_idx,
                        "day_name": d_name,
                        "period": p,
                        "tu": 2,
                        "is_break": False,
                        "start": times[4],
                        "end": times[5],
                        "half_label": "2nd Half (TU 2)",
                    })
                else:
                    units.append({
                        "key": f"{d_idx}|{p}|0",
                        "day_idx": d_idx,
                        "day_name": d_name,
                        "period": p,
                        "tu": 0,
                        "is_break": is_break,
                        "start": times[0],
                        "end": times[1],
                        "half_label": "Break" if is_break else "Full Quarter",
                    })
        return units

    @staticmethod
    def _hhmm_to_min(hhmm: str) -> int:
        m = SessionPlanTimezone._time_str_to_minutes(hhmm or "")
        return m if m is not None else 0

    def _resolve_now_context(self) -> dict:
        """
        Determines live vs manual target timeline unit, status banner, remaining/countdown minutes,
        immediate next 1–2 upcoming periods today, and remaining periods for that day.
        """
        units = self._build_now_timeline_units()
        if not units:
            return {"units": [], "target": None, "mode": "empty", "status_pill": "📭 Empty", "banner": "No scheduled slots."}

        venue_dt = SessionPlanTimezone.get_venue_current_datetime(self.tz_offset_min)
        cur_hhmm = venue_dt.strftime("%H:%M")
        cur_min = venue_dt.hour * 60 + venue_dt.minute

        day_diff: Optional[int] = None
        base_date = None
        try:
            base_date = datetime.strptime(self.start_date_str, "%Y-%m-%d").date()
            day_diff = (venue_dt.date() - base_date).days
        except Exception:
            pass

        live_unit = None
        next_live_unit = None
        live_mode = "outside"

        if day_diff is not None and 0 <= day_diff <= 4:
            today_units = [u for u in units if u["day_idx"] == day_diff]
            for u in today_units:
                if u["start"] <= cur_hhmm < u["end"]:
                    live_unit = u
                    live_mode = "live"
                    break
            if live_unit is None:
                for u in today_units:
                    if u["start"] > cur_hhmm:
                        next_live_unit = u
                        live_mode = "upcoming_today"
                        break
                if next_live_unit is None:
                    after_units = [u for u in units if u["day_idx"] > day_diff]
                    if after_units:
                        next_live_unit = after_units[0]
                        live_mode = "next_day"
                    else:
                        live_mode = "after_meeting"
        elif day_diff is not None and day_diff < 0:
            live_mode = "before_meeting"
        elif day_diff is not None and day_diff > 4:
            live_mode = "after_meeting"

        auto_target = live_unit or next_live_unit or units[0]
        target = auto_target
        effective_mode = live_mode

        if not self.now_auto_follow and self.now_selected_unit_key:
            manual = next((u for u in units if u["key"] == self.now_selected_unit_key), None)
            if manual is not None:
                target = manual
                if live_unit and manual["key"] == live_unit["key"]:
                    effective_mode = "live"
                else:
                    effective_mode = "manual"
        else:
            self.now_selected_unit_key = target["key"]

        date_lbl = ""
        if base_date:
            d_obj = base_date + timedelta(days=target["day_idx"])
            date_lbl = f" ({d_obj.strftime('%d %b')})"

        venue_t = format_venue_time_range(target["start"], target["end"])

        if effective_mode == "live":
            rem_min = max(0, self._hhmm_to_min(target["end"]) - cur_min)
            status_pill = "🟢 HAPPENING NOW"
            banner = f"{target['day_name']}{date_lbl} · {target['period']} · {target['half_label']} — {venue_t} Venue ({rem_min} min remaining)"
        elif effective_mode == "upcoming_today":
            wait_min = max(0, self._hhmm_to_min(target["start"]) - cur_min)
            status_pill = f"⏳ UP NEXT IN {wait_min} MIN"
            banner = f"Currently between sessions ({cur_hhmm} Venue) — Next: {target['day_name']}{date_lbl} · {target['period']} · {target['half_label']} ({venue_t})"
        elif effective_mode == "next_day":
            status_pill = "🌙 NEXT DAY PREVIEW"
            banner = f"Today's sessions concluded ({cur_hhmm} Venue) — Next: {target['day_name']}{date_lbl} · {target['period']} ({venue_t})"
        elif effective_mode == "before_meeting":
            days_until = max(1, -(day_diff or 0))
            status_pill = f"📅 STARTS IN {days_until}D"
            banner = f"Before meeting week — Previewing {target['day_name']}{date_lbl} · {target['period']} ({venue_t})"
        elif effective_mode == "after_meeting":
            status_pill = "🏁 MEETING CONCLUDED"
            banner = f"Meeting week concluded — Previewing {target['day_name']}{date_lbl} · {target['period']} ({venue_t})"
        elif effective_mode == "manual":
            status_pill = "🔎 MANUAL PREVIEW"
            banner = f"Inspecting {target['day_name']}{date_lbl} · {target['period']} · {target['half_label']} — {venue_t} Venue"
        else:
            status_pill = "📅 SCHEDULE PREVIEW"
            banner = f"{target['day_name']}{date_lbl} · {target['period']} · {target['half_label']} — {venue_t} Venue"

        # Collect distinct periods on target day after target period for "Up Next" (1-2) and "Rest of Today"
        all_periods = list(PERIOD_ORDER)
        for sl in self.active_snapshot.slots:
            if sl.quarter and sl.quarter not in all_periods:
                all_periods.append(sl.quarter)

        try:
            cur_p_idx = all_periods.index(target["period"])
        except ValueError:
            cur_p_idx = -1

        later_periods: List[dict] = []
        for p_idx in range(cur_p_idx + 1, len(all_periods)):
            p_name = all_periods[p_idx]
            p_slots = [
                sl for sl in self.active_snapshot.slots
                if sl.day_index == target["day_idx"]
                and sl.quarter == p_name
                and sl.raw_text.strip()
                and sl.raw_text.strip() != "-"
            ]
            if p_slots:
                t_rng = self._get_period_effective_times(p_name)
                later_periods.append({
                    "period": p_name,
                    "start": t_rng[0],
                    "end": t_rng[1],
                    "is_break": any(b in p_name.lower() for b in ("break", "coffee", "lunch")),
                    "slots": sorted(p_slots, key=lambda s: (s.stream_number, s.sub_slot)),
                })

        return {
            "units": units,
            "target": target,
            "live_unit": live_unit,
            "mode": effective_mode,
            "status_pill": status_pill,
            "banner": banner,
            "up_next_periods": later_periods[:2],
            "rest_of_day_periods": later_periods[2:],
        }

    def _update_now_tab_badge(self):
        """Updates the 🔴 Now tab label with active quarter indicator when live."""
        if not hasattr(self, "btn_now"):
            return
        ctx = self._resolve_now_context()
        live_u = ctx.get("live_unit")
        if live_u:
            self.btn_now.setText(f"🟢 Now ({live_u['period']})")
        else:
            self.btn_now.setText("🔴 Now")

    def _slots_for_now_unit(self, unit: dict) -> List[SessionSlot]:
        """Returns active slots in the snapshot matching `(day_idx, period, tu)` sorted by stream."""
        if not self.active_snapshot or not unit:
            return []
        d_idx = unit["day_idx"]
        p = unit["period"]
        tu = unit["tu"]
        matched = []
        for sl in self.active_snapshot.slots:
            if not sl.raw_text.strip() or sl.raw_text.strip() == "-":
                continue
            if sl.day_index != d_idx or sl.quarter != p:
                continue
            if tu in (1, 2) and sl.sub_slot not in (0, tu):
                continue
            matched.append(sl)
        return sorted(matched, key=lambda s: (s.stream_number, s.sub_slot))

    def _step_now_unit(self, delta: int):
        units = self._build_now_timeline_units()
        if not units:
            return
        cur_key = self.now_selected_unit_key
        idx = next((i for i, u in enumerate(units) if u["key"] == cur_key), 0)
        new_idx = (idx + delta) % len(units)
        self.now_auto_follow = False
        self.now_selected_unit_key = units[new_idx]["key"]
        self._render_now_view()

    def _on_now_combo_changed(self, idx: int):
        if idx < 0 or not hasattr(self, "now_slot_combo"):
            return
        key = self.now_slot_combo.currentData()
        if key:
            self.now_auto_follow = False
            self.now_selected_unit_key = key
            self._render_now_view()

    def _reset_now_to_live(self):
        self.now_auto_follow = True
        self.now_selected_unit_key = None
        self._render_now_view()
