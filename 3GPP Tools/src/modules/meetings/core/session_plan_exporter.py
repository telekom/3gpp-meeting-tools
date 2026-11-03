# --- File: src/modules/meetings/core/session_plan_exporter.py ---
"""
Multi-format exporter for 3GPP SA2 Meeting Session Plans.
1. Mobile-First Standalone HTML: Zero-dependency, offline-ready, responsive day/quarter cards.
2. iCalendar (.ics): RFC 5545 UTC-anchored events for Outlook / Apple / Google Calendar.
3. Colleague Email Summary: Pre-formatted executive HTML/text for Outlook paste.
"""
import html
import json
import logging
import re
from datetime import datetime, timezone, timedelta
from pathlib import Path
from typing import Dict, List, Optional, Set, Tuple

from modules.meetings.core.session_plan_diff import SessionPlanDiffResult
from modules.meetings.core.session_plan_models import (
    DelegationMember,
    DelegationPlan,
    SessionAppointment,
    SessionPlanSnapshot,
    SessionSlot,
    TopicOwnership,
    UserSessionAnnotation,
    resolve_slot_effective_delegates,
    compute_unique_initials,
)
from modules.meetings.core.session_plan_parser import (
    DAY_NAMES,
    PERIOD_ORDER,
    PERIOD_TIMES,
)
from modules.meetings.core.session_plan_timezone import SessionPlanTimezone

logger = logging.getLogger(__name__)


class SessionPlanExporter:
    """Generates mobile HTML, .ics calendar, and email summaries."""

    @staticmethod
    def _is_slot_tracked(slot: SessionSlot, tracked_tokens: Set[str]) -> bool:
        """Determines whether a slot matches any user-tracked topic or agenda item."""
        if not tracked_tokens:
            return False
        raw_lower = slot.raw_text.lower()
        if any(tk in raw_lower for tk in tracked_tokens if tk):
            return True
        for it in slot.items:
            if it.ai_number.lower() in tracked_tokens or any(tk in it.topic_label.lower() for tk in tracked_tokens if tk):
                return True
        return False

    @staticmethod
    def get_slot_effective_delegates(slot: SessionSlot, plan: Optional[DelegationPlan]) -> Tuple[List[str], bool, bool]:
        """
        Resolves effective delegates list, is_coverage flag, and is_all flag: (delegates, is_cov, is_all).
        Delegates to the shared, boundary-aware resolver in session_plan_models so that the
        UI grid, clash detection, .ics and email exports always agree.
        """
        return resolve_slot_effective_delegates(slot, plan)

    @classmethod
    def get_slot_effective_delegate(cls, slot: SessionSlot, plan: Optional[DelegationPlan]) -> Tuple[Optional[str], bool]:
        """
        Backward-compatible single-string delegate helper.
        """
        dels, is_cov, is_all = cls.get_slot_effective_delegates(slot, plan)
        if not dels:
            return None, False
        if is_all:
            return "ALL", is_cov
        return ", ".join(dels), is_cov

    @classmethod
    def export_mobile_html(
        cls,
        snapshot: SessionPlanSnapshot,
        output_path: Path,
        meeting_tz_offset_min: int = 120,
        tz_name: str = "Venue Time",
        meeting_dates_str: str = "",
        annotations: Optional[Dict[str, UserSessionAnnotation]] = None,
        tracked_tokens: Optional[Set[str]] = None,
        delegation_plan: Optional[DelegationPlan] = None,
        agenda_map: Optional[dict] = None,
    ) -> Path:
        """
        Exports a self-contained, responsive, 100% offline HTML view featuring both:
        1. 🗓️ 5-Day Spatial Overview Grid (identical format to the desktop Overview Grid)
        2. 📱 Mobile Day Cards (day-by-day touch-friendly quarter view)
        With interactive delegate toggling, text filtering, stream toggles, and rich hover tooltips.
        No external CDN dependencies (fonts and CSS are purely embedded).
        """
        annotations = annotations or {}
        tracked_tokens = {t.lower() for t in (tracked_tokens or set())}
        agenda_map = agenda_map or {}
        output_path = Path(output_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)

        laptop_offset_min = SessionPlanTimezone.get_laptop_tz_offset_minutes()
        is_tz_different = (laptop_offset_min != meeting_tz_offset_min)
        tz_diff_hours = (laptop_offset_min - meeting_tz_offset_min) // 60
        tz_diff_str = ("+" if tz_diff_hours >= 0 else "") + str(tz_diff_hours) + "h"

        base_date = None
        if meeting_dates_str:
            try:
                base_date = datetime.strptime(meeting_dates_str, "%Y-%m-%d").date()
            except Exception:
                m_date = re.search(r"\b(\d{4}-\d{2}-\d{2})\b", meeting_dates_str)
                if m_date:
                    try:
                        base_date = datetime.strptime(m_date.group(1), "%Y-%m-%d").date()
                    except Exception:
                        pass

        venue_today = SessionPlanTimezone.get_venue_current_datetime(meeting_tz_offset_min).date()

        members = delegation_plan.members if delegation_plan else {}
        initials_map = compute_unique_initials(members.keys()) if members else {}

        # ------------------------------------------------------------------
        # Helper: Rich Tooltip Content Builder
        # ------------------------------------------------------------------
        def build_tooltip_html(slot: SessionSlot, day_idx: int, period: str, tu: int) -> str:
            day_name = DAY_NAMES[day_idx] if day_idx < len(DAY_NAMES) else f"Day {day_idx+1}"
            date_lbl = ""
            if base_date:
                cur_d = base_date + timedelta(days=day_idx)
                date_lbl = f" ({cur_d.strftime('%d %b')})"

            times = PERIOD_TIMES.get(period, ("09:00", "10:30", "09:00", "09:45", "09:45", "10:30"))
            if tu == 1:
                s_t, e_t = times[2], times[3]
                half_lbl = "1st Half (TU 1)"
            elif tu == 2:
                s_t, e_t = times[4], times[5]
                half_lbl = "2nd Half (TU 2)"
            else:
                s_t, e_t = times[0], times[1]
                half_lbl = "Full Quarter"

            time_dual = SessionPlanTimezone.format_dual_time_range(s_t, e_t, meeting_tz_offset_min)

            s_info = snapshot.streams.get(slot.stream_number)
            r_name = getattr(slot, "room_override", "") or (s_info.room_name if s_info else f"Stream {slot.stream_number}")
            cap = f" ({s_info.capacity} seats)" if s_info and s_info.capacity else ""
            convener = s_info.convener if s_info else ("SA2 Chair / MCC" if slot.stream_number == 1 else "")
            tohru = s_info.tohru_room if s_info else ""

            meta_parts = [f"🏢 <b>Stream {slot.stream_number}:</b> {html.escape(r_name)}{cap}"]
            if convener:
                meta_parts.append(f"👤 <b>Chair:</b> {html.escape(convener)}")
            if tohru:
                meta_parts.append(f"📡 <b>TOHRU:</b> {html.escape(tohru)}")

            ai_bullets = []
            for it in slot.items:
                if it.ai_number:
                    ai_obj = agenda_map.get(it.ai_number)
                    desc = getattr(ai_obj, "description", None) if ai_obj else it.topic_label
                    tdoc_str = f" [{it.tdoc_count_str}]" if it.tdoc_count_str else ""
                    ai_bullets.append(f"• <b>AI {it.ai_number}:</b> {html.escape(desc or it.topic_label)}{tdoc_str}")

            del_names, is_cov, is_all = cls.get_slot_effective_delegates(slot, delegation_plan)
            del_html = ""
            if is_all:
                cov_lbl = " (Ad-hoc coverage)" if is_cov else " (General / All Team Topic)"
                del_html = f"<div class='tt-section'>👥 <b>Delegation:</b> All Team{cov_lbl}</div>"
            elif del_names:
                role_lbl = " (Ad-hoc coverage)" if is_cov else " (Topic Owner)"
                badges = []
                for dn in del_names:
                    m = members.get(dn)
                    c = m.color if m else "#2563EB"
                    bg = m.bg_tint if m else "#EFF6FF"
                    bdr = m.border_tint if m else "#BFDBFE"
                    badges.append(f"<span style='background:{bg}; color:{c}; border:1px solid {bdr}; padding:1px 5px; border-radius:3px; font-weight:700;'>● {html.escape(dn)}</span>")
                del_html = f"<div class='tt-section'>👤 <b>Delegation:</b> {' '.join(badges)}{role_lbl}</div>"

            ann = annotations.get(slot.slot_key)
            note_html = ""
            if ann and (ann.comment or ann.assigned_colleague):
                parts = []
                if ann.comment:
                    parts.append(f"“{html.escape(ann.comment)}”")
                if ann.assigned_colleague:
                    parts.append(f"(Assigned: <b>{html.escape(ann.assigned_colleague)}</b>)")
                note_html = f"<div class='tt-section tt-note'>📝 <b>Notes:</b> {' '.join(parts)}</div>"

            tt = f"""<div class="tt-header">📅 {day_name}{date_lbl} · {period} · {half_lbl}</div>
<div class="tt-time">🕒 <b>Time:</b> {time_dual}</div>
<div class="tt-meta">{' &nbsp;|&nbsp; '.join(meta_parts)}</div>
<hr class="tt-sep">
<div class="tt-body"><b>Topic:</b> {html.escape(slot.raw_text)}</div>"""
            if ai_bullets:
                tt += f"<div class='tt-ais'>📑 <b>Agenda Items:</b><br>{'<br>'.join(ai_bullets)}</div>"
            if del_html:
                tt += del_html
            if note_html:
                tt += note_html
            return tt

        # ------------------------------------------------------------------
        # Data Mappings for Overview Grid & Mobile Day Cards
        # ------------------------------------------------------------------
        all_periods = list(PERIOD_ORDER)
        for s in snapshot.slots:
            if s.quarter and s.quarter not in all_periods:
                all_periods.append(s.quarter)

        slot_map: Dict[Tuple[int, str, int], List[SessionSlot]] = {}
        break_slot_map: Dict[Tuple[int, str], List[SessionSlot]] = {}
        slots_by_day_q: Dict[str, Dict[str, List[SessionSlot]]] = {}

        for s in snapshot.slots:
            slot_map.setdefault((s.day_index, s.quarter, s.stream_number), []).append(s)
            if any(b in s.quarter.lower() for b in ["break", "coffee", "lunch"]):
                break_slot_map.setdefault((s.day_index, s.quarter), []).append(s)
            slots_by_day_q.setdefault(s.day_name, {}).setdefault(s.quarter, []).append(s)

        active_break_periods = {
            s.quarter for s in snapshot.slots
            if any(b in s.quarter.lower() for b in ["break", "coffee", "lunch"])
            and s.raw_text.strip() and s.raw_text.strip() != "-"
        }

        # Build clean JSON events dataset for client-side .ics export & live status
        cal_events_data = []
        for s in snapshot.slots:
            if not s.raw_text.strip() or s.raw_text.strip() == "-":
                continue
            eff_dels, is_cov, is_all = cls.get_slot_effective_delegates(s, delegation_plan)
            s_info = snapshot.streams.get(s.stream_number)
            r_name = getattr(s, "room_override", "") or (s_info.room_name if s_info else f"Stream {s.stream_number}")
            convener = s_info.convener if s_info else ""
            ann = annotations.get(s.slot_key)
            note_str = ann.comment if ann else ""
            cal_events_data.append({
                "day_idx": s.day_index,
                "quarter": s.quarter,
                "sub_slot": s.sub_slot,
                "start_local": s.start_time_local,
                "end_local": s.end_time_local,
                "stream_num": s.stream_number,
                "room": r_name,
                "convener": convener,
                "title": s.raw_text,
                "is_drafting": s.is_drafting,
                "is_all": is_all,
                "is_cov": is_cov,
                "delegates": eff_dels,
                "note": note_str,
                "slot_key": s.slot_key,
            })
        cal_events_json = json.dumps(cal_events_data, ensure_ascii=False)
        base_date_js = base_date.strftime("%Y-%m-%d") if base_date else ""

        # ------------------------------------------------------------------
        # Helper: Render Overview Grid Slot Cell
        # ------------------------------------------------------------------
        def render_overview_cell(slot: SessionSlot, colspan: int, d_idx: int, p: str, s_num: int, tu: int) -> str:
            is_tracked = cls._is_slot_tracked(slot, tracked_tokens)
            eff_dels, is_cov, is_all = cls.get_slot_effective_delegates(slot, delegation_plan)
            dels_joined = ",".join(eff_dels)

            search_tokens = [slot.raw_text.lower()]
            for it in slot.items:
                if it.ai_number:
                    search_tokens.append(it.ai_number.lower())
                    search_tokens.append(f"ai {it.ai_number.lower()}")
                if it.topic_label:
                    search_tokens.append(it.topic_label.lower())
            ann = annotations.get(slot.slot_key)
            if ann and ann.comment:
                search_tokens.append(ann.comment.lower())

            raw_esc = html.escape(slot.raw_text)
            tt_html = build_tooltip_html(slot, d_idx, p, tu)

            cell_classes = ["grid-cell", "slot-cell", f"stream-{s_num}"]
            if slot.is_drafting:
                cell_classes.append("is-drafting")
            if is_tracked:
                cell_classes.append("is-tracked")

            hdr_badges = []
            if slot.is_drafting:
                hdr_badges.append('<span class="cell-badge badge-drafting">🛠️ Draft</span>')
            else:
                hdr_badges.append(f'<span class="cell-badge badge-str str-{s_num}">Str {s_num}</span>')

            if tu == 1:
                hdr_badges.append('<span class="cell-badge badge-tu">TU 1</span>')
            elif tu == 2:
                hdr_badges.append('<span class="cell-badge badge-tu">TU 2</span>')

            s_info = snapshot.streams.get(s_num)
            r_short = (s_info.room_name.split("(")[0].strip() if s_info else f"Str {s_num}")
            hdr_badges.append(f'<span class="cell-room">📍 {html.escape(r_short)}</span>')

            chips = []
            for it in slot.items:
                ai_lbl = f"AI {it.ai_number}" if it.ai_number else it.topic_label
                cnt = f'<span class="chip-count">[{it.tdoc_count_str}]</span>' if it.tdoc_count_str else ""
                chips.append(f'<span class="ai-chip">{html.escape(ai_lbl)}{cnt}</span>')

            del_badges = []
            if is_all:
                del_badges.append('<span class="del-pill-badge is-all" title="All Team">👥 All</span>')
            elif eff_dels:
                for dn in eff_dels:
                    m = members.get(dn)
                    ini = initials_map.get(dn, dn[:2].upper())
                    col = m.color if m else "#2563EB"
                    bg = m.bg_tint if m else "#EFF6FF"
                    bdr = m.border_tint if m else "#BFDBFE"
                    bdr_style = "dashed" if is_cov else "solid"
                    cov_note = " (Ad-hoc cover)" if is_cov else " (Owner)"
                    del_badges.append(f'<span class="del-pill-badge" style="background:{bg}; color:{col}; border:1px {bdr_style} {bdr};" title="{html.escape(dn)}{cov_note}">● {ini}</span>')

            note_badge = ""
            if ann and ann.comment:
                note_badge = f'<div class="cell-note-pill" title="{html.escape(ann.comment)}">📝 {html.escape(ann.comment[:35])}{"…" if len(ann.comment) > 35 else ""}</div>'

            return f"""<td class="{' '.join(cell_classes)}" colspan="{colspan}" data-stream="{s_num}" data-period="{p}" data-day="{d_idx}" data-delegates="{html.escape(dels_joined)}" data-is-all="{1 if is_all else 0}" data-is-cov="{1 if is_cov else 0}" data-tracked="{1 if is_tracked else 0}" data-drafting="{1 if slot.is_drafting else 0}" data-search="{html.escape(' '.join(search_tokens))}">
  <div class="cell-top-row">{' '.join(hdr_badges)}</div>
  <div class="cell-title">{raw_esc}</div>
  {f'<div class="cell-chips-row">{"".join(chips)}</div>' if chips else ''}
  {f'<div class="cell-del-row">{"".join(del_badges)}</div>' if del_badges else ''}
  {note_badge}
  <div class="slot-tooltip-data" style="display:none;">{tt_html}</div>
</td>"""

        # ------------------------------------------------------------------
        # Overview Grid Table HTML Assembly
        # ------------------------------------------------------------------
        overview_html = []
        overview_html.append('<div class="overview-grid-container"><table class="overview-table">\n')

        # Header Row 1: Day Names & Dates (colspan=2 per day)
        overview_html.append('<thead><tr class="hdr-days-row">\n')
        overview_html.append('<th class="corner-hdr" rowspan="2">Period / Stream<br><span class="hdr-sub">Venue & Laptop Times</span></th>\n')
        for d_idx, d_name in enumerate(DAY_NAMES):
            date_str = ""
            today_badge = ""
            if base_date:
                cur_d = base_date + timedelta(days=d_idx)
                date_str = f" <span class='hdr-date'>({cur_d.strftime('%d %b')})</span>"
                if cur_d == venue_today:
                    today_badge = " <span class='badge-today'>🟢 Today</span>"
            overview_html.append(f'<th class="col-day-hdr" colspan="2">{d_name}{date_str}{today_badge}</th>\n')
        overview_html.append('</tr>\n')

        # Header Row 2: Half-Quarter TU1 / TU2
        overview_html.append('<tr class="hdr-tu-row">\n')
        for d_idx in range(5):
            overview_html.append('<th class="col-tu-hdr">1st Half (TU 1)</th>\n<th class="col-tu-hdr">2nd Half (TU 2)</th>\n')
        overview_html.append('</tr></thead>\n<tbody>\n')

        # Body Rows
        streams = snapshot.streams
        for p in all_periods:
            times = PERIOD_TIMES.get(p, ("09:00", "10:30", "09:00", "09:45", "09:45", "10:30"))
            if any(b in p.lower() for b in ["break", "coffee", "lunch"]):
                clean_name = (
                    p.replace(" 1", "")
                    .replace(" 2", "")
                    .replace("Coffee", "☕ Coffee Break")
                    .replace("Break", "☕ Short Break")
                    .replace("Lunch", "🍱 Lunch Break")
                )
                is_empty_brk = (p not in active_break_periods)
                brk_cls = "grid-row break-row" + (" empty-break" if is_empty_brk else "")
                overview_html.append(f'<tr class="{brk_cls}" data-period="{p}">\n')
                overview_html.append(f'<th class="row-hdr break-hdr">{clean_name}<br><span class="hdr-times">{times[0]} – {times[1]}</span></th>\n')

                for d_idx in range(5):
                    matching_b = break_slot_map.get((d_idx, p), [])
                    valid_b = [s for s in matching_b if s.raw_text.strip() and s.raw_text.strip() != "-"]
                    if valid_b:
                        overview_html.append(render_overview_cell(valid_b[0], colspan=2, d_idx=d_idx, p=p, s_num=valid_b[0].stream_number or 1, tu=0))
                    else:
                        overview_html.append(f'<td class="grid-cell empty-break-cell" colspan="2" data-day="{d_idx}" data-period="{p}">—</td>\n')
                overview_html.append('</tr>\n')
            else:
                for s_num in [1, 2, 3]:
                    s_info = streams.get(s_num)
                    r_name = s_info.room_name if s_info else ("Main" if s_num == 1 else f"Brk {s_num-1}")
                    short_r = r_name.split("(")[0].strip()

                    overview_html.append(f'<tr class="grid-row stream-row stream-row-{s_num}" data-stream="{s_num}" data-period="{p}">\n')
                    overview_html.append(f'<th class="row-hdr stream-hdr-s{s_num}">{p} · Str {s_num} <span class="hdr-room">({html.escape(short_r)})</span><br><span class="hdr-times">{times[0]} – {times[1]}</span></th>\n')

                    for d_idx in range(5):
                        matching = slot_map.get((d_idx, p, s_num), [])
                        full_slot = next((s for s in matching if s.sub_slot == 0), None)
                        tu1_slot = next((s for s in matching if s.sub_slot == 1), None)
                        tu2_slot = next((s for s in matching if s.sub_slot == 2), None)

                        if full_slot:
                            overview_html.append(render_overview_cell(full_slot, colspan=2, d_idx=d_idx, p=p, s_num=s_num, tu=0))
                        elif tu1_slot or tu2_slot:
                            if tu1_slot:
                                overview_html.append(render_overview_cell(tu1_slot, colspan=1, d_idx=d_idx, p=p, s_num=s_num, tu=1))
                            else:
                                overview_html.append(f'<td class="grid-cell empty-slot-cell" colspan="1" data-day="{d_idx}" data-period="{p}" data-stream="{s_num}">—</td>\n')

                            if tu2_slot:
                                overview_html.append(render_overview_cell(tu2_slot, colspan=1, d_idx=d_idx, p=p, s_num=s_num, tu=2))
                            else:
                                overview_html.append(f'<td class="grid-cell empty-slot-cell" colspan="1" data-day="{d_idx}" data-period="{p}" data-stream="{s_num}">—</td>\n')
                        else:
                            overview_html.append(f'<td class="grid-cell empty-slot-cell" colspan="2" data-day="{d_idx}" data-period="{p}" data-stream="{s_num}">—</td>\n')
                    overview_html.append('</tr>\n')

        overview_html.append('</tbody></table></div>\n')

        # ------------------------------------------------------------------
        # Mobile Day Cards View HTML Assembly
        # ------------------------------------------------------------------
        cards_html = []
        active_days = [d for d in DAY_NAMES if d in slots_by_day_q]

        cards_html.append('<div class="day-tabs">\n')
        for idx, d_name in enumerate(active_days):
            active_cls = " active" if idx == 0 else ""
            cards_html.append(f'<button class="day-tab{active_cls}" onclick="switchDay(\'{d_name}\')">{d_name[:3]}</button>\n')
        cards_html.append('</div>\n<div class="cards-content">\n')

        for idx, d_name in enumerate(active_days):
            d_idx = DAY_NAMES.index(d_name) if d_name in DAY_NAMES else idx
            day_display = "" if idx == 0 else ' style="display:none;"'
            cards_html.append(f'<div id="day-{d_name}" class="day-section"{day_display}>\n')

            q_map = slots_by_day_q.get(d_name, {})
            for q in PERIOD_ORDER:
                if q not in q_map:
                    continue
                slots_in_q = q_map[q]
                if not slots_in_q:
                    continue

                sample_s = slots_in_q[0]
                q_time_str = SessionPlanTimezone.format_dual_time_range(
                    sample_s.start_time_local, sample_s.end_time_local, meeting_tz_offset_min
                )

                cards_html.append(f"""  <div class="quarter-card">
    <div class="quarter-header">
      <span class="quarter-title">{q}</span>
      <span class="quarter-time">{q_time_str}</span>
    </div>
""")

                for s in slots_in_q:
                    is_tracked = cls._is_slot_tracked(s, tracked_tokens)
                    eff_dels, is_cov, is_all = cls.get_slot_effective_delegates(s, delegation_plan)
                    dels_joined = ",".join(eff_dels)

                    search_tokens = [s.raw_text.lower()]
                    for it in s.items:
                        if it.ai_number:
                            search_tokens.append(it.ai_number.lower())
                        if it.topic_label:
                            search_tokens.append(it.topic_label.lower())
                    ann = annotations.get(s.slot_key)
                    if ann and ann.comment:
                        search_tokens.append(ann.comment.lower())

                    raw_esc = html.escape(s.raw_text)
                    tt_html = build_tooltip_html(s, d_idx, q, s.sub_slot)

                    s_info = streams.get(s.stream_number)
                    r_name = getattr(s, "room_override", "") or (s_info.room_name if s_info else f"Stream {s.stream_number}")
                    convener = s_info.convener if s_info else ""

                    sub_tag = ""
                    if s.sub_slot == 1:
                        sub_tag = f'<span class="sub-slot-badge">TU 1 ({s.start_time_local}–{s.end_time_local})</span>'
                    elif s.sub_slot == 2:
                        sub_tag = f'<span class="sub-slot-badge">TU 2 ({s.start_time_local}–{s.end_time_local})</span>'

                    chips = []
                    for it in s.items:
                        cnt = f'<span class="count-chip">[{it.tdoc_count_str}]</span>' if it.tdoc_count_str else ""
                        ai_lbl = f"AI {it.ai_number}" if it.ai_number else it.topic_label
                        chips.append(f'<span class="ai-chip">{html.escape(ai_lbl)}{cnt}</span>')

                    del_badge = ""
                    if is_all:
                        cov_lbl = " (cov)" if is_cov else ""
                        del_badge = f'<span class="del-pill-badge is-all">👥 All Team{cov_lbl}</span>'
                    elif eff_dels:
                        cov_lbl = " (cov)" if is_cov else ""
                        b_list = []
                        for ed in eff_dels:
                            m = members.get(ed)
                            col = m.color if m else "#2563EB"
                            bg = m.bg_tint if m else "#EFF6FF"
                            bdr = m.border_tint if m else "#BFDBFE"
                            b_list.append(f'<span class="del-pill-badge" style="background:{bg}; color:{col}; border:1px solid {bdr};">● {html.escape(ed)}{cov_lbl}</span>')
                        del_badge = "".join(b_list)

                    note_html = ""
                    if ann and ann.comment:
                        note_html = f'<div class="note-chip">📝 <b>Note:</b> {html.escape(ann.comment)}</div>'

                    row_classes = ["cards-slot-row"]
                    if s.is_drafting:
                        row_classes.append("drafting")
                    if is_tracked:
                        row_classes.append("tracked")

                    cards_html.append(f"""    <div class="{' '.join(row_classes)}" data-stream="{s.stream_number}" data-period="{q}" data-day="{d_idx}" data-delegates="{html.escape(dels_joined)}" data-is-all="{1 if is_all else 0}" data-is-cov="{1 if is_cov else 0}" data-tracked="{1 if is_tracked else 0}" data-drafting="{1 if s.is_drafting else 0}" data-search="{html.escape(' '.join(search_tokens))}">
      <div class="slot-header-line">
        <span class="stream-badge str-{s.stream_number}">Stream {s.stream_number}</span>
        <span class="room-pill">📍 {html.escape(r_name)}{" · " + html.escape(convener) if convener else ""}</span>
        {sub_tag}
        {del_badge}
      </div>
      <div class="slot-content">{raw_esc}</div>
      {f'<div class="chips-container">{"".join(chips)}</div>' if chips else ''}
      {note_html}
      <div class="slot-tooltip-data" style="display:none;">{tt_html}</div>
    </div>
""")

                cards_html.append("  </div>\n")  # Close quarter-card
            cards_html.append("</div>\n")  # Close day-section
        cards_html.append("</div>\n")  # Close cards-content

        # ------------------------------------------------------------------
        # Delegate Selector Buttons HTML
        # ------------------------------------------------------------------
        del_filter_buttons = []
        for name, m in members.items():
            del_filter_buttons.append(
                f'<button class="del-filter-btn" data-del="{html.escape(name)}" '
                f'style="--del-color:{m.color}; --del-bg:{m.bg_tint}; --del-bdr:{m.border_tint};" '
                f'onclick="selectDelegate(\'{html.escape(name)}\')">● {html.escape(name)}</button>'
            )
        del_filter_html = "\n".join(del_filter_buttons)

        # ------------------------------------------------------------------
        # Document Synthesis
        # ------------------------------------------------------------------
        css_style = """
  :root {
    --bg: #F8FAFC;
    --card: #FFFFFF;
    --text: #0F172A;
    --subtext: #475569;
    --primary: #2563EB;
    --primary-light: #EFF6FF;
    --stream1: #4F46E5;
    --stream2: #059669;
    --stream3: #D97706;
    --drafting: #DC2626;
    --drafting-bg: #FEF2F2;
    --tracked: #7C3AED;
    --tracked-bg: #F5F3FF;
    --border: #E2E8F0;
    --grid-border: #CBD5E1;
  }
  * { box-sizing: border-box; -webkit-tap-highlight-color: transparent; }
  body {
    margin: 0; padding: 0; background: var(--bg); color: var(--text);
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
    font-size: 13px; line-height: 1.4;
  }
  header {
    position: sticky; top: 0; z-index: 100; background: #FFFFFF;
    border-bottom: 1px solid var(--border); padding: 8px 14px;
    box-shadow: 0 1px 3px rgba(0,0,0,0.06);
  }
  .header-top {
    display: flex; justify-content: space-between; align-items: center; margin-bottom: 6px; flex-wrap: wrap; gap: 8px;
  }
  .title-block { display: flex; align-items: center; gap: 8px; flex-wrap: wrap; }
  .title { font-size: 16px; font-weight: 700; color: #1E293B; margin: 0; }
  .badge-tdoc { font-size: 11px; background: #E0E7FF; color: #3730A3; padding: 2px 8px; border-radius: 999px; font-weight: 700; }
  .badge-snap { font-size: 11px; background: #FEF3C7; color: #92400E; padding: 2px 7px; border-radius: 999px; font-weight: 600; }
  .header-actions { display: flex; align-items: center; gap: 8px; flex-wrap: wrap; }
  .btn-cal-download {
    padding: 5px 12px; font-size: 12px; font-weight: 600; border-radius: 6px;
    background: #059669; color: #FFFFFF; border: 1px solid #047857; cursor: pointer;
    transition: all 0.15s ease; display: inline-flex; align-items: center; gap: 4px;
  }
  .btn-cal-download:hover { background: #047857; }
  .badge-clash {
    font-size: 11px; font-weight: 700; background: #FEF2F2; color: #B91C1C;
    border: 1px solid #FECACA; padding: 3px 8px; border-radius: 999px;
    display: inline-flex; align-items: center; gap: 4px;
  }
  .live-banner {
    display: flex; align-items: center; gap: 8px; font-size: 11.5px; font-weight: 600;
    padding: 5px 12px; border-radius: 6px; background: #ECFDF5; color: #065F46;
    border: 1px solid #A7F3D0; margin-top: 4px; margin-bottom: 4px;
  }
  .pulse-dot {
    width: 8px; height: 8px; border-radius: 50%; background: #10B981;
    display: inline-block; animation: pulseLive 1.8s infinite;
  }
  @keyframes pulseLive {
    0% { transform: scale(0.95); box-shadow: 0 0 0 0 rgba(16, 185, 129, 0.7); }
    70% { transform: scale(1.1); box-shadow: 0 0 0 6px rgba(16, 185, 129, 0); }
    100% { transform: scale(0.95); box-shadow: 0 0 0 0 rgba(16, 185, 129, 0); }
  }
  .grid-cell.has-clash {
    outline: 2px dashed #EF4444 !important; outline-offset: -2px;
  }
  .cards-slot-row.has-clash {
    border-left: 4px solid #EF4444 !important; background: #FFF1F2 !important;
  }
  .grid-cell.is-live-now {
    outline: 2px solid #10B981 !important; outline-offset: -2px; background: #F0FDF4 !important;
  }
  .cards-slot-row.is-live-now {
    border-left: 4px solid #10B981 !important; background: #F0FDF4 !important;
  }
  .view-switcher {
    display: inline-flex; background: #F1F5F9; padding: 2px; border-radius: 8px; border: 1px solid var(--border);
  }
  .view-btn {
    padding: 5px 12px; font-size: 12px; font-weight: 600; border: none; background: transparent;
    color: var(--subtext); border-radius: 6px; cursor: pointer; transition: all 0.15s ease;
  }
  .view-btn.active {
    background: #FFFFFF; color: var(--primary); box-shadow: 0 1px 3px rgba(0,0,0,0.1);
  }
  .tz-banner { font-size: 11px; color: var(--subtext); margin-bottom: 6px; }
  
  /* Delegate Selector Strip */
  .del-strip {
    display: flex; gap: 5px; align-items: center; overflow-x: auto; padding: 4px 0 6px 0;
    -webkit-overflow-scrolling: touch;
  }
  .strip-label { font-size: 11px; font-weight: 700; color: #64748B; text-transform: uppercase; margin-right: 2px; }
  .del-filter-btn {
    padding: 3px 10px; font-size: 11.5px; font-weight: 600; border-radius: 999px; cursor: pointer;
    background: #F1F5F9; color: #334155; border: 1px solid var(--border); white-space: nowrap;
    transition: all 0.15s ease;
  }
  .del-filter-btn[data-del="ALL"].active {
    background: #2563EB; color: #FFFFFF; border-color: #2563EB;
  }
  .del-filter-btn[data-del]:not([data-del="ALL"]) {
    background: var(--del-bg, #F1F5F9); color: var(--del-color, #334155); border-color: var(--del-bdr, var(--border));
  }
  .del-filter-btn[data-del]:not([data-del="ALL"]).active {
    background: var(--del-color, #2563EB); color: #FFFFFF; border-color: var(--del-color, #2563EB);
    box-shadow: 0 2px 6px rgba(0,0,0,0.15);
  }
  .strip-sep { width: 1px; height: 18px; background: var(--border); margin: 0 3px; }
  .btn-toggle { background: #FFFFFF; }
  .btn-toggle.active { background: #7C3AED; color: #FFFFFF; border-color: #7C3AED; }
  #btnDrafting.active { background: #DC2626; color: #FFFFFF; border-color: #DC2626; }

  /* Filter Controls Row */
  .filter-controls-row {
    display: flex; gap: 8px; align-items: center; flex-wrap: wrap; margin-top: 4px;
  }
  .search-box {
    flex: 1; min-width: 200px; padding: 6px 10px; border: 1px solid var(--border); border-radius: 6px;
    font-size: 12.5px; outline: none; background: #F8FAFC;
  }
  .search-box:focus { background: #FFFFFF; border-color: var(--primary); }
  .grid-row-controls { display: flex; gap: 8px; align-items: center; font-size: 11.5px; }
  .toggle-label { display: inline-flex; align-items: center; gap: 3px; cursor: pointer; color: var(--subtext); font-weight: 500; }
  .toggle-label input { cursor: pointer; }

  /* Overview Grid View Styles */
  .overview-grid-container {
    padding: 10px 14px 60px 14px; overflow-x: auto; -webkit-overflow-scrolling: touch;
  }
  .overview-table {
    border-collapse: separate; border-spacing: 0; width: 100%; min-width: 1100px;
    background: #FFFFFF; border: 1px solid var(--grid-border); border-radius: 8px;
    box-shadow: 0 1px 3px rgba(0,0,0,0.04);
  }
  .overview-table th, .overview-table td {
    border-right: 1px solid var(--border); border-bottom: 1px solid var(--border);
    padding: 6px 8px; vertical-align: top; font-size: 11.5px;
  }
  .overview-table th:last-child, .overview-table td:last-child { border-right: none; }
  .overview-table tr:last-child td { border-bottom: none; }

  .hdr-days-row th {
    background: #F1F5F9; color: #1E293B; font-weight: 700; text-align: center;
    position: sticky; top: 0; z-index: 10; border-bottom: 1px solid var(--grid-border);
  }
  .corner-hdr {
    background: #E2E8F0 !important; text-align: left !important; font-size: 11px;
    position: sticky; left: 0; z-index: 20 !important; width: 160px; min-width: 160px;
  }
  .hdr-sub { font-size: 9.5px; font-weight: 500; color: #64748B; }
  .hdr-date { font-size: 10.5px; font-weight: 600; color: #475569; }
  .badge-today {
    display: inline-block; font-size: 9px; font-weight: 700; background: #DCFCE7;
    color: #15803D; padding: 1px 5px; border-radius: 4px; margin-left: 4px;
  }
  .hdr-tu-row th {
    background: #F8FAFC; color: #64748B; font-size: 10px; font-weight: 600; text-align: center;
    border-bottom: 1px solid var(--grid-border);
  }
  .row-hdr {
    background: #F8FAFC; color: #1E293B; font-weight: 600; text-align: left;
    position: sticky; left: 0; z-index: 5; width: 160px; min-width: 160px; font-size: 11px;
  }
  .hdr-room { font-size: 10px; font-weight: 500; color: #64748B; }
  .hdr-times { font-size: 9.5px; font-weight: 500; color: #64748B; }
  .stream-hdr-s1 { border-left: 3px solid var(--stream1); }
  .stream-hdr-s2 { border-left: 3px solid var(--stream2); }
  .stream-hdr-s3 { border-left: 3px solid var(--stream3); }
  .break-hdr { background: #F8FAFC; font-weight: 500; color: #475569; border-left: 3px solid #94A3B8; }

  /* Grid Cells */
  .grid-cell {
    background: #FFFFFF; transition: opacity 0.2s ease, filter 0.2s ease, box-shadow 0.2s ease;
    cursor: default; position: relative;
  }
  .grid-cell.slot-cell:hover {
    background: #F8FAFC;
  }
  .grid-cell.empty-slot-cell {
    text-align: center; color: #CBD5E1; font-weight: 500; background: #FAFAFA;
  }
  .grid-cell.empty-break-cell {
    text-align: center; color: #CBD5E1; font-weight: 500; background: #F8FAFC;
  }
  .grid-cell.is-drafting {
    background: var(--drafting-bg); border-left: 3px solid var(--drafting) !important;
  }
  .grid-cell.is-tracked {
    background: var(--tracked-bg); border-left: 3px solid var(--tracked) !important;
  }

  /* Delegate Dimming / Highlighting */
  .grid-cell.dimmed {
    opacity: 0.20; filter: grayscale(85%);
  }
  .grid-cell.highlighted {
    box-shadow: inset 0 0 0 2px var(--primary);
    background: #FFFFFF !important;
  }

  /* Cell Contents */
  .cell-top-row { display: flex; gap: 4px; align-items: center; margin-bottom: 3px; flex-wrap: wrap; }
  .cell-badge {
    font-size: 9px; font-weight: 700; padding: 1px 4px; border-radius: 3px; text-transform: uppercase;
  }
  .badge-str.str-1 { background: #EEF2FF; color: var(--stream1); }
  .badge-str.str-2 { background: #ECFDF5; color: var(--stream2); }
  .badge-str.str-3 { background: #FFFBEB; color: var(--stream3); }
  .badge-drafting { background: #FEE2E2; color: var(--drafting); }
  .badge-tu { background: #E2E8F0; color: #334155; }
  .cell-room { font-size: 10px; color: #64748B; font-weight: 500; }
  .cell-title { font-size: 11.5px; font-weight: 600; color: #0F172A; line-height: 1.35; margin-bottom: 3px; }
  .cell-chips-row { display: flex; gap: 3px; flex-wrap: wrap; margin-bottom: 3px; }
  .ai-chip {
    font-size: 10px; font-weight: 600; padding: 1px 5px; border-radius: 3px; background: #E0E7FF; color: #3730A3;
  }
  .chip-count { font-size: 9px; font-weight: 700; color: #475569; margin-left: 2px; }
  .cell-del-row { display: flex; gap: 3px; flex-wrap: wrap; margin-top: 2px; }
  .del-pill-badge {
    font-size: 9px; font-weight: 700; padding: 1px 5px; border-radius: 4px; display: inline-flex; align-items: center;
  }
  .del-pill-badge.is-all {
    background: #EEF2FF; color: #4338CA; border: 1px solid #C7D2FE;
  }
  .cell-note-pill {
    margin-top: 3px; font-size: 9.5px; font-weight: 500; color: #92400E; background: #FEF3C7;
    padding: 1px 4px; border-radius: 3px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;
  }

  /* Mobile Day Cards View Styles */
  .day-tabs {
    display: flex; gap: 6px; overflow-x: auto; padding: 8px 14px; background: #FFFFFF;
    border-bottom: 1px solid var(--border); -webkit-overflow-scrolling: touch;
  }
  .day-tab {
    padding: 6px 14px; border-radius: 999px; font-weight: 600; font-size: 13px;
    background: #F1F5F9; color: var(--subtext); border: none; cursor: pointer; white-space: nowrap;
  }
  .day-tab.active { background: var(--primary); color: #FFFFFF; }
  .cards-content { padding: 12px 14px 60px 14px; }
  .quarter-card {
    background: var(--card); border: 1px solid var(--border); border-radius: 12px;
    margin-bottom: 12px; overflow: hidden; box-shadow: 0 1px 2px rgba(0,0,0,0.03);
  }
  .quarter-header {
    background: #F8FAFC; padding: 8px 12px; border-bottom: 1px solid var(--border);
    display: flex; justify-content: space-between; align-items: center;
  }
  .quarter-title { font-size: 13px; font-weight: 700; color: #1E293B; }
  .quarter-time { font-size: 12px; color: var(--subtext); }
  .cards-slot-row {
    padding: 10px 12px; border-bottom: 1px solid #F1F5F9; transition: background 0.15s ease;
  }
  .cards-slot-row:last-child { border-bottom: none; }
  .cards-slot-row.drafting { background: var(--drafting-bg); border-left: 4px solid var(--drafting); }
  .cards-slot-row.tracked { background: var(--tracked-bg); border-left: 4px solid var(--tracked); }
  .slot-header-line { display: flex; gap: 6px; align-items: center; flex-wrap: wrap; margin-bottom: 4px; }
  .stream-badge {
    font-size: 10px; font-weight: 700; padding: 2px 6px; border-radius: 4px; text-transform: uppercase;
  }
  .stream-badge.str-1 { background: #EEF2FF; color: var(--stream1); }
  .stream-badge.str-2 { background: #ECFDF5; color: var(--stream2); }
  .stream-badge.str-3 { background: #FFFBEB; color: var(--stream3); }
  .room-pill { font-size: 11px; color: var(--subtext); font-weight: 500; }
  .sub-slot-badge {
    font-size: 10px; font-weight: 600; padding: 1px 5px; border-radius: 3px; background: #E2E8F0; color: #334155;
  }
  .slot-content { font-size: 13px; font-weight: 500; color: #1E293B; margin-top: 3px; }
  .chips-container { margin-top: 4px; display: flex; gap: 4px; flex-wrap: wrap; }
  .note-chip {
    display: block; margin-top: 5px; padding: 4px 8px; border-radius: 6px;
    background: #FEF3C7; color: #92400E; font-size: 11px; font-weight: 500;
  }
  .hidden { display: none !important; }

  /* Floating Rich Tooltip Popover */
  .floating-tooltip {
    position: fixed; z-index: 9999; pointer-events: none; max-width: 420px;
    background: #FFFFFF; border: 1px solid #CBD5E1; border-radius: 8px;
    box-shadow: 0 10px 25px rgba(0,0,0,0.18); padding: 10px 14px;
    font-family: inherit; font-size: 12px; color: #0F172A; display: none;
  }
  .tt-header { font-weight: 700; font-size: 13px; color: #0F172A; margin-bottom: 3px; }
  .tt-time { color: var(--subtext); font-size: 11.5px; margin-bottom: 2px; }
  .tt-meta { color: var(--subtext); font-size: 11px; margin-bottom: 4px; }
  .tt-sep { border: none; border-top: 1px solid #E2E8F0; margin: 6px 0; }
  .tt-body { font-weight: 600; font-size: 12px; margin-bottom: 4px; }
  .tt-ais { font-size: 11px; color: #334155; line-height: 1.4; margin-top: 4px; }
  .tt-section { margin-top: 4px; font-size: 11px; }
  .tt-note { background: #FEF3C7; color: #92400E; padding: 3px 6px; border-radius: 4px; }

  /* Print Stylesheet */
  @media print {
    header, .del-strip, .filter-controls-row, .view-switcher, .day-tabs { display: none !important; }
    #cardsView { display: none !important; }
    #overviewView { display: block !important; padding: 0 !important; }
    .overview-table { width: 100% !important; min-width: 0 !important; font-size: 8.5px !important; box-shadow: none !important; }
    .overview-table th, .overview-table td { padding: 3px 4px !important; }
    .slot-cell { opacity: 1 !important; filter: none !important; box-shadow: none !important; }
    @page { size: landscape; margin: 8mm; }
  }
"""

        js_script = """
  let currentView = 'grid';
  let selectedDelegate = 'ALL';
  let onlyTracked = false;
  let onlyDrafting = false;

  function switchView(viewName) {
    currentView = viewName;
    document.getElementById('btnViewGrid').classList.toggle('active', viewName === 'grid');
    document.getElementById('btnViewCards').classList.toggle('active', viewName === 'cards');
    document.getElementById('overviewView').style.display = (viewName === 'grid') ? 'block' : 'none';
    document.getElementById('cardsView').style.display = (viewName === 'cards') ? 'block' : 'none';
    document.getElementById('gridControls').style.display = (viewName === 'grid') ? 'flex' : 'none';
    applyFilters();
  }

  function selectDelegate(delName) {
    selectedDelegate = delName;
    document.querySelectorAll('.del-filter-btn[data-del]').forEach(b => {
      b.classList.toggle('active', b.getAttribute('data-del') === delName);
    });
    const dlBtn = document.getElementById('btnDownloadIcs');
    if (dlBtn) {
      dlBtn.textContent = (delName === 'ALL') ? '📅 Download Team .ics' : ('📅 ' + delName + ' .ics');
    }
    applyFilters();
    detectDelegateClashes(delName);
  }

  function toggleTracked() {
    onlyTracked = !onlyTracked;
    document.getElementById('btnTracked').classList.toggle('active', onlyTracked);
    applyFilters();
  }

  function toggleDrafting() {
    onlyDrafting = !onlyDrafting;
    document.getElementById('btnDrafting').classList.toggle('active', onlyDrafting);
    applyFilters();
  }

  function toggleStream(num) {
    const chk = document.getElementById('chkStr' + num);
    const rows = document.querySelectorAll('.stream-row-' + num);
    rows.forEach(r => r.style.display = chk.checked ? '' : 'none');
  }

  function toggleBreaks() {
    const chk = document.getElementById('chkBreaks');
    const rows = document.querySelectorAll('.break-row');
    rows.forEach(r => r.style.display = chk.checked ? '' : 'none');
  }

  function toggleEmptyBreaks() {
    const chk = document.getElementById('chkEmptyBreaks');
    const rows = document.querySelectorAll('.break-row.empty-break');
    rows.forEach(r => {
      if (chk.checked) {
        r.style.display = 'none';
      } else {
        const showAll = document.getElementById('chkBreaks').checked;
        r.style.display = showAll ? '' : 'none';
      }
    });
  }

  function switchDay(dayName) {
    document.querySelectorAll('.day-tab').forEach(b => b.classList.remove('active'));
    document.querySelectorAll('.day-section').forEach(s => s.style.display = 'none');
    event.target.classList.add('active');
    const target = document.getElementById('day-' + dayName);
    if (target) target.style.display = 'block';
  }

  function applyFilters() {
    const query = (document.getElementById('searchInput').value || '').toLowerCase().trim();

    // 1. Filter Overview Grid Cells
    const gridCells = document.querySelectorAll('.grid-cell.slot-cell');
    gridCells.forEach(cell => {
      const sTxt = cell.getAttribute('data-search') || '';
      const isTrk = cell.getAttribute('data-tracked') === '1';
      const isDrf = cell.getAttribute('data-drafting') === '1';
      const dels = (cell.getAttribute('data-delegates') || '').split(',').map(d => d.trim().toLowerCase());
      const isAll = cell.getAttribute('data-is-all') === '1';

      const matchesSearch = !query || sTxt.includes(query);
      const matchesTracked = !onlyTracked || isTrk;
      const matchesDrafting = !onlyDrafting || isDrf;
      const matchesDelegate = (selectedDelegate === 'ALL') || isAll || dels.includes(selectedDelegate.toLowerCase());

      const isMatch = matchesSearch && matchesTracked && matchesDrafting && matchesDelegate;

      if (selectedDelegate === 'ALL' && !query && !onlyTracked && !onlyDrafting) {
        cell.classList.remove('dimmed', 'highlighted');
      } else if (isMatch) {
        cell.classList.remove('dimmed');
        cell.classList.add('highlighted');
      } else {
        cell.classList.add('dimmed');
        cell.classList.remove('highlighted');
      }
    });

    // 2. Filter Cards View Rows
    const cardRows = document.querySelectorAll('.cards-slot-row');
    cardRows.forEach(row => {
      const sTxt = row.getAttribute('data-search') || '';
      const isTrk = row.getAttribute('data-tracked') === '1';
      const isDrf = row.getAttribute('data-drafting') === '1';
      const dels = (row.getAttribute('data-delegates') || '').split(',').map(d => d.trim().toLowerCase());
      const isAll = row.getAttribute('data-is-all') === '1';

      const matchesSearch = !query || sTxt.includes(query);
      const matchesTracked = !onlyTracked || isTrk;
      const matchesDrafting = !onlyDrafting || isDrf;
      const matchesDelegate = (selectedDelegate === 'ALL') || isAll || dels.includes(selectedDelegate.toLowerCase());

      if (matchesSearch && matchesTracked && matchesDrafting && matchesDelegate) {
        row.classList.remove('hidden');
      } else {
        row.classList.add('hidden');
      }
    });

    // Hide empty quarter cards
    document.querySelectorAll('.quarter-card').forEach(card => {
      const visible = card.querySelectorAll('.cards-slot-row:not(.hidden)');
      card.style.display = (visible.length > 0) ? 'block' : 'none';
    });
  }

  // Floating Rich Tooltip Implementation
  const tooltip = document.getElementById('floatingTooltip');

  document.querySelectorAll('.slot-cell, .cards-slot-row').forEach(cell => {
    const dataEl = cell.querySelector('.slot-tooltip-data');
    if (!dataEl) return;

    cell.addEventListener('mouseenter', e => {
      tooltip.innerHTML = dataEl.innerHTML;
      tooltip.style.display = 'block';
      positionTooltip(e);
    });

    cell.addEventListener('mousemove', e => {
      positionTooltip(e);
    });

    cell.addEventListener('mouseleave', () => {
      tooltip.style.display = 'none';
    });
  });

  function positionTooltip(e) {
    const w = tooltip.offsetWidth;
    const h = tooltip.offsetHeight;
    let x = e.clientX + 14;
    let y = e.clientY + 14;
    if (x + w > window.innerWidth - 12) x = window.innerWidth - w - 12;
    if (y + h > window.innerHeight - 12) y = window.innerHeight - h - 12;
    if (x < 10) x = 10;
    if (y < 10) y = 10;
    tooltip.style.left = x + 'px';
    tooltip.style.top = y + 'px';
  }

  // RFC 5545 Client-Side .ics Calendar Generation & Download
  function venueToUtcIso(baseDateStr, dayIdx, timeStr, offsetMin) {
    if (!baseDateStr) return null;
    const parts = baseDateStr.split('-');
    const year = parseInt(parts[0], 10);
    const month = parseInt(parts[1], 10) - 1;
    const day = parseInt(parts[2], 10) + dayIdx;
    const tParts = (timeStr || '09:00').split(':');
    const h = parseInt(tParts[0], 10);
    const m = parseInt(tParts[1], 10);

    const utcMs = Date.UTC(year, month, day, h, m) - (offsetMin * 60000);
    const d = new Date(utcMs);
    const pad = n => String(n).padStart(2, '0');
    return d.getUTCFullYear() +
           pad(d.getUTCMonth() + 1) +
           pad(d.getUTCDate()) + 'T' +
           pad(d.getUTCHours()) +
           pad(d.getUTCMinutes()) + '00Z';
  }

  function downloadSelectedCalendar() {
    if (!MEETING_BASE_DATE) {
      alert('Meeting start date is not configured for this snapshot; cannot calculate physical calendar dates.');
      return;
    }
    const isAll = (selectedDelegate === 'ALL');
    const filterName = isAll ? 'Whole Team' : selectedDelegate;
    const safeName = isAll ? 'Team' : selectedDelegate.replace(/[^a-zA-Z0-9_-]/g, '_');
    const nowUtc = new Date().toISOString().replace(/[-:]/g, '').split('.')[0] + 'Z';

    const calLines = [
      'BEGIN:VCALENDAR',
      'VERSION:2.0',
      'PRODID:-//3GPP Meeting Tools//SA2 #' + MEETING_NUMBER + ' Session Plan//EN',
      'CALSCALE:GREGORIAN',
      'METHOD:PUBLISH',
      'X-WR-CALNAME:SA2 #' + MEETING_NUMBER + ' - ' + filterName + ' (' + TDOC_ID + ')'
    ];

    let count = 0;
    for (const ev of CAL_EVENTS) {
      const isSlotAll = ev.is_all;
      const effDels = ev.delegates || [];
      const delMatch = isAll || isSlotAll || effDels.some(d => d.toLowerCase() === selectedDelegate.toLowerCase());
      if (!delMatch) continue;

      const dtStart = venueToUtcIso(MEETING_BASE_DATE, ev.day_idx, ev.start_local, MEETING_TZ_OFFSET_MIN);
      const dtEnd = venueToUtcIso(MEETING_BASE_DATE, ev.day_idx, ev.end_local, MEETING_TZ_OFFSET_MIN);
      if (!dtStart || !dtEnd) continue;

      const safeKey = (ev.slot_key || ('slot-' + ev.day_idx + '-' + ev.quarter + '-' + ev.stream_num)).replace(/[^a-zA-Z0-9_-]/g, '_');
      const uid = 'sa2-' + MEETING_NUMBER + '-' + safeName + '-' + safeKey + '-' + dtStart + '@3gpp.tools';

      let delTag = '';
      if (isSlotAll) {
        delTag = ev.is_cov ? '[All Team (cov)] ' : '[All Team] ';
      } else if (effDels.length > 0) {
        delTag = '[' + effDels.join(', ') + (ev.is_cov ? ' (cov)' : '') + '] ';
      }

      const summaryPrefix = ev.is_drafting ? '🛠️ [Drafting] ' + delTag : '[Str ' + ev.stream_num + '] ' + delTag;
      const summary = summaryPrefix + (ev.title || '').substring(0, 60);

      const descParts = [
        'Session: SA2 #' + MEETING_NUMBER + ' (' + TDOC_ID + ')',
        'Room: ' + ev.room,
        'Convener: ' + (ev.convener || 'N/A'),
        'Scheduled: ' + ev.title
      ];
      if (ev.note) descParts.push('Note: ' + ev.note);

      calLines.push('BEGIN:VEVENT');
      calLines.push('UID:' + uid);
      calLines.push('DTSTAMP:' + nowUtc);
      calLines.push('DTSTART:' + dtStart);
      calLines.push('DTEND:' + dtEnd);
      calLines.push('SUMMARY:' + summary.replace(/[\r\n]+/g, ' '));
      calLines.push('LOCATION:' + (ev.room || ''));
      calLines.push('DESCRIPTION:' + descParts.join('\\n').replace(/[\r\n]+/g, ' '));
      calLines.push('END:VEVENT');
      count++;
    }

    calLines.push('END:VCALENDAR');
    const icsContent = calLines.join('\r\n');

    const blob = new Blob([icsContent], { type: 'text/calendar;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'SA2_' + MEETING_NUMBER + '_' + safeName + '_' + TDOC_ID + '.ics';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  // Conflict & Parallel Stream Clash Detection
  function detectDelegateClashes(delName) {
    const clashBadge = document.getElementById('clashWarningBadge');
    if (!clashBadge) return;

    document.querySelectorAll('.has-clash').forEach(el => el.classList.remove('has-clash'));

    if (delName === 'ALL') {
      const memberClashes = {};
      let totalClashes = 0;
      const allMembers = [];
      document.querySelectorAll('.del-filter-btn[data-del]').forEach(b => {
        const d = b.getAttribute('data-del');
        if (d && d !== 'ALL') allMembers.push(d);
      });

      for (const member of allMembers) {
        const c = getClashesForMember(member);
        if (c.length > 0) {
          memberClashes[member] = c.length;
          totalClashes += c.length;
          c.forEach(ev => markClashCells(ev));
        }
      }

      if (totalClashes > 0) {
        const summary = Object.entries(memberClashes).map(([m, n]) => m + ' (' + n + ')').join(', ');
        clashBadge.innerHTML = '⚠️ ' + totalClashes + ' Delegate Clash' + (totalClashes > 1 ? 'es' : '') + ': ' + summary;
        clashBadge.style.display = 'inline-flex';
      } else {
        clashBadge.style.display = 'none';
      }
      return;
    }

    const clashes = getClashesForMember(delName);
    if (clashes.length > 0) {
      clashBadge.innerHTML = '⚠️ ' + clashes.length + ' Parallel Clash' + (clashes.length > 1 ? 'es' : '') + ' for ' + delName;
      clashBadge.style.display = 'inline-flex';
      clashes.forEach(ev => markClashCells(ev));
    } else {
      clashBadge.style.display = 'none';
    }
  }

  function getClashesForMember(memberName) {
    const mLower = memberName.toLowerCase();
    const tuOccupancy = {};

    for (const ev of CAL_EVENTS) {
      const effDels = (ev.delegates || []).map(d => d.toLowerCase());
      if (!effDels.includes(mLower)) continue;

      const tus = (ev.sub_slot === 1) ? [1] : (ev.sub_slot === 2) ? [2] : [1, 2];
      for (const tu of tus) {
        const key = ev.day_idx + '|' + ev.quarter + '|' + tu;
        if (!tuOccupancy[key]) tuOccupancy[key] = [];
        tuOccupancy[key].push(ev);
      }
    }

    const clashingEvents = new Set();
    for (const key in tuOccupancy) {
      const list = tuOccupancy[key];
      const streams = new Set(list.map(e => e.stream_num));
      if (streams.size > 1) {
        list.forEach(e => clashingEvents.add(e));
      }
    }
    return Array.from(clashingEvents);
  }

  function markClashCells(ev) {
    const sDay = String(ev.day_idx);
    const sPeriod = String(ev.quarter);
    const sStr = String(ev.stream_num);
    document.querySelectorAll('.grid-cell[data-day="' + sDay + '"][data-period="' + sPeriod + '"][data-stream="' + sStr + '"]').forEach(el => el.classList.add('has-clash'));
    document.querySelectorAll('.cards-slot-row[data-day="' + sDay + '"][data-period="' + sPeriod + '"][data-stream="' + sStr + '"]').forEach(el => el.classList.add('has-clash'));
  }

  // Live Venue Status Monitor ("Now / Up Next")
  function checkLiveVenueStatus() {
    const banner = document.getElementById('liveVenueBanner');
    if (!banner || !MEETING_BASE_DATE) return;

    document.querySelectorAll('.is-live-now').forEach(el => el.classList.remove('is-live-now'));

    const now = new Date();
    const venueMs = now.getTime() + (MEETING_TZ_OFFSET_MIN * 60000);
    const vDate = new Date(venueMs);

    const pad = n => String(n).padStart(2, '0');
    const curTime = pad(vDate.getUTCHours()) + ':' + pad(vDate.getUTCMinutes());

    const parts = MEETING_BASE_DATE.split('-');
    const baseMs = Date.UTC(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
    const vDayMs = Date.UTC(vDate.getUTCFullYear(), vDate.getUTCMonth(), vDate.getUTCDate());
    const dayDiff = Math.floor((vDayMs - baseMs) / 86400000);

    if (dayDiff < 0) {
      const daysUntil = -dayDiff;
      banner.innerHTML = '📅 Meeting starts in ' + daysUntil + ' day' + (daysUntil > 1 ? 's' : '') + ' (' + MEETING_BASE_DATE + ')';
      banner.style.display = 'flex';
      return;
    }
    if (dayDiff > 4) {
      banner.innerHTML = '🏁 Meeting concluded on Friday (' + MEETING_BASE_DATE + ')';
      banner.style.display = 'flex';
      return;
    }

    const todayEvents = CAL_EVENTS.filter(e => e.day_idx === dayDiff);
    const liveNow = todayEvents.filter(e => e.start_local <= curTime && curTime < e.end_local);
    const upcoming = todayEvents.filter(e => e.start_local > curTime).sort((a, b) => a.start_local.localeCompare(b.start_local));

    if (liveNow.length > 0) {
      liveNow.forEach(ev => {
        const sDay = String(ev.day_idx);
        const sPeriod = String(ev.quarter);
        const sStr = String(ev.stream_num);
        document.querySelectorAll('.grid-cell[data-day="' + sDay + '"][data-period="' + sPeriod + '"][data-stream="' + sStr + '"]').forEach(el => el.classList.add('is-live-now'));
        document.querySelectorAll('.cards-slot-row[data-day="' + sDay + '"][data-period="' + sPeriod + '"][data-stream="' + sStr + '"]').forEach(el => el.classList.add('is-live-now'));
      });

      const liveTitles = liveNow.map(e => (e.is_drafting ? '🛠️ Draft: ' : ('Str ' + e.stream_num + ': ')) + (e.title || '').substring(0, 40)).join(' | ');
      let nextMsg = '';
      if (upcoming.length > 0) {
        const n = upcoming[0];
        nextMsg = ' &bull; Next: ' + n.start_local + ' (' + (n.title || '').substring(0, 30) + ')';
      }
      banner.innerHTML = '<span class="pulse-dot"></span> <span><strong>LIVE NOW</strong> (Venue ' + curTime + '): ' + liveTitles + nextMsg + '</span>';
      banner.style.display = 'flex';
    } else if (upcoming.length > 0) {
      const n = upcoming[0];
      banner.innerHTML = '⏳ <strong>Venue Time ' + curTime + '</strong> &bull; Next session at ' + n.start_local + ': ' + (n.title || '').substring(0, 45);
      banner.style.display = 'flex';
    } else {
      banner.innerHTML = '🌙 <strong>Venue Time ' + curTime + '</strong> &bull; Today\'s sessions concluded. Resumes tomorrow at 09:00.';
      banner.style.display = 'flex';
    }
  }

  // Auto-switch to mobile cards on narrow screens & apply initial empty break hiding
  if (window.innerWidth <= 768) {
    switchView('cards');
  }
  toggleEmptyBreaks();
  detectDelegateClashes('ALL');
  checkLiveVenueStatus();
  setInterval(checkLiveVenueStatus, 30000);
"""

        snap_label_html = f'<span class="badge-snap">{html.escape(snapshot.snapshot_label)}</span>' if snapshot.snapshot_label else ''
        tz_extra = f" | 💻 Austrian/Laptop Time: {tz_diff_str}" if is_tz_different else ""

        full_html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
<title>SA2 #{snapshot.meeting_number} Session Plan ({snapshot.tdoc_id})</title>
<style>
{css_style}
</style>
</head>
<body>

<header>
  <div class="header-top">
    <div class="title-block">
      <h1 class="title">SA2 #{snapshot.meeting_number} Session Plan</h1>
      <span class="badge-tdoc">{snapshot.tdoc_id}</span>
      {snap_label_html}
    </div>
    <div class="header-actions">
      <span id="clashWarningBadge" class="badge-clash" style="display:none;"></span>
      <button id="btnDownloadIcs" class="btn-cal-download" onclick="downloadSelectedCalendar()" title="Download RFC 5545 calendar for selected delegate or whole team">📅 Download Team .ics</button>
      <div class="view-switcher">
        <button id="btnViewGrid" class="view-btn active" onclick="switchView('grid')">🗓️ Overview Grid</button>
        <button id="btnViewCards" class="view-btn" onclick="switchView('cards')">📱 Day Cards</button>
      </div>
    </div>
  </div>

  <div class="tz-banner">
    📍 {tz_name}{tz_extra}
  </div>

  <div id="liveVenueBanner" class="live-banner" style="display:none;"></div>

  <div class="del-strip">
    <span class="strip-label">Delegates:</span>
    <button class="del-filter-btn active" data-del="ALL" onclick="selectDelegate('ALL')">👥 All Team</button>
    {del_filter_html}
    <span class="strip-sep"></span>
    <button id="btnTracked" class="del-filter-btn btn-toggle" onclick="toggleTracked()">⭐ Starred Topics</button>
    <button id="btnDrafting" class="del-filter-btn btn-toggle" onclick="toggleDrafting()">🛠️ Drafting Only</button>
  </div>

  <div class="filter-controls-row">
    <input type="text" id="searchInput" class="search-box" placeholder="🔍 Search AI number, topic, convener, notes..." oninput="applyFilters()">
    <div id="gridControls" class="grid-row-controls">
      <label class="toggle-label"><input type="checkbox" id="chkStr1" checked onchange="toggleStream(1)"> Str 1</label>
      <label class="toggle-label"><input type="checkbox" id="chkStr2" checked onchange="toggleStream(2)"> Str 2</label>
      <label class="toggle-label"><input type="checkbox" id="chkStr3" checked onchange="toggleStream(3)"> Str 3</label>
      <label class="toggle-label"><input type="checkbox" id="chkBreaks" checked onchange="toggleBreaks()"> Breaks</label>
      <label class="toggle-label"><input type="checkbox" id="chkEmptyBreaks" checked onchange="toggleEmptyBreaks()"> Hide Empty</label>
    </div>
  </div>
</header>

<!-- Main View 1: 5-Day Spatial Overview Grid -->
<div id="overviewView">
  {''.join(overview_html)}
</div>

<!-- Main View 2: Responsive Mobile Day Cards -->
<div id="cardsView" style="display:none;">
  {''.join(cards_html)}
</div>

<!-- Floating Rich Tooltip Element -->
<div id="floatingTooltip" class="floating-tooltip"></div>

<script>
  const CAL_EVENTS = {cal_events_json};
  const MEETING_BASE_DATE = "{base_date_js}";
  const MEETING_TZ_OFFSET_MIN = {meeting_tz_offset_min};
  const MEETING_NUMBER = "{snapshot.meeting_number}";
  const TDOC_ID = "{snapshot.tdoc_id}";
</script>
<script>
{js_script}
</script>
</body>
</html>
"""

        output_path.write_text(full_html, encoding="utf-8")
        logger.info(f"[SessionPlanExporter] Generated interactive Session Plan HTML: {output_path} ({output_path.stat().st_size} bytes)")
        return output_path

    @classmethod
    def export_ics(
        cls,
        snapshot: SessionPlanSnapshot,
        output_path: Path,
        meeting_start_date_str: str,
        meeting_tz_offset_min: int = 120,
        only_tracked: bool = False,
        only_drafting: bool = False,
        tracked_tokens: Optional[Set[str]] = None,
        annotations: Optional[Dict[str, UserSessionAnnotation]] = None,
        delegation_plan: Optional[DelegationPlan] = None,
        target_delegate: Optional[str] = None,
    ) -> Path:
        """
        Exports an RFC 5545 .ics Calendar file strictly anchored in UTC.
        Guarantees correct calendar reminder alerts across all devices.
        Optionally scopes to a specific delegate and annotates assignments.
        """
        tracked_tokens = {t.lower() for t in (tracked_tokens or set())}
        annotations = annotations or {}
        output_path = Path(output_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)

        base_date = datetime.strptime(meeting_start_date_str, "%Y-%m-%d").date()

        cal_name = (
            f"SA2 #{snapshot.meeting_number} - {target_delegate} ({snapshot.tdoc_id})"
            if target_delegate
            else f"SA2 #{snapshot.meeting_number} ({snapshot.tdoc_id})"
        )

        lines = [
            "BEGIN:VCALENDAR",
            "VERSION:2.0",
            f"PRODID:-//3GPP Meeting Tools//SA2 #{snapshot.meeting_number} Session Plan//EN",
            "CALSCALE:GREGORIAN",
            "METHOD:PUBLISH",
            f"X-WR-CALNAME:{cal_name}",
        ]

        streams = snapshot.streams

        for s in snapshot.slots:
            if not s.raw_text.strip() or s.raw_text.strip() == "-":
                continue

            if only_drafting and not s.is_drafting:
                continue

            # Check if tracked
            is_tracked = cls._is_slot_tracked(s, tracked_tokens)

            if only_tracked and not is_tracked and not s.is_drafting:
                continue

            # Check delegate assignment / target delegate filtering
            eff_dels, is_cov, is_all = cls.get_slot_effective_delegates(s, delegation_plan)
            if target_delegate:
                td_lower = target_delegate.strip().lower()
                if not (is_all or any(ed.strip().lower() == td_lower for ed in eff_dels)):
                    continue

            # Calculate physical date
            from datetime import timedelta
            slot_date = base_date + timedelta(days=s.day_index)
            slot_date_str = slot_date.strftime("%Y-%m-%d")

            utc_start = SessionPlanTimezone.venue_local_to_utc_dt(
                slot_date_str, s.start_time_local, meeting_tz_offset_min
            )
            utc_end = SessionPlanTimezone.venue_local_to_utc_dt(
                slot_date_str, s.end_time_local, meeting_tz_offset_min
            )

            if not utc_start or not utc_end:
                continue

            dt_stamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
            dt_start = utc_start.strftime("%Y%m%dT%H%M%SZ")
            dt_end = utc_end.strftime("%Y%m%dT%H%M%SZ")
            # Hoisted out of the f-string: backslashes inside f-string expressions are a
            # SyntaxError before Python 3.12 (project supports >= 3.8).
            safe_target = ("-" + re.sub(r"[^\w\-]", "_", target_delegate)) if target_delegate else ""
            uid = f"sa2-{snapshot.meeting_number}{safe_target}-{s.slot_key}-{dt_start}@3gpp.tools"

            s_info = streams.get(s.stream_number)
            r_name = getattr(s, "room_override", "") or (s_info.room_name if s_info else f"Stream {s.stream_number}")
            convener = s_info.convener if s_info else ""
            tohru = s_info.tohru_room if s_info else ""

            del_tag = ""
            if is_all:
                del_tag = "[All Team (cov)] " if is_cov else "[All Team] "
            elif eff_dels:
                cov_s = " (cov)" if is_cov else ""
                del_tag = f"[{', '.join(eff_dels)}{cov_s}] "

            summary_prefix = f"🛠️ [Drafting] {del_tag}" if s.is_drafting else f"[Str {s.stream_number}] {del_tag}"
            summary = f"{summary_prefix}{s.raw_text[:50]}"
            location = f"{r_name}"

            desc_lines = [
                f"Session: SA2 #{snapshot.meeting_number} ({snapshot.tdoc_id})",
                f"Room: {r_name}",
                f"Convener: {convener}",
                f"TOHRU: https://tohru.3gpp.org/ (Room: {tohru})",
                f"Scheduled: {s.raw_text}",
            ]

            if is_all:
                desc_lines.append(f"Delegates: All Team ({'Ad-hoc coverage' if is_cov else 'General / All Team Topic'})")
            elif len(eff_dels) == 1:
                desc_lines.append(f"Delegate: {eff_dels[0]} ({'Ad-hoc coverage' if is_cov else 'Topic Owner'})")
            elif eff_dels:
                desc_lines.append(f"Delegates: {', '.join(eff_dels)} ({'Ad-hoc coverage' if is_cov else 'Topic Owner'})")

            note_obj = annotations.get(s.slot_key)
            if note_obj and note_obj.comment:
                desc_lines.append(f"My Notes: {note_obj.comment}")

            description = "\\n".join(desc_lines)

            lines.extend([
                "BEGIN:VEVENT",
                f"UID:{uid}",
                f"DTSTAMP:{dt_stamp}",
                f"DTSTART:{dt_start}",
                f"DTEND:{dt_end}",
                f"SUMMARY:{summary}",
                f"LOCATION:{location}",
                f"DESCRIPTION:{description}",
                "STATUS:CONFIRMED",
                "END:VEVENT",
            ])

        lines.append("END:VCALENDAR")
        output_path.write_text("\r\n".join(lines), encoding="utf-8")
        logger.info(f"[SessionPlanExporter] Generated .ics calendar: {output_path}")
        return output_path

    @classmethod
    def generate_email_summary(
        cls,
        snapshot: SessionPlanSnapshot,
        meeting_tz_offset_min: int,
        tz_name: str,
        diff_result: Optional[SessionPlanDiffResult] = None,
        tracked_tokens: Optional[Set[str]] = None,
        annotations: Optional[Dict[str, UserSessionAnnotation]] = None,
    ) -> Tuple[str, str]:
        """
        Generates pre-formatted rich HTML and plain text email bodies ready for Outlook.
        Returns: (subject, html_body)
        """
        annotations = annotations or {}
        tracked_tokens = {t.lower() for t in (tracked_tokens or set())}
        subject = f"[SA2#{snapshot.meeting_number}] Session Plan Update: {snapshot.tdoc_id}"

        # 1. Changes
        changes_html = ""
        if diff_result and diff_result.has_changes:
            bullets = "".join(f"<li>{html.escape(b)}</li>" for b in diff_result.changelog_bullets[:12])
            changes_html = f"""
            <div style="background:#F0FDF4; border:1px solid #BBF7D0; border-radius:6px; padding:10px 14px; margin-bottom:14px;">
              <h3 style="margin-top:0; color:#166534; font-size:14px;">🔄 Key Changes from Previous Plan ({diff_result.old_snapshot_label}):</h3>
              <ul style="margin-bottom:0; padding-left:18px; font-size:13px; color:#14532D;">
                {bullets}
              </ul>
            </div>
            """

        # 2. Drafting sessions
        drafting_slots = [s for s in snapshot.slots if s.is_drafting]
        drafting_rows = ""
        for s in drafting_slots:
            time_str = SessionPlanTimezone.format_dual_time_range(s.start_time_local, s.end_time_local, meeting_tz_offset_min)
            s_info = snapshot.streams.get(s.stream_number)
            r_name = getattr(s, "room_override", "") or (s_info.room_name if s_info else f"Stream {s.stream_number}")
            drafting_rows += f"<tr><td style='padding:4px 8px; border:1px solid #E2E8F0;'><b>{s.day_name}</b> {s.quarter}</td><td style='padding:4px 8px; border:1px solid #E2E8F0;'>{time_str}</td><td style='padding:4px 8px; border:1px solid #E2E8F0;'>{html.escape(r_name)}</td><td style='padding:4px 8px; border:1px solid #E2E8F0;'>{html.escape(s.raw_text)}</td></tr>"

        drafting_html = f"""
        <h3 style="color:#991B1B; font-size:14px; margin-bottom:6px;">🛠️ Upcoming Drafting Sessions:</h3>
        <table style="border-collapse:collapse; width:100%; font-size:12px; margin-bottom:16px;">
          <tr style="background:#FEF2F2; color:#991B1B;">
            <th style="padding:4px 8px; border:1px solid #E2E8F0; text-align:left;">Slot</th>
            <th style="padding:4px 8px; border:1px solid #E2E8F0; text-align:left;">Time (Venue / Vienna)</th>
            <th style="padding:4px 8px; border:1px solid #E2E8F0; text-align:left;">Room</th>
            <th style="padding:4px 8px; border:1px solid #E2E8F0; text-align:left;">Topic</th>
          </tr>
          {drafting_rows}
        </table>
        """ if drafting_slots else ""

        # 3. Tracked topics
        tracked_slots = [s for s in snapshot.slots if cls._is_slot_tracked(s, tracked_tokens)]

        tracked_rows = ""
        for s in tracked_slots:
            time_str = SessionPlanTimezone.format_dual_time_range(s.start_time_local, s.end_time_local, meeting_tz_offset_min)
            s_info = snapshot.streams.get(s.stream_number)
            r_name = getattr(s, "room_override", "") or (s_info.room_name if s_info else f"Stream {s.stream_number}")
            tracked_rows += f"<tr><td style='padding:4px 8px; border:1px solid #E2E8F0;'><b>{s.day_name}</b> {s.quarter}</td><td style='padding:4px 8px; border:1px solid #E2E8F0;'>{time_str}</td><td style='padding:4px 8px; border:1px solid #E2E8F0;'>{html.escape(r_name)}</td><td style='padding:4px 8px; border:1px solid #E2E8F0;'>{html.escape(s.raw_text)}</td></tr>"

        tracked_html = f"""
        <h3 style="color:#5B21B6; font-size:14px; margin-bottom:6px;">🎯 Our Tracked Topics Schedule:</h3>
        <table style="border-collapse:collapse; width:100%; font-size:12px; margin-bottom:16px;">
          <tr style="background:#F5F3FF; color:#5B21B6;">
            <th style="padding:4px 8px; border:1px solid #E2E8F0; text-align:left;">Slot</th>
            <th style="padding:4px 8px; border:1px solid #E2E8F0; text-align:left;">Time (Venue / Vienna)</th>
            <th style="padding:4px 8px; border:1px solid #E2E8F0; text-align:left;">Room</th>
            <th style="padding:4px 8px; border:1px solid #E2E8F0; text-align:left;">Topic</th>
          </tr>
          {tracked_rows}
        </table>
        """ if tracked_slots else ""

        html_body = f"""
        <div style="font-family:Segoe UI, Helvetica, Arial, sans-serif; color:#1E293B; line-height:1.5;">
          <p>Hi Team,</p>
          <p>Here is the latest session plan update (<b>{snapshot.tdoc_id}</b>) for <b>SA2 #{snapshot.meeting_number}</b>.</p>
          <p style="font-size:12px; color:#64748B;">📍 Venue Timezone: <b>{tz_name}</b></p>
          {changes_html}
          {drafting_html}
          {tracked_html}
          <p style="font-size:12px; color:#64748B;"><i>Generated via 3GPP Meeting Tools.</i></p>
        </div>
        """

        return subject, html_body

    # --- Delegation Plan Serialization & Email Generation ---

    @staticmethod
    def serialize_delegation_plan_json(plan: DelegationPlan) -> str:
        """Serializes a DelegationPlan to formatted JSON string."""
        data = {
            "3gpp_delegation_plan_version": 1,
            "meeting_number": str(plan.meeting_number),
            "members": [
                {
                    "name": m.name,
                    "email": m.email,
                    "color": m.color,
                    "is_self": m.is_self,
                    "notes": m.notes,
                }
                for m in plan.members.values()
            ],
            "topic_owners": {
                ai: {"owner": to.owner_name, "notes": to.notes}
                for ai, to in plan.topic_owners.items()
            },
            "appointments": {
                sk: {
                    "appointed_delegate": ap.appointed_delegate,
                    "original_owner": ap.original_owner,
                    "reason": ap.reason,
                }
                for sk, ap in plan.appointments.items()
            },
        }
        return json.dumps(data, indent=2, ensure_ascii=False)

    @staticmethod
    def parse_delegation_plan_json(raw_text: str) -> Optional[dict]:
        """
        Parses delegation JSON from raw string, file content, or email text with embedded comment.
        Looks for <!-- 3GPP-DELEGATION-PLAN: {...} --> or top-level JSON.
        """
        if not raw_text or not raw_text.strip():
            return None

        # 1. Look for embedded HTML comment payload
        match = re.search(r"<!--\s*3GPP-DELEGATION-PLAN:\s*(\{.+?\})\s*-->", raw_text, re.DOTALL)
        if match:
            try:
                data = json.loads(match.group(1))
                if isinstance(data, dict) and "members" in data:
                    return data
            except Exception:
                pass

        # 2. Look for JSON markdown block or raw JSON object
        json_match = re.search(r"```(?:json)?\s*(\{.+?\})\s*```", raw_text, re.DOTALL)
        candidate = json_match.group(1) if json_match else raw_text.strip()

        try:
            data = json.loads(candidate)
            if isinstance(data, dict) and ("members" in data or "topic_owners" in data):
                return data
        except Exception:
            pass

        return None

    @classmethod
    def generate_delegation_email_summary(
        cls,
        snapshot: SessionPlanSnapshot,
        plan: DelegationPlan,
        meeting_tz_offset_min: int,
        tz_name: str,
        agenda_map: Optional[dict] = None,
    ) -> Tuple[str, str, str]:
        """
        Generates executive delegation email summary with:
        1. Roster & Workload Breakdown (Color badges)
        2. Master Topic Ownership Table
        3. Ad-Hoc Session Coverages (Collision Appointments)
        4. Individual Per-Delegate Daily Agendas
        5. Embedded machine-readable JSON payload for 1-click import.
        Returns: (subject, html_body, plain_text)
        """
        agenda_map = agenda_map or {}
        members = plan.members
        topic_owners = plan.topic_owners
        appointments = plan.appointments

        subject = f"[SA2#{snapshot.meeting_number}] Delegation Plan & AI Distribution ({snapshot.tdoc_id})"

        # Compute effective assignments for each slot in snapshot: slot_key -> List of delegate names
        slot_delegate_map: Dict[str, List[str]] = {}
        for s in snapshot.slots:
            dels, _, is_all = cls.get_slot_effective_delegates(s, plan)
            slot_delegate_map[s.slot_key] = dels

        # Calculate quarters workload per member
        member_workload: Dict[str, int] = {name: 0 for name in members}
        for s in snapshot.slots:
            for del_name in slot_delegate_map.get(s.slot_key, []):
                if del_name in member_workload:
                    member_workload[del_name] += 1

        # 1. Executive Roster Table
        roster_rows = ""
        plain_roster = []
        for name, m in members.items():
            wl = member_workload.get(name, 0)
            owned_count = sum(1 for to in topic_owners.values() if name in to.get_owners(list(members.keys())))
            cov_count = sum(1 for ap in appointments.values() if name in ap.get_appointed_delegates(list(members.keys())))
            role_str = m.notes or "Delegate"
            roster_rows += f"""<tr>
              <td style='padding:6px 10px; border:1px solid #E2E8F0;'><span style='display:inline-block; width:10px; height:10px; background-color:{m.color}; border-radius:50%; margin-right:6px;'></span><b>{html.escape(name)}</b></td>
              <td style='padding:6px 10px; border:1px solid #E2E8F0;'>{html.escape(role_str)}</td>
              <td style='padding:6px 10px; border:1px solid #E2E8F0; text-align:center;'>{owned_count} AIs</td>
              <td style='padding:6px 10px; border:1px solid #E2E8F0; text-align:center;'>{cov_count} slots</td>
              <td style='padding:6px 10px; border:1px solid #E2E8F0; text-align:center;'><b>{wl}</b> slots</td>
            </tr>"""
            plain_roster.append(f"- {name} ({role_str}): {owned_count} AIs owned, {wl} slots scheduled")

        # 2. Topic Ownership Table
        topic_rows = ""
        plain_topics = []
        for ai_num in sorted(topic_owners.keys()):
            to = topic_owners[ai_num]
            ai_info = agenda_map.get(ai_num)
            desc = ai_info.description if ai_info else f"Agenda Item {ai_num}"
            if to.is_all():
                badge_html = "<span style='display:inline-block; padding:2px 8px; border-radius:4px; background-color:#EEF2FF; border:1px solid #C7D2FE; color:#4338CA; font-weight:600;'>👥 All Team</span>"
                plain_owner = "All Team"
            else:
                assigned = to.get_owners(list(members.keys()))
                badges = []
                for oname in assigned:
                    m = members.get(oname)
                    col = m.color if m else "#2563EB"
                    bg_col = m.bg_tint if m else "#E4ECFC"
                    bdr_col = m.border_tint if m else "#B2C8F8"
                    badges.append(f"<span style='display:inline-block; padding:2px 8px; border-radius:4px; background-color:{bg_col}; border:1px solid {bdr_col}; color:{col}; font-weight:600; margin-right:4px;'>● {html.escape(oname)}</span>")
                badge_html = "".join(badges) or "—"
                plain_owner = ", ".join(assigned) or "Unassigned"

            topic_rows += f"""<tr>
              <td style='padding:6px 10px; border:1px solid #E2E8F0;'><b>{html.escape(ai_num)}</b></td>
              <td style='padding:6px 10px; border:1px solid #E2E8F0;'>{html.escape(desc)}</td>
              <td style='padding:6px 10px; border:1px solid #E2E8F0;'>{badge_html}</td>
            </tr>"""
            plain_topics.append(f"- AI {ai_num}: {desc} -> {plain_owner}")

        # 3. Ad-Hoc Session Coverages (Appointments)
        appt_rows = ""
        plain_appts = []
        for sk in sorted(appointments.keys()):
            ap = appointments[sk]
            parts = sk.split("|")
            day_name = parts[0] if len(parts) > 0 else ""
            q_name = parts[1] if len(parts) > 1 else ""
            str_num = parts[3] if len(parts) > 3 else "1"
            orig_txt = f"for {ap.original_owner}" if ap.original_owner else ""
            reason_txt = f" ({ap.reason})" if ap.reason else ""

            if ap.is_all():
                badge_app = "<span style='display:inline-block; padding:2px 8px; border-radius:4px; background-color:#EEF2FF; border:1px solid #C7D2FE; color:#4338CA; font-weight:600;'>👥 All Team</span>"
                plain_app = "All Team"
            else:
                assigned = ap.get_appointed_delegates(list(members.keys()))
                badges = []
                for oname in assigned:
                    m_app = members.get(oname)
                    col_app = m_app.color if m_app else "#047857"
                    bg_app = m_app.bg_tint if m_app else "#E0EEEA"
                    bdr_app = m_app.border_tint if m_app else "#A7CFC4"
                    badges.append(f"<span style='display:inline-block; padding:2px 8px; border-radius:4px; background-color:{bg_app}; border:1px solid {bdr_app}; color:{col_app}; font-weight:600; margin-right:4px;'>● {html.escape(oname)}</span>")
                badge_app = "".join(badges) or "—"
                plain_app = ", ".join(assigned) or "Unassigned"

            appt_rows += f"""<tr>
              <td style='padding:6px 10px; border:1px solid #E2E8F0;'><b>{day_name}</b> {q_name} (Str {str_num})</td>
              <td style='padding:6px 10px; border:1px solid #E2E8F0;'>{badge_app} {html.escape(orig_txt)}</td>
              <td style='padding:6px 10px; border:1px solid #E2E8F0; color:#64748B;'>{html.escape(ap.reason or 'Ad-hoc cover')}</td>
            </tr>"""
            plain_appts.append(f"- {day_name} {q_name} (Stream {str_num}): {plain_app} covering {orig_txt}{reason_txt}")

        appt_html = f"""
        <div style="margin-bottom:20px;">
          <h3 style="color:#B45309; font-size:14px; margin-bottom:8px;">⚡ Ad-Hoc Session Coverages & Clash Resolutions:</h3>
          <table style="border-collapse:collapse; width:100%; font-size:12px; margin-bottom:12px;">
            <tr style="background:#FFFBEB; color:#92400E;">
              <th style="padding:6px 10px; border:1px solid #E2E8F0; text-align:left;">Slot</th>
              <th style="padding:6px 10px; border:1px solid #E2E8F0; text-align:left;">Covering Delegate</th>
              <th style="padding:6px 10px; border:1px solid #E2E8F0; text-align:left;">Reason</th>
            </tr>
            {appt_rows}
          </table>
        </div>
        """ if appointments else ""

        # 4. Individual Timetables (per delegate)
        individual_sections_html = ""
        plain_schedules = []
        for name, m in members.items():
            del_slots = [s for s in snapshot.slots if name in slot_delegate_map.get(s.slot_key, [])]
            if not del_slots:
                continue

            del_rows = ""
            for s in del_slots:
                time_str = SessionPlanTimezone.format_dual_time_range(
                    s.start_time_local, s.end_time_local, meeting_tz_offset_min
                )
                s_info = snapshot.streams.get(s.stream_number)
                r_name = getattr(s, "room_override", "") or (s_info.room_name if s_info else f"Stream {s.stream_number}")
                is_cov = (
                    s.slot_key in appointments
                    and name in appointments[s.slot_key].get_appointed_delegates(list(members.keys()))
                )
                role_badge = "<span style='background:#FEF3C7; color:#92400E; padding:2px 6px; border-radius:4px; font-size:10px;'>Coverage</span>" if is_cov else "<span style='background:#EFF6FF; color:#1E40AF; padding:2px 6px; border-radius:4px; font-size:10px;'>Owner</span>"

                del_rows += f"""<tr>
                  <td style='padding:5px 8px; border:1px solid #E2E8F0;'><b>{s.day_name}</b> {s.quarter}</td>
                  <td style='padding:5px 8px; border:1px solid #E2E8F0;'>{time_str}</td>
                  <td style='padding:5px 8px; border:1px solid #E2E8F0;'>Str {s.stream_number} ({html.escape(r_name)})</td>
                  <td style='padding:5px 8px; border:1px solid #E2E8F0;'>{role_badge} {html.escape(s.raw_text)}</td>
                </tr>"""

            individual_sections_html += f"""
            <div style="margin-bottom:16px;">
              <h4 style="color:{m.color}; margin-bottom:6px; font-size:13px;">📅 {html.escape(name)}'s Schedule ({len(del_slots)} sessions):</h4>
              <table style="border-collapse:collapse; width:100%; font-size:11px; margin-bottom:8px;">
                <tr style="background:#F8FAFC;">
                  <th style="padding:4px 8px; border:1px solid #E2E8F0; text-align:left;">Slot</th>
                  <th style="padding:4px 8px; border:1px solid #E2E8F0; text-align:left;">Time (Venue / Vienna)</th>
                  <th style="padding:4px 8px; border:1px solid #E2E8F0; text-align:left;">Room</th>
                  <th style="padding:4px 8px; border:1px solid #E2E8F0; text-align:left;">Topic</th>
                </tr>
                {del_rows}
              </table>
            </div>
            """

        # 5. Embedded Machine-Readable JSON Payload
        json_payload = cls.serialize_delegation_plan_json(plan)
        embedded_comment = f"\n<!-- 3GPP-DELEGATION-PLAN: {json_payload} -->\n"

        html_body = f"""
        <div style="font-family:Segoe UI, Helvetica, Arial, sans-serif; color:#1E293B; line-height:1.5;">
          <p>Hi Team,</p>
          <p>Here is the delegation plan and topic allocation for <b>SA2 #{snapshot.meeting_number}</b> (Session Plan: <b>{snapshot.tdoc_id}</b>).</p>
          <p style="font-size:12px; color:#64748B;">📍 Venue Timezone: <b>{tz_name}</b></p>

          <h3 style="color:#1E40AF; font-size:14px; margin-bottom:8px;">👥 Delegation Roster & Workload:</h3>
          <table style="border-collapse:collapse; width:100%; font-size:12px; margin-bottom:16px;">
            <tr style="background:#F1F5F9; color:#334155;">
              <th style="padding:6px 10px; border:1px solid #E2E8F0; text-align:left;">Delegate</th>
              <th style="padding:6px 10px; border:1px solid #E2E8F0; text-align:left;">Role</th>
              <th style="padding:6px 10px; border:1px solid #E2E8F0; text-align:center;">Owned Topics</th>
              <th style="padding:6px 10px; border:1px solid #E2E8F0; text-align:center;">Coverages</th>
              <th style="padding:6px 10px; border:1px solid #E2E8F0; text-align:center;">Total Quarters</th>
            </tr>
            {roster_rows}
          </table>

          {appt_html}

          <h3 style="color:#1E40AF; font-size:14px; margin-bottom:8px;">📋 Master Topic Ownership:</h3>
          <table style="border-collapse:collapse; width:100%; font-size:12px; margin-bottom:16px;">
            <tr style="background:#F1F5F9; color:#334155;">
              <th style="padding:6px 10px; border:1px solid #E2E8F0; text-align:left;">AI</th>
              <th style="padding:6px 10px; border:1px solid #E2E8F0; text-align:left;">Topic</th>
              <th style="padding:6px 10px; border:1px solid #E2E8F0; text-align:left;">Lead Owner</th>
            </tr>
            {topic_rows}
          </table>

          <h3 style="color:#1E40AF; font-size:14px; margin-bottom:8px;">📅 Individual Delegate Schedules:</h3>
          {individual_sections_html}

          <p style="font-size:11px; color:#94A3B8; margin-top:20px;">
            <i>Tip: To import this plan into your 3GPP Tools, copy this email text and click <b>Import Team Plan</b> in the Session Plan window.</i>
          </p>
        </div>
        {embedded_comment}
        """

        plain_text = f"SA2 #{snapshot.meeting_number} Delegation Plan ({snapshot.tdoc_id})\n"
        plain_text += f"Venue Timezone: {tz_name}\n\n"
        plain_text += "--- ROSTER ---\n" + "\n".join(plain_roster) + "\n\n"
        if plain_appts:
            plain_text += "--- AD-HOC COVERAGES ---\n" + "\n".join(plain_appts) + "\n\n"
        plain_text += "--- TOPIC OWNERSHIP ---\n" + "\n".join(plain_topics) + "\n"

        return subject, html_body, plain_text
