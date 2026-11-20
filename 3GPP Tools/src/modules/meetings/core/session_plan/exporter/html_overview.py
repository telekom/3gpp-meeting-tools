# --- File: src/modules/meetings/core/session_plan/exporter/html_overview.py ---
"""
HTML builders for the 5-day spatial Overview Grid, slot hover tooltips, and delegate filter buttons.
"""
import html
from datetime import timedelta

from modules.meetings.core.session_plan.exporter.helpers import (
    get_period_effective_times,
    get_slot_effective_delegates,
    is_slot_unassigned,
)
from modules.meetings.core.session_plan.exporter.html_data import MobileHtmlContext
from modules.meetings.core.session_plan.session_plan_models import (
    DelegationMember,
    SessionSlot,
    build_slot_ai_bullets_html,
    expand_ai_number,
    format_break_period_label,
    format_slot_display_text,
    format_venue_time_range,
    get_slot_unassigned_ais,
    is_slot_not_followed,
    resolve_effective_room_and_location,
)
from modules.meetings.core.session_plan.session_plan_parser import DAY_NAMES


def build_tooltip_html(ctx: MobileHtmlContext, slot: SessionSlot, day_idx: int, period: str, tu: int) -> str:
    """Builds rich HTML tooltip / drawer payload for a scheduled session slot."""
    snapshot = ctx.snapshot
    day_name = DAY_NAMES[day_idx] if day_idx < len(DAY_NAMES) else f"Day {day_idx+1}"
    date_lbl = ""
    if ctx.base_date:
        cur_d = ctx.base_date + timedelta(days=day_idx)
        date_lbl = f" ({cur_d.strftime('%d %b')})"

    times = get_period_effective_times(snapshot, period)
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

    time_venue = f"{format_venue_time_range(s_t, e_t)} (Venue Time)"

    s_info = snapshot.streams.get(slot.stream_number)
    r_name, loc_note, is_swapped = resolve_effective_room_and_location(
        slot, streams=snapshot.streams, room_locations=ctx.room_locations, slot_locations=ctx.slot_locations
    )
    cap = f" ({s_info.capacity} seats)" if (not is_swapped and s_info and s_info.capacity) else ""
    convener = s_info.convener if s_info else ("SA2 Chair / MCC" if slot.stream_number == 1 else "")
    tohru = s_info.tohru_room if s_info else ""

    swap_tag = " <i>(Swapped)</i>" if is_swapped else ""
    meta_parts = [f"🏢 <b>Stream {slot.stream_number}:</b> {html.escape(r_name)}{cap}{swap_tag}"]
    if loc_note:
        meta_parts.append(f"🧭 <b>Location:</b> {html.escape(loc_note)}")
    if convener:
        meta_parts.append(f"👤 <b>Chair:</b> {html.escape(convener)}")
    if tohru:
        meta_parts.append(f"📡 <b>TOHRU:</b> {html.escape(tohru)}")

    ai_bullets = build_slot_ai_bullets_html(
        slot, agenda_map=ctx.agenda_map, plan=ctx.delegation_plan
    )

    del_names, is_cov, is_all = get_slot_effective_delegates(slot, ctx.delegation_plan, agenda_map=ctx.agenda_map)
    del_html = ""
    if is_all:
        cov_lbl = " (Ad-hoc coverage)" if is_cov else " (General / All Team Topic)"
        del_html = f"<div class='tt-section'>👥 <b>Delegation:</b> All Team{cov_lbl}</div>"
    elif del_names:
        role_lbl = " (Ad-hoc coverage)" if is_cov else " (Topic Owner)"
        badges = []
        for dn in del_names:
            m = ctx.members.get(dn)
            c = m.color if m else "#2563EB"
            bg = m.bg_tint if m else "#EFF6FF"
            bdr = m.border_tint if m else "#BFDBFE"
            badges.append(f"<span style='background:{bg}; color:{c}; border:1px solid {bdr}; padding:1px 5px; border-radius:3px; font-weight:700;'>● {html.escape(dn)}</span>")
        del_html = f"<div class='tt-section'>👤 <b>Delegation:</b> {' '.join(badges)}{role_lbl}</div>"

    ann = ctx.annotations.get(slot.slot_key)
    note_html = ""
    if ann and (ann.comment or ann.assigned_colleague):
        parts = []
        if ann.comment:
            parts.append(f"“{html.escape(ann.comment)}”")
        if ann.assigned_colleague:
            parts.append(f"(Assigned: <b>{html.escape(ann.assigned_colleague)}</b>)")
        note_html = f"<div class='tt-section tt-note'>📝 <b>Notes:</b> {' '.join(parts)}</div>"

    tt = f"""<div class="tt-header">📅 {day_name}{date_lbl} · {period} · {half_lbl}</div>
<div class="tt-time">🕒 <b>Time:</b> {time_venue}</div>
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


def render_overview_cell(
    ctx: MobileHtmlContext,
    slot: SessionSlot,
    colspan: int,
    d_idx: int,
    p: str,
    s_num: int,
    tu: int,
) -> str:
    """Renders a single `<td>` cell in the Overview Grid table."""
    eff_dels, is_cov, is_all = get_slot_effective_delegates(slot, ctx.delegation_plan, agenda_map=ctx.agenda_map)
    is_unassigned = is_slot_unassigned(slot, ctx.delegation_plan, eff_dels, is_all, agenda_map=ctx.agenda_map)
    is_not_followed = is_slot_not_followed(slot, ctx.delegation_plan, agenda_map=ctx.agenda_map)
    unassigned_ais = get_slot_unassigned_ais(slot, ctx.delegation_plan, agenda_map=ctx.agenda_map)
    dels_joined = ",".join(eff_dels)

    search_tokens = [slot.raw_text.lower()]
    for it in slot.items:
        if it.ai_number:
            search_tokens.append(it.ai_number.lower())
            search_tokens.append(f"ai {it.ai_number.lower()}")
            for sub_ai in expand_ai_number(it.ai_number, ctx.agenda_map):
                search_tokens.append(sub_ai.lower())
                search_tokens.append(f"ai {sub_ai.lower()}")
        if it.topic_label:
            search_tokens.append(it.topic_label.lower())
    ann = ctx.annotations.get(slot.slot_key)
    if ann and ann.comment:
        search_tokens.append(ann.comment.lower())

    room_ov = getattr(slot, "room_override", "") or ""
    clean_display_text, is_empty_drafting = format_slot_display_text(slot.raw_text, room_ov)
    raw_esc = html.escape(clean_display_text).replace("\n", "<br>")
    tt_html = build_tooltip_html(ctx, slot, d_idx, p, tu)

    day_cls = ctx.day_parity_classes[d_idx] if 0 <= d_idx < len(ctx.day_parity_classes) else "day-even"
    bdr_cls = "tu-split-col" if (colspan == 1 and tu == 1) else "day-end-col"

    # Quiet dash for bare empty "Drafting:" placeholders without notes or delegates
    if is_empty_drafting and not (ann and ann.comment) and not is_all and not eff_dels:
        return (
            f'<td class="grid-cell empty-slot-cell is-empty-drafting stream-{s_num} {day_cls} {bdr_cls}" '
            f'colspan="{colspan}" data-stream="{s_num}" data-period="{p}" data-tu="{tu}" data-day="{d_idx}" '
            f'data-delegates="" data-is-all="0" data-is-cov="0" data-unassigned="0" '
            f'data-not-followed="0" data-drafting="1" data-search="{html.escape(" ".join(search_tokens))}">—'
            f'<div class="slot-tooltip-data" style="display:none;">{tt_html}</div></td>'
        )

    cell_classes = ["grid-cell", "slot-cell", f"stream-{s_num}", day_cls, bdr_cls]
    if is_empty_drafting:
        cell_classes.append("is-empty-drafting")
    elif slot.is_drafting:
        cell_classes.append("is-drafting")
    if is_unassigned:
        cell_classes.append("is-unassigned")
    elif is_not_followed:
        cell_classes.append("is-not-followed")

    # Unified rounded slot card styling matching the Desktop UI
    card_inline_style = ""
    if is_all:
        card_inline_style = ' style="background:#EEF2FF; border:1px solid #C7D2FE; border-left:4px solid #4338CA;"'
    elif eff_dels:
        pm = ctx.members.get(eff_dels[0])
        p_col = pm.color if pm else "#2563EB"
        p_tint = DelegationMember.compute_tint(p_col, 0.07)
        p_bdr = pm.border_tint if pm else "#BFDBFE"
        card_inline_style = f' style="background:{p_tint}; border:1px solid {p_bdr}; border-left:4px solid {p_col};"'

    # Effective slot time pill + optional room override pill (mirroring Desktop GUI):
    # Omit redundant time pill on standard Full-Quarter slots since the row header shows it.
    p_times = get_period_effective_times(ctx.snapshot, p)
    if tu == 1:
        s_t = slot.start_time_local or p_times[2]
        e_t = slot.end_time_local or p_times[3]
        is_custom_t = (s_t, e_t) != (p_times[2], p_times[3])
        time_pill_cls = "is-custom-time" if is_custom_t else "is-half"
    elif tu == 2:
        s_t = slot.start_time_local or p_times[4]
        e_t = slot.end_time_local or p_times[5]
        is_custom_t = (s_t, e_t) != (p_times[4], p_times[5])
        time_pill_cls = "is-custom-time" if is_custom_t else "is-half"
    else:
        s_t = slot.start_time_local or p_times[0]
        e_t = slot.end_time_local or p_times[1]
        is_custom_t = (s_t, e_t) != (p_times[0], p_times[1])
        time_pill_cls = "is-custom-time" if is_custom_t else "is-full"

    time_pill_html = ""
    if (not is_empty_drafting) and s_t and e_t and (tu in (1, 2) or is_custom_t):
        time_pill_html = f'<span class="cell-time-pill {time_pill_cls}">🕒 {html.escape(s_t)}–{html.escape(e_t)}</span>'

    room_pill_html = f'<span class="cell-room-override">📍 {html.escape(room_ov)}</span>' if room_ov else ""
    meta_row_html = f'<div class="cell-meta-row">{time_pill_html}{room_pill_html}</div>' if (time_pill_html or room_pill_html) else ""

    has_unassigned_ai = bool(not is_all and not is_not_followed and unassigned_ais)

    del_badges = []
    has_ai_items = any(it.ai_number for it in slot.items)
    if is_all:
        bdr_style = "dashed" if is_cov else "solid"
        cov_note = " (Ad-hoc coverage)" if is_cov else " (All Team)"
        del_badges.append(
            f'<span class="del-pill-badge is-all" style="border:1px {bdr_style} #6366F1;" '
            f'title="Whole delegation{cov_note}">👥 ALL</span>'
        )
    elif eff_dels:
        max_badges = 3
        shown_dels = eff_dels if len(eff_dels) <= max_badges else eff_dels[: max_badges - 1]
        for dn in shown_dels:
            m = ctx.members.get(dn)
            ini = ctx.initials_map.get(dn, dn[:2].upper())
            col = m.color if m else "#2563EB"
            bg = m.bg_tint if m else "#EFF6FF"
            bdr = m.border_tint if m else "#BFDBFE"
            bdr_style = "dashed" if is_cov else "solid"
            cov_note = " (Ad-hoc cover)" if is_cov else " (Owner)"
            del_badges.append(
                f'<span class="del-pill-badge" data-del-name="{html.escape(dn)}" '
                f'style="background:{bg}; color:{col}; border:1px {bdr_style} {bdr};" '
                f'title="{html.escape(dn)}{cov_note}">{ini}</span>'
            )
        if len(eff_dels) > len(shown_dels):
            rest = eff_dels[len(shown_dels):]
            del_badges.append(
                f'<span class="del-pill-badge is-more-pill" title="{html.escape(", ".join(rest))}">+{len(rest)}</span>'
            )
    elif is_not_followed:
        del_badges.append('<span class="del-pill-badge is-not-followed-pill" title="Explicitly not followed (Out of scope)">⛔</span>')

    if has_unassigned_ai:
        un_lbl = ", ".join(unassigned_ais)
        del_badges.append(
            f'<span class="del-pill-badge is-unassigned-pill" '
            f'title="Unassigned Agenda Item(s): {html.escape(un_lbl)}">❓</span>'
        )
    elif not is_all and not eff_dels and not is_not_followed and (has_ai_items or unassigned_ais):
        del_badges.append('<span class="del-pill-badge is-unassigned-pill" title="Unassigned Agenda Item">❓</span>')

    if is_all or eff_dels:
        del_badges.append('<span class="del-pill-badge is-clash-pill" title="Parallel-stream clash">⚠</span>')

    note_badge = ""
    if ann and ann.comment:
        note_badge = f'<div class="cell-note-pill" title="{html.escape(ann.comment)}">📝 {html.escape(ann.comment[:35])}{"…" if len(ann.comment) > 35 else ""}</div>'

    return f"""<td class="{' '.join(cell_classes)}" colspan="{colspan}" data-stream="{s_num}" data-period="{p}" data-tu="{tu}" data-day="{d_idx}" data-delegates="{html.escape(dels_joined)}" data-is-all="{1 if is_all else 0}" data-is-cov="{1 if is_cov else 0}" data-unassigned="{1 if is_unassigned else 0}" data-not-followed="{1 if is_not_followed else 0}" data-drafting="{1 if slot.is_drafting else 0}" data-search="{html.escape(' '.join(search_tokens))}">
  <div class="slot-card"{card_inline_style}>
    <div class="slot-card-body">
      {meta_row_html}
      {f'<div class="cell-title">{raw_esc}</div>' if raw_esc else ''}
      {note_badge}
    </div>
    {f'<div class="cell-del-float">{"".join(del_badges)}</div>' if del_badges else ''}
  </div>
  <div class="slot-tooltip-data" style="display:none;">{tt_html}</div>
</td>"""


def build_overview_grid_html(ctx: MobileHtmlContext) -> str:
    """Assembles the 5-day spatial Overview Grid HTML table."""
    overview_html = []
    overview_html.append('<div class="overview-grid-container"><table class="overview-table">\n')

    # Header Row 1: Day Names & Dates (colspan=2 per day)
    overview_html.append('<thead><tr class="hdr-days-row">\n')
    overview_html.append('<th class="corner-hdr" rowspan="2">Period / Stream<br><span class="hdr-sub">Venue Times</span></th>\n')
    for d_idx, d_name in enumerate(DAY_NAMES):
        day_cls = ctx.day_parity_classes[d_idx]
        date_str = ""
        today_badge = ""
        if ctx.base_date:
            cur_d = ctx.base_date + timedelta(days=d_idx)
            date_str = f" <span class='hdr-date'>{cur_d.strftime('%d %b')}</span>"
            if cur_d == ctx.venue_today:
                today_badge = " <span class='badge-today'>Today</span>"
        overview_html.append(f'<th class="col-day-hdr {day_cls} day-end-col" colspan="2" data-day="{d_idx}">{d_name[:3]}{date_str}{today_badge}</th>\n')
    overview_html.append('</tr>\n')

    # Header Row 2: Half-Quarter TU1 / TU2
    overview_html.append('<tr class="hdr-tu-row">\n')
    for d_idx in range(5):
        day_cls = ctx.day_parity_classes[d_idx]
        overview_html.append(
            f'<th class="col-tu-hdr {day_cls} tu-split-col" data-day="{d_idx}" data-tu="1" title="1st Half (TU 1)">1st Half</th>\n'
            f'<th class="col-tu-hdr {day_cls} day-end-col" data-day="{d_idx}" data-tu="2" title="2nd Half (TU 2)">2nd Half</th>\n'
        )
    overview_html.append('</tr></thead>\n<tbody>\n')

    # Body Rows
    streams = ctx.snapshot.streams
    for p in ctx.all_periods:
        times = get_period_effective_times(ctx.snapshot, p)
        if any(b in p.lower() for b in ["break", "coffee", "lunch"]):
            clean_name = format_break_period_label(p)
            brk_cls = "grid-row break-row period-end-row"
            overview_html.append(f'<tr class="{brk_cls}" data-period="{p}">\n')
            overview_html.append(
                f'<th class="row-hdr break-hdr" data-period="{p}">'
                f'<div class="hdr-time-line"><span class="hdr-break-title">{clean_name}</span><span class="hdr-now-tag">NOW</span></div>'
                f'<div class="hdr-time-break">⏰ {times[0]} – {times[1]}</div>'
                f'</th>\n'
            )

            for d_idx in range(5):
                day_cls = ctx.day_parity_classes[d_idx]
                matching_b = ctx.break_slot_map.get((d_idx, p), [])
                valid_b = [s for s in matching_b if s.raw_text.strip() and s.raw_text.strip() != "-"]
                if valid_b:
                    overview_html.append(render_overview_cell(ctx, valid_b[0], colspan=2, d_idx=d_idx, p=p, s_num=valid_b[0].stream_number or 1, tu=0))
                else:
                    overview_html.append(f'<td class="grid-cell empty-break-cell {day_cls} day-end-col" colspan="2" data-day="{d_idx}" data-period="{p}">—</td>\n')
            overview_html.append('</tr>\n')
        else:
            for s_num in [1, 2, 3]:
                s_info = streams.get(s_num)
                r_name = s_info.room_name if s_info else ("Main" if s_num == 1 else f"Brk {s_num-1}")
                short_r = r_name.split("(")[0].strip()

                p_end_cls = " period-end-row" if s_num == 3 else ""
                overview_html.append(f'<tr class="grid-row stream-row stream-row-{s_num}{p_end_cls}" data-stream="{s_num}" data-period="{p}">\n')
                overview_html.append(
                    f'<th class="row-hdr stream-hdr-s{s_num}" data-period="{p}" '
                    f'title="{p}: {times[0]}–{times[1]} · 1st Half {times[2]}–{times[3]} | 2nd Half {times[4]}–{times[5]}">'
                    f'<div class="hdr-time-line"><span class="hdr-time-primary">{p} · {times[0]} – {times[1]}</span><span class="hdr-now-tag">NOW</span></div>'
                    f'<div class="hdr-stream-line"><span class="str-dot s{s_num}"></span> <span class="hdr-stream-sub">Str {s_num} · {html.escape(short_r)}</span></div>'
                    f'</th>\n'
                )

                for d_idx in range(5):
                    day_cls = ctx.day_parity_classes[d_idx]
                    matching = ctx.slot_map.get((d_idx, p, s_num), [])
                    full_slot = next((s for s in matching if s.sub_slot == 0), None)
                    tu1_slot = next((s for s in matching if s.sub_slot == 1), None)
                    tu2_slot = next((s for s in matching if s.sub_slot == 2), None)

                    if full_slot:
                        overview_html.append(render_overview_cell(ctx, full_slot, colspan=2, d_idx=d_idx, p=p, s_num=s_num, tu=0))
                    elif tu1_slot or tu2_slot:
                        if tu1_slot:
                            overview_html.append(render_overview_cell(ctx, tu1_slot, colspan=1, d_idx=d_idx, p=p, s_num=s_num, tu=1))
                        else:
                            overview_html.append(f'<td class="grid-cell empty-slot-cell {day_cls} tu-split-col" colspan="1" data-day="{d_idx}" data-period="{p}" data-stream="{s_num}">—</td>\n')

                        if tu2_slot:
                            overview_html.append(render_overview_cell(ctx, tu2_slot, colspan=1, d_idx=d_idx, p=p, s_num=s_num, tu=2))
                        else:
                            overview_html.append(f'<td class="grid-cell empty-slot-cell {day_cls} day-end-col" colspan="1" data-day="{d_idx}" data-period="{p}" data-stream="{s_num}">—</td>\n')
                    else:
                        overview_html.append(f'<td class="grid-cell empty-slot-cell {day_cls} day-end-col" colspan="2" data-day="{d_idx}" data-period="{p}" data-stream="{s_num}">—</td>\n')
                overview_html.append('</tr>\n')

    overview_html.append('</tbody></table></div>\n')
    return "".join(overview_html)


def build_delegate_filter_buttons_html(ctx: MobileHtmlContext) -> str:
    """Renders per-member delegate focus buttons for the header control strip."""
    del_filter_buttons = []
    for name, m in ctx.members.items():
        safe_name_attr = html.escape(name)
        del_filter_buttons.append(
            f'<button class="del-filter-btn" data-del="{safe_name_attr}" '
            f'style="--del-color:{m.color}; --del-bg:{m.bg_tint}; --del-bdr:{m.border_tint};" '
            f'onclick="selectDelegate(this.getAttribute(\'data-del\'))">● {safe_name_attr}</button>'
        )
    return "\n".join(del_filter_buttons)
