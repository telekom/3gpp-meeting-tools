# --- File: src/modules/meetings/core/session_plan/exporter/html_cards.py ---
"""
HTML builder for the responsive Mobile Day Cards view of the exported Session Plan.
"""
import html

from modules.meetings.core.session_plan.exporter.helpers import (
    get_period_effective_times,
    get_slot_effective_delegates,
    is_slot_unassigned,
)
from modules.meetings.core.session_plan.exporter.html_data import MobileHtmlContext
from modules.meetings.core.session_plan.exporter.html_overview import build_tooltip_html
from modules.meetings.core.session_plan.session_plan_models import (
    _matching_topic_ownerships,
    expand_ai_number,
    format_venue_time_range,
    get_slot_unassigned_ais,
    is_slot_not_followed,
    resolve_effective_room_and_location,
)
from modules.meetings.core.session_plan.session_plan_parser import DAY_NAMES, PERIOD_ORDER


def build_mobile_cards_html(ctx: MobileHtmlContext) -> str:
    """Assembles the day-by-day touch-friendly Mobile Day Cards HTML view."""
    cards_html = []
    streams = ctx.snapshot.streams
    active_days = [d for d in DAY_NAMES if d in ctx.slots_by_day_q]

    cards_html.append('<div class="day-tabs">\n')
    for idx, d_name in enumerate(active_days):
        active_cls = " active" if idx == 0 else ""
        cards_html.append(f'<button class="day-tab{active_cls}" data-day="{d_name}" onclick="switchDay(\'{d_name}\', this)">{d_name[:3]}</button>\n')
    cards_html.append('</div>\n<div class="cards-content">\n')

    for idx, d_name in enumerate(active_days):
        d_idx = DAY_NAMES.index(d_name) if d_name in DAY_NAMES else idx
        day_display = "" if idx == 0 else ' style="display:none;"'
        cards_html.append(f'<div id="day-{d_name}" class="day-section"{day_display}>\n')

        q_map = ctx.slots_by_day_q.get(d_name, {})
        for q in PERIOD_ORDER:
            if q not in q_map:
                continue
            slots_in_q = q_map[q]
            if not slots_in_q:
                continue

            q_eff_times = get_period_effective_times(ctx.snapshot, q)
            q_time_str = format_venue_time_range(q_eff_times[0], q_eff_times[1])

            cards_html.append(f"""  <div class="quarter-card">
    <div class="quarter-header">
      <span class="quarter-title">{q}</span>
      <span class="quarter-time">{q_time_str}</span>
    </div>
""")

            for s in slots_in_q:
                eff_dels, is_cov, is_all = get_slot_effective_delegates(s, ctx.delegation_plan, agenda_map=ctx.agenda_map)
                is_unassigned = is_slot_unassigned(s, ctx.delegation_plan, eff_dels, is_all, agenda_map=ctx.agenda_map)
                is_not_followed = is_slot_not_followed(s, ctx.delegation_plan, agenda_map=ctx.agenda_map)
                unassigned_ais = get_slot_unassigned_ais(s, ctx.delegation_plan, agenda_map=ctx.agenda_map)
                dels_joined = ",".join(eff_dels)

                search_tokens = [s.raw_text.lower()]
                for it in s.items:
                    if it.ai_number:
                        search_tokens.append(it.ai_number.lower())
                        for sub_ai in expand_ai_number(it.ai_number, ctx.agenda_map):
                            search_tokens.append(sub_ai.lower())
                    if it.topic_label:
                        search_tokens.append(it.topic_label.lower())
                ann = ctx.annotations.get(s.slot_key)
                if ann and ann.comment:
                    search_tokens.append(ann.comment.lower())

                raw_esc = html.escape(s.raw_text)
                tt_html = build_tooltip_html(ctx, s, d_idx, q, s.sub_slot)

                s_info = streams.get(s.stream_number)
                r_name, loc_note, _ = resolve_effective_room_and_location(
                    s, streams=streams, room_locations=ctx.room_locations, slot_locations=ctx.slot_locations
                )
                convener = s_info.convener if s_info else ""
                loc_suffix = f" · 🧭 {html.escape(loc_note)}" if loc_note else ""

                q_t = q_eff_times
                if s.sub_slot == 1:
                    st_s = s.start_time_local or q_t[2]
                    st_e = s.end_time_local or q_t[3]
                    sub_tag = f'<span class="sub-slot-badge is-half">🕒 {html.escape(st_s)}–{html.escape(st_e)} [TU 1]</span>'
                elif s.sub_slot == 2:
                    st_s = s.start_time_local or q_t[4]
                    st_e = s.end_time_local or q_t[5]
                    sub_tag = f'<span class="sub-slot-badge is-half">🕒 {html.escape(st_s)}–{html.escape(st_e)} [TU 2]</span>'
                else:
                    st_s = s.start_time_local or q_t[0]
                    st_e = s.end_time_local or q_t[1]
                    sub_tag = f'<span class="sub-slot-badge">🕒 {html.escape(st_s)}–{html.escape(st_e)}</span>' if (st_s and st_e) else ""

                chips = []
                for it in s.items:
                    cnt = f'<span class="count-chip">[{it.tdoc_count_str}]</span>' if it.tdoc_count_str else ""
                    if not it.ai_number:
                        chips.append(f'<span class="ai-chip">{html.escape(it.topic_label)}{cnt}</span>')
                        continue
                    expanded_ais = expand_ai_number(it.ai_number, ctx.agenda_map)
                    for idx_ai, target_ai in enumerate(expanded_ais):
                        ai_lbl = f"AI {target_ai}"
                        is_ai_unassigned = target_ai in unassigned_ais
                        sub_matches = _matching_topic_ownerships(target_ai, ctx.topic_owners_map, agenda_map=ctx.agenda_map)
                        is_ai_not_followed = bool(
                            sub_matches and all(m.is_not_followed() for m in sub_matches)
                        )
                        if is_ai_unassigned:
                            chip_cls = "ai-chip unassigned-chip"
                            prefix = "❓ "
                        elif is_ai_not_followed:
                            chip_cls = "ai-chip not-followed-chip"
                            prefix = "⛔ "
                        else:
                            chip_cls = "ai-chip"
                            prefix = ""
                        chip_cnt = cnt if idx_ai == 0 else ""
                        chips.append(f'<span class="{chip_cls}">{prefix}{html.escape(ai_lbl)}{chip_cnt}</span>')

                del_badge = ""
                if is_all:
                    cov_lbl = " (cov)" if is_cov else ""
                    del_badge = f'<span class="del-pill-badge is-all">👥 All Team{cov_lbl}</span>'
                elif eff_dels:
                    cov_lbl = " (cov)" if is_cov else ""
                    b_list = []
                    for ed in eff_dels:
                        m = ctx.members.get(ed)
                        col = m.color if m else "#2563EB"
                        bg = m.bg_tint if m else "#EFF6FF"
                        bdr = m.border_tint if m else "#BFDBFE"
                        b_list.append(f'<span class="del-pill-badge" style="background:{bg}; color:{col}; border:1px solid {bdr};">● {html.escape(ed)}{cov_lbl}</span>')
                    del_badge = "".join(b_list)
                elif is_not_followed:
                    del_badge = '<span class="del-pill-badge is-not-followed-pill">⛔ Not Followed</span>'
                else:
                    del_badge = '<span class="del-pill-badge is-unassigned-pill">❓ Unassigned</span>'

                note_html = ""
                if ann and ann.comment:
                    note_html = f'<div class="note-chip">📝 <b>Note:</b> {html.escape(ann.comment)}</div>'

                row_classes = ["cards-slot-row"]
                if s.is_drafting:
                    row_classes.append("drafting")
                elif is_unassigned:
                    row_classes.append("unassigned")
                elif is_not_followed:
                    row_classes.append("not-followed")

                cards_html.append(f"""    <div class="{' '.join(row_classes)}" data-stream="{s.stream_number}" data-period="{q}" data-tu="{s.sub_slot}" data-day="{d_idx}" data-delegates="{html.escape(dels_joined)}" data-is-all="{1 if is_all else 0}" data-is-cov="{1 if is_cov else 0}" data-unassigned="{1 if is_unassigned else 0}" data-not-followed="{1 if is_not_followed else 0}" data-drafting="{1 if s.is_drafting else 0}" data-search="{html.escape(' '.join(search_tokens))}">
      <div class="slot-header-line">
        <span class="stream-badge str-{s.stream_number}">Stream {s.stream_number}</span>
        <span class="room-pill">📍 {html.escape(r_name)}{loc_suffix}{" · " + html.escape(convener) if convener else ""}</span>
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
    return "".join(cards_html)
