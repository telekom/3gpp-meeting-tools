# --- File: src/modules/meetings/ui/session_plan/window/mixins/slot_cells.py ---
"""
Overview grid cell widgets: rich slot tooltip, delegate badges, slot and break cards.
UI changes here must be mirrored in SessionPlanExporter.export_mobile_html.

Mixin of SessionPlanWindow: relies on attributes set in SessionPlanWindow.__init__.
"""
import html
from datetime import datetime, timedelta
from typing import (
    Dict,
    List,
    Optional,
)
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QFrame,
    QHBoxLayout,
    QLabel,
    QSizePolicy,
    QVBoxLayout,
    QWidget,
)
from modules.meetings.core.session_plan.session_plan_models import (
    DelegationMember,
    SessionSlot,
    build_slot_ai_bullets_html,
    format_slot_display_text,
    format_venue_time_range,
    get_slot_unassigned_ais,
    is_slot_not_followed,
    resolve_effective_room_and_location,
)
from core.ui.theme import C
from modules.meetings.core.session_plan.session_plan_parser import PERIOD_TIMES
from modules.meetings.ui.session_plan.window.constants import DAY_NAMES


class SlotCellsMixin:
    """Overview grid cell widgets: rich slot tooltip, delegate badges, slot and break cards."""

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

        venue_range = format_venue_time_range(s_t, e_t)
        time_display = f"{venue_range} (Venue Time)"

        streams = self.active_snapshot.streams if self.active_snapshot else None
        s_info = streams.get(stream_num) if streams else None
        r_name, loc_note, is_swapped = resolve_effective_room_and_location(
            slot, streams=streams, room_locations=self.room_locations, slot_locations=self.slot_locations
        )
        cap = f" ({s_info.capacity} seats)" if (not is_swapped and s_info and s_info.capacity) else ""
        convener = s_info.convener if s_info else ("SA2 Chair / MCC" if stream_num == 1 else "")
        tohru = s_info.tohru_room if s_info else ""

        swap_tag = " <i>(Swapped)</i>" if is_swapped else ""
        meta_parts = [f"🏢 <b>Stream {stream_num}:</b> {html.escape(r_name)}{cap}{swap_tag}"]
        if loc_note:
            meta_parts.append(f"🧭 <b>Location:</b> {html.escape(loc_note)}")
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
            cov_lbl = f" <span style='color:{C.INDIGO_ACCENT}; font-weight:normal;'>(Ad-Hoc Coverage)</span>" if is_cov else f" <span style='color:{C.INDIGO_ACCENT}; font-weight:normal;'>(General / All Team Topic)</span>"
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
            f"QLabel {{ background: {member.bg_tint}; color: {member.color}; font-weight: 700; font-size: 9px; "
            f"border: {bdr}; border-radius: 8px; padding: 1px 5px; min-width: 16px; }}"
        )
        return lbl

    def _create_slot_cell_widget(
        self, slot: SessionSlot, day_idx: int, period: str, stream_num: int, is_half: bool = False, tu: int = 0
    ) -> QWidget:
        """Creates a compact unified card for a session slot with right-aligned pills and neutral unassigned styling."""
        w = QWidget()
        w.setStyleSheet(".QWidget { background: transparent; }")
        outer_lay = QVBoxLayout(w)
        outer_lay.setContentsMargins(3, 3, 3, 3)
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

        # Left column: explicit time badge + optional room override chip + wrapped topic text
        body_lay = QVBoxLayout()
        body_lay.setContentsMargins(0, 0, 0, 0)
        body_lay.setSpacing(2)

        times = self._get_period_effective_times(period)
        if tu == 1:
            s_t = slot.start_time_local or times[2]
            e_t = slot.end_time_local or times[3]
            is_custom_time = (s_t, e_t) != (times[2], times[3])
        elif tu == 2:
            s_t = slot.start_time_local or times[4]
            e_t = slot.end_time_local or times[5]
            is_custom_time = (s_t, e_t) != (times[4], times[5])
        else:
            s_t = slot.start_time_local or times[0]
            e_t = slot.end_time_local or times[1]
            is_custom_time = (s_t, e_t) != (times[0], times[1])

        time_chip = None
        if not is_empty_drafting and s_t and e_t and (is_half or tu in (1, 2) or is_custom_time):
            time_chip = QLabel(f"🕒 {s_t}–{e_t}")
            if is_dimmed:
                tc_bg, tc_fg, tc_bdr = C.BG_SURFACE, C.TEXT_FAINT, C.BORDER_SUBTLE
            elif is_custom_time:
                tc_bg, tc_fg, tc_bdr = C.WARNING_SOFT, C.WARNING_DARK, C.WARNING_BORDER
            else:
                tc_bg, tc_fg, tc_bdr = C.INFO_SOFT, C.INFO_DARK, C.INFO_BORDER
            time_chip.setStyleSheet(
                f"QLabel {{ background: {tc_bg}; color: {tc_fg}; font-size: 9.5px; "
                f"font-weight: 700; border: 1px solid {tc_bdr}; border-radius: 3px; padding: 0px 4px; }}"
            )
            time_chip.setSizePolicy(QSizePolicy.Maximum, QSizePolicy.Fixed)

        room_chip = None
        if room_ov:
            room_chip = QLabel(f"📍 {room_ov}")
            room_chip.setStyleSheet(
                f"QLabel {{ background: {C.BG_SURFACE}; color: {C.TEXT_SECONDARY}; font-size: 9.5px; "
                f"font-weight: 700; border: 1px solid {C.BORDER_SUBTLE}; border-radius: 3px; padding: 0px 4px; }}"
            )
            room_chip.setSizePolicy(QSizePolicy.Maximum, QSizePolicy.Fixed)

        if time_chip and room_chip and not is_half:
            meta_row = QHBoxLayout()
            meta_row.setContentsMargins(0, 0, 0, 0)
            meta_row.setSpacing(4)
            meta_row.addWidget(time_chip, 0, Qt.AlignLeft)
            meta_row.addWidget(room_chip, 0, Qt.AlignLeft)
            meta_row.addStretch(1)
            body_lay.addLayout(meta_row)
        else:
            if time_chip:
                body_lay.addWidget(time_chip, 0, Qt.AlignLeft)
            if room_chip:
                body_lay.addWidget(room_chip, 0, Qt.AlignLeft)

        if clean_text:
            lbl = QLabel(clean_text)
            lbl.setWordWrap(True)
            lbl.setAlignment(Qt.AlignTop | Qt.AlignLeft)
            lbl.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.MinimumExpanding)
            lbl.setStyleSheet(f"QLabel {{ background: transparent; border: none; color: {text_color}; font-weight: {font_weight}; }}")
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
                f"QLabel {{ background: {C.WG_SA_BG}; color: {C.INDIGO_DARK}; font-weight: 700; font-size: 8.5px; "
                f"border: 1px {border_style} {C.INDIGO_ACCENT}; border-radius: 7px; padding: 1px 4px; }}"
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
                    f"QLabel {{ background: {C.BG_CANVAS}; color: {C.TEXT_BODY}; font-weight: 700; font-size: 8.5px; "
                    f"border: 1px solid {C.BORDER_STRONG}; border-radius: 7px; padding: 1px 4px; }}"
                )
                pill_widgets.append(more)
        elif is_not_followed_slot:
            nf_pill = QLabel("⛔")
            nf_pill.setToolTip("Explicitly Not Followed (Out of scope)")
            nf_pill.setStyleSheet(
                f"QLabel {{ background: {C.BG_CANVAS}; color: {C.TEXT_MUTED}; font-size: 9px; "
                f"border: 1px solid {C.BORDER_STRONG}; border-radius: 7px; padding: 0px 4px; }}"
            )
            pill_widgets.append(nf_pill)
        if has_unassigned_ai:
            un_pill = QLabel("❓")
            un_lbl = ", ".join(unassigned_ais)
            un_pill.setToolTip(f"Unassigned Agenda Item(s): {un_lbl} — right-click to assign")
            un_pill.setStyleSheet(
                f"QLabel {{ background: {C.WARNING_SOFT_ALT}; color: {C.WARNING}; font-weight: 700; font-size: 9px; "
                f"border: 1px dashed {C.WARNING_ACCENT}; border-radius: 7px; padding: 0px 4px; }}"
            )
            pill_widgets.append(un_pill)

        if is_collision:
            warn_pill = QLabel("⚠")
            who = ", ".join(colliding_members)
            warn_pill.setToolTip(f"Parallel-stream clash for: {who}")
            warn_pill.setStyleSheet(
                f"QLabel {{ background: {C.WARNING_SOFT_ALT}; color: {C.WARNING}; font-weight: bold; font-size: 9.5px; "
                f"border: 1px solid {C.WARNING_ACCENT}; border-radius: 7px; padding: 0px 4px; }}"
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
        w.setStyleSheet(".QWidget { background: transparent; }")
        lay = QVBoxLayout(w)
        lay.setContentsMargins(4, 3, 4, 3)

        lbl = QLabel(slot.raw_text)
        lbl.setWordWrap(True)
        lbl.setAlignment(Qt.AlignTop | Qt.AlignLeft)

        if slot.is_drafting:
            lbl.setStyleSheet(f"QLabel {{ background: {C.DANGER_SOFT}; color: {C.DANGER_DARK}; border: 1px solid {C.DANGER_BORDER}; border-radius: 4px; padding: 4px; font-weight: 600; }}")
        elif slot.is_plenary:
            lbl.setStyleSheet(f"QLabel {{ background: {C.INFO_SOFT}; color: {C.INFO_DARK}; border: 1px solid {C.INFO_BORDER}; border-radius: 4px; padding: 4px; font-weight: 600; }}")
        else:
            lbl.setStyleSheet(f"QLabel {{ background: {C.BG_SURFACE}; color: {C.TEXT_BODY}; border: 1px solid {C.BORDER_SUBTLE}; border-radius: 4px; padding: 4px; }}")

        lay.addWidget(lbl)

        day_name = DAY_NAMES[day_idx]
        times = PERIOD_TIMES.get(period, ("12:30", "14:00"))
        time_display = f"{format_venue_time_range(times[0], times[1])} (Venue Time)"
        tt = f"""<div style='font-family:Segoe UI, sans-serif; font-size:12px; min-width:220px; padding:4px;'>
            <div style='font-weight:bold; font-size:13px; color:{C.TEXT_STRONG}; margin-bottom:4px;'>
                ☕ {day_name} · {period}
            </div>
            <div style='color:{C.TEXT_SECONDARY}; margin-bottom:6px;'>
                🕒 <b>Time:</b> {time_display}
            </div>
            <hr style='border:none; border-top:1px solid {C.BORDER_STRONG}; margin:6px 0;'>
            <div style='font-weight:bold; color:{C.TEXT_PRIMARY};'>
                {html.escape(slot.raw_text)}
            </div>
        </div>"""
        w.setToolTip(tt)
        lbl.setToolTip(tt)
        return w
