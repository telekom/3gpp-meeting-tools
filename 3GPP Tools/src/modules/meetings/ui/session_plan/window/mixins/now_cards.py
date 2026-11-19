# --- File: src/modules/meetings/ui/session_plan/window/mixins/now_cards.py ---
"""
NOW view rendering: stream cards, upcoming period card and compact timeline rows.

Mixin of SessionPlanWindow: relies on attributes set in SessionPlanWindow.__init__.
"""
import html
from typing import Dict, Optional
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QFrame,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QVBoxLayout,
)
from modules.meetings.core.session_plan.session_plan_models import (
    SessionSlot,
    build_slot_ai_bullets_html,
    format_slot_display_text,
    format_venue_time_range,
    is_slot_not_followed,
    lookup_room_location_note,
    resolve_effective_room_and_location,
)
from core.ui.theme import C


class NowCardsMixin:
    """NOW view rendering: stream cards, upcoming period card and compact timeline rows."""

    def _render_now_view(self):
        """Renders the room-centric 'Happening Now', 'Up Next', and 'Rest of Today' dashboard."""
        if not hasattr(self, "now_content_layout"):
            return

        self._update_now_tab_badge()
        self._clear_now_layout()

        ctx = self._resolve_now_context()
        units = ctx["units"]
        target = ctx["target"]

        # Update stepper dropdown cleanly
        if hasattr(self, "now_slot_combo"):
            self.now_slot_combo.blockSignals(True)
            self.now_slot_combo.clear()
            for u in units:
                lbl = f"{u['day_name'][:3]} · {u['period']} · {u['half_label']} ({u['start']}–{u['end']})"
                self.now_slot_combo.addItem(lbl, u["key"])
            if target:
                idx = self.now_slot_combo.findData(target["key"])
                if idx >= 0:
                    self.now_slot_combo.setCurrentIndex(idx)
            self.now_slot_combo.blockSignals(False)

        mode = ctx["mode"]
        if mode == "live":
            self.now_status_pill.setStyleSheet(
                f"background: {C.SUCCESS_SOFT}; color: {C.SUCCESS}; border: 1px solid {C.SUCCESS_BORDER}; "
                f"border-radius: 999px; padding: 3px 10px; font-weight: 700; font-size: 11.5px;"
            )
        elif mode == "manual":
            self.now_status_pill.setStyleSheet(
                f"background: {C.WARNING_SOFT}; color: {C.WARNING_DARK}; border: 1px solid {C.WARNING_BORDER}; "
                f"border-radius: 999px; padding: 3px 10px; font-weight: 700; font-size: 11.5px;"
            )
        else:
            self.now_status_pill.setStyleSheet(
                f"background: {C.INFO_SOFT}; color: {C.INFO}; border: 1px solid {C.INFO_BORDER}; "
                f"border-radius: 999px; padding: 3px 10px; font-weight: 700; font-size: 11.5px;"
            )
        self.now_status_pill.setText(ctx["status_pill"])
        self.now_banner_lbl.setText(ctx["banner"])

        if not target:
            empty = QLabel("📭 No scheduled sessions available in this snapshot.")
            empty.setStyleSheet(f"color: {C.TEXT_MUTED}; font-size: 13px; padding: 24px;")
            self.now_content_layout.addWidget(empty)
            self.now_content_layout.addStretch(1)
            return

        # --- Section 1: Happening Right Now (Side-by-Side Streams / Rooms) ---
        sec1_title = "🔴 Happening Now — Rooms & Active Streams" if mode == "live" else f"📍 Active Window — {target['day_name']} · {target['period']} ({target['half_label']})"
        s1_lbl = QLabel(f"<b>{sec1_title}</b>")
        s1_lbl.setStyleSheet(f"font-size: 13px; color: {C.TEXT_PRIMARY};")
        self.now_content_layout.addWidget(s1_lbl)

        active_slots = self._slots_for_now_unit(target)
        # Look ahead for TU 2 changes when viewing TU 1
        tu2_by_stream: Dict[int, SessionSlot] = {}
        if target["tu"] == 1:
            for sl in self.active_snapshot.slots:
                if sl.day_index == target["day_idx"] and sl.quarter == target["period"] and sl.sub_slot == 2:
                    tu2_by_stream[sl.stream_number] = sl

        if target["is_break"] or (len(active_slots) == 1 and active_slots[0].stream_number == 0):
            # Single wide break/plenary row
            cards_row = QHBoxLayout()
            cards_row.setSpacing(10)
            for sl in active_slots:
                cards_row.addWidget(self._make_now_stream_card(sl.stream_number, sl, target, tu2_slot=None), 1)
            self.now_content_layout.addLayout(cards_row)
        else:
            # Side-by-side cards for Streams 1, 2, 3
            cards_row = QHBoxLayout()
            cards_row.setSpacing(10)
            by_str = {sl.stream_number: sl for sl in active_slots}
            for s_num in (1, 2, 3):
                sl = by_str.get(s_num)
                tu2_sl = tu2_by_stream.get(s_num)
                if sl and tu2_sl and tu2_sl.slot_key == sl.slot_key:
                    tu2_sl = None
                cards_row.addWidget(self._make_now_stream_card(s_num, sl, target, tu2_slot=tu2_sl), 1)
            self.now_content_layout.addLayout(cards_row)

        # --- Section 2: Up Next Today (Next 1-2 Periods) ---
        up_next = ctx["up_next_periods"]
        s2_lbl = QLabel(f"<b>⏭️ Up Next Today ({target['day_name']})</b>")
        s2_lbl.setStyleSheet(f"font-size: 13px; color: {C.TEXT_PRIMARY}; margin-top: 4px;")
        self.now_content_layout.addWidget(s2_lbl)

        if not up_next:
            no_next = QLabel(f"✅ No further sessions scheduled later on {target['day_name']}.")
            no_next.setStyleSheet(
                f"background: {C.BG_SURFACE}; border: 1px dashed {C.BORDER_STRONG}; border-radius: 8px; "
                f"padding: 12px; color: {C.TEXT_MUTED}; font-size: 12px;"
            )
            self.now_content_layout.addWidget(no_next)
        else:
            for p_info in up_next:
                self.now_content_layout.addWidget(self._make_upcoming_period_card(target["day_idx"], p_info))

        # --- Section 3: Rest of Today's Schedule ---
        rest = ctx["rest_of_day_periods"]
        if rest:
            s3_lbl = QLabel(f"<b>📅 Rest of {target['day_name']} Timeline</b>")
            s3_lbl.setStyleSheet(f"font-size: 13px; color: {C.TEXT_PRIMARY}; margin-top: 4px;")
            self.now_content_layout.addWidget(s3_lbl)
            for p_info in rest:
                self.now_content_layout.addWidget(self._make_compact_timeline_row(target["day_idx"], p_info))

        self.now_content_layout.addStretch(1)

    def _make_now_stream_card(
        self,
        stream_num: int,
        slot: Optional[SessionSlot],
        unit: dict,
        tu2_slot: Optional[SessionSlot] = None,
    ) -> QFrame:
        """Creates a rich, room-centric card for a stream in the 'Happening Now' section."""
        streams = self.active_snapshot.streams if self.active_snapshot else {}
        s_info = streams.get(stream_num)

        stream_accents = {
            1: (C.INDIGO_DARK, C.WG_SA_BG, C.WG_SA_BORDER, "Stream 1"),
            2: (C.EMERALD_DARK, C.WG_RAN_BG, C.WG_RAN_BORDER, "Stream 2"),
            3: (C.WARNING, C.WARNING_SOFT, C.WARNING_BORDER, "Stream 3"),
        }
        accent_col, accent_bg, accent_bdr, stream_title = stream_accents.get(
            stream_num, (C.TEXT_SECONDARY, C.BG_SURFACE, C.BORDER_STRONG, "Break / Plenary" if stream_num == 0 else f"Stream {stream_num}")
        )

        card = QFrame()
        card.setFrameShape(QFrame.StyledPanel)
        if slot is None:
            card.setStyleSheet(
                f"QFrame {{ background: {C.BG_SURFACE}; border: 1px dashed {C.BORDER_STRONG}; "
                f"border-radius: 8px; }}"
            )
            lay = QVBoxLayout(card)
            lay.setContentsMargins(12, 10, 12, 10)
            lay.setSpacing(6)
            hdr = QLabel(f"<b style='color:{accent_col};'>🏢 {stream_title}</b>")
            hdr.setStyleSheet("border: none; background: transparent;")
            lay.addWidget(hdr)
            default_room = s_info.room_name.split("(")[0].strip() if s_info else f"Stream {stream_num}"
            loc_note = lookup_room_location_note(default_room, self.room_locations)
            room_txt = f"📍 {default_room}" + (f" · 🧭 {loc_note}" if loc_note else "")
            rm_lbl = QLabel(room_txt)
            rm_lbl.setStyleSheet(f"color: {C.TEXT_MUTED}; font-size: 11px; border: none; background: transparent;")
            lay.addWidget(rm_lbl)
            idle_lbl = QLabel("— No active session scheduled in this stream —")
            idle_lbl.setStyleSheet(f"color: {C.TEXT_FAINT}; font-style: italic; font-size: 12px; padding: 12px 0; border: none; background: transparent;")
            lay.addWidget(idle_lbl)
            lay.addStretch(1)
            return card

        card.setStyleSheet(
            f"QFrame {{ background: {C.BG_CARD}; border: 1px solid {C.BORDER_STRONG}; "
            f"border-top: 4px solid {accent_col}; border-radius: 8px; }}"
        )
        lay = QVBoxLayout(card)
        lay.setContentsMargins(12, 10, 12, 10)
        lay.setSpacing(6)

        # Row 1: Stream pill + subtle delegate awareness badges
        top_row = QHBoxLayout()
        top_row.setSpacing(6)
        str_pill = QLabel(f"🏢 {stream_title}" + (" · Drafting" if slot.is_drafting else ""))
        str_pill.setStyleSheet(
            f"background: {accent_bg}; color: {accent_col}; border: 1px solid {accent_bdr}; "
            f"border-radius: 5px; padding: 2px 8px; font-weight: 700; font-size: 11px;"
        )
        top_row.addWidget(str_pill)
        top_row.addStretch(1)

        del_names, is_cov, is_all = self._get_slot_effective_delegates(slot)
        if is_all:
            p_lbl = QLabel("👥 ALL")
            p_lbl.setStyleSheet(
                f"background: {C.WG_SA_BG}; color: {C.INDIGO_DARK}; border: 1px solid #6366F1; "
                f"border-radius: 7px; padding: 1px 6px; font-weight: 700; font-size: 9.5px;"
            )
            top_row.addWidget(p_lbl)
        elif del_names:
            initials = self._delegate_initials_map()
            for dn in del_names[:3]:
                m = self.delegation_plan.members.get(dn) if self.delegation_plan else None
                if m:
                    top_row.addWidget(self._make_delegate_badge(m, initials.get(dn, m.initials), is_cov=is_cov))
        elif is_slot_not_followed(slot, self.delegation_plan, agenda_map=self.agenda_map):
            nf = QLabel("⛔ Out of Scope")
            nf.setStyleSheet(
                f"background: {C.BG_SURFACE}; color: {C.TEXT_MUTED}; border: 1px solid {C.BORDER_STRONG}; "
                f"border-radius: 7px; padding: 1px 6px; font-size: 9.5px; font-weight: 600;"
            )
            top_row.addWidget(nf)
        lay.addLayout(top_row)

        # Row 2: Prominent Physical Room + Hand-Made Location Note + Edit Pencil
        phys_room, loc_note, is_swapped = resolve_effective_room_and_location(
            slot, streams=streams, room_locations=self.room_locations, slot_locations=self.slot_locations
        )
        room_box = QFrame()
        room_box.setStyleSheet(
            f"QFrame {{ background: {C.BG_SURFACE}; border: 1px solid {C.BORDER_SUBTLE}; border-radius: 6px; padding: 4px 6px; }}"
        )
        rb_lay = QVBoxLayout(room_box)
        rb_lay.setContentsMargins(6, 4, 6, 4)
        rb_lay.setSpacing(3)

        rm_top = QHBoxLayout()
        rm_top.setSpacing(6)
        rm_lbl = QLabel(f"<b>📍 {html.escape(phys_room)}</b>")
        rm_lbl.setStyleSheet(f"color: {C.TEXT_PRIMARY}; font-size: 12.5px; border: none; background: transparent;")
        rm_top.addWidget(rm_lbl)

        if is_swapped:
            def_r = s_info.room_name.split("(")[0].strip() if s_info else ""
            sw = QLabel(f"⇄ Moved from {def_r}" if def_r else "⇄ Room Override")
            sw.setStyleSheet(
                f"background: {C.WARNING_SOFT}; color: {C.WARNING_DARK}; border: 1px solid {C.WARNING_BORDER}; "
                f"border-radius: 4px; padding: 1px 5px; font-size: 10px; font-weight: 700;"
            )
            rm_top.addWidget(sw)

        rm_top.addStretch(1)

        btn_edit_loc = QPushButton("✏️")
        btn_edit_loc.setFixedSize(24, 22)
        btn_edit_loc.setCursor(Qt.PointingHandCursor)
        btn_edit_loc.setToolTip(f"Edit floor / location note for '{phys_room}' or this slot")
        btn_edit_loc.setStyleSheet(
            f"QPushButton {{ border: 1px solid {C.BORDER_STRONG}; border-radius: 4px; background: {C.BG_CARD}; font-size: 11px; }}"
            f"QPushButton:hover {{ background: {C.INFO_SOFT}; border-color: {C.INFO_ACCENT}; }}"
        )
        btn_edit_loc.clicked.connect(lambda _, pr=phys_room, sl=slot: self._edit_slot_or_room_location(pr, sl))
        rm_top.addWidget(btn_edit_loc)
        rb_lay.addLayout(rm_top)

        if loc_note:
            loc_lbl = QLabel(f"🧭 <b>{html.escape(loc_note)}</b>")
            loc_lbl.setWordWrap(True)
            loc_lbl.setStyleSheet(
                f"color: {C.INFO_DARK}; background: {C.INFO_SOFT}; border: 1px solid {C.INFO_BORDER}; "
                f"border-radius: 4px; padding: 2px 6px; font-size: 11px;"
            )
            rb_lay.addWidget(loc_lbl)
        else:
            btn_add_hint = QPushButton("+ Add floor / room directions...")
            btn_add_hint.setCursor(Qt.PointingHandCursor)
            btn_add_hint.setStyleSheet(
                f"QPushButton {{ text-align: left; border: none; background: transparent; color: {C.TEXT_MUTED}; font-size: 10.5px; padding: 0; }}"
                f"QPushButton:hover {{ color: {C.INFO}; text-decoration: underline; }}"
            )
            btn_add_hint.clicked.connect(lambda _, pr=phys_room, sl=slot: self._edit_slot_or_room_location(pr, sl))
            rb_lay.addWidget(btn_add_hint)

        meta_bits = []
        if s_info and s_info.convener and not is_swapped:
            meta_bits.append(f"👤 {s_info.convener}")
        elif s_info and s_info.convener:
            meta_bits.append(f"👤 {s_info.convener}")
        if s_info and s_info.tohru_room:
            meta_bits.append(f"📡 {s_info.tohru_room}")
        if meta_bits:
            meta_lbl = QLabel(" · ".join(meta_bits))
            meta_lbl.setStyleSheet(f"color: {C.TEXT_MUTED}; font-size: 10.5px; border: none; background: transparent;")
            rb_lay.addWidget(meta_lbl)

        lay.addWidget(room_box)

        # Row 3: Topic Title & Resolved Agenda Items
        clean_txt, _ = format_slot_display_text(slot.raw_text, getattr(slot, "room_override", "") or "")
        topic_lbl = QLabel(clean_txt or slot.raw_text)
        topic_lbl.setWordWrap(True)
        topic_lbl.setStyleSheet(f"color: {C.TEXT_PRIMARY}; font-weight: 600; font-size: 12.5px; border: none; background: transparent;")
        lay.addWidget(topic_lbl)

        ai_bullets = build_slot_ai_bullets_html(slot, agenda_map=self.agenda_map, plan=self.delegation_plan, max_expanded=6)
        if ai_bullets:
            ai_lbl = QLabel("<br>".join(ai_bullets))
            ai_lbl.setWordWrap(True)
            ai_lbl.setTextFormat(Qt.RichText)
            ai_lbl.setStyleSheet(f"color: {C.TEXT_BODY}; font-size: 11px; line-height: 1.35; border: none; background: transparent;")
            lay.addWidget(ai_lbl)

        ann = self.annotations.get(slot.slot_key)
        if ann and ann.comment:
            note_lbl = QLabel(f"📝 {ann.comment}")
            note_lbl.setWordWrap(True)
            note_lbl.setStyleSheet(
                f"background: {C.WARNING_SOFT_ALT}; color: {C.WARNING_DARK}; border-radius: 4px; padding: 3px 6px; font-size: 11px;"
            )
            lay.addWidget(note_lbl)

        lay.addStretch(1)

        # Optional Footer: Half-Quarter TU 2 preview + Show in Grid button
        if tu2_slot:
            tu2_room, tu2_loc, _ = resolve_effective_room_and_location(
                tu2_slot, streams=streams, room_locations=self.room_locations, slot_locations=self.slot_locations
            )
            tu2_clean, _ = format_slot_display_text(tu2_slot.raw_text, getattr(tu2_slot, "room_override", "") or "")
            tu2_loc_s = f" ({tu2_loc})" if tu2_loc else ""
            tu2_bar = QLabel(
                f"⏭️ <b>2nd Half ({tu2_slot.start_time_local}–{tu2_slot.end_time_local}):</b> "
                f"📍 {html.escape(tu2_room)}{html.escape(tu2_loc_s)} — {html.escape(tu2_clean[:85])}"
            )
            tu2_bar.setWordWrap(True)
            tu2_bar.setStyleSheet(
                f"background: {C.BG_SURFACE}; color: {C.TEXT_SECONDARY}; border-top: 1px dashed {C.BORDER_STRONG}; "
                f"padding-top: 4px; font-size: 10.5px;"
            )
            lay.addWidget(tu2_bar)

        foot_row = QHBoxLayout()
        foot_row.addStretch(1)
        btn_grid_jump = QPushButton("🔍 Show in Grid")
        btn_grid_jump.setCursor(Qt.PointingHandCursor)
        btn_grid_jump.setStyleSheet(
            f"QPushButton {{ padding: 2px 8px; font-size: 10.5px; font-weight: 600; border-radius: 4px; "
            f"background: {C.INFO_SOFT}; color: {C.INFO}; border: 1px solid {C.INFO_BORDER}; }}"
            f"QPushButton:hover {{ background: #DBEAFE; }}"
        )
        btn_grid_jump.clicked.connect(lambda _, sl=slot: self._jump_to_slot_in_grid(sl.day_index, sl.quarter, sl.stream_number, sl.sub_slot))
        foot_row.addWidget(btn_grid_jump)
        lay.addLayout(foot_row)

        return card

    def _make_upcoming_period_card(self, day_idx: int, period_info: dict) -> QFrame:
        """Creates an upcoming period summary card showing rooms, custom locations, and topics across streams."""
        p_name = period_info["period"]
        venue_t = format_venue_time_range(period_info["start"], period_info["end"])
        streams = self.active_snapshot.streams if self.active_snapshot else {}

        frame = QFrame()
        frame.setStyleSheet(
            f"QFrame {{ background: {C.BG_CARD}; border: 1px solid {C.BORDER_SUBTLE}; border-radius: 8px; }}"
        )
        lay = QVBoxLayout(frame)
        lay.setContentsMargins(10, 8, 10, 8)
        lay.setSpacing(6)

        hdr = QHBoxLayout()
        p_icon = "☕" if period_info["is_break"] else "⏭️"
        hdr_lbl = QLabel(f"<b>{p_icon} {html.escape(p_name)}</b> &nbsp;·&nbsp; <span style='color:{C.TEXT_SECONDARY};'>🕒 {html.escape(venue_t)}</span>")
        hdr_lbl.setStyleSheet(f"color: {C.TEXT_PRIMARY}; font-size: 12px; border: none; background: transparent;")
        hdr.addWidget(hdr_lbl)
        hdr.addStretch(1)
        lay.addLayout(hdr)

        slots_row = QHBoxLayout()
        slots_row.setSpacing(8)
        for sl in period_info["slots"]:
            phys_room, loc_note, is_swapped = resolve_effective_room_and_location(
                sl, streams=streams, room_locations=self.room_locations, slot_locations=self.slot_locations
            )
            clean_txt, _ = format_slot_display_text(sl.raw_text, getattr(sl, "room_override", "") or "")
            tu_tag = f" [TU {sl.sub_slot}]" if sl.sub_slot in (1, 2) else ""
            sub_card = QFrame()
            sub_card.setStyleSheet(
                f"QFrame {{ background: {C.BG_SURFACE}; border: 1px solid {C.BORDER_SUBTLE}; border-radius: 6px; padding: 5px 7px; }}"
            )
            sc_lay = QVBoxLayout(sub_card)
            sc_lay.setContentsMargins(4, 3, 4, 3)
            sc_lay.setSpacing(2)

            loc_s = f" · 🧭 {html.escape(loc_note)}" if loc_note else ""
            sw_s = " ⇄" if is_swapped else ""
            rm_lbl = QLabel(f"<b>Str {sl.stream_number}{tu_tag} · 📍 {html.escape(phys_room)}{sw_s}</b>{loc_s}")
            rm_lbl.setWordWrap(True)
            rm_lbl.setStyleSheet(f"color: {C.INFO_DARK}; font-size: 11px; border: none; background: transparent;")
            sc_lay.addWidget(rm_lbl)

            tp_lbl = QLabel(clean_txt or sl.raw_text)
            tp_lbl.setWordWrap(True)
            tp_lbl.setStyleSheet(f"color: {C.TEXT_PRIMARY}; font-size: 11px; font-weight: 500; border: none; background: transparent;")
            sc_lay.addWidget(tp_lbl)
            slots_row.addWidget(sub_card, 1)

        lay.addLayout(slots_row)
        return frame

    def _make_compact_timeline_row(self, day_idx: int, period_info: dict) -> QFrame:
        """Creates a compact rest-of-day chronological strip row."""
        streams = self.active_snapshot.streams if self.active_snapshot else {}
        p_name = period_info["period"]
        venue_t = format_venue_time_range(period_info["start"], period_info["end"])

        row_frame = QFrame()
        row_frame.setStyleSheet(
            f"QFrame {{ background: {C.BG_SURFACE}; border: 1px solid {C.BORDER_SUBTLE}; border-radius: 6px; }}"
        )
        r_lay = QHBoxLayout(row_frame)
        r_lay.setContentsMargins(10, 6, 10, 6)
        r_lay.setSpacing(10)

        left_lbl = QLabel(f"<b>{html.escape(p_name)}</b><br><span style='color:{C.TEXT_MUTED}; font-size:10.5px;'>{html.escape(venue_t)}</span>")
        left_lbl.setFixedWidth(185)
        left_lbl.setStyleSheet("border: none; background: transparent;")
        r_lay.addWidget(left_lbl)

        parts = []
        for sl in period_info["slots"]:
            phys_room, loc_note, _ = resolve_effective_room_and_location(
                sl, streams=streams, room_locations=self.room_locations, slot_locations=self.slot_locations
            )
            clean_txt, _ = format_slot_display_text(sl.raw_text, getattr(sl, "room_override", "") or "")
            loc_s = f" ({html.escape(loc_note)})" if loc_note else ""
            tu_s = f" TU{sl.sub_slot}" if sl.sub_slot in (1, 2) else ""
            parts.append(f"<b>Str {sl.stream_number}{tu_s} [📍 {html.escape(phys_room)}{loc_s}]:</b> {html.escape(clean_txt[:75])}")

        summary_lbl = QLabel(" &nbsp;|&nbsp; ".join(parts))
        summary_lbl.setWordWrap(True)
        summary_lbl.setStyleSheet(f"color: {C.TEXT_BODY}; font-size: 11px; border: none; background: transparent;")
        r_lay.addWidget(summary_lbl, 1)

        first_sl = period_info["slots"][0] if period_info["slots"] else None
        if first_sl:
            btn = QPushButton("🔍 Grid")
            btn.setCursor(Qt.PointingHandCursor)
            btn.clicked.connect(lambda _, sl=first_sl: self._jump_to_slot_in_grid(sl.day_index, sl.quarter, sl.stream_number, sl.sub_slot))
            r_lay.addWidget(btn)

        return row_frame
