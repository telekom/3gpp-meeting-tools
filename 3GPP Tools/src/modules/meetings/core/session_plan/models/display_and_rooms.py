"""Slot display normalization, HTML bullet builders, and physical room/location resolution."""

from __future__ import annotations

import html
import re
from typing import Any, Dict, List, Optional, Tuple

from modules.meetings.core.session_plan.models.dataclasses import (
    DelegationPlan,
    MeetingStreamInfo,
    SessionPlanSnapshot,
    SessionSlot,
)
from modules.meetings.core.session_plan.models.delegation_resolver import (
    _resolve_single_ai_ownerships,
    expand_ai_number,
)

_ROOM_PREFIX_RE = re.compile(
    r"^\s*(?:\(?\s*in\s+(?:the\s+)?)?(main room|breakout\s*1|breakout\s*2|bmo room \d+|palomino [a-z\-]+|percheron [a-z\-]+)\s*\)?[\s:.\-–—]*",
    re.I,
)
_TRAILING_COUNT_RE = re.compile(r"^\[\s*[\d\+\,\s]+\s*\]$")
_EMPTY_DRAFTING_RE = re.compile(r"^\s*drafting\s*:?\s*$", re.I)
_PAREN_CAP_RE = re.compile(r"\s*\([^)]*\)\s*$")


def format_slot_display_text(raw_text: str, room_override: str = "") -> Tuple[str, bool]:
    """
    Normalizes session slot text for compact card display in Desktop & HTML grids:
    - Detects empty placeholder cells like ``Drafting:`` (``is_empty_drafting=True``).
    - Strips redundant leading room-override lines when ``room_override`` is shown as a badge.
    - Merges standalone trailing ``[TDocCount]`` lines onto the preceding topic line.
    """
    raw = (raw_text or "").strip()
    is_empty_drafting = bool(_EMPTY_DRAFTING_RE.match(raw))
    lines = [ln.strip() for ln in raw.splitlines() if ln.strip()]

    cleaned_lines: List[str] = []
    for ln in lines:
        if room_override:
            m_room = _ROOM_PREFIX_RE.match(ln)
            if m_room:
                rest = ln[m_room.end():].strip()
                if rest:
                    cleaned_lines.append(rest)
                continue
        if cleaned_lines and _TRAILING_COUNT_RE.match(ln):
            cleaned_lines[-1] = f"{cleaned_lines[-1]} {ln}"
            continue
        cleaned_lines.append(ln)

    return "\n".join(cleaned_lines), is_empty_drafting


def format_break_period_label(period: str) -> str:
    """Normalizes internal break period codes (`Coffee 1`, `Break 1`, `Lunch`) into clean display titles."""
    p_low = (period or "").strip().lower()
    if "coffee" in p_low:
        return "☕ Coffee Break"
    if "lunch" in p_low:
        return "🍱 Lunch Break"
    if "break" in p_low:
        return "☕ Short Break"
    return (period or "").strip()


def format_venue_time_range(start_hhmm: str, end_hhmm: str) -> str:
    """Formats a venue local start/end HH:MM time range without secondary timezone suffixes."""
    s = (start_hhmm or "").strip()
    e = (end_hhmm or "").strip()
    if s and e:
        return f"{s} – {e}"
    return s or e


def build_slot_ai_bullets_html(
    slot: SessionSlot,
    agenda_map: Optional[Dict[str, Any]] = None,
    plan: Optional[DelegationPlan] = None,
    max_expanded: int = 12,
) -> List[str]:
    """
    Formats HTML bullet lines for a slot's Agenda Items, expanding ``.x`` wildcards
    into their concrete sub-AIs from ``agenda_map`` (with per-sub-AI delegation status
    when ``plan`` is provided and no slot-level override masks them).
    """
    bullets: List[str] = []
    has_override = bool(plan and slot.slot_key in plan.appointments)
    topic_owners = plan.topic_owners if plan else {}

    for it in slot.items:
        if not it.ai_number:
            continue
        tdoc_str = f" [{it.tdoc_count_str}]" if it.tdoc_count_str else ""
        is_wild = it.ai_number.strip().lower().endswith(".x")
        expanded = expand_ai_number(it.ai_number, agenda_map) if agenda_map else [it.ai_number]
        did_expand = bool(
            agenda_map
            and is_wild
            and expanded
            and not (len(expanded) == 1 and expanded[0].strip().lower() == it.ai_number.strip().lower())
        )
        if not did_expand:
            ai_info = agenda_map.get(it.ai_number) if agenda_map else None
            desc = (getattr(ai_info, "description", None) if ai_info else None) or it.topic_label
            bullets.append(
                f"• <b>AI {html.escape(it.ai_number)}:</b> {html.escape(desc)}{html.escape(tdoc_str)}"
            )
            continue

        bullets.append(
            f"• <b>AI {html.escape(it.ai_number)}:</b> {html.escape(it.topic_label)}{html.escape(tdoc_str)} "
            f"<span style='color:#64748B;'>(covers {len(expanded)} agenda item(s))</span>"
        )
        parent_prefix = it.ai_number.strip()[:-2].strip()
        for sub_ai in expanded[:max_expanded]:
            sub_info = agenda_map.get(sub_ai) if agenda_map else None
            sub_desc = (getattr(sub_info, "description", None) if sub_info else None) or f"Agenda Item {sub_ai}"
            status_tag = ""
            if plan and not has_override:
                matches = _resolve_single_ai_ownerships(sub_ai, topic_owners, include_parent_prefix=parent_prefix)
                if matches:
                    m_to = matches[0]
                    if m_to.is_not_followed():
                        status_tag = " <span style='color:#64748B; font-weight:600;'>[⛔ Not Followed]</span>"
                    elif m_to.is_all():
                        status_tag = " <span style='color:#4338CA; font-weight:600;'>[👥 All Team]</span>"
                    elif m_to.owner_name.strip():
                        status_tag = f" <span style='color:#1E40AF; font-weight:600;'>[● {html.escape(m_to.owner_name.strip())}]</span>"
                    else:
                        status_tag = " <span style='color:#B45309; font-weight:600;'>[❓ Unassigned]</span>"
                else:
                    status_tag = " <span style='color:#B45309; font-weight:600;'>[❓ Unassigned]</span>"
            bullets.append(
                f"&nbsp;&nbsp;&nbsp;◦ <b>{html.escape(sub_ai)}:</b> {html.escape(sub_desc)}{status_tag}"
            )
        if len(expanded) > max_expanded:
            bullets.append(
                f"&nbsp;&nbsp;&nbsp;◦ <i>… and {len(expanded) - max_expanded} more sub-item(s)</i>"
            )
    return bullets


def normalize_room_key(room_name: str) -> str:
    """Normalizes a physical room name for case-insensitive matching (stripping seat counts)."""
    cleaned = _PAREN_CAP_RE.sub("", (room_name or "").strip())
    return re.sub(r"\s+", " ", cleaned).strip().lower()


def collect_meeting_physical_rooms(
    snapshot: Optional[SessionPlanSnapshot],
    existing_room_locations: Optional[Dict[str, str]] = None,
) -> List[str]:
    """
    Collects all unique physical rooms referenced by streams, slot room_overrides,
    and existing user location notes in deterministic order.
    """
    rooms: List[str] = []
    seen_keys: set = set()

    def _add(name: str):
        raw = (name or "").strip()
        if not raw:
            return
        display = _PAREN_CAP_RE.sub("", raw).strip() or raw
        key = normalize_room_key(display)
        if key and key not in seen_keys:
            seen_keys.add(key)
            rooms.append(display)

    if snapshot:
        for s_num in sorted(snapshot.streams.keys()):
            _add(snapshot.streams[s_num].room_name)
        for sl in snapshot.slots:
            if getattr(sl, "room_override", ""):
                _add(sl.room_override)

    if existing_room_locations:
        for r_name in existing_room_locations.keys():
            _add(r_name)

    return rooms


def lookup_room_location_note(
    physical_room: str,
    room_locations: Optional[Dict[str, str]],
) -> str:
    """Finds a custom location note for a physical room name (exact or normalized match)."""
    if not physical_room or not room_locations:
        return ""
    if physical_room in room_locations and room_locations[physical_room].strip():
        return room_locations[physical_room].strip()
    target_key = normalize_room_key(physical_room)
    if not target_key:
        return ""
    for k, val in room_locations.items():
        if val and val.strip() and normalize_room_key(k) == target_key:
            return val.strip()
    return ""


def resolve_effective_room_and_location(
    slot: SessionSlot,
    streams: Optional[Dict[int, MeetingStreamInfo]] = None,
    room_locations: Optional[Dict[str, str]] = None,
    slot_locations: Optional[Dict[str, str]] = None,
) -> Tuple[str, str, bool]:
    """
    Resolves `(physical_room, location_note, is_room_swapped)` for a slot.
    - Physical room uses `slot.room_override` when set, otherwise `streams[stream_number].room_name`.
    - Location note checks per-slot override `slot_locations[slot.slot_key]` first,
      then automatically inherits the note for `physical_room` from `room_locations`.
    """
    room_ov = (getattr(slot, "room_override", "") or "").strip()
    is_swapped = bool(room_ov)
    if room_ov:
        physical_room = room_ov
    elif streams and slot.stream_number in streams:
        raw_r = streams[slot.stream_number].room_name or ""
        physical_room = _PAREN_CAP_RE.sub("", raw_r).strip() or raw_r or f"Stream {slot.stream_number}"
    else:
        physical_room = f"Stream {slot.stream_number}" if slot.stream_number > 0 else "Plenary / Break"

    loc_note = ""
    if slot_locations:
        loc_note = (slot_locations.get(slot.slot_key) or "").strip()
    if not loc_note:
        loc_note = lookup_room_location_note(physical_room, room_locations)

    return physical_room, loc_note, is_swapped
