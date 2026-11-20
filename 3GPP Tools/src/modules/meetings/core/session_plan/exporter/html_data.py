# --- File: src/modules/meetings/core/session_plan/exporter/html_data.py ---
"""
Data preparation for standalone mobile HTML Session Plan exports:
timezone calculations, day/quarter slot maps, unassigned AI catalogs, and serialized JS payloads.
"""
import json
import re
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Dict, List, Optional, Set, Tuple

from modules.meetings.core.session_plan.exporter.helpers import (
    get_slot_effective_delegates,
    is_slot_unassigned,
    load_folder_locations_fallback,
)
from modules.meetings.core.session_plan.session_plan_models import (
    DelegationMember,
    DelegationPlan,
    SessionPlanSnapshot,
    SessionSlot,
    TopicOwnership,
    UserSessionAnnotation,
    _matching_topic_ownerships,
    ai_numbers_match,
    ai_sort_key,
    compute_unique_initials,
    expand_slot_ai_numbers,
    get_slot_unassigned_ais,
    is_slot_not_followed,
    lookup_room_location_note,
    resolve_effective_room_and_location,
)
from modules.meetings.core.session_plan.session_plan_parser import DAY_NAMES, PERIOD_ORDER
from modules.meetings.core.session_plan.session_plan_timezone import SessionPlanTimezone


@dataclass
class MobileHtmlContext:
    """Prepared data and serialized JSON payloads consumed by HTML renderers."""

    snapshot: SessionPlanSnapshot
    meeting_tz_offset_min: int
    tz_name: str
    annotations: Dict[str, UserSessionAnnotation]
    tracked_tokens: Set[str]
    delegation_plan: Optional[DelegationPlan]
    agenda_map: dict
    room_locations: Dict[str, str]
    slot_locations: Dict[str, str]
    is_tz_different: bool
    tz_diff_str: str
    venue_utc_str: str
    base_date: Optional[date]
    venue_today: date
    base_date_js: str
    members: Dict[str, DelegationMember]
    initials_map: Dict[str, str]
    topic_owners_map: Dict[str, TopicOwnership]
    all_periods: List[str]
    slot_map: Dict[Tuple[int, str, int], List[SessionSlot]]
    break_slot_map: Dict[Tuple[int, str], List[SessionSlot]]
    slots_by_day_q: Dict[str, Dict[str, List[SessionSlot]]]
    active_break_periods: Set[str]
    day_parity_classes: List[str]
    today_day_idx: int
    cal_events_json: str
    roster_members_json: str
    room_locations_json: str
    streams_meta_json: str
    unassigned_ais_json: str
    unassigned_slots_json: str
    unassigned_slot_count: int
    unassigned_ai_count: int


def build_mobile_html_context(
    snapshot: SessionPlanSnapshot,
    output_path: Path,
    meeting_tz_offset_min: int = 120,
    tz_name: str = "Venue Time",
    meeting_dates_str: str = "",
    annotations: Optional[Dict[str, UserSessionAnnotation]] = None,
    tracked_tokens: Optional[Set[str]] = None,
    delegation_plan: Optional[DelegationPlan] = None,
    agenda_map: Optional[dict] = None,
    room_locations: Optional[Dict[str, str]] = None,
    slot_locations: Optional[Dict[str, str]] = None,
) -> MobileHtmlContext:
    """Normalizes inputs and builds slot maps, unassigned AI catalogs, and JSON payloads."""
    annotations = annotations or {}
    tracked_tokens = {t.lower() for t in (tracked_tokens or set())}
    agenda_map = agenda_map or {}
    output_path = Path(output_path)
    if room_locations is None or slot_locations is None:
        fb_r, fb_s = load_folder_locations_fallback(output_path)
        if room_locations is None:
            room_locations = fb_r
        if slot_locations is None:
            slot_locations = fb_s
    room_locations = dict(room_locations or {})
    slot_locations = dict(slot_locations or {})

    laptop_offset_min = SessionPlanTimezone.get_laptop_tz_offset_minutes()
    is_tz_different = (laptop_offset_min != meeting_tz_offset_min)
    tz_diff_min = laptop_offset_min - meeting_tz_offset_min
    tz_sign = "+" if tz_diff_min >= 0 else "-"
    tz_abs_h, tz_abs_m = divmod(abs(tz_diff_min), 60)
    tz_diff_str = f"{tz_sign}{tz_abs_h}h" if tz_abs_m == 0 else f"{tz_sign}{tz_abs_h}h{tz_abs_m:02d}m"

    v_sign = "+" if meeting_tz_offset_min >= 0 else "-"
    v_abs_h, v_abs_m = divmod(abs(meeting_tz_offset_min), 60)
    venue_utc_str = f"UTC{v_sign}{v_abs_h:02d}:{v_abs_m:02d}"

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
    topic_owners_map = delegation_plan.topic_owners if delegation_plan else {}

    unassigned_slots_data = []
    scheduled_ai_slots: Dict[str, List[dict]] = {}

    for s in snapshot.slots:
        if not s.raw_text.strip() or s.raw_text.strip() == "-":
            continue
        eff_dels, is_cov, is_all = get_slot_effective_delegates(s, delegation_plan, agenda_map=agenda_map)
        unassigned_ais = get_slot_unassigned_ais(s, delegation_plan, agenda_map=agenda_map)
        is_unassigned_slot = is_slot_unassigned(s, delegation_plan, eff_dels, is_all, agenda_map=agenda_map)

        s_info = snapshot.streams.get(s.stream_number)
        r_name, loc_note, is_swapped = resolve_effective_room_and_location(
            s, streams=snapshot.streams, room_locations=room_locations, slot_locations=slot_locations
        )
        convener = s_info.convener if s_info else ""
        tohru = s_info.tohru_room if s_info else ""
        ann = annotations.get(s.slot_key)
        note_str = ann.comment if ann else ""
        ai_nums = list(dict.fromkeys([it.ai_number for it in s.items if it.ai_number] + expand_slot_ai_numbers(s, agenda_map)))

        ev_dict = {
            "day_idx": s.day_index,
            "day_name": s.day_name or (DAY_NAMES[s.day_index] if s.day_index < len(DAY_NAMES) else f"Day {s.day_index+1}"),
            "quarter": s.quarter,
            "sub_slot": s.sub_slot,
            "start_local": s.start_time_local,
            "end_local": s.end_time_local,
            "stream_num": s.stream_number,
            "room": r_name,
            "location_note": loc_note,
            "room_swapped": is_swapped,
            "convener": convener,
            "tohru": tohru,
            "is_break": bool(any(b in (s.quarter or "").lower() for b in ("break", "coffee", "lunch"))),
            "title": s.raw_text,
            "is_drafting": s.is_drafting,
            "is_all": is_all,
            "is_cov": is_cov,
            "delegates": eff_dels,
            "unassigned": is_unassigned_slot,
            "unassigned_ais": unassigned_ais,
            "not_followed": is_slot_not_followed(s, delegation_plan, agenda_map=agenda_map),
            "ai_numbers": ai_nums,
            "note": note_str,
            "slot_key": s.slot_key,
        }
        cal_events_data.append(ev_dict)

        for it in s.items:
            if it.ai_number:
                scheduled_ai_slots.setdefault(it.ai_number, []).append({
                    "day_idx": s.day_index,
                    "day_name": ev_dict["day_name"],
                    "quarter": s.quarter,
                    "sub_slot": s.sub_slot,
                    "stream_num": s.stream_number,
                    "room": r_name,
                    "start_local": s.start_time_local,
                    "end_local": s.end_time_local,
                    "tdoc_count": it.tdoc_count_str,
                    "raw_text": s.raw_text,
                })

        if is_unassigned_slot:
            unassigned_slots_data.append(ev_dict)

    # Build Unassigned Agenda Items catalog (both scheduled and unscheduled AIs)
    all_known_ais: Dict[str, dict] = {}
    for ai_num, ai_obj in agenda_map.items():
        desc = getattr(ai_obj, "description", "") or str(ai_obj)
        all_known_ais[ai_num] = {
            "ai_number": ai_num,
            "title": desc,
            "tdoc_count": "",
        }
    for s in snapshot.slots:
        for it in s.items:
            if it.ai_number:
                existing = all_known_ais.get(it.ai_number)
                if not existing:
                    all_known_ais[it.ai_number] = {
                        "ai_number": it.ai_number,
                        "title": it.topic_label or s.raw_text,
                        "tdoc_count": it.tdoc_count_str or "",
                    }
                elif it.tdoc_count_str and not existing.get("tdoc_count"):
                    existing["tdoc_count"] = it.tdoc_count_str

    unassigned_ais_catalog = []
    for ai_num in sorted(all_known_ais.keys(), key=ai_sort_key):
        matches = _matching_topic_ownerships(ai_num, topic_owners_map, agenda_map=agenda_map)
        if any(to.is_not_followed() or to.owner_name.strip() for to in matches):
            continue
        if ai_num.lower().endswith(".x") and any(
            k != ai_num and ai_numbers_match(ai_num, k) for k in agenda_map
        ):
            continue
        info = all_known_ais[ai_num]
        sched_list = scheduled_ai_slots.get(ai_num)
        if sched_list is None:
            sched_list = []
            for sched_ai, sls in scheduled_ai_slots.items():
                if ai_numbers_match(ai_num, sched_ai):
                    sched_list.extend(sls)
        unassigned_ais_catalog.append({
            "ai_number": ai_num,
            "title": info["title"],
            "tdoc_count": info["tdoc_count"],
            "scheduled_count": len(sched_list),
            "slots": sched_list,
        })

    roster_members_data = [
        {
            "name": m.name,
            "initials": initials_map.get(m.name, m.name[:2].upper()),
            "color": m.color,
            "bg_tint": m.bg_tint,
            "border_tint": m.border_tint,
            "role": getattr(m, "role", "") or "",
        }
        for m in members.values()
    ]

    cal_events_json = json.dumps(cal_events_data, ensure_ascii=False)
    roster_members_json = json.dumps(roster_members_data, ensure_ascii=False)
    room_locations_json = json.dumps(room_locations, ensure_ascii=False)
    streams_meta_data = {
        str(sn): {
            "stream_num": sn,
            "room_name": si.room_name.split("(")[0].strip() or si.room_name,
            "location_note": lookup_room_location_note(si.room_name.split("(")[0].strip() or si.room_name, room_locations),
            "convener": si.convener,
            "tohru": si.tohru_room,
        }
        for sn, si in sorted(snapshot.streams.items())
    }
    streams_meta_json = json.dumps(streams_meta_data, ensure_ascii=False)
    unassigned_ais_json = json.dumps(unassigned_ais_catalog, ensure_ascii=False)
    unassigned_slots_json = json.dumps(unassigned_slots_data, ensure_ascii=False)
    base_date_js = base_date.strftime("%Y-%m-%d") if base_date else ""

    day_parity_classes = []
    today_day_idx = -1
    for d_i in range(5):
        if base_date and (base_date + timedelta(days=d_i)) == venue_today:
            day_parity_classes.append("day-today")
            today_day_idx = d_i
        elif d_i % 2 == 1:
            day_parity_classes.append("day-odd")
        else:
            day_parity_classes.append("day-even")

    return MobileHtmlContext(
        snapshot=snapshot,
        meeting_tz_offset_min=meeting_tz_offset_min,
        tz_name=tz_name,
        annotations=annotations,
        tracked_tokens=tracked_tokens,
        delegation_plan=delegation_plan,
        agenda_map=agenda_map,
        room_locations=room_locations,
        slot_locations=slot_locations,
        is_tz_different=is_tz_different,
        tz_diff_str=tz_diff_str,
        venue_utc_str=venue_utc_str,
        base_date=base_date,
        venue_today=venue_today,
        base_date_js=base_date_js,
        members=members,
        initials_map=initials_map,
        topic_owners_map=topic_owners_map,
        all_periods=all_periods,
        slot_map=slot_map,
        break_slot_map=break_slot_map,
        slots_by_day_q=slots_by_day_q,
        active_break_periods=active_break_periods,
        day_parity_classes=day_parity_classes,
        today_day_idx=today_day_idx,
        cal_events_json=cal_events_json,
        roster_members_json=roster_members_json,
        room_locations_json=room_locations_json,
        streams_meta_json=streams_meta_json,
        unassigned_ais_json=unassigned_ais_json,
        unassigned_slots_json=unassigned_slots_json,
        unassigned_slot_count=len(unassigned_slots_data),
        unassigned_ai_count=len(unassigned_ais_catalog),
    )
