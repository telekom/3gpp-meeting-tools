# --- File: src/modules/meetings/core/session_plan/exporter/ics_exporter.py ---
"""
RFC 5545 iCalendar (.ics) exporter for 3GPP SA2 Session Plans.
Produces UTC-anchored VEVENT entries for full delegation or individual delegates.
"""
import re
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Dict, Optional, Set

from modules.meetings.core.session_plan.exporter.helpers import (
    get_slot_effective_delegates,
    is_slot_tracked,
    load_folder_locations_fallback,
    logger,
)
from modules.meetings.core.session_plan.session_plan_models import (
    DelegationPlan,
    SessionPlanSnapshot,
    UserSessionAnnotation,
    resolve_effective_room_and_location,
)
from modules.meetings.core.session_plan.session_plan_timezone import SessionPlanTimezone


def export_ics(
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
    agenda_map: Optional[dict] = None,
    room_locations: Optional[Dict[str, str]] = None,
    slot_locations: Optional[Dict[str, str]] = None,
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
    if room_locations is None or slot_locations is None:
        fb_r, fb_s = load_folder_locations_fallback(output_path)
        if room_locations is None:
            room_locations = fb_r
        if slot_locations is None:
            slot_locations = fb_s

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
        is_tracked = is_slot_tracked(s, tracked_tokens, agenda_map=agenda_map)

        if only_tracked and not is_tracked and not s.is_drafting:
            continue

        # Check delegate assignment / target delegate filtering
        eff_dels, is_cov, is_all = get_slot_effective_delegates(s, delegation_plan, agenda_map=agenda_map)
        if target_delegate:
            td_lower = target_delegate.strip().lower()
            if not (is_all or any(ed.strip().lower() == td_lower for ed in eff_dels)):
                continue

        # Calculate physical date
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
        r_name, loc_note, _ = resolve_effective_room_and_location(
            s, streams=streams, room_locations=room_locations, slot_locations=slot_locations
        )
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
        location = f"{r_name} ({loc_note})" if loc_note else f"{r_name}"

        desc_lines = [
            f"Session: SA2 #{snapshot.meeting_number} ({snapshot.tdoc_id})",
            f"Room: {location}",
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
