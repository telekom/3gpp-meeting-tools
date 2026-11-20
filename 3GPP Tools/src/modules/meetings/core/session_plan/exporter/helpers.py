# --- File: src/modules/meetings/core/session_plan/exporter/helpers.py ---
"""
Shared pure helpers for Session Plan HTML, ICS, and e-mail exporters:
delegate resolution, tracked/unassigned checks, period timings, and folder location fallback.
"""
import json
import logging
from pathlib import Path
from typing import Dict, List, Optional, Set, Tuple

from modules.meetings.core.session_plan.session_plan_models import (
    DelegationPlan,
    SessionPlanSnapshot,
    SessionSlot,
    expand_ai_number,
    get_slot_unassigned_ais,
    is_slot_not_followed,
    resolve_slot_effective_delegates,
)
from modules.meetings.core.session_plan.session_plan_parser import PERIOD_TIMES, SessionPlanParser

logger = logging.getLogger("modules.meetings.core.session_plan.session_plan_exporter")


def is_slot_tracked(
    slot: SessionSlot,
    tracked_tokens: Set[str],
    agenda_map: Optional[dict] = None,
) -> bool:
    """Determines whether a slot matches any user-tracked topic or agenda item."""
    if not tracked_tokens:
        return False
    raw_lower = slot.raw_text.lower()
    if any(tk in raw_lower for tk in tracked_tokens if tk):
        return True
    for it in slot.items:
        if it.ai_number.lower() in tracked_tokens or any(tk in it.topic_label.lower() for tk in tracked_tokens if tk):
            return True
        if it.ai_number:
            for sub_ai in expand_ai_number(it.ai_number, agenda_map):
                if sub_ai.lower() in tracked_tokens:
                    return True
    return False


def get_slot_effective_delegates(
    slot: SessionSlot,
    plan: Optional[DelegationPlan],
    agenda_map: Optional[dict] = None,
) -> Tuple[List[str], bool, bool]:
    """
    Resolves effective delegates list, is_coverage flag, and is_all flag: (delegates, is_cov, is_all).
    Delegates to the shared, boundary-aware resolver in session_plan_models so that the
    UI grid, clash detection, .ics and email exports always agree.
    """
    return resolve_slot_effective_delegates(slot, plan, agenda_map=agenda_map)


def get_slot_effective_delegate(
    slot: SessionSlot,
    plan: Optional[DelegationPlan],
    agenda_map: Optional[dict] = None,
) -> Tuple[Optional[str], bool]:
    """Backward-compatible single-string delegate helper."""
    dels, is_cov, is_all = get_slot_effective_delegates(slot, plan, agenda_map=agenda_map)
    if not dels:
        return None, False
    if is_all:
        return "ALL", is_cov
    return ", ".join(dels), is_cov


def load_folder_locations_fallback(output_path: Path) -> Tuple[Dict[str, str], Dict[str, str]]:
    """Loads `<meeting_dir>/meeting_locations.json` when exporting without explicit location maps."""
    candidates = [
        output_path.parent / "meeting_locations.json",
        output_path.parent.parent / "meeting_locations.json",
    ]
    for cand in candidates:
        if cand.is_file():
            try:
                raw = json.loads(cand.read_text(encoding="utf-8"))
                r_map = {str(k).strip(): str(v).strip() for k, v in (raw.get("room_locations") or {}).items() if str(k).strip() and str(v).strip()}
                s_map = {str(k).strip(): str(v).strip() for k, v in (raw.get("slot_locations") or {}).items() if str(k).strip() and str(v).strip()}
                return r_map, s_map
            except Exception:
                pass
    return {}, {}


def get_period_effective_times(
    snapshot: SessionPlanSnapshot,
    period: str,
) -> Tuple[str, str, str, str, str, str]:
    """Returns effective (full_s, full_e, tu1_s, tu1_e, tu2_s, tu2_e) times for a period."""
    default_t = PERIOD_TIMES.get(period, ("09:00", "10:30", "09:00", "09:45", "09:45", "10:30"))
    p_slots = [s for s in snapshot.slots if s.quarter == period]
    if not p_slots:
        return default_t
    full_slots = [s for s in p_slots if s.sub_slot == 0 and s.start_time_local and s.end_time_local]
    if full_slots:
        return SessionPlanParser._resolve_quarter_times(period, (full_slots[0].start_time_local, full_slots[0].end_time_local))
    tu1_starts = [s.start_time_local for s in p_slots if s.sub_slot == 1 and s.start_time_local]
    tu2_ends = [s.end_time_local for s in p_slots if s.sub_slot == 2 and s.end_time_local]
    s_str = min(tu1_starts) if tu1_starts else default_t[0]
    e_str = max(tu2_ends) if tu2_ends else default_t[1]
    return SessionPlanParser._resolve_quarter_times(period, (s_str, e_str))


def is_slot_unassigned(
    slot_obj: SessionSlot,
    delegation_plan: Optional[DelegationPlan],
    eff_dels_list: List[str],
    is_all_flag: bool,
    agenda_map: Optional[dict] = None,
) -> bool:
    """True if the slot has no assigned delegates at all, or contains an explicitly unassigned AI without slot override."""
    if not slot_obj.raw_text.strip() or slot_obj.raw_text.strip() == "-":
        return False
    if is_slot_not_followed(slot_obj, delegation_plan, agenda_map=agenda_map):
        return False
    if not eff_dels_list and not is_all_flag:
        return True
    if get_slot_unassigned_ais(slot_obj, delegation_plan, agenda_map=agenda_map):
        return True
    return False
