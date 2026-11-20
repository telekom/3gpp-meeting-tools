# --- File: src/modules/meetings/core/session_plan/exporter/facade.py ---
"""
Backward-compatible `SessionPlanExporter` class facade delegating to the
specialized HTML, ICS, and e-mail exporter modules.
"""
from pathlib import Path
from typing import Dict, List, Optional, Set, Tuple

from modules.meetings.core.session_plan.exporter.email_exporter import (
    generate_delegation_email_summary as _generate_delegation_email_summary,
    generate_email_summary as _generate_email_summary,
    parse_delegation_plan_json as _parse_delegation_plan_json,
    serialize_delegation_plan_json as _serialize_delegation_plan_json,
)
from modules.meetings.core.session_plan.exporter.helpers import (
    get_slot_effective_delegate as _get_slot_effective_delegate,
    get_slot_effective_delegates as _get_slot_effective_delegates,
    is_slot_tracked as _is_slot_tracked_fn,
    load_folder_locations_fallback as _load_folder_locations_fallback_fn,
)
from modules.meetings.core.session_plan.exporter.html_exporter import (
    export_mobile_html as _export_mobile_html,
)
from modules.meetings.core.session_plan.exporter.ics_exporter import (
    export_ics as _export_ics,
)
from modules.meetings.core.session_plan.session_plan_diff import SessionPlanDiffResult
from modules.meetings.core.session_plan.session_plan_models import (
    DelegationPlan,
    SessionPlanSnapshot,
    SessionSlot,
    UserSessionAnnotation,
)


class SessionPlanExporter:
    """Generates mobile HTML, .ics calendar, and email summaries."""

    @staticmethod
    def _is_slot_tracked(slot: SessionSlot, tracked_tokens: Set[str], agenda_map: Optional[dict] = None) -> bool:
        """Determines whether a slot matches any user-tracked topic or agenda item."""
        return _is_slot_tracked_fn(slot, tracked_tokens, agenda_map=agenda_map)

    @staticmethod
    def get_slot_effective_delegates(
        slot: SessionSlot, plan: Optional[DelegationPlan], agenda_map: Optional[dict] = None
    ) -> Tuple[List[str], bool, bool]:
        """Resolves effective delegates list, is_coverage flag, and is_all flag."""
        return _get_slot_effective_delegates(slot, plan, agenda_map=agenda_map)

    @classmethod
    def get_slot_effective_delegate(
        cls, slot: SessionSlot, plan: Optional[DelegationPlan], agenda_map: Optional[dict] = None
    ) -> Tuple[Optional[str], bool]:
        """Backward-compatible single-string delegate helper."""
        return _get_slot_effective_delegate(slot, plan, agenda_map=agenda_map)

    @staticmethod
    def _load_folder_locations_fallback(output_path: Path) -> Tuple[Dict[str, str], Dict[str, str]]:
        """Loads `<meeting_dir>/meeting_locations.json` when exporting without explicit location maps."""
        return _load_folder_locations_fallback_fn(output_path)

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
        room_locations: Optional[Dict[str, str]] = None,
        slot_locations: Optional[Dict[str, str]] = None,
    ) -> Path:
        """Exports a self-contained, responsive, 100% offline HTML view."""
        return _export_mobile_html(
            snapshot=snapshot,
            output_path=output_path,
            meeting_tz_offset_min=meeting_tz_offset_min,
            tz_name=tz_name,
            meeting_dates_str=meeting_dates_str,
            annotations=annotations,
            tracked_tokens=tracked_tokens,
            delegation_plan=delegation_plan,
            agenda_map=agenda_map,
            room_locations=room_locations,
            slot_locations=slot_locations,
        )

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
        agenda_map: Optional[dict] = None,
        room_locations: Optional[Dict[str, str]] = None,
        slot_locations: Optional[Dict[str, str]] = None,
    ) -> Path:
        """Exports an RFC 5545 .ics Calendar file strictly anchored in UTC."""
        return _export_ics(
            snapshot=snapshot,
            output_path=output_path,
            meeting_start_date_str=meeting_start_date_str,
            meeting_tz_offset_min=meeting_tz_offset_min,
            only_tracked=only_tracked,
            only_drafting=only_drafting,
            tracked_tokens=tracked_tokens,
            annotations=annotations,
            delegation_plan=delegation_plan,
            target_delegate=target_delegate,
            agenda_map=agenda_map,
            room_locations=room_locations,
            slot_locations=slot_locations,
        )

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
        """Generates pre-formatted rich HTML email body ready for Outlook."""
        return _generate_email_summary(
            snapshot=snapshot,
            meeting_tz_offset_min=meeting_tz_offset_min,
            tz_name=tz_name,
            diff_result=diff_result,
            tracked_tokens=tracked_tokens,
            annotations=annotations,
        )

    @staticmethod
    def serialize_delegation_plan_json(plan: DelegationPlan) -> str:
        """Serializes a DelegationPlan to formatted JSON string."""
        return _serialize_delegation_plan_json(plan)

    @staticmethod
    def parse_delegation_plan_json(raw_text: str) -> Optional[dict]:
        """Parses delegation JSON from raw string, file content, or email text."""
        return _parse_delegation_plan_json(raw_text)

    @classmethod
    def generate_delegation_email_summary(
        cls,
        snapshot: SessionPlanSnapshot,
        plan: DelegationPlan,
        meeting_tz_offset_min: int,
        tz_name: str,
        agenda_map: Optional[dict] = None,
    ) -> Tuple[str, str, str]:
        """Generates executive delegation email summary (subject, html_body, plain_text)."""
        return _generate_delegation_email_summary(
            snapshot=snapshot,
            plan=plan,
            meeting_tz_offset_min=meeting_tz_offset_min,
            tz_name=tz_name,
            agenda_map=agenda_map,
        )
