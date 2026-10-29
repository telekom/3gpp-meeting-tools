# --- File: src/modules/meetings/core/session_plan_models.py ---
"""
Domain models and dataclasses for 3GPP SA2 Meeting Session Plans.
Supports Half-Quarter Sub-Slots (Time Units), Stream/Room topologies,
AI Range Expansions, and User Annotations.
"""
from dataclasses import dataclass, field
from datetime import datetime
from typing import Dict, List, Optional


@dataclass
class AgendaSlotItem:
    """An individual Agenda Item or topic reference extracted within a Quarter/TU."""
    ai_number: str                 # e.g. '20.6.18', '20.10.1', '5.1'
    display_expr: str              # Original token, e.g. 'KI#18 (AI for Network) (20.6.18) [5+15]'
    topic_label: str               # e.g. 'KI#18 (AI for Network)'
    tdoc_count_str: str = ""       # e.g. '5+15', '7', '18'
    is_drafting: bool = False
    is_tentative: bool = False     # If marked with '(?)'


@dataclass
class SessionSlot:
    """
    A scheduled slot in the meeting grid.
    sub_slot:
      0 = Full Quarter (colspan=2 across TU1 and TU2, 90 mins or full Q0)
      1 = First Half-Quarter / TU 1 (e.g. 09:00 - 09:45)
      2 = Second Half-Quarter / TU 2 (e.g. 09:45 - 10:30)
    """
    id: Optional[int] = None
    snapshot_id: Optional[int] = None
    day_index: int = 0             # 0=Monday, 1=Tuesday, 2=Wednesday, 3=Thursday, 4=Friday
    day_name: str = "Monday"       # 'Monday', 'Tuesday', ...
    day_date: str = ""             # '2026-10-12'
    quarter: str = "Q1"            # 'Q0', 'Q1', 'Q2', 'Lunch', 'Q3', 'Q4', 'Q5'
    sub_slot: int = 0              # 0=Full, 1=TU1, 2=TU2
    start_time_local: str = "09:00"
    end_time_local: str = "10:30"
    stream_number: int = 1         # 1=Main Room, 2=Breakout 1, 3=Breakout 2, 0=All Streams (Lunch/Plenary)
    chair_override: str = ""       # Optional moderator/chair override
    raw_text: str = ""             # Exact text in the cell
    is_drafting: bool = False
    is_tentative: bool = False     # Marked with (?)
    is_single_stream: bool = False # Single stream restriction (other streams marked '-')
    is_plenary: bool = False
    items: List[AgendaSlotItem] = field(default_factory=list)

    @property
    def slot_key(self) -> str:
        """Deterministic unique key for persistent annotations across revisions."""
        return f"{self.day_name}|{self.quarter}|{self.sub_slot}|{self.stream_number}"


@dataclass
class MeetingStreamInfo:
    """Stream-to-room topology and leadership mapping."""
    stream_number: int             # 1, 2, 3
    stream_name: str               # 'Stream 1', 'Stream 2', 'Stream 3'
    room_name: str                 # 'B2 (Floor 2)'
    capacity: int = 0              # 300
    convener: str = ""             # 'SA2 Chair / MCC', 'Mike', 'Wanqiang'
    tohru_room: str = ""           # 'SA2_Main', 'SA2_Brk1', 'SA2_Brk2'


@dataclass
class SessionPlanSnapshot:
    """An imported session plan document revision."""
    id: Optional[int] = None
    meeting_number: str = ""       # e.g. '177' or 'SA2#177'
    tdoc_id: str = ""              # e.g. 'S2-2610195'
    snapshot_label: str = ""       # e.g. 'S2-2610195 (Snap 1 - Mon PM)'
    content_hash: str = ""         # SHA-256 of cell matrix to avoid duplicates
    cached_file_path: str = ""     # Relative/absolute path in Agenda/SessionPlans/
    source_url: str = ""           # Remote download origin if applicable
    imported_at: Optional[datetime] = None
    is_active: bool = True
    streams: Dict[int, MeetingStreamInfo] = field(default_factory=dict)
    slots: List[SessionSlot] = field(default_factory=list)


@dataclass
class UserSessionAnnotation:
    """User notes and status badges attached to slots or AIs."""
    target_type: str               # 'slot' (slot_key) or 'ai' (ai_number)
    target_key: str                # e.g. 'Tuesday|Q1|1|1' or '20.6.18'
    status_badge: str = "normal"   # 'must_attend', 'colleague', 'tohru', 'done'
    assigned_colleague: str = ""   # e.g. 'Martin'
    comment: str = ""
    updated_at: Optional[datetime] = None


# Harmonious, modern, high-contrast delegate palette (10 distinct hues, WCAG AA compliant >= 4.5:1 on white)
DELEGATE_PALETTE: List[str] = [
    "#2563EB",  # Royal Blue
    "#047857",  # Emerald Forest
    "#7C3AED",  # Deep Violet
    "#C2410C",  # Warm Terracotta
    "#0E7490",  # Cyan Ocean
    "#BE123C",  # Ruby Rose
    "#4338CA",  # Indigo Night
    "#A21CAF",  # Plum Orchid
    "#B45309",  # Warm Bronze
    "#475569",  # Slate Graphite
]


@dataclass
class DelegationMember:
    """A colleague/delegate within the delegation."""
    name: str                           # e.g. 'Martin'
    email: str = ""                     # e.g. 'martin@telekom.com'
    color: str = "#2563EB"              # Hex color code from DELEGATE_PALETTE or custom
    is_self: bool = False               # True if this delegate represents the local user
    notes: str = ""                     # Optional role / comments

    @property
    def initials(self) -> str:
        """Returns 2-letter uppercase initials for badge pills."""
        parts = self.name.strip().split()
        if not parts:
            return "??"
        if len(parts) == 1:
            return parts[0][:2].upper()
        return (parts[0][0] + parts[-1][0]).upper()

    @property
    def bg_tint(self) -> str:
        """Returns a soft 12% pastel background tint (#RRGGBB) blended on white."""
        return self.compute_tint(self.color, 0.12)

    @property
    def border_tint(self) -> str:
        """Returns a soft 35% border outline tint (#RRGGBB) blended on white."""
        return self.compute_tint(self.color, 0.35)

    @staticmethod
    def compute_tint(hex_code: str, factor: float = 0.12) -> str:
        """Blends a hex color with #FFFFFF to produce a 100% solid, cross-platform pastel hex."""
        try:
            h = hex_code.lstrip("#")
            if len(h) >= 6:
                r = int(h[0:2], 16)
                g = int(h[2:4], 16)
                b = int(h[4:6], 16)
                r_out = int(r * factor + 255 * (1.0 - factor))
                g_out = int(g * factor + 255 * (1.0 - factor))
                b_out = int(b * factor + 255 * (1.0 - factor))
                return f"#{r_out:02x}{g_out:02x}{b_out:02x}"
        except Exception:
            pass
        return "#F1F5F9"


@dataclass
class TopicOwnership:
    """Designates a default topic owner for an Agenda Item."""
    meeting_number: str                 # e.g. '177'
    ai_number: str                      # e.g. '20.6.18'
    owner_name: str                     # e.g. 'Martin'
    notes: str = ""


@dataclass
class SessionAppointment:
    """A session-specific coverage appointment (resolves collisions or ad-hoc coverage)."""
    meeting_number: str                 # e.g. '177'
    slot_key: str                       # e.g. 'Tuesday|Q2|0|2' (Day|Quarter|SubSlot|Stream)
    appointed_delegate: str             # e.g. 'Thomas'
    original_owner: str = ""            # e.g. 'Martin'
    reason: str = ""                    # e.g. 'Collision in Stream 1'


@dataclass
class DelegationPlan:
    """Encapsulates the complete delegation state for a meeting."""
    meeting_number: str = ""
    members: Dict[str, DelegationMember] = field(default_factory=dict)
    topic_owners: Dict[str, TopicOwnership] = field(default_factory=dict)
    appointments: Dict[str, SessionAppointment] = field(default_factory=dict)

    @property
    def session_appointments(self) -> Dict[str, SessionAppointment]:
        """Backward-compatible alias for appointments dictionary."""
        return self.appointments

