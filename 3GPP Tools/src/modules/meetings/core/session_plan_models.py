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
