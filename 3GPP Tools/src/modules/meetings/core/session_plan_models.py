# --- File: src/modules/meetings/core/session_plan_models.py ---
"""
Domain models and dataclasses for 3GPP SA2 Meeting Session Plans.
Supports Half-Quarter Sub-Slots (Time Units), Stream/Room topologies,
AI Range Expansions, and User Annotations.
"""
import re
from dataclasses import dataclass, field
from datetime import datetime
from typing import Dict, Iterable, List, Optional, Tuple


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
    owner_name: str                     # e.g. 'Martin', 'Martin, Anna', or 'ALL'
    notes: str = ""

    def is_all(self) -> bool:
        """Returns True if this topic is assigned to the entire delegation."""
        if not self.owner_name:
            return False
        clean = self.owner_name.strip().upper()
        return clean in ("ALL", "ALL TEAM", "ALL DELEGATES", "👥 ALL TEAM", "👥 ALL")

    def get_owners(self, all_roster_names: Optional[List[str]] = None) -> List[str]:
        """Returns list of assigned delegate names, expanding 'ALL' if roster names are provided."""
        if not self.owner_name:
            return []
        if self.is_all():
            return list(all_roster_names) if all_roster_names else ["ALL"]
        return [p.strip() for p in self.owner_name.split(",") if p.strip()]


@dataclass
class SessionAppointment:
    """A session-specific coverage appointment (resolves collisions or ad-hoc coverage)."""
    meeting_number: str                 # e.g. '177'
    slot_key: str                       # e.g. 'Tuesday|Q2|0|2' (Day|Quarter|SubSlot|Stream)
    appointed_delegate: str             # e.g. 'Thomas', 'Thomas, Anna', or 'ALL'
    original_owner: str = ""            # e.g. 'Martin'
    reason: str = ""                    # e.g. 'Collision in Stream 1'

    def is_all(self) -> bool:
        """Returns True if this session is appointed to the entire delegation."""
        if not self.appointed_delegate:
            return False
        clean = self.appointed_delegate.strip().upper()
        return clean in ("ALL", "ALL TEAM", "ALL DELEGATES", "👥 ALL TEAM", "👥 ALL")

    def get_appointed_delegates(self, all_roster_names: Optional[List[str]] = None) -> List[str]:
        """Returns list of appointed delegate names, expanding 'ALL' if roster names are provided."""
        if not self.appointed_delegate:
            return []
        if self.is_all():
            return list(all_roster_names) if all_roster_names else ["ALL"]
        return [p.strip() for p in self.appointed_delegate.split(",") if p.strip()]


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


# ---------------------------------------------------------------------------
# Shared delegate resolution helpers (single source of truth for UI + exporters)
# ---------------------------------------------------------------------------

def ai_number_in_text(ai_number: str, text: str) -> bool:
    """
    Boundary-aware Agenda Item lookup inside free text.

    A naive substring test (``"20.6.2" in "KI#24 (IoT) (20.6.24)"``) yields false
    positives for every AI sharing a numeric prefix (20.6.2 -> 20.6.20 … 20.6.29,
    20.6.1 -> 20.6.10 … 20.6.19). The AI must therefore not be preceded by a digit
    or dot, and not be followed by another digit or a deeper ``.N`` level.
    """
    if not ai_number or not text:
        return False
    pattern = r"(?<![\d.])" + re.escape(ai_number.strip()) + r"(?!\d|\.\d)"
    return re.search(pattern, text) is not None


def resolve_slot_effective_delegates(
    slot: SessionSlot, plan: Optional[DelegationPlan]
) -> Tuple[List[str], bool, bool]:
    """
    Returns ``(delegates, is_appointment_override, is_all)`` for a session slot.

    Priority:
    1. Explicit session appointment (slot-level override).
    2. Topic ownership of the AI numbers extracted by the parser (``slot.items``).
    3. Fallback: boundary-aware match of owned AI numbers in the raw cell text
       (covers cells the parser could not fully tokenise).
    """
    if not plan:
        return [], False, False

    roster_names = list(plan.members.keys())

    # 1. Slot override appointment
    app = plan.appointments.get(slot.slot_key)
    if app:
        return app.get_appointed_delegates(roster_names), True, app.is_all()

    found: List[str] = []
    is_all_flag = False

    def _absorb(to: TopicOwnership) -> None:
        nonlocal is_all_flag
        if to.is_all():
            is_all_flag = True
            owners: Iterable[str] = roster_names
        else:
            owners = to.get_owners(roster_names)
        for owner in owners:
            if owner not in found:
                found.append(owner)

    # 2. Topic-level ownership from parsed slot items
    for it in slot.items:
        if it.ai_number and it.ai_number in plan.topic_owners:
            _absorb(plan.topic_owners[it.ai_number])
    if found:
        return found, False, is_all_flag

    # 3. Boundary-aware raw text fallback
    for ai_num, to in plan.topic_owners.items():
        if ai_number_in_text(ai_num, slot.raw_text):
            _absorb(to)

    return found, False, is_all_flag


def compute_unique_initials(names: Iterable[str]) -> Dict[str, str]:
    """
    Computes short, roster-unique badge initials (e.g. Josep -> 'JO', Dieter -> 'DI').

    Starts from the 2-letter ``DelegationMember.initials`` convention and, if two
    delegates collide (e.g. 'Josep' and 'Jordi'), progressively disambiguates them
    with additional letters ('JOS' / 'JOR'), capped at 3 characters.
    """
    name_list = [n for n in names if n and n.strip()]
    result: Dict[str, str] = {n: DelegationMember(name=n).initials for n in name_list}

    # Disambiguate colliding 2-letter initials with a third letter
    groups: Dict[str, List[str]] = {}
    for n, ini in result.items():
        groups.setdefault(ini, []).append(n)
    for dupes in groups.values():
        if len(dupes) < 2:
            continue
        for n in dupes:
            parts = n.strip().split()
            if len(parts) > 1:
                cand = (parts[0][:2] + parts[-1][0]).upper()
            else:
                cand = re.sub(r"[^A-Za-z0-9]", "", n)[:3].upper()
            result[n] = cand or result[n]

    # Final guard: still-duplicated initials get a numeric suffix
    seen: Dict[str, int] = {}
    for n in name_list:
        ini = result[n]
        if ini in seen:
            seen[ini] += 1
            result[n] = f"{ini[:2]}{seen[ini]}"
        else:
            seen[ini] = 1
    return result

