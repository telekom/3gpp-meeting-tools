# --- File: src/modules/meetings/core/session_plan_models.py ---
"""
Domain models and dataclasses for 3GPP SA2 Meeting Session Plans.
Supports Half-Quarter Sub-Slots (Time Units), Stream/Room topologies,
AI Range Expansions, and User Annotations.
"""
import html
import re
from dataclasses import dataclass, field
from datetime import datetime
from typing import Any, Dict, Iterable, List, Optional, Tuple


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
    room_override: str = ""        # Optional room override, e.g. 'Breakout 1', 'Main Room'
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


NOT_FOLLOWED_TOKENS = (
    "NOT_FOLLOWED",
    "NOT FOLLOWED",
    "NOT_FOLLOW",
    "NOT FOLLOW",
    "⛔ NOT FOLLOWED",
    "⛔ NOT_FOLLOWED",
    "⚪ NOT FOLLOWED",
    "EXCLUDE",
    "IGNORED",
)


@dataclass
class TopicOwnership:
    """Designates a default topic owner for an Agenda Item."""
    meeting_number: str                 # e.g. '177'
    ai_number: str                      # e.g. '20.6.18'
    owner_name: str                     # e.g. 'Martin', 'Martin, Anna', 'ALL', or 'NOT_FOLLOWED'
    notes: str = ""

    def is_not_followed(self) -> bool:
        """Returns True if this topic is explicitly marked as not followed (out of scope)."""
        if not self.owner_name:
            return False
        clean = self.owner_name.strip().upper()
        return clean in NOT_FOLLOWED_TOKENS

    def is_all(self) -> bool:
        """Returns True if this topic is assigned to the entire delegation."""
        if not self.owner_name or self.is_not_followed():
            return False
        clean = self.owner_name.strip().upper()
        return clean in ("ALL", "ALL TEAM", "ALL DELEGATES", "👥 ALL TEAM", "👥 ALL")

    def get_owners(self, all_roster_names: Optional[List[str]] = None) -> List[str]:
        """Returns list of assigned delegate names, expanding 'ALL' if roster names are provided."""
        if not self.owner_name or self.is_not_followed():
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

def ai_numbers_match(ai_a: str, ai_b: str) -> bool:
    """
    Checks whether two Agenda Item numbers match, supporting ``.x`` wildcards
    (e.g. ``30.x`` matches ``30.x``, ``30``, ``30.1``, ``30.2``, and ``20.6.x`` matches
    ``20.6`` and ``20.6.18``, without matching ``300.1`` or ``3.1``).
    """
    a = (ai_a or "").strip().lower()
    b = (ai_b or "").strip().lower()
    if not a or not b:
        return False
    if a == b:
        return True

    def _wild_matches(wild: str, concrete: str) -> bool:
        if not wild.endswith(".x"):
            return False
        prefix = wild[:-2]
        if not prefix:
            return False
        return concrete == prefix or concrete.startswith(prefix + ".")

    return _wild_matches(a, b) or _wild_matches(b, a)


def ai_sort_key(ai_str: str) -> List[Any]:
    """Natural sort key for dotted Agenda Item numbers (e.g. '30.2' < '30.10')."""
    return [int(p) if p.isdigit() else p.lower() for p in re.split(r"(\d+)", str(ai_str or ""))]


def expand_ai_number(ai_num: str, agenda_map: Optional[Dict[str, Any]] = None) -> List[str]:
    """
    Expands a ``.x`` wildcard AI number (e.g. ``30.x``, ``20.6.x``) into concrete
    AI keys from ``agenda_map``.

    Prefers strict child AIs under ``<prefix>.`` (ignoring the bare chapter/section
    header row ``<prefix>`` when child items exist); falls back to ``<prefix>`` if
    it is the only matching entry in ``agenda_map``, or ``[ai_num]`` when
    ``agenda_map`` is unavailable or contains no match.
    """
    clean = (ai_num or "").strip()
    if not clean:
        return []
    if not clean.lower().endswith(".x") or not agenda_map:
        return [clean]
    prefix = clean[:-2].strip()
    if not prefix:
        return [clean]
    prefix_dot = prefix.lower() + "."
    children = [k for k in agenda_map if str(k).strip().lower().startswith(prefix_dot)]
    if children:
        return sorted(children, key=ai_sort_key)
    prefix_low = prefix.lower()
    parents = [k for k in agenda_map if str(k).strip().lower() == prefix_low]
    if parents:
        return parents
    return [clean]


def expand_slot_ai_numbers(slot: SessionSlot, agenda_map: Optional[Dict[str, Any]] = None) -> List[str]:
    """Returns deduplicated AI numbers for a slot, expanding ``.x`` wildcards against ``agenda_map``."""
    expanded: List[str] = []
    seen = set()
    for it in slot.items:
        if not it.ai_number:
            continue
        for sub_ai in expand_ai_number(it.ai_number, agenda_map):
            key = sub_ai.strip().lower()
            if key and key not in seen:
                seen.add(key)
                expanded.append(sub_ai)
    return expanded


def _resolve_single_ai_ownerships(
    ai_num: str, topic_owners: Dict[str, TopicOwnership], include_parent_prefix: Optional[str] = None
) -> List[TopicOwnership]:
    """Looks up exact, case-insensitive, wildcard, or explicit parent ownership for one concrete AI."""
    if not ai_num or not topic_owners:
        return []
    if ai_num in topic_owners:
        return [topic_owners[ai_num]]
    ai_low = ai_num.strip().lower()
    for k, to in topic_owners.items():
        if k.strip().lower() == ai_low:
            return [to]
    wild_matches = [to for k, to in topic_owners.items() if ai_numbers_match(ai_num, k)]
    if wild_matches:
        return wild_matches
    if include_parent_prefix:
        p_low = include_parent_prefix.strip().lower()
        for k, to in topic_owners.items():
            if k.strip().lower() == p_low:
                return [to]
    return []


def _matching_topic_ownerships(
    ai_num: str,
    topic_owners: Dict[str, TopicOwnership],
    agenda_map: Optional[Dict[str, Any]] = None,
) -> List[TopicOwnership]:
    """
    Finds TopicOwnership entries matching ``ai_num``.
    When ``agenda_map`` is supplied and ``ai_num`` is a ``.x`` wildcard, expands it to
    its concrete child AIs in ``agenda_map`` so per-sub-AI ownerships (and unassigned
    gaps) are resolved accurately instead of stopping at a single partial match.
    """
    if not ai_num or not topic_owners:
        return []
    clean = ai_num.strip()
    if agenda_map and clean.lower().endswith(".x"):
        expanded = expand_ai_number(clean, agenda_map)
        if expanded and not (len(expanded) == 1 and expanded[0].strip().lower() == clean.lower()):
            parent_prefix = clean[:-2].strip()
            results: List[TopicOwnership] = []
            seen_keys = set()
            for sub_ai in expanded:
                for to in _resolve_single_ai_ownerships(sub_ai, topic_owners, include_parent_prefix=parent_prefix):
                    k = to.ai_number.strip().lower()
                    if k not in seen_keys:
                        seen_keys.add(k)
                        results.append(to)
            return results
    return _resolve_single_ai_ownerships(clean, topic_owners)


def ai_number_in_text(ai_number: str, text: str) -> bool:
    """
    Boundary-aware Agenda Item lookup inside free text, supporting ``.x`` wildcards.

    A naive substring test (``"20.6.2" in "KI#24 (IoT) (20.6.24)"``) yields false
    positives for every AI sharing a numeric prefix (20.6.2 -> 20.6.20 … 20.6.29,
    20.6.1 -> 20.6.10 … 20.6.19). The AI must therefore not be preceded by a digit
    or dot, and not be followed by another digit or a deeper ``.N`` level.
    """
    if not ai_number or not text:
        return False
    ai_clean = ai_number.strip()
    if ai_clean.lower().endswith(".x"):
        prefix = ai_clean[:-2]
        if prefix:
            w_pat = r"(?<![\d.])" + re.escape(prefix) + r"(?:\.(?:\d+|[xX]))?(?![\d.])"
            return re.search(w_pat, text) is not None

    pattern = r"(?<![\d.])" + re.escape(ai_clean) + r"(?!\d|\.\d)"
    if re.search(pattern, text, re.I if ai_clean.lower().endswith(".x") else 0) is not None:
        return True

    # Concrete AI (e.g. '30.1' or '30' or '20.6.18') matches explicit '<prefix>.x' in text
    parent = ai_clean.rsplit(".", 1)[0] if "." in ai_clean else ai_clean
    if parent:
        wild_in_text = r"(?<![\d.])" + re.escape(parent) + r"\.[xX](?![\d\w])"
        if re.search(wild_in_text, text) is not None:
            return True
    return False


def resolve_slot_effective_delegates(
    slot: SessionSlot,
    plan: Optional[DelegationPlan],
    agenda_map: Optional[Dict[str, Any]] = None,
) -> Tuple[List[str], bool, bool]:
    """
    Returns ``(delegates, is_appointment_override, is_all)`` for a session slot.

    Priority:
    1. Explicit session appointment (slot-level override).
    2. Topic ownership of the AI numbers extracted by the parser (``slot.items``),
       expanding ``.x`` wildcards against ``agenda_map`` when available.
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
        if to.is_not_followed():
            return
        if to.is_all():
            is_all_flag = True
            owners: Iterable[str] = roster_names
        else:
            owners = to.get_owners(roster_names)
        for owner in owners:
            if owner not in found:
                found.append(owner)

    # 2. Topic-level ownership from parsed slot items (including .x wildcards)
    for it in slot.items:
        if it.ai_number:
            for to in _matching_topic_ownerships(it.ai_number, plan.topic_owners, agenda_map=agenda_map):
                _absorb(to)
    if found:
        return found, False, is_all_flag

    # 3. Boundary-aware raw text fallback
    for ai_num, to in plan.topic_owners.items():
        if ai_number_in_text(ai_num, slot.raw_text):
            _absorb(to)

    return found, False, is_all_flag


def get_slot_unassigned_ais(
    slot: SessionSlot,
    plan: Optional[DelegationPlan],
    agenda_map: Optional[Dict[str, Any]] = None,
) -> List[str]:
    """
    Returns concrete AI numbers in ``slot`` that have neither an owner nor
    ``NOT_FOLLOWED`` status (expanding ``.x`` wildcards against ``agenda_map``),
    unless the slot itself has an active session-level appointment override.
    """
    if plan and slot.slot_key in plan.appointments:
        app = plan.appointments[slot.slot_key]
        if (app.appointed_delegate or "").strip():
            return []

    topic_owners = plan.topic_owners if plan else {}
    unassigned: List[str] = []
    for it in slot.items:
        if not it.ai_number:
            continue
        is_wild = it.ai_number.strip().lower().endswith(".x")
        parent_prefix = it.ai_number.strip()[:-2].strip() if is_wild else None
        for sub_ai in expand_ai_number(it.ai_number, agenda_map):
            if agenda_map and is_wild and sub_ai.strip().lower() != it.ai_number.strip().lower():
                matches = _resolve_single_ai_ownerships(
                    sub_ai, topic_owners, include_parent_prefix=parent_prefix
                )
            else:
                matches = _matching_topic_ownerships(sub_ai, topic_owners, agenda_map=agenda_map)
            if any(to.is_not_followed() or to.owner_name.strip() for to in matches):
                continue
            if sub_ai not in unassigned:
                unassigned.append(sub_ai)
    return unassigned


def is_slot_not_followed(
    slot: SessionSlot,
    plan: Optional[DelegationPlan],
    agenda_map: Optional[Dict[str, Any]] = None,
) -> bool:
    """
    Returns True if this session slot is explicitly marked as Not Followed
    (i.e. has no appointment override and all its expanded agenda items are Not Followed).
    """
    if not plan or not plan.topic_owners:
        return False
    # Explicit slot-level session appointment override takes precedence
    if slot.slot_key in plan.appointments:
        return False
    raw_ai_nums = [it.ai_number for it in slot.items if it.ai_number]
    if raw_ai_nums:
        for raw_ai in raw_ai_nums:
            is_wild = raw_ai.strip().lower().endswith(".x")
            parent_prefix = raw_ai.strip()[:-2].strip() if is_wild else None
            for sub_ai in expand_ai_number(raw_ai, agenda_map):
                if agenda_map and is_wild and sub_ai.strip().lower() != raw_ai.strip().lower():
                    matches = _resolve_single_ai_ownerships(
                        sub_ai, plan.topic_owners, include_parent_prefix=parent_prefix
                    )
                else:
                    matches = _matching_topic_ownerships(sub_ai, plan.topic_owners, agenda_map=agenda_map)
                if not matches or not all(m.is_not_followed() for m in matches):
                    return False
        return True
    # Fallback to checking raw cell text for boundary-matched not-followed AIs
    for ai_num, to in plan.topic_owners.items():
        if to.is_not_followed() and ai_number_in_text(ai_num, slot.raw_text):
            return True
    return False


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


_ROOM_PREFIX_RE = re.compile(
    r"^\s*(?:\(?\s*in\s+(?:the\s+)?)?(main room|breakout\s*1|breakout\s*2|bmo room \d+|palomino [a-z\-]+|percheron [a-z\-]+)\s*\)?[\s:.\-–—]*",
    re.I,
)
_TRAILING_COUNT_RE = re.compile(r"^\[\s*[\d\+\,\s]+\s*\]$")
_EMPTY_DRAFTING_RE = re.compile(r"^\s*drafting\s*:?\s*$", re.I)


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


