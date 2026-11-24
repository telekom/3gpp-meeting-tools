"""Agenda Item expansion, boundary-aware matching, and delegate resolution helpers."""

from __future__ import annotations

import re
from typing import Any, Dict, Iterable, List, Optional, Tuple

from modules.meetings.core.session_plan.models.dataclasses import (
    DelegationMember,
    DelegationPlan,
    SessionSlot,
    TopicOwnership,
)

_RANGE_DASH_RE = re.compile(r"[\u2010\u2011\u2012\u2013\u2014\u2015\u2212\-]")
_AI_RANGE_TOKEN_RE = re.compile(
    r"(\d{1,2}(?:\.\d+)+)\s*[\u2010-\u2015\u2212\-]\s*(?:(\d{1,2}(?:\.\d+)*)\.)?(\d+)\b"
)


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


def expand_numeric_ai_range(range_str: str) -> List[str]:
    """
    Expands shorthand or full AI ranges (supporting ASCII hyphen and Unicode dashes):
    - '20.6.12 - 16' / '20.6.12 – 16' -> ['20.6.12', '20.6.13', '20.6.14', '20.6.15', '20.6.16']
    - '20.6.12 – 20.6.16' -> ['20.6.12', '20.6.13', '20.6.14', '20.6.15', '20.6.16']
    """
    clean = _RANGE_DASH_RE.sub("-", (range_str or "").strip())
    m = re.match(r"^(\d+(?:\.\d+)*)\s*-\s*(?:(\d+(?:\.\d+)*)\.)?(\d+)$", clean)
    if not m:
        return [range_str.strip()] if (range_str or "").strip() else []
    prefix_part, right_base, end_num_str = m.group(1), m.group(2), m.group(3)
    parts = prefix_part.split(".")
    if len(parts) < 2:
        return [range_str.strip()]

    start_num = int(parts[-1])
    end_num = int(end_num_str)
    base = ".".join(parts[:-1])
    if right_base and right_base != base:
        return [range_str.strip()]

    if start_num <= end_num and end_num - start_num <= 50:
        return [f"{base}.{i}" for i in range(start_num, end_num + 1)]
    return [range_str.strip()]


def expand_ai_number(ai_num: str, agenda_map: Optional[Dict[str, Any]] = None) -> List[str]:
    """
    Expands a range (e.g. ``20.6.12 – 16``), ``.x`` wildcard (e.g. ``30.x``, ``20.6.x``),
    or unlisted parent prefix (e.g. ``20.6`` when only ``20.6.*`` exist in ``agenda_map``)
    into concrete AI keys.

    Prefers strict child AIs under ``<prefix>.`` (ignoring the bare chapter/section
    header row ``<prefix>`` when child items exist); falls back to ``<prefix>`` if
    it is the only matching entry in ``agenda_map``, or ``[ai_num]`` when
    ``agenda_map`` is unavailable or contains no match.
    """
    clean = (ai_num or "").strip()
    if not clean:
        return []
    if _RANGE_DASH_RE.search(clean) and re.search(r"\d\s*[\u2010-\u2015\u2212\-]\s*\d", clean):
        return expand_numeric_ai_range(clean)
    if not agenda_map:
        return [clean]
    if clean.lower().endswith(".x"):
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

    clean_low = clean.lower()
    if any(str(k).strip().lower() == clean_low for k in agenda_map):
        return [clean]
    children = [k for k in agenda_map if str(k).strip().lower().startswith(clean_low + ".")]
    if children:
        return sorted(children, key=ai_sort_key)
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

    # Concrete AI (e.g. '20.6.14') matches numeric AI ranges in text (e.g. '20.6.12 – 16')
    for m_rng in _AI_RANGE_TOKEN_RE.finditer(text):
        for expanded_ai in expand_numeric_ai_range(m_rng.group(0)):
            if expanded_ai.strip().lower() == ai_clean.lower():
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
