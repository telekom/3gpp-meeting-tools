# --- File: src/modules/meetings/core/session_plan_diff.py ---
"""
Version comparison engine for 3GPP SA2 Session Plans.
Compares two snapshots (e.g. S2-2609251 vs S2-2610195, or intra-day updates)
and detects Added, Removed, Rescheduled/Moved, and TDoc Volume changes.
"""
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Set, Tuple

from modules.meetings.core.session_plan_models import (
    AgendaSlotItem,
    SessionPlanSnapshot,
    SessionSlot,
)


@dataclass
class MovedItemDelta:
    item: AgendaSlotItem
    from_slot_key: str
    from_slot_desc: str
    to_slot_key: str
    to_slot_desc: str


@dataclass
class VolumeChangeDelta:
    item: AgendaSlotItem
    slot_key: str
    slot_desc: str
    old_count: str
    new_count: str


@dataclass
class SessionPlanDiffResult:
    old_snapshot_label: str
    new_snapshot_label: str
    added_slots: List[Tuple[SessionSlot, List[AgendaSlotItem]]] = field(default_factory=list)
    removed_slots: List[Tuple[SessionSlot, List[AgendaSlotItem]]] = field(default_factory=list)
    moved_items: List[MovedItemDelta] = field(default_factory=list)
    volume_changes: List[VolumeChangeDelta] = field(default_factory=list)
    slot_status_map: Dict[str, str] = field(default_factory=dict)  # slot_key -> 'added' | 'removed' | 'moved' | 'modified' | 'unchanged'
    changelog_bullets: List[str] = field(default_factory=list)

    @property
    def has_changes(self) -> bool:
        return bool(
            self.added_slots
            or self.removed_slots
            or self.moved_items
            or self.volume_changes
            or any(s in ("added", "removed", "moved", "modified") for s in self.slot_status_map.values())
        )


class SessionPlanDiffEngine:
    """Computes semantic deltas between two session plan snapshots."""

    @classmethod
    def compare_snapshots(
        cls, snap_a: SessionPlanSnapshot, snap_b: SessionPlanSnapshot
    ) -> SessionPlanDiffResult:
        """
        snap_a: Baseline (older) snapshot
        snap_b: Target (newer) snapshot
        """
        result = SessionPlanDiffResult(
            old_snapshot_label=snap_a.snapshot_label,
            new_snapshot_label=snap_b.snapshot_label,
        )

        slots_a: Dict[str, SessionSlot] = {s.slot_key: s for s in snap_a.slots}
        slots_b: Dict[str, SessionSlot] = {s.slot_key: s for s in snap_b.slots}

        # Track items by identifier to detect moves
        # Map: norm_key -> List[(slot_key, slot, item)]
        items_in_a: Dict[str, List[Tuple[str, SessionSlot, AgendaSlotItem]]] = {}
        for s in snap_a.slots:
            for it in s.items:
                k = cls._normalize_item_key(it)
                if k:
                    items_in_a.setdefault(k, []).append((s.slot_key, s, it))

        items_in_b: Dict[str, List[Tuple[str, SessionSlot, AgendaSlotItem]]] = {}
        for s in snap_b.slots:
            for it in s.items:
                k = cls._normalize_item_key(it)
                if k:
                    items_in_b.setdefault(k, []).append((s.slot_key, s, it))

        all_slot_keys = set(slots_a.keys()) | set(slots_b.keys())

        matched_moves: Set[Tuple[str, str, str]] = set()  # (norm_key, from_key, to_key)

        # Detect moves across the grid
        for norm_k, b_entries in items_in_b.items():
            if norm_k in items_in_a:
                a_slots = {e[0] for e in items_in_a[norm_k]}
                b_slots = {e[0] for e in b_entries}

                # Slots in B but not in A
                new_in_b = b_slots - a_slots
                old_in_a = a_slots - b_slots

                for to_slot_key in new_in_b:
                    if old_in_a:
                        from_slot_key = old_in_a.pop()
                        slot_from = slots_a[from_slot_key]
                        slot_to = slots_b[to_slot_key]
                        item_obj = next(e[2] for e in b_entries if e[0] == to_slot_key)
                        move_delta = MovedItemDelta(
                            item=item_obj,
                            from_slot_key=from_slot_key,
                            from_slot_desc=f"{slot_from.day_name} {slot_from.quarter} Stream {slot_from.stream_number}",
                            to_slot_key=to_slot_key,
                            to_slot_desc=f"{slot_to.day_name} {slot_to.quarter} Stream {slot_to.stream_number}",
                        )
                        result.moved_items.append(move_delta)
                        matched_moves.add((norm_k, from_slot_key, to_slot_key))

        # Slot-by-slot comparison
        for sk in all_slot_keys:
            s_a = slots_a.get(sk)
            s_b = slots_b.get(sk)

            if s_a and not s_b:
                result.removed_slots.append((s_a, s_a.items))
                result.slot_status_map[sk] = "removed"
                continue

            if not s_a and s_b:
                result.added_slots.append((s_b, s_b.items))
                result.slot_status_map[sk] = "added"
                continue

            # Both exist: check content differences
            raw_a = s_a.raw_text.strip()
            raw_b = s_b.raw_text.strip()

            if raw_a == raw_b:
                result.slot_status_map[sk] = "unchanged"
                continue

            # Check if this slot was involved in a move
            is_move_target = any(m.to_slot_key == sk for m in result.moved_items)
            is_move_source = any(m.from_slot_key == sk for m in result.moved_items)

            # Check for volume changes
            items_a_map = {cls._normalize_item_key(it): it for it in s_a.items}
            items_b_map = {cls._normalize_item_key(it): it for it in s_b.items}

            slot_vol_changed = False
            for k, it_b in items_b_map.items():
                if k in items_a_map:
                    it_a = items_a_map[k]
                    if it_a.tdoc_count_str != it_b.tdoc_count_str and (it_a.tdoc_count_str or it_b.tdoc_count_str):
                        result.volume_changes.append(
                            VolumeChangeDelta(
                                item=it_b,
                                slot_key=sk,
                                slot_desc=f"{s_b.day_name} {s_b.quarter} Stream {s_b.stream_number}",
                                old_count=it_a.tdoc_count_str or "0",
                                new_count=it_b.tdoc_count_str or "0",
                            )
                        )
                        slot_vol_changed = True

            if is_move_target or is_move_source:
                result.slot_status_map[sk] = "moved"
            elif slot_vol_changed:
                result.slot_status_map[sk] = "modified"
            else:
                # Was previous slot a generic 'Revisions'?
                if "revisions" in raw_a.lower() and "revisions" not in raw_b.lower():
                    result.slot_status_map[sk] = "added"
                elif "revisions" not in raw_a.lower() and "revisions" in raw_b.lower():
                    result.slot_status_map[sk] = "removed"
                else:
                    result.slot_status_map[sk] = "modified"

        # Generate human-readable changelog bullets
        cls._generate_changelog_bullets(result, slots_a, slots_b)

        return result

    @classmethod
    def _normalize_item_key(cls, item: AgendaSlotItem) -> str:
        """Normalized identity key for matching an item across slots."""
        if item.ai_number:
            return f"ai:{item.ai_number.lower()}"
        if item.topic_label:
            clean = "".join(c for c in item.topic_label.lower() if c.isalnum())
            if len(clean) > 3:
                return f"topic:{clean}"
        return ""

    @classmethod
    def _generate_changelog_bullets(
        cls,
        result: SessionPlanDiffResult,
        slots_a: Dict[str, SessionSlot],
        slots_b: Dict[str, SessionSlot],
    ):
        """Constructs concise bulleted changelog entries for UI and email dispatch."""
        bullets = []

        # 1. Moved items
        for m in result.moved_items:
            name = m.item.topic_label or f"AI {m.item.ai_number}"
            bullets.append(f"↔ Rescheduled: '{name}' moved from {m.from_slot_desc} ➔ {m.to_slot_desc}")

        # 2. Volume changes
        for v in result.volume_changes:
            name = v.item.topic_label or f"AI {v.item.ai_number}"
            bullets.append(f"Δ Volume: '{name}' in {v.slot_desc} updated from [{v.old_count}] to [{v.new_count}]")

        # 3. Newly added drafting sessions
        for sk, status in result.slot_status_map.items():
            if status == "added":
                s_b = slots_b.get(sk)
                if s_b:
                    if s_b.is_drafting:
                        bullets.append(f"+ Added Drafting: {s_b.raw_text[:60]} ({s_b.day_name} {s_b.quarter} Stream {s_b.stream_number})")
                    elif s_b.raw_text.strip() and s_b.raw_text.strip() != "-":
                        bullets.append(f"+ Scheduled: {s_b.raw_text[:60]} ({s_b.day_name} {s_b.quarter} Stream {s_b.stream_number})")

        # 4. Removed or postponed sessions
        for sk, status in result.slot_status_map.items():
            if status == "removed":
                s_a = slots_a.get(sk)
                if s_a and s_a.raw_text.strip() and s_a.raw_text.strip() != "-":
                    bullets.append(f"- Removed / Postponed: {s_a.raw_text[:60]} ({s_a.day_name} {s_a.quarter} Stream {s_a.stream_number})")

        result.changelog_bullets = bullets
