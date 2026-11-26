# --- File: src/modules/meetings/core/session_plan/db/mixins/snapshots_and_annotations.py ---
"""Snapshots CRUD, auto-repair of cached agenda slot items, annotations, and tracked topics."""

import logging
from typing import Dict, List, Optional, Tuple

from modules.meetings.core.session_plan.session_plan_models import (
    AgendaSlotItem,
    MeetingStreamInfo,
    SessionPlanSnapshot,
    SessionSlot,
    UserSessionAnnotation,
)

logger = logging.getLogger("modules.meetings.core.session_plan.session_plan_db")


class SessionPlanSnapshotsAnnotationsMixin:
    """Manages snapshots, self-healing slot items, user annotations, and tracked topics."""

    def save_snapshot(self, snapshot: SessionPlanSnapshot) -> int:
        """
        Saves a session plan snapshot.
        If a snapshot with identical content_hash exists, updates metadata and marks it active.
        Otherwise creates a new snapshot and sets it active.
        """
        with self._get_connection() as conn:
            cursor = conn.cursor()

            # Check if this exact content hash is already saved
            existing = cursor.execute(
                "SELECT id FROM session_plan_snapshots WHERE meeting_number = ? AND content_hash = ?",
                (snapshot.meeting_number, snapshot.content_hash),
            ).fetchone()

            if existing:
                snap_id = existing["id"]
                # Deactivate all, activate this one
                cursor.execute(
                    "UPDATE session_plan_snapshots SET is_active = 0 WHERE meeting_number = ?",
                    (snapshot.meeting_number,),
                )
                cursor.execute(
                    "UPDATE session_plan_snapshots SET is_active = 1, cached_file_path = COALESCE(?, cached_file_path) WHERE id = ?",
                    (snapshot.cached_file_path, snap_id),
                )
                logger.info(f"[SessionPlanDB] Snapshot with hash {snapshot.content_hash[:8]} already exists as ID {snap_id}. Activated.")
                return snap_id

            # Deactivate older snapshots
            cursor.execute(
                "UPDATE session_plan_snapshots SET is_active = 0 WHERE meeting_number = ?",
                (snapshot.meeting_number,),
            )

            # Insert new snapshot
            cursor.execute(
                """
                INSERT INTO session_plan_snapshots (meeting_number, tdoc_id, snapshot_label, content_hash, cached_file_path, source_url, is_active)
                VALUES (?, ?, ?, ?, ?, ?, 1)
                """,
                (
                    snapshot.meeting_number,
                    snapshot.tdoc_id,
                    snapshot.snapshot_label,
                    snapshot.content_hash,
                    snapshot.cached_file_path,
                    snapshot.source_url,
                ),
            )
            snap_id = cursor.lastrowid

            # Insert streams
            for _s_num, s_info in snapshot.streams.items():
                cursor.execute(
                    """
                    INSERT INTO meeting_streams (snapshot_id, stream_number, stream_name, room_name, capacity, convener, tohru_room)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        snap_id,
                        s_info.stream_number,
                        s_info.stream_name,
                        s_info.room_name,
                        s_info.capacity,
                        s_info.convener,
                        s_info.tohru_room,
                    ),
                )

            # Insert slots and their items
            for slot in snapshot.slots:
                cursor.execute(
                    """
                    INSERT INTO session_slots (snapshot_id, day_index, day_name, day_date, quarter, sub_slot,
                                               start_time_local, end_time_local, stream_number, chair_override,
                                               room_override, raw_text, is_drafting, is_tentative, is_single_stream, is_plenary)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        snap_id,
                        slot.day_index,
                        slot.day_name,
                        slot.day_date,
                        slot.quarter,
                        slot.sub_slot,
                        slot.start_time_local,
                        slot.end_time_local,
                        slot.stream_number,
                        slot.chair_override,
                        slot.room_override,
                        slot.raw_text,
                        1 if slot.is_drafting else 0,
                        1 if slot.is_tentative else 0,
                        1 if slot.is_single_stream else 0,
                        1 if slot.is_plenary else 0,
                    ),
                )
                slot_id = cursor.lastrowid

                for item in slot.items:
                    cursor.execute(
                        """
                        INSERT INTO slot_agenda_items (slot_id, ai_number, display_expr, topic_label, tdoc_count_str, is_drafting)
                        VALUES (?, ?, ?, ?, ?, ?)
                        """,
                        (
                            slot_id,
                            item.ai_number,
                            item.display_expr,
                            item.topic_label,
                            item.tdoc_count_str,
                            1 if item.is_drafting else 0,
                        ),
                    )

            logger.info(f"[SessionPlanDB] Saved new snapshot ID {snap_id} ('{snapshot.snapshot_label}') with {len(snapshot.slots)} slots.")
            return snap_id

    def list_snapshots(self, meeting_number: str) -> List[Dict]:
        with self._get_connection() as conn:
            rows = conn.execute(
                """
                SELECT id, meeting_number, tdoc_id, snapshot_label, content_hash, cached_file_path, source_url, imported_at, is_active
                FROM session_plan_snapshots
                WHERE meeting_number = ?
                ORDER BY id DESC
                """,
                (str(meeting_number),),
            ).fetchall()
            return [dict(r) for r in rows]

    def set_active_snapshot(self, snapshot_id: int):
        with self._get_connection() as conn:
            row = conn.execute("SELECT meeting_number FROM session_plan_snapshots WHERE id = ?", (snapshot_id,)).fetchone()
            if row:
                conn.execute(
                    "UPDATE session_plan_snapshots SET is_active = 0 WHERE meeting_number = ?",
                    (row["meeting_number"],),
                )
                conn.execute("UPDATE session_plan_snapshots SET is_active = 1 WHERE id = ?", (snapshot_id,))

    def get_active_snapshot(self, meeting_number: str) -> Optional[SessionPlanSnapshot]:
        with self._get_connection() as conn:
            row = conn.execute(
                "SELECT id FROM session_plan_snapshots WHERE meeting_number = ? AND is_active = 1 LIMIT 1",
                (str(meeting_number),),
            ).fetchone()
            if not row:
                # Fallback to latest imported
                row = conn.execute(
                    "SELECT id FROM session_plan_snapshots WHERE meeting_number = ? ORDER BY id DESC LIMIT 1",
                    (str(meeting_number),),
                ).fetchone()

            if not row:
                return None
            return self.get_snapshot_by_id(row["id"])

    def get_snapshot_by_id(self, snapshot_id: int) -> Optional[SessionPlanSnapshot]:
        with self._get_connection() as conn:
            snap_row = conn.execute(
                "SELECT * FROM session_plan_snapshots WHERE id = ?", (snapshot_id,)
            ).fetchone()
            if not snap_row:
                return None

            snapshot = SessionPlanSnapshot(
                id=snap_row["id"],
                meeting_number=snap_row["meeting_number"],
                tdoc_id=snap_row["tdoc_id"],
                snapshot_label=snap_row["snapshot_label"],
                content_hash=snap_row["content_hash"],
                cached_file_path=snap_row["cached_file_path"] or "",
                source_url=snap_row["source_url"] or "",
                is_active=bool(snap_row["is_active"]),
            )

            # Streams
            stream_rows = conn.execute(
                "SELECT * FROM meeting_streams WHERE snapshot_id = ? ORDER BY stream_number ASC",
                (snapshot_id,),
            ).fetchall()
            for sr in stream_rows:
                snapshot.streams[sr["stream_number"]] = MeetingStreamInfo(
                    stream_number=sr["stream_number"],
                    stream_name=sr["stream_name"],
                    room_name=sr["room_name"],
                    capacity=sr["capacity"] or 0,
                    convener=sr["convener"] or "",
                    tohru_room=sr["tohru_room"] or "",
                )

            # Slots and items
            slot_rows = conn.execute(
                "SELECT * FROM session_slots WHERE snapshot_id = ? ORDER BY day_index ASC, quarter ASC, sub_slot ASC, stream_number ASC",
                (snapshot_id,),
            ).fetchall()

            slot_ids = [sr["id"] for sr in slot_rows]
            items_by_slot: Dict[int, List[AgendaSlotItem]] = {sid: [] for sid in slot_ids}

            if slot_ids:
                # Chunk in 500 if large
                item_rows = conn.execute(
                    f"SELECT * FROM slot_agenda_items WHERE slot_id IN ({','.join(['?']*len(slot_ids))})",
                    slot_ids,
                ).fetchall()
                for ir in item_rows:
                    items_by_slot[ir["slot_id"]].append(
                        AgendaSlotItem(
                            ai_number=ir["ai_number"],
                            display_expr=ir["display_expr"] or "",
                            topic_label=ir["topic_label"] or "",
                            tdoc_count_str=ir["tdoc_count_str"] or "",
                            is_drafting=bool(ir["is_drafting"]),
                        )
                    )

            for sr in slot_rows:
                s = SessionSlot(
                    id=sr["id"],
                    snapshot_id=sr["snapshot_id"],
                    day_index=sr["day_index"],
                    day_name=sr["day_name"],
                    day_date=sr["day_date"] or "",
                    quarter=sr["quarter"],
                    sub_slot=sr["sub_slot"],
                    start_time_local=sr["start_time_local"],
                    end_time_local=sr["end_time_local"],
                    stream_number=sr["stream_number"],
                    chair_override=sr["chair_override"] or "",
                    room_override=sr["room_override"] if "room_override" in sr.keys() and sr["room_override"] else "",
                    raw_text=sr["raw_text"] or "",
                    is_drafting=bool(sr["is_drafting"]),
                    is_tentative=bool(sr["is_tentative"]),
                    is_single_stream=bool(sr["is_single_stream"]),
                    is_plenary=bool(sr["is_plenary"]),
                    items=items_by_slot.get(sr["id"], []),
                )
                snapshot.slots.append(s)

            self._repair_stale_slot_items(conn, snapshot)
            return snapshot

    @staticmethod
    def _repair_stale_slot_items(conn, snapshot: SessionPlanSnapshot):
        """
        Re-parses cached slot_agenda_items when parser improvements (e.g. Unicode dash
        ranges like `(20.6.12 – 16)` or multi-KI lists) extract richer or corrected AIs
        compared to an earlier import stored in SQLite. Runs idempotently once per revision.
        """
        if not snapshot.slots:
            return
        try:
            from modules.meetings.core.session_plan.session_plan_parser import PERIOD_ORDER, SessionPlanParser
        except Exception:
            return

        period_rank = {p: idx for idx, p in enumerate(PERIOD_ORDER)}
        chronological = sorted(
            snapshot.slots,
            key=lambda sl: (sl.day_index, period_rank.get(sl.quarter, 99), sl.sub_slot, sl.stream_number),
        )

        def _items_sig(items_list: List[AgendaSlotItem]) -> List[Tuple[str, str, str]]:
            return [(it.ai_number or "", it.topic_label or "", it.tdoc_count_str or "") for it in items_list]

        ki_map: Dict[str, str] = {}
        repaired_count = 0
        for s in chronological:
            if not s.items or not s.raw_text or s.raw_text.strip() in ("-", "—", "Revisions", "Revisions (tbd)"):
                continue
            fresh_room_ov, fresh_items = SessionPlanParser._extract_agenda_items(
                s.raw_text,
                default_drafting=s.is_drafting,
                default_tentative=s.is_tentative,
                ki_map=ki_map,
            )
            if _items_sig(s.items) != _items_sig(fresh_items):
                s.items = fresh_items
                if fresh_room_ov and not s.room_override:
                    s.room_override = fresh_room_ov
                    if s.id is not None:
                        conn.execute(
                            "UPDATE session_slots SET room_override = ? WHERE id = ?",
                            (fresh_room_ov, s.id),
                        )
                if s.id is not None:
                    conn.execute("DELETE FROM slot_agenda_items WHERE slot_id = ?", (s.id,))
                    for item in fresh_items:
                        conn.execute(
                            """
                            INSERT INTO slot_agenda_items (slot_id, ai_number, display_expr, topic_label, tdoc_count_str, is_drafting)
                            VALUES (?, ?, ?, ?, ?, ?)
                            """,
                            (
                                s.id,
                                item.ai_number,
                                item.display_expr,
                                item.topic_label,
                                item.tdoc_count_str,
                                1 if item.is_drafting else 0,
                            ),
                        )
                repaired_count += 1

        if repaired_count:
            logger.info(
                f"[SessionPlanDB] Auto-repaired {repaired_count} stale slot item list(s) in snapshot {snapshot.id}."
            )

    # --- User Annotations ---

    def get_annotations(self, meeting_number: str) -> Dict[str, UserSessionAnnotation]:
        """Returns map of target_key -> UserSessionAnnotation."""
        with self._get_connection() as conn:
            rows = conn.execute(
                "SELECT * FROM user_session_annotations WHERE meeting_number = ?",
                (str(meeting_number),),
            ).fetchall()
            result = {}
            for r in rows:
                result[r["target_key"]] = UserSessionAnnotation(
                    target_type=r["target_type"],
                    target_key=r["target_key"],
                    status_badge=r["status_badge"] or "normal",
                    assigned_colleague=r["assigned_colleague"] or "",
                    comment=r["comment"] or "",
                )
            return result

    def set_annotation(
        self,
        meeting_number: str,
        target_type: str,
        target_key: str,
        status_badge: str = "normal",
        assigned_colleague: str = "",
        comment: str = "",
    ):
        with self._get_connection() as conn:
            conn.execute(
                """
                INSERT INTO user_session_annotations (meeting_number, target_type, target_key, status_badge, assigned_colleague, comment)
                VALUES (?, ?, ?, ?, ?, ?)
                ON CONFLICT(meeting_number, target_type, target_key) DO UPDATE SET
                    status_badge = excluded.status_badge,
                    assigned_colleague = excluded.assigned_colleague,
                    comment = excluded.comment,
                    updated_at = CURRENT_TIMESTAMP
                """,
                (str(meeting_number), target_type, target_key, status_badge, assigned_colleague, comment),
            )

    # --- User Tracked Topics ---

    def get_tracked_topics(self, meeting_number: str) -> Dict[str, str]:
        """Returns dict of topic_token -> topic_label."""
        with self._get_connection() as conn:
            rows = conn.execute(
                "SELECT topic_token, topic_label FROM user_tracked_topics WHERE meeting_number = ? ORDER BY id ASC",
                (str(meeting_number),),
            ).fetchall()
            return {r["topic_token"]: (r["topic_label"] or r["topic_token"]) for r in rows}

    def add_tracked_topic(self, meeting_number: str, token: str, label: str = ""):
        clean_tok = str(token).lower().strip()
        if not clean_tok:
            return
        with self._get_connection() as conn:
            conn.execute(
                """
                INSERT INTO user_tracked_topics (meeting_number, topic_token, topic_label)
                VALUES (?, ?, ?)
                ON CONFLICT(meeting_number, topic_token) DO UPDATE SET
                    topic_label = COALESCE(NULLIF(excluded.topic_label, ''), topic_label)
                """,
                (str(meeting_number), clean_tok, label or token),
            )

    def remove_tracked_topic(self, meeting_number: str, token: str):
        clean_tok = str(token).lower().strip()
        with self._get_connection() as conn:
            conn.execute(
                "DELETE FROM user_tracked_topics WHERE meeting_number = ? AND topic_token = ?",
                (str(meeting_number), clean_tok),
            )

    def set_tracked_topics(self, meeting_number: str, topics_dict: Dict[str, str]):
        """Replaces tracked topics for the meeting."""
        with self._get_connection() as conn:
            conn.execute("DELETE FROM user_tracked_topics WHERE meeting_number = ?", (str(meeting_number),))
            for tok, lbl in topics_dict.items():
                clean_tok = str(tok).lower().strip()
                if clean_tok:
                    conn.execute(
                        "INSERT INTO user_tracked_topics (meeting_number, topic_token, topic_label) VALUES (?, ?, ?)",
                        (str(meeting_number), clean_tok, lbl or tok),
                    )
