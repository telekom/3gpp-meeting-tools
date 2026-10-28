# --- File: src/modules/meetings/core/session_plan_db.py ---
"""
Persistent SQLite database interface for meeting session plans.
Stores snapshots, stream topology, sub-slots, and user annotations.
Located at: <meeting_dir>/Agenda/session_plans.db
"""
import contextlib
import logging
import sqlite3
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from modules.meetings.core.session_plan_models import (
    AgendaSlotItem,
    MeetingStreamInfo,
    SessionPlanSnapshot,
    SessionSlot,
    UserSessionAnnotation,
)

logger = logging.getLogger(__name__)


class SessionPlanDatabase:
    def __init__(self, db_path: Path):
        self.db_path = Path(db_path)
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        self._init_tables()

    @contextlib.contextmanager
    def _get_connection(self):
        conn = sqlite3.connect(self.db_path, timeout=30.0)
        conn.row_factory = sqlite3.Row
        try:
            conn.execute("PRAGMA journal_mode=WAL;")
            conn.execute("PRAGMA busy_timeout=30000;")
            with conn:
                yield conn
        finally:
            conn.close()

    def _init_tables(self):
        with self._get_connection() as conn:
            cursor = conn.cursor()

            # 1. Meeting Metadata & Confirmed Timezone
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS meeting_meta (
                    meeting_number TEXT PRIMARY KEY,
                    location TEXT,
                    start_date TEXT,
                    end_date TEXT,
                    tz_offset_minutes INTEGER,
                    tz_name TEXT,
                    tz_confirmed INTEGER DEFAULT 0
                )
            """)

            # 2. Session Plan Snapshots
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS session_plan_snapshots (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    meeting_number TEXT NOT NULL,
                    tdoc_id TEXT NOT NULL,
                    snapshot_label TEXT NOT NULL,
                    content_hash TEXT NOT NULL,
                    cached_file_path TEXT,
                    source_url TEXT,
                    imported_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    is_active INTEGER DEFAULT 0
                )
            """)

            # 3. Stream & Room Definitions
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS meeting_streams (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    snapshot_id INTEGER NOT NULL,
                    stream_number INTEGER NOT NULL,
                    stream_name TEXT NOT NULL,
                    room_name TEXT NOT NULL,
                    capacity INTEGER DEFAULT 0,
                    convener TEXT,
                    tohru_room TEXT,
                    FOREIGN KEY (snapshot_id) REFERENCES session_plan_snapshots (id) ON DELETE CASCADE
                )
            """)

            # 4. Scheduled Sub-Slots (TU 1 / TU 2)
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS session_slots (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    snapshot_id INTEGER NOT NULL,
                    day_index INTEGER NOT NULL,
                    day_name TEXT NOT NULL,
                    day_date TEXT,
                    quarter TEXT NOT NULL,
                    sub_slot INTEGER NOT NULL,
                    start_time_local TEXT NOT NULL,
                    end_time_local TEXT NOT NULL,
                    stream_number INTEGER NOT NULL,
                    chair_override TEXT,
                    raw_text TEXT NOT NULL,
                    is_drafting INTEGER DEFAULT 0,
                    is_tentative INTEGER DEFAULT 0,
                    is_single_stream INTEGER DEFAULT 0,
                    is_plenary INTEGER DEFAULT 0,
                    FOREIGN KEY (snapshot_id) REFERENCES session_plan_snapshots (id) ON DELETE CASCADE
                )
            """)

            # 5. Extracted Agenda Items per Slot
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS slot_agenda_items (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    slot_id INTEGER NOT NULL,
                    ai_number TEXT NOT NULL,
                    display_expr TEXT,
                    topic_label TEXT,
                    tdoc_count_str TEXT,
                    is_drafting INTEGER DEFAULT 0,
                    FOREIGN KEY (slot_id) REFERENCES session_slots (id) ON DELETE CASCADE
                )
            """)

            # 6. Persistent User Marks & Notes (Survives across plan snapshots!)
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS user_session_annotations (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    meeting_number TEXT NOT NULL,
                    target_type TEXT NOT NULL,
                    target_key TEXT NOT NULL,
                    status_badge TEXT DEFAULT 'normal',
                    assigned_colleague TEXT,
                    comment TEXT,
                    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    UNIQUE(meeting_number, target_type, target_key)
                )
            """)

            cursor.execute("CREATE INDEX IF NOT EXISTS idx_slots_snap ON session_slots(snapshot_id);")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_items_slot ON slot_agenda_items(slot_id);")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_ann_mtg ON user_session_annotations(meeting_number, target_key);")

    # --- Meeting Metadata & Timezone ---

    def get_meeting_meta(self, meeting_number: str) -> Optional[Dict]:
        with self._get_connection() as conn:
            row = conn.execute(
                "SELECT * FROM meeting_meta WHERE meeting_number = ?", (str(meeting_number),)
            ).fetchone()
            return dict(row) if row else None

    def set_meeting_meta(
        self,
        meeting_number: str,
        location: str = "",
        start_date: str = "",
        end_date: str = "",
        tz_offset_minutes: Optional[int] = None,
        tz_name: str = "",
        tz_confirmed: int = 1,
    ):
        with self._get_connection() as conn:
            conn.execute(
                """
                INSERT INTO meeting_meta (meeting_number, location, start_date, end_date, tz_offset_minutes, tz_name, tz_confirmed)
                VALUES (?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(meeting_number) DO UPDATE SET
                    location = COALESCE(NULLIF(excluded.location, ''), location),
                    start_date = COALESCE(NULLIF(excluded.start_date, ''), start_date),
                    end_date = COALESCE(NULLIF(excluded.end_date, ''), end_date),
                    tz_offset_minutes = COALESCE(excluded.tz_offset_minutes, tz_offset_minutes),
                    tz_name = COALESCE(NULLIF(excluded.tz_name, ''), tz_name),
                    tz_confirmed = excluded.tz_confirmed
                """,
                (str(meeting_number), location, start_date, end_date, tz_offset_minutes, tz_name, tz_confirmed),
            )

    # --- Snapshots CRUD ---

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
            for s_num, s_info in snapshot.streams.items():
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
                                               raw_text, is_drafting, is_tentative, is_single_stream, is_plenary)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
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
                    raw_text=sr["raw_text"] or "",
                    is_drafting=bool(sr["is_drafting"]),
                    is_tentative=bool(sr["is_tentative"]),
                    is_single_stream=bool(sr["is_single_stream"]),
                    is_plenary=bool(sr["is_plenary"]),
                    items=items_by_slot.get(sr["id"], []),
                )
                snapshot.slots.append(s)

            return snapshot

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
