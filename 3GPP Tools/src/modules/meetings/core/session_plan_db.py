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
    DELEGATE_PALETTE,
    AgendaSlotItem,
    DelegationMember,
    DelegationPlan,
    MeetingStreamInfo,
    SessionAppointment,
    SessionPlanSnapshot,
    SessionSlot,
    TopicOwnership,
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

            # 7. Persistent User Tracked Topics for this meeting
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS user_tracked_topics (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    meeting_number TEXT NOT NULL,
                    topic_token TEXT NOT NULL,
                    topic_label TEXT,
                    added_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    UNIQUE(meeting_number, topic_token)
                )
            """)

            # 8. Delegation Roster (Colleagues / Team Members)
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS delegation_roster (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    meeting_number TEXT NOT NULL,
                    name TEXT NOT NULL,
                    email TEXT DEFAULT '',
                    color TEXT DEFAULT '#2563EB',
                    is_self INTEGER DEFAULT 0,
                    notes TEXT DEFAULT '',
                    UNIQUE(meeting_number, name)
                )
            """)

            # 9. Topic Ownership (AI / Agenda Item to Lead Delegate)
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS topic_owners (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    meeting_number TEXT NOT NULL,
                    ai_number TEXT NOT NULL,
                    owner_name TEXT NOT NULL,
                    notes TEXT DEFAULT '',
                    UNIQUE(meeting_number, ai_number)
                )
            """)

            # 10. Session-Specific Appointments (Collision Overrides & Ad-hoc Coverage)
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS session_appointments (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    meeting_number TEXT NOT NULL,
                    slot_key TEXT NOT NULL,
                    appointed_delegate TEXT NOT NULL,
                    original_owner TEXT DEFAULT '',
                    reason TEXT DEFAULT '',
                    UNIQUE(meeting_number, slot_key)
                )
            """)

            cursor.execute("CREATE INDEX IF NOT EXISTS idx_slots_snap ON session_slots(snapshot_id);")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_items_slot ON slot_agenda_items(slot_id);")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_ann_mtg ON user_session_annotations(meeting_number, target_key);")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_trk_mtg ON user_tracked_topics(meeting_number);")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_roster_mtg ON delegation_roster(meeting_number);")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_town_mtg ON topic_owners(meeting_number);")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_sapp_mtg ON session_appointments(meeting_number);")

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

    # --- Delegation Roster CRUD ---

    def get_delegation_members(self, meeting_number: str) -> Dict[str, DelegationMember]:
        """Returns map of name -> DelegationMember."""
        with self._get_connection() as conn:
            rows = conn.execute(
                "SELECT name, email, color, is_self, notes FROM delegation_roster WHERE meeting_number = ? ORDER BY id ASC",
                (str(meeting_number),),
            ).fetchall()
            result = {}
            for r in rows:
                result[r["name"]] = DelegationMember(
                    name=r["name"],
                    email=r["email"] or "",
                    color=r["color"] or DELEGATE_PALETTE[len(result) % len(DELEGATE_PALETTE)],
                    is_self=bool(r["is_self"]),
                    notes=r["notes"] or "",
                )
            return result

    def save_delegation_member(self, meeting_number: str, member: DelegationMember):
        """Adds or updates a delegation member."""
        with self._get_connection() as conn:
            # If marking as self, unmark any previous self in this meeting
            if member.is_self:
                conn.execute(
                    "UPDATE delegation_roster SET is_self = 0 WHERE meeting_number = ?",
                    (str(meeting_number),),
                )
            conn.execute(
                """
                INSERT INTO delegation_roster (meeting_number, name, email, color, is_self, notes)
                VALUES (?, ?, ?, ?, ?, ?)
                ON CONFLICT(meeting_number, name) DO UPDATE SET
                    email = excluded.email,
                    color = excluded.color,
                    is_self = excluded.is_self,
                    notes = excluded.notes
                """,
                (
                    str(meeting_number),
                    member.name.strip(),
                    member.email.strip(),
                    member.color.strip() or "#2563EB",
                    1 if member.is_self else 0,
                    member.notes.strip(),
                ),
            )

    def delete_delegation_member(self, meeting_number: str, name: str):
        """Deletes a delegation member and unassigns any of their topics/appointments."""
        with self._get_connection() as conn:
            clean_name = name.strip()
            conn.execute(
                "DELETE FROM delegation_roster WHERE meeting_number = ? AND name = ?",
                (str(meeting_number), clean_name),
            )
            conn.execute(
                "DELETE FROM topic_owners WHERE meeting_number = ? AND owner_name = ?",
                (str(meeting_number), clean_name),
            )
            conn.execute(
                "DELETE FROM session_appointments WHERE meeting_number = ? AND appointed_delegate = ?",
                (str(meeting_number), clean_name),
            )

    def set_self_delegate(self, meeting_number: str, name: str):
        """Designates the given member name as the local user ('is_self')."""
        with self._get_connection() as conn:
            conn.execute(
                "UPDATE delegation_roster SET is_self = 0 WHERE meeting_number = ?",
                (str(meeting_number),),
            )
            conn.execute(
                "UPDATE delegation_roster SET is_self = 1 WHERE meeting_number = ? AND name = ?",
                (str(meeting_number), name.strip()),
            )

    # --- Topic Ownership CRUD ---

    def get_topic_owners(self, meeting_number: str) -> Dict[str, TopicOwnership]:
        """Returns map of ai_number -> TopicOwnership."""
        with self._get_connection() as conn:
            rows = conn.execute(
                "SELECT ai_number, owner_name, notes FROM topic_owners WHERE meeting_number = ? ORDER BY id ASC",
                (str(meeting_number),),
            ).fetchall()
            return {
                r["ai_number"]: TopicOwnership(
                    meeting_number=str(meeting_number),
                    ai_number=r["ai_number"],
                    owner_name=r["owner_name"],
                    notes=r["notes"] or "",
                )
                for r in rows
            }

    def set_topic_owner(self, meeting_number: str, ai_number: str, owner_name: str, notes: str = ""):
        clean_ai = ai_number.strip()
        clean_owner = owner_name.strip()
        with self._get_connection() as conn:
            if not clean_owner:
                conn.execute(
                    "DELETE FROM topic_owners WHERE meeting_number = ? AND ai_number = ?",
                    (str(meeting_number), clean_ai),
                )
            else:
                conn.execute(
                    """
                    INSERT INTO topic_owners (meeting_number, ai_number, owner_name, notes)
                    VALUES (?, ?, ?, ?)
                    ON CONFLICT(meeting_number, ai_number) DO UPDATE SET
                        owner_name = excluded.owner_name,
                        notes = excluded.notes
                    """,
                    (str(meeting_number), clean_ai, clean_owner, notes.strip()),
                )

    def delete_topic_owner(self, meeting_number: str, ai_number: str):
        with self._get_connection() as conn:
            conn.execute(
                "DELETE FROM topic_owners WHERE meeting_number = ? AND ai_number = ?",
                (str(meeting_number), ai_number.strip()),
            )

    # --- Session Appointments CRUD (Collision / Ad-Hoc Overrides) ---

    def get_session_appointments(self, meeting_number: str) -> Dict[str, SessionAppointment]:
        """Returns map of slot_key -> SessionAppointment."""
        with self._get_connection() as conn:
            rows = conn.execute(
                "SELECT slot_key, appointed_delegate, original_owner, reason FROM session_appointments WHERE meeting_number = ?",
                (str(meeting_number),),
            ).fetchall()
            return {
                r["slot_key"]: SessionAppointment(
                    meeting_number=str(meeting_number),
                    slot_key=r["slot_key"],
                    appointed_delegate=r["appointed_delegate"],
                    original_owner=r["original_owner"] or "",
                    reason=r["reason"] or "",
                )
                for r in rows
            }

    def set_session_appointment(
        self,
        meeting_number: str,
        slot_key: str,
        appointed_delegate: str,
        original_owner: str = "",
        reason: str = "",
    ):
        clean_delegate = appointed_delegate.strip()
        with self._get_connection() as conn:
            if not clean_delegate:
                conn.execute(
                    "DELETE FROM session_appointments WHERE meeting_number = ? AND slot_key = ?",
                    (str(meeting_number), slot_key.strip()),
                )
            else:
                conn.execute(
                    """
                    INSERT INTO session_appointments (meeting_number, slot_key, appointed_delegate, original_owner, reason)
                    VALUES (?, ?, ?, ?, ?)
                    ON CONFLICT(meeting_number, slot_key) DO UPDATE SET
                        appointed_delegate = excluded.appointed_delegate,
                        original_owner = excluded.original_owner,
                        reason = excluded.reason
                    """,
                    (str(meeting_number), slot_key.strip(), clean_delegate, original_owner.strip(), reason.strip()),
                )

    def delete_session_appointment(self, meeting_number: str, slot_key: str):
        with self._get_connection() as conn:
            conn.execute(
                "DELETE FROM session_appointments WHERE meeting_number = ? AND slot_key = ?",
                (str(meeting_number), slot_key.strip()),
            )

    # --- Aggregate Delegation Plan ---

    def get_delegation_plan(self, meeting_number: str) -> DelegationPlan:
        return DelegationPlan(
            meeting_number=str(meeting_number),
            members=self.get_delegation_members(meeting_number),
            topic_owners=self.get_topic_owners(meeting_number),
            appointments=self.get_session_appointments(meeting_number),
        )

    def export_delegation_plan_dict(self, meeting_number: str) -> dict:
        """Serializes delegation state to JSON-serializable dictionary."""
        plan = self.get_delegation_plan(meeting_number)
        return {
            "3gpp_delegation_plan_version": 1,
            "meeting_number": str(meeting_number),
            "members": [
                {
                    "name": m.name,
                    "email": m.email,
                    "color": m.color,
                    "is_self": m.is_self,
                    "notes": m.notes,
                }
                for m in plan.members.values()
            ],
            "topic_owners": {
                ai: {"owner": to.owner_name, "notes": to.notes}
                for ai, to in plan.topic_owners.items()
            },
            "appointments": {
                sk: {
                    "appointed_delegate": ap.appointed_delegate,
                    "original_owner": ap.original_owner,
                    "reason": ap.reason,
                }
                for sk, ap in plan.appointments.items()
            },
        }

    def import_delegation_plan_dict(self, meeting_number: str, data: dict, merge: bool = False) -> int:
        """
        Imports delegation state from dictionary.
        If merge=False, clears existing roster, topic owners, and appointments for this meeting.
        Returns count of imported/updated items.
        """
        imported_count = 0
        with self._get_connection() as conn:
            m_num = str(meeting_number)
            if not merge:
                conn.execute("DELETE FROM delegation_roster WHERE meeting_number = ?", (m_num,))
                conn.execute("DELETE FROM topic_owners WHERE meeting_number = ?", (m_num,))
                conn.execute("DELETE FROM session_appointments WHERE meeting_number = ?", (m_num,))

            # 1. Members
            for m_data in data.get("members", []):
                name = m_data.get("name", "").strip()
                if not name:
                    continue
                email = m_data.get("email", "").strip()
                color = m_data.get("color", "").strip() or "#2563EB"
                is_self = 1 if m_data.get("is_self") else 0
                notes = m_data.get("notes", "").strip()
                conn.execute(
                    """
                    INSERT INTO delegation_roster (meeting_number, name, email, color, is_self, notes)
                    VALUES (?, ?, ?, ?, ?, ?)
                    ON CONFLICT(meeting_number, name) DO UPDATE SET
                        email = excluded.email,
                        color = excluded.color,
                        is_self = excluded.is_self,
                        notes = excluded.notes
                    """,
                    (m_num, name, email, color, is_self, notes),
                )
                imported_count += 1

            # 2. Topic Owners
            for ai, to_data in data.get("topic_owners", {}).items():
                owner = to_data.get("owner", "") if isinstance(to_data, dict) else str(to_data)
                notes = to_data.get("notes", "") if isinstance(to_data, dict) else ""
                if ai and owner:
                    conn.execute(
                        """
                        INSERT INTO topic_owners (meeting_number, ai_number, owner_name, notes)
                        VALUES (?, ?, ?, ?)
                        ON CONFLICT(meeting_number, ai_number) DO UPDATE SET
                            owner_name = excluded.owner_name,
                            notes = excluded.notes
                        """,
                        (m_num, str(ai).strip(), str(owner).strip(), str(notes).strip()),
                    )
                    imported_count += 1

            # 3. Appointments
            for sk, ap_data in data.get("appointments", {}).items():
                if isinstance(ap_data, dict):
                    app_del = ap_data.get("appointed_delegate", "").strip()
                    orig_ow = ap_data.get("original_owner", "").strip()
                    reason = ap_data.get("reason", "").strip()
                else:
                    app_del = str(ap_data).strip()
                    orig_ow = ""
                    reason = ""
                if sk and app_del:
                    conn.execute(
                        """
                        INSERT INTO session_appointments (meeting_number, slot_key, appointed_delegate, original_owner, reason)
                        VALUES (?, ?, ?, ?, ?)
                        ON CONFLICT(meeting_number, slot_key) DO UPDATE SET
                            appointed_delegate = excluded.appointed_delegate,
                            original_owner = excluded.original_owner,
                            reason = excluded.reason
                        """,
                        (m_num, str(sk).strip(), app_del, orig_ow, reason),
                    )
                    imported_count += 1

        return imported_count
