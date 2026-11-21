# --- File: src/modules/meetings/core/session_plan/db/mixins/schema_and_defaults.py ---
"""SQLite table initialization, schema migrations, cross-meeting defaults sync, and metadata CRUD."""

import json
import logging
from typing import Dict, Optional

logger = logging.getLogger("modules.meetings.core.session_plan.session_plan_db")


class SessionPlanSchemaDefaultsMixin:
    """Handles SQLite schema creation, global delegation_defaults.json sync, and meeting_meta."""

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
                    room_override TEXT,
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

            # 11. Cross-Meeting Persistent Default Delegation Roster
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS default_delegation_roster (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    name TEXT NOT NULL UNIQUE,
                    email TEXT DEFAULT '',
                    color TEXT DEFAULT '#2563EB',
                    is_self INTEGER DEFAULT 0,
                    notes TEXT DEFAULT ''
                )
            """)

            # 12. Cross-Meeting Persistent Default Topic Ownership (AI -> Default Owner)
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS default_topic_owners (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    ai_number TEXT NOT NULL UNIQUE,
                    owner_name TEXT NOT NULL,
                    notes TEXT DEFAULT ''
                )
            """)

            # 13. Physical Room Custom Location Notes (Floor / Wing / Hall)
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS room_locations (
                    meeting_number TEXT NOT NULL,
                    room_name TEXT NOT NULL,
                    location_note TEXT NOT NULL,
                    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    PRIMARY KEY (meeting_number, room_name)
                )
            """)

            # 14. Optional Per-Slot Location Overrides (e.g. ad-hoc drafting foyer/table)
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS slot_locations (
                    meeting_number TEXT NOT NULL,
                    slot_key TEXT NOT NULL,
                    location_note TEXT NOT NULL,
                    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    PRIMARY KEY (meeting_number, slot_key)
                )
            """)

            cursor.execute("CREATE INDEX IF NOT EXISTS idx_slots_snap ON session_slots(snapshot_id);")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_items_slot ON slot_agenda_items(slot_id);")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_ann_mtg ON user_session_annotations(meeting_number, target_key);")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_trk_mtg ON user_tracked_topics(meeting_number);")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_roster_mtg ON delegation_roster(meeting_number);")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_town_mtg ON topic_owners(meeting_number);")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_sapp_mtg ON session_appointments(meeting_number);")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_rloc_mtg ON room_locations(meeting_number);")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_sloc_mtg ON slot_locations(meeting_number);")

            # Schema migration: room_override column
            try:
                cursor.execute("ALTER TABLE session_slots ADD COLUMN room_override TEXT;")
            except Exception:
                pass

    # --- Cross-Meeting Persistent Defaults Sync ---

    def _sync_from_global_defaults(self):
        """
        Synchronizes `default_delegation_roster` and `default_topic_owners` with the
        shared cross-meeting `delegation_defaults.json` file.
        If no global file exists yet, seeds defaults from any existing meeting in this DB.
        """
        try:
            if self.global_defaults_path and self.global_defaults_path.is_file():
                raw = json.loads(self.global_defaults_path.read_text(encoding="utf-8"))
                with self._get_connection() as conn:
                    conn.execute("DELETE FROM default_delegation_roster")
                    for m_data in raw.get("members", []):
                        name = str(m_data.get("name", "")).strip()
                        if not name:
                            continue
                        conn.execute(
                            """
                            INSERT INTO default_delegation_roster (name, email, color, is_self, notes)
                            VALUES (?, ?, ?, ?, ?)
                            ON CONFLICT(name) DO UPDATE SET
                                email = excluded.email,
                                color = excluded.color,
                                is_self = excluded.is_self,
                                notes = excluded.notes
                            """,
                            (
                                name,
                                str(m_data.get("email", "")).strip(),
                                str(m_data.get("color", "")).strip() or "#2563EB",
                                1 if m_data.get("is_self") else 0,
                                str(m_data.get("notes", "")).strip(),
                            ),
                        )
                    conn.execute("DELETE FROM default_topic_owners")
                    for ai_num, to_data in raw.get("topic_owners", {}).items():
                        owner = to_data.get("owner", "") if isinstance(to_data, dict) else str(to_data)
                        notes = to_data.get("notes", "") if isinstance(to_data, dict) else ""
                        clean_ai = str(ai_num).strip()
                        clean_owner = str(owner).strip()
                        if clean_ai and clean_owner:
                            conn.execute(
                                """
                                INSERT INTO default_topic_owners (ai_number, owner_name, notes)
                                VALUES (?, ?, ?)
                                ON CONFLICT(ai_number) DO UPDATE SET
                                    owner_name = excluded.owner_name,
                                    notes = excluded.notes
                                """,
                                (clean_ai, clean_owner, str(notes).strip()),
                            )
                return

            # Seed defaults from existing meeting rows if global file does not exist yet
            with self._get_connection() as conn:
                has_def_roster = conn.execute("SELECT COUNT(*) FROM default_delegation_roster").fetchone()[0] > 0
                has_def_topics = conn.execute("SELECT COUNT(*) FROM default_topic_owners").fetchone()[0] > 0
                seeded = False
                if not has_def_roster:
                    rows = conn.execute("SELECT name, email, color, is_self, notes FROM delegation_roster ORDER BY id ASC").fetchall()
                    for r in rows:
                        conn.execute(
                            """
                            INSERT INTO default_delegation_roster (name, email, color, is_self, notes)
                            VALUES (?, ?, ?, ?, ?)
                            ON CONFLICT(name) DO UPDATE SET
                                email = excluded.email,
                                color = excluded.color,
                                is_self = excluded.is_self,
                                notes = excluded.notes
                            """,
                            (r["name"], r["email"] or "", r["color"] or "#2563EB", r["is_self"] or 0, r["notes"] or ""),
                        )
                        seeded = True
                if not has_def_topics:
                    rows = conn.execute("SELECT ai_number, owner_name, notes FROM topic_owners ORDER BY id ASC").fetchall()
                    for r in rows:
                        conn.execute(
                            """
                            INSERT INTO default_topic_owners (ai_number, owner_name, notes)
                            VALUES (?, ?, ?)
                            ON CONFLICT(ai_number) DO UPDATE SET
                                owner_name = excluded.owner_name,
                                notes = excluded.notes
                            """,
                            (r["ai_number"], r["owner_name"], r["notes"] or ""),
                        )
                        seeded = True
                if seeded:
                    self._flush_global_defaults_conn(conn)
        except Exception as e:
            logger.debug(f"[SessionPlanDB] Could not sync global defaults: {e}")

    def _flush_global_defaults_conn(self, conn) -> None:
        """Writes current `default_delegation_roster` and `default_topic_owners` to `self.global_defaults_path`."""
        if not self.global_defaults_path:
            return
        try:
            m_rows = conn.execute("SELECT name, email, color, is_self, notes FROM default_delegation_roster ORDER BY id ASC").fetchall()
            t_rows = conn.execute("SELECT ai_number, owner_name, notes FROM default_topic_owners ORDER BY id ASC").fetchall()
            payload = {
                "3gpp_delegation_defaults_version": 1,
                "members": [
                    {
                        "name": r["name"],
                        "email": r["email"] or "",
                        "color": r["color"] or "#2563EB",
                        "is_self": bool(r["is_self"]),
                        "notes": r["notes"] or "",
                    }
                    for r in m_rows
                ],
                "topic_owners": {
                    r["ai_number"]: {"owner": r["owner_name"], "notes": r["notes"] or ""}
                    for r in t_rows
                },
            }
            self.global_defaults_path.parent.mkdir(parents=True, exist_ok=True)
            self.global_defaults_path.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")
        except Exception as e:
            logger.debug(f"[SessionPlanDB] Could not write global defaults file: {e}")

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
