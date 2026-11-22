# --- File: src/modules/emails/core/tracked_threads_db.py ---
import contextlib
import json
import sqlite3
import uuid
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional


class TrackedThreadsDatabase:
    """SQLite persistence for pinned/tracked 3GPP email threads and their messages."""

    def __init__(self, db_path: Path):
        self.db_path = Path(db_path)
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        self._init_db()

    @contextlib.contextmanager
    def _get_connection(self):
        conn = sqlite3.connect(self.db_path, timeout=30.0, check_same_thread=False)
        conn.row_factory = sqlite3.Row
        try:
            conn.execute("PRAGMA journal_mode=WAL;")
            conn.execute("PRAGMA busy_timeout=30000;")
            conn.execute("PRAGMA foreign_keys=ON;")
            with conn:
                yield conn
        finally:
            conn.close()

    def _init_db(self):
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS tracked_threads (
                    thread_id TEXT PRIMARY KEY,
                    title TEXT NOT NULL,
                    normalized_subject TEXT NOT NULL,
                    conversation_ids TEXT DEFAULT '[]',
                    subject_patterns TEXT DEFAULT '[]',
                    outlook_folders TEXT DEFAULT '[]',
                    meeting_id TEXT DEFAULT '',
                    wg TEXT DEFAULT '',
                    linked_ais TEXT DEFAULT '[]',
                    linked_wis TEXT DEFAULT '[]',
                    linked_tdocs TEXT DEFAULT '[]',
                    linked_sub_items TEXT DEFAULT '[]',
                    user_status TEXT DEFAULT '⚪ Neutral',
                    user_notes TEXT DEFAULT '',
                    user_tags TEXT DEFAULT '[]',
                    unread_count INTEGER DEFAULT 0,
                    flagged_count INTEGER DEFAULT 0,
                    message_count INTEGER DEFAULT 0,
                    last_message_date TEXT DEFAULT '',
                    last_synced_at TEXT DEFAULT '',
                    created_at TEXT DEFAULT ''
                )
            """)
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS tracked_thread_messages (
                    id TEXT PRIMARY KEY,
                    thread_id TEXT NOT NULL,
                    subject TEXT DEFAULT '',
                    sender_name TEXT DEFAULT '',
                    sender_email TEXT DEFAULT '',
                    company TEXT DEFAULT '',
                    date_sent TEXT DEFAULT '',
                    body_text TEXT DEFAULT '',
                    clean_body TEXT DEFAULT '',
                    message_id_header TEXT DEFAULT '',
                    in_reply_to TEXT DEFAULT '',
                    conversation_id TEXT DEFAULT '',
                    is_read INTEGER DEFAULT 0,
                    flag_status INTEGER DEFAULT 0,
                    attachments TEXT DEFAULT '[]',
                    FOREIGN KEY(thread_id) REFERENCES tracked_threads(thread_id) ON DELETE CASCADE
                )
            """)
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_tt_meeting ON tracked_threads(meeting_id)")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_tt_norm_subj ON tracked_threads(normalized_subject)")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_ttm_thread ON tracked_thread_messages(thread_id)")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_ttm_date ON tracked_thread_messages(date_sent)")

            # Schema migration: add linked_sub_items, annotations, and flag columns for existing DBs
            cursor.execute("PRAGMA table_info(tracked_threads)")
            cols = [info[1] for info in cursor.fetchall()]
            if "linked_sub_items" not in cols:
                cursor.execute("ALTER TABLE tracked_threads ADD COLUMN linked_sub_items TEXT DEFAULT '[]'")
            if "user_status" not in cols:
                cursor.execute("ALTER TABLE tracked_threads ADD COLUMN user_status TEXT DEFAULT '⚪ Neutral'")
            if "user_notes" not in cols:
                cursor.execute("ALTER TABLE tracked_threads ADD COLUMN user_notes TEXT DEFAULT ''")
            if "user_tags" not in cols:
                cursor.execute("ALTER TABLE tracked_threads ADD COLUMN user_tags TEXT DEFAULT '[]'")
            if "flagged_count" not in cols:
                cursor.execute("ALTER TABLE tracked_threads ADD COLUMN flagged_count INTEGER DEFAULT 0")
            if "is_objected" not in cols:
                cursor.execute("ALTER TABLE tracked_threads ADD COLUMN is_objected INTEGER DEFAULT 0")
            if "objected_at_rev" not in cols:
                cursor.execute("ALTER TABLE tracked_threads ADD COLUMN objected_at_rev TEXT DEFAULT ''")
            cursor.execute("PRAGMA table_info(tracked_thread_messages)")
            msg_cols = [info[1] for info in cursor.fetchall()]
            if "flag_status" not in msg_cols:
                cursor.execute("ALTER TABLE tracked_thread_messages ADD COLUMN flag_status INTEGER DEFAULT 0")
            conn.commit()
            from modules.emails.core.email_deduplicator import deduplicate_tracked_thread_messages
            removed_msgs, affected_tids = deduplicate_tracked_thread_messages(conn)
            if removed_msgs and affected_tids:
                for tid in affected_tids:
                    self._recompute_thread_stats(cursor, tid)
                conn.commit()

        from modules.emails.core.email_deduplicator import deduplicate_tracked_threads_table
        deduplicate_tracked_threads_table(self)

    @staticmethod
    def _row_to_thread(row: sqlite3.Row) -> Dict[str, Any]:
        d = dict(row)
        for key in ("conversation_ids", "subject_patterns", "outlook_folders", "linked_ais", "linked_wis", "linked_tdocs", "linked_sub_items", "user_tags"):
            try:
                d[key] = json.loads(d.get(key) or "[]")
            except Exception:
                d[key] = []
        d.setdefault("user_status", "⚪ Neutral")
        d.setdefault("user_notes", "")
        d["is_objected"] = bool(int(d.get("is_objected", 0) or 0))
        d["objected_at_rev"] = str(d.get("objected_at_rev") or "").strip()
        return d

    @staticmethod
    def _row_to_message(row: sqlite3.Row) -> Dict[str, Any]:
        d = dict(row)
        try:
            d["attachments"] = json.loads(d.get("attachments") or "[]")
        except Exception:
            d["attachments"] = []
        return d

    def create_or_get_thread(
        self,
        title: str,
        normalized_subject: str,
        conversation_id: str = "",
        outlook_folders: Optional[List[str]] = None,
        meeting_id: str = "",
        wg: str = "",
        linked_ais: Optional[List[str]] = None,
        linked_wis: Optional[List[str]] = None,
        linked_tdocs: Optional[List[str]] = None,
        linked_sub_items: Optional[List[str]] = None,
    ) -> str:
        """Creates a new tracked thread or returns an existing matching thread_id."""
        with self._get_connection() as conn:
            cursor = conn.cursor()
            if conversation_id:
                cursor.execute(
                    "SELECT thread_id, conversation_ids FROM tracked_threads WHERE meeting_id = ?",
                    (meeting_id,),
                )
                for row in cursor.fetchall():
                    cids = json.loads(row["conversation_ids"] or "[]")
                    if conversation_id in cids:
                        return row["thread_id"]

            cursor.execute(
                "SELECT thread_id FROM tracked_threads WHERE normalized_subject = ? AND meeting_id = ?",
                (normalized_subject, meeting_id),
            )
            existing = cursor.fetchone()
            if existing:
                return existing["thread_id"]

            thread_id = f"THR-{uuid.uuid4().hex[:10].upper()}"
            now = datetime.now().isoformat(timespec="seconds")
            cids = [conversation_id] if conversation_id else []
            patterns = [normalized_subject] if normalized_subject else []
            cursor.execute(
                """
                INSERT INTO tracked_threads (
                    thread_id, title, normalized_subject, conversation_ids, subject_patterns,
                    outlook_folders, meeting_id, wg, linked_ais, linked_wis, linked_tdocs, linked_sub_items,
                    unread_count, message_count, last_message_date, last_synced_at, created_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 0, 0, '', ?, ?)
                """,
                (
                    thread_id,
                    title,
                    normalized_subject,
                    json.dumps(cids),
                    json.dumps(patterns),
                    json.dumps(outlook_folders or []),
                    meeting_id,
                    wg,
                    json.dumps(linked_ais or []),
                    json.dumps(linked_wis or []),
                    json.dumps(linked_tdocs or []),
                    json.dumps(linked_sub_items or []),
                    now,
                    now,
                ),
            )
            conn.commit()
            return thread_id

    def get_existing_message_states(self, thread_id: str = "") -> Dict[str, Dict[str, Any]]:
        """Returns {msg_id: {'is_read': int, 'flag_status': int, 'has_body': bool}} for cached messages."""
        with self._get_connection() as conn:
            cursor = conn.cursor()
            if thread_id:
                cursor.execute(
                    "SELECT id, is_read, COALESCE(flag_status, 0), body_text FROM tracked_thread_messages WHERE thread_id = ?",
                    (thread_id,),
                )
            else:
                cursor.execute(
                    "SELECT id, is_read, COALESCE(flag_status, 0), body_text FROM tracked_thread_messages"
                )
            return {
                row[0]: {
                    "is_read": int(row[1] or 0),
                    "flag_status": 2 if int(row[2] or 0) == 2 else 0,
                    "has_body": bool(str(row[3] or "").strip()),
                }
                for row in cursor.fetchall()
            }

    def update_message_flags_batch(self, thread_id: str, states: List[tuple]) -> int:
        """Applies light read/flag status updates [(msg_id, is_read, flag_status)] for existing Outlook items."""
        if not states:
            return 0
        updated = 0
        with self._get_connection() as conn:
            cursor = conn.cursor()
            for msg_id, is_read, flag_status in states:
                if not msg_id or str(msg_id).startswith("MSG-"):
                    continue
                ir = 1 if int(is_read or 0) else 0
                fs = 2 if int(flag_status or 0) == 2 else 0
                cursor.execute(
                    """
                    UPDATE tracked_thread_messages
                    SET is_read = ?, flag_status = ?
                    WHERE id = ? AND thread_id = ? AND (is_read != ? OR COALESCE(flag_status, 0) != ?)
                    """,
                    (ir, fs, msg_id, thread_id, ir, fs),
                )
                updated += cursor.rowcount
            self._recompute_thread_stats(cursor, thread_id)
            conn.commit()
        return updated

    def upsert_messages(self, thread_id: str, messages: List[Dict[str, Any]], sync_unread: bool = False) -> int:
        """Inserts/updates messages for a thread. When sync_unread=True, mirrors Outlook read/unread states too."""
        if not messages:
            return 0
        read_expr = (
            "CASE WHEN excluded.id LIKE 'MSG-%' THEN tracked_thread_messages.is_read ELSE excluded.is_read END"
            if sync_unread
            else "CASE WHEN excluded.is_read = 1 THEN 1 ELSE tracked_thread_messages.is_read END"
        )
        new_count = 0
        with self._get_connection() as conn:
            cursor = conn.cursor()
            for msg in messages:
                msg_id = msg["id"]
                cursor.execute("SELECT 1 FROM tracked_thread_messages WHERE id = ?", (msg_id,))
                exists = cursor.fetchone() is not None
                if not exists:
                    new_count += 1
                cursor.execute(
                    f"""
                    INSERT INTO tracked_thread_messages (
                        id, thread_id, subject, sender_name, sender_email, company,
                        date_sent, body_text, clean_body, message_id_header,
                        in_reply_to, conversation_id, is_read, flag_status, attachments
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    ON CONFLICT(id) DO UPDATE SET
                        subject = excluded.subject,
                        body_text = CASE WHEN excluded.body_text != '' THEN excluded.body_text ELSE tracked_thread_messages.body_text END,
                        clean_body = CASE WHEN excluded.clean_body != '' THEN excluded.clean_body ELSE tracked_thread_messages.clean_body END,
                        attachments = CASE WHEN excluded.attachments != '[]' THEN excluded.attachments ELSE tracked_thread_messages.attachments END,
                        is_read = {read_expr},
                        flag_status = CASE WHEN excluded.id LIKE 'MSG-%' THEN tracked_thread_messages.flag_status ELSE excluded.flag_status END
                    """,
                    (
                        msg_id,
                        thread_id,
                        msg.get("subject", ""),
                        msg.get("sender_name", ""),
                        msg.get("sender_email", ""),
                        msg.get("company", ""),
                        msg.get("date_sent", ""),
                        msg.get("body_text", ""),
                        msg.get("clean_body", ""),
                        msg.get("message_id_header", ""),
                        msg.get("in_reply_to", ""),
                        msg.get("conversation_id", ""),
                        int(msg.get("is_read", 0)),
                        2 if int(msg.get("flag_status", 0) or 0) == 2 else 0,
                        json.dumps(msg.get("attachments", [])),
                    ),
                )
            self._recompute_thread_stats(cursor, thread_id)
            conn.commit()
            from modules.emails.core.email_deduplicator import deduplicate_tracked_thread_messages
            removed_msgs, affected_tids = deduplicate_tracked_thread_messages(conn)
            if removed_msgs and affected_tids:
                for tid in affected_tids:
                    self._recompute_thread_stats(cursor, tid)
                conn.commit()
        return new_count

    def _recompute_thread_stats(self, cursor: sqlite3.Cursor, thread_id: str):
        cursor.execute(
            """
            SELECT COUNT(*) as total,
                   SUM(CASE WHEN is_read = 0 THEN 1 ELSE 0 END) as unread,
                   SUM(CASE WHEN flag_status = 2 THEN 1 ELSE 0 END) as flagged,
                   MAX(date_sent) as latest
            FROM tracked_thread_messages
            WHERE thread_id = ?
            """,
            (thread_id,),
        )
        row = cursor.fetchone()
        total = row["total"] or 0
        unread = row["unread"] or 0
        flagged = row["flagged"] or 0
        latest = row["latest"] or ""
        now = datetime.now().isoformat(timespec="seconds")
        cursor.execute(
            """
            UPDATE tracked_threads
            SET message_count = ?, unread_count = ?, flagged_count = ?, last_message_date = ?, last_synced_at = ?
            WHERE thread_id = ?
            """,
            (total, unread, flagged, latest, now, thread_id),
        )

    def get_threads_for_meeting(self, meeting_id: str = "") -> List[Dict[str, Any]]:
        with self._get_connection() as conn:
            cursor = conn.cursor()
            if meeting_id:
                cursor.execute(
                    "SELECT * FROM tracked_threads WHERE meeting_id = ? ORDER BY last_message_date DESC",
                    (meeting_id,),
                )
            else:
                cursor.execute("SELECT * FROM tracked_threads ORDER BY last_message_date DESC")
            return [self._row_to_thread(r) for r in cursor.fetchall()]

    def get_thread(self, thread_id: str) -> Optional[Dict[str, Any]]:
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM tracked_threads WHERE thread_id = ?", (thread_id,))
            row = cursor.fetchone()
            return self._row_to_thread(row) if row else None

    def get_thread_messages(self, thread_id: str) -> List[Dict[str, Any]]:
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(
                "SELECT * FROM tracked_thread_messages WHERE thread_id = ? ORDER BY date_sent ASC",
                (thread_id,),
            )
            return [self._row_to_message(r) for r in cursor.fetchall()]

    def update_thread_links(
        self,
        thread_id: str,
        linked_ais: Optional[List[str]] = None,
        linked_wis: Optional[List[str]] = None,
        linked_tdocs: Optional[List[str]] = None,
        linked_sub_items: Optional[List[str]] = None,
    ):
        with self._get_connection() as conn:
            cursor = conn.cursor()
            if linked_ais is not None:
                cursor.execute("UPDATE tracked_threads SET linked_ais = ? WHERE thread_id = ?", (json.dumps(linked_ais), thread_id))
            if linked_wis is not None:
                cursor.execute("UPDATE tracked_threads SET linked_wis = ? WHERE thread_id = ?", (json.dumps(linked_wis), thread_id))
            if linked_tdocs is not None:
                cursor.execute("UPDATE tracked_threads SET linked_tdocs = ? WHERE thread_id = ?", (json.dumps(linked_tdocs), thread_id))
            if linked_sub_items is not None:
                cursor.execute("UPDATE tracked_threads SET linked_sub_items = ? WHERE thread_id = ?", (json.dumps(linked_sub_items), thread_id))
            conn.commit()

    def update_thread_annotations(
        self,
        thread_id: str,
        user_status: Optional[str] = None,
        user_notes: Optional[str] = None,
        user_tags: Optional[List[str]] = None,
        is_objected: Optional[bool] = None,
        objected_at_rev: Optional[str] = None,
    ):
        """Persists personal stance, comments, tags, and smart-parking objection state onto a tracked email thread."""
        with self._get_connection() as conn:
            cursor = conn.cursor()
            if user_status is not None:
                cursor.execute("UPDATE tracked_threads SET user_status = ? WHERE thread_id = ?", (user_status, thread_id))
            if user_notes is not None:
                cursor.execute("UPDATE tracked_threads SET user_notes = ? WHERE thread_id = ?", (user_notes, thread_id))
            if user_tags is not None:
                cursor.execute("UPDATE tracked_threads SET user_tags = ? WHERE thread_id = ?", (json.dumps(user_tags), thread_id))
            if is_objected is not None:
                cursor.execute("UPDATE tracked_threads SET is_objected = ? WHERE thread_id = ?", (1 if is_objected else 0, thread_id))
            if objected_at_rev is not None:
                cursor.execute("UPDATE tracked_threads SET objected_at_rev = ? WHERE thread_id = ?", (str(objected_at_rev or "").strip(), thread_id))
            conn.commit()

    def add_conversation_ids(self, thread_id: str, cids: List[str]):
        """Persists newly discovered Outlook ConversationIDs onto an existing tracked thread."""
        clean_cids = [str(c).strip() for c in (cids or []) if str(c).strip()]
        if not clean_cids:
            return
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT conversation_ids FROM tracked_threads WHERE thread_id = ?", (thread_id,))
            row = cursor.fetchone()
            if not row:
                return
            try:
                existing = json.loads(row["conversation_ids"] or "[]")
            except Exception:
                existing = []
            merged = list(dict.fromkeys(existing + clean_cids))
            if merged != existing:
                cursor.execute(
                    "UPDATE tracked_threads SET conversation_ids = ? WHERE thread_id = ?",
                    (json.dumps(merged), thread_id),
                )
                conn.commit()

    def mark_thread_read(self, thread_id: str):
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("UPDATE tracked_thread_messages SET is_read = 1 WHERE thread_id = ?", (thread_id,))
            self._recompute_thread_stats(cursor, thread_id)
            conn.commit()

    def _update_message_status_col(self, message_id: str, column: str, value: int) -> Optional[str]:
        msg_id = str(message_id or "").strip()
        if not msg_id or column not in ("is_read", "flag_status"):
            return None
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT thread_id FROM tracked_thread_messages WHERE id = ?", (msg_id,))
            row = cursor.fetchone()
            if not row:
                return None
            thread_id = row["thread_id"]
            cursor.execute(f"UPDATE tracked_thread_messages SET {column} = ? WHERE id = ?", (value, msg_id))
            self._recompute_thread_stats(cursor, thread_id)
            conn.commit()
            return thread_id

    def set_message_read_status(self, message_id: str, is_read: bool) -> Optional[str]:
        """Marks a single tracked thread message read/unread and recomputes thread unread count."""
        return self._update_message_status_col(message_id, "is_read", 1 if is_read else 0)

    def set_message_flag_status(self, message_id: str, is_flagged: bool) -> Optional[str]:
        """Sets 2-state Outlook-compatible flag_status (2=flagged, 0=unflagged) on a single message."""
        return self._update_message_status_col(message_id, "flag_status", 2 if is_flagged else 0)

    def merge_threads(self, target_thread_id: str, source_thread_ids: List[str]):
        """Merges source threads into target_thread_id (repairing Outlook-split threads)."""
        if not source_thread_ids:
            return
        with self._get_connection() as conn:
            cursor = conn.cursor()
            target = self.get_thread(target_thread_id)
            if not target:
                return
            cids = set(target["conversation_ids"])
            patterns = set(target["subject_patterns"])
            folders = set(target.get("outlook_folders") or [])
            ais = set(target["linked_ais"])
            wis = set(target["linked_wis"])
            tdocs = set(target["linked_tdocs"])
            sub_items = set(target.get("linked_sub_items") or [])
            tags = list(dict.fromkeys(target.get("user_tags") or []))
            seen_tags = set(tags)
            status = str(target.get("user_status") or "").strip()
            note_parts = [str(target.get("user_notes") or "").strip()] if str(target.get("user_notes") or "").strip() else []

            def _is_neutral_status(val: str) -> bool:
                v = str(val or "").strip().lower()
                return not v or v in ("neutral", "⚪ neutral")

            for src_id in source_thread_ids:
                if src_id == target_thread_id:
                    continue
                src = self.get_thread(src_id)
                if not src:
                    continue
                cids.update(src["conversation_ids"])
                patterns.update(src["subject_patterns"])
                folders.update(src.get("outlook_folders") or [])
                ais.update(src["linked_ais"])
                wis.update(src["linked_wis"])
                tdocs.update(src["linked_tdocs"])
                sub_items.update(src.get("linked_sub_items") or [])
                for tg in (src.get("user_tags") or []):
                    if tg and tg not in seen_tags:
                        seen_tags.add(tg)
                        tags.append(tg)
                src_status = str(src.get("user_status") or "").strip()
                if _is_neutral_status(status) and not _is_neutral_status(src_status):
                    status = src_status
                src_note = str(src.get("user_notes") or "").strip()
                if src_note and src_note not in note_parts:
                    note_parts.append(src_note)
                cursor.execute(
                    "UPDATE tracked_thread_messages SET thread_id = ? WHERE thread_id = ?",
                    (target_thread_id, src_id),
                )
                cursor.execute("DELETE FROM tracked_threads WHERE thread_id = ?", (src_id,))

            merged_notes = "\n\n".join(note_parts)
            cursor.execute(
                """
                UPDATE tracked_threads
                SET conversation_ids = ?, subject_patterns = ?, outlook_folders = ?,
                    linked_ais = ?, linked_wis = ?, linked_tdocs = ?, linked_sub_items = ?,
                    user_status = ?, user_notes = ?, user_tags = ?
                WHERE thread_id = ?
                """,
                (
                    json.dumps(sorted(cids)),
                    json.dumps(sorted(patterns)),
                    json.dumps(sorted(folders)),
                    json.dumps(sorted(ais)),
                    json.dumps(sorted(wis)),
                    json.dumps(sorted(tdocs)),
                    json.dumps(sorted(sub_items)),
                    status or "⚪ Neutral",
                    merged_notes,
                    json.dumps(tags),
                    target_thread_id,
                ),
            )
            self._recompute_thread_stats(cursor, target_thread_id)
            conn.commit()
            from modules.emails.core.email_deduplicator import deduplicate_tracked_thread_messages
            removed_msgs, _ = deduplicate_tracked_thread_messages(conn)
            if removed_msgs:
                self._recompute_thread_stats(cursor, target_thread_id)
                conn.commit()

    def delete_thread(self, thread_id: str):
        with self._get_connection() as conn:
            conn.execute("DELETE FROM tracked_threads WHERE thread_id = ?", (thread_id,))
            conn.commit()
