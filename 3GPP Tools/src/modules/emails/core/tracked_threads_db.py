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
                    attachments TEXT DEFAULT '[]',
                    FOREIGN KEY(thread_id) REFERENCES tracked_threads(thread_id) ON DELETE CASCADE
                )
            """)
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_tt_meeting ON tracked_threads(meeting_id)")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_tt_norm_subj ON tracked_threads(normalized_subject)")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_ttm_thread ON tracked_thread_messages(thread_id)")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_ttm_date ON tracked_thread_messages(date_sent)")

            # Schema migration: add linked_sub_items and user annotations for existing DBs
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
            conn.commit()

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
                    outlook_folders, meeting_id, wg, linked_ais, linked_wis, linked_tdocs,
                    unread_count, message_count, last_message_date, last_synced_at, created_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 0, 0, '', ?, ?)
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
                    now,
                    now,
                ),
            )
            conn.commit()
            return thread_id

    def upsert_messages(self, thread_id: str, messages: List[Dict[str, Any]]) -> int:
        """Inserts new messages for a thread without overwriting existing is_read states. Returns count of newly added messages."""
        if not messages:
            return 0
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
                    """
                    INSERT INTO tracked_thread_messages (
                        id, thread_id, subject, sender_name, sender_email, company,
                        date_sent, body_text, clean_body, message_id_header,
                        in_reply_to, conversation_id, is_read, attachments
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    ON CONFLICT(id) DO UPDATE SET
                        subject = excluded.subject,
                        body_text = excluded.body_text,
                        clean_body = excluded.clean_body,
                        attachments = excluded.attachments,
                        is_read = CASE WHEN excluded.is_read = 1 THEN 1 ELSE tracked_thread_messages.is_read END
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
                        json.dumps(msg.get("attachments", [])),
                    ),
                )
            self._recompute_thread_stats(cursor, thread_id)
            conn.commit()
        return new_count

    def _recompute_thread_stats(self, cursor: sqlite3.Cursor, thread_id: str):
        cursor.execute(
            """
            SELECT COUNT(*) as total,
                   SUM(CASE WHEN is_read = 0 THEN 1 ELSE 0 END) as unread,
                   MAX(date_sent) as latest
            FROM tracked_thread_messages
            WHERE thread_id = ?
            """,
            (thread_id,),
        )
        row = cursor.fetchone()
        total = row["total"] or 0
        unread = row["unread"] or 0
        latest = row["latest"] or ""
        now = datetime.now().isoformat(timespec="seconds")
        cursor.execute(
            """
            UPDATE tracked_threads
            SET message_count = ?, unread_count = ?, last_message_date = ?, last_synced_at = ?
            WHERE thread_id = ?
            """,
            (total, unread, latest, now, thread_id),
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
    ):
        """Persists personal stance, comments, and tags onto a tracked email thread."""
        with self._get_connection() as conn:
            cursor = conn.cursor()
            if user_status is not None:
                cursor.execute("UPDATE tracked_threads SET user_status = ? WHERE thread_id = ?", (user_status, thread_id))
            if user_notes is not None:
                cursor.execute("UPDATE tracked_threads SET user_notes = ? WHERE thread_id = ?", (user_notes, thread_id))
            if user_tags is not None:
                cursor.execute("UPDATE tracked_threads SET user_tags = ? WHERE thread_id = ?", (json.dumps(user_tags), thread_id))
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
            ais = set(target["linked_ais"])
            wis = set(target["linked_wis"])
            tdocs = set(target["linked_tdocs"])
            sub_items = set(target.get("linked_sub_items") or [])

            for src_id in source_thread_ids:
                if src_id == target_thread_id:
                    continue
                src = self.get_thread(src_id)
                if not src:
                    continue
                cids.update(src["conversation_ids"])
                patterns.update(src["subject_patterns"])
                ais.update(src["linked_ais"])
                wis.update(src["linked_wis"])
                tdocs.update(src["linked_tdocs"])
                sub_items.update(src.get("linked_sub_items") or [])
                cursor.execute(
                    "UPDATE tracked_thread_messages SET thread_id = ? WHERE thread_id = ?",
                    (target_thread_id, src_id),
                )
                cursor.execute("DELETE FROM tracked_threads WHERE thread_id = ?", (src_id,))

            cursor.execute(
                """
                UPDATE tracked_threads
                SET conversation_ids = ?, subject_patterns = ?, linked_ais = ?, linked_wis = ?, linked_tdocs = ?, linked_sub_items = ?
                WHERE thread_id = ?
                """,
                (
                    json.dumps(sorted(cids)),
                    json.dumps(sorted(patterns)),
                    json.dumps(sorted(ais)),
                    json.dumps(sorted(wis)),
                    json.dumps(sorted(tdocs)),
                    json.dumps(sorted(sub_items)),
                    target_thread_id,
                ),
            )
            self._recompute_thread_stats(cursor, target_thread_id)
            conn.commit()

    def delete_thread(self, thread_id: str):
        with self._get_connection() as conn:
            conn.execute("DELETE FROM tracked_threads WHERE thread_id = ?", (thread_id,))
            conn.commit()
