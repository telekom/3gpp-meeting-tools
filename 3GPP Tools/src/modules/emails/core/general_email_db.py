# --- File: src/modules/emails/core/general_email_db.py ---
import sqlite3
from pathlib import Path
from typing import Dict, List, Set


class GeneralEmailDatabase:
    def __init__(self, db_path: Path):
        self.db_path = Path(db_path)
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        self._init_db()

    def _init_db(self):
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS general_emails (
                    id TEXT PRIMARY KEY,
                    folder_tag TEXT,
                    subject TEXT,
                    sender_name TEXT,
                    sender_email TEXT,
                    company TEXT,
                    date_received TEXT,
                    body_text TEXT,
                    is_read INTEGER DEFAULT 0,
                    is_ignored INTEGER DEFAULT 0
                )
            """)
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS general_email_tdoc_matches (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    email_id TEXT,
                    tdoc_id TEXT,
                    rev_matched TEXT,
                    match_location TEXT,
                    FOREIGN KEY(email_id) REFERENCES general_emails(id) ON DELETE CASCADE
                )
            """)

            # Schema migration: add is_ignored if existing DB lacks it
            cursor.execute("PRAGMA table_info(general_emails)")
            columns = [info[1] for info in cursor.fetchall()]
            if 'is_ignored' not in columns:
                cursor.execute("ALTER TABLE general_emails ADD COLUMN is_ignored INTEGER DEFAULT 0")

            cursor.execute("CREATE INDEX IF NOT EXISTS idx_gen_tdoc ON general_email_tdoc_matches(tdoc_id)")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_gen_email_id ON general_email_tdoc_matches(email_id)")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_gen_read ON general_emails(is_read)")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_gen_ignored ON general_emails(is_ignored)")
            conn.commit()

    def save_emails_batch(self, emails_data: List[dict], matches_data: List[dict]):
        if not emails_data:
            return
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            email_tuples = [
                (
                    e['id'], e.get('folder_tag', ''), e.get('subject', ''),
                    e.get('sender_name', ''), e.get('sender_email', ''),
                    e.get('company', ''), e.get('date_received', ''),
                    e.get('body_text', ''), e.get('is_read', 0),
                    e.get('is_ignored', 0)
                )
                for e in emails_data
            ]
            # Preserve existing is_ignored and is_read states during re-sync
            cursor.executemany("""
                INSERT INTO general_emails (id, folder_tag, subject, sender_name, sender_email, company, date_received, body_text, is_read, is_ignored)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(id) DO UPDATE SET
                    folder_tag = excluded.folder_tag,
                    subject = excluded.subject,
                    sender_name = excluded.sender_name,
                    sender_email = excluded.sender_email,
                    company = excluded.company,
                    date_received = excluded.date_received,
                    body_text = excluded.body_text
            """, email_tuples)

            email_ids = [e['id'] for e in emails_data]
            cursor.executemany("DELETE FROM general_email_tdoc_matches WHERE email_id = ?", [(eid,) for eid in email_ids])

            match_tuples = [
                (m['email_id'], m['tdoc_id'], m.get('rev_matched', ''), m.get('match_location', 'Body'))
                for m in matches_data
            ]
            cursor.executemany("""
                INSERT INTO general_email_tdoc_matches (email_id, tdoc_id, rev_matched, match_location)
                VALUES (?, ?, ?, ?)
            """, match_tuples)
            conn.commit()

    def get_email_counts_per_tdoc(self) -> Dict[str, Dict[str, int]]:
        """Returns {tdoc_id: {'total': int, 'unread': int}} excluding ignored emails."""
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute("""
                SELECT m.tdoc_id,
                       COUNT(DISTINCT e.id) AS total_count,
                       SUM(CASE WHEN e.is_read = 0 THEN 1 ELSE 0 END) AS unread_count
                FROM general_email_tdoc_matches m
                JOIN general_emails e ON m.email_id = e.id
                WHERE e.is_ignored = 0
                GROUP BY m.tdoc_id
            """)
            return {
                row[0]: {'total': row[1], 'unread': row[2] or 0}
                for row in cursor.fetchall()
            }

    def get_emails_for_tdocs(self, tdoc_ids: Set[str], show_ignored: bool = False) -> List[dict]:
        if not tdoc_ids:
            return []
        placeholders = ",".join(["?"] * len(tdoc_ids))
        ignored_clause = "" if show_ignored else "AND e.is_ignored = 0"
        with sqlite3.connect(self.db_path) as conn:
            conn.row_factory = sqlite3.Row
            cursor = conn.cursor()
            cursor.execute(f"""
                SELECT DISTINCT e.*, m.tdoc_id as matched_tdoc, m.rev_matched, m.match_location
                FROM general_emails e
                JOIN general_email_tdoc_matches m ON e.id = m.email_id
                WHERE m.tdoc_id IN ({placeholders}) {ignored_clause}
                ORDER BY e.date_received DESC
            """, list(tdoc_ids))
            return [dict(row) for row in cursor.fetchall()]

    def set_emails_read_status(self, email_ids: List[str], is_read: bool):
        if not email_ids:
            return
        with sqlite3.connect(self.db_path) as conn:
            conn.executemany("UPDATE general_emails SET is_read = ? WHERE id = ?",
                             [(1 if is_read else 0, eid) for eid in email_ids])
            conn.commit()

    def set_emails_ignored_status(self, email_ids: List[str], is_ignored: bool):
        if not email_ids:
            return
        with sqlite3.connect(self.db_path) as conn:
            conn.executemany("UPDATE general_emails SET is_ignored = ? WHERE id = ?",
                             [(1 if is_ignored else 0, eid) for eid in email_ids])
            conn.commit()

    def delete_emails(self, email_ids: List[str]):
        if not email_ids:
            return
        with sqlite3.connect(self.db_path) as conn:
            conn.executemany("DELETE FROM general_email_tdoc_matches WHERE email_id = ?", [(eid,) for eid in email_ids])
            conn.executemany("DELETE FROM general_emails WHERE id = ?", [(eid,) for eid in email_ids])
            conn.commit()

    def set_tdocs_read_status(self, tdoc_ids: Set[str], is_read: bool):
        if not tdoc_ids:
            return
        placeholders = ",".join(["?"] * len(tdoc_ids))
        with sqlite3.connect(self.db_path) as conn:
            conn.execute(f"""
                UPDATE general_emails 
                SET is_read = ?
                WHERE id IN (
                    SELECT email_id FROM general_email_tdoc_matches WHERE tdoc_id IN ({placeholders})
                )
            """, [1 if is_read else 0] + list(tdoc_ids))
            conn.commit()

    def mark_all_read(self):
        with sqlite3.connect(self.db_path) as conn:
            conn.execute("UPDATE general_emails SET is_read = 1 WHERE is_ignored = 0")
            conn.commit()

    def wipe_generic_emails(self):
        """High-speed purge of generic tables only; preserves SA2 eMeeting tables."""
        with sqlite3.connect(self.db_path) as conn:
            conn.execute("DELETE FROM general_email_tdoc_matches")
            conn.execute("DELETE FROM general_emails")
            conn.execute("PRAGMA optimize")
            conn.commit()