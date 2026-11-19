# --- File: src/modules/emails/core/general_email_db.py ---
import contextlib
import sqlite3
from pathlib import Path
from typing import Dict, List, Set


class GeneralEmailDatabase:
    def __init__(self, db_path: Path):
        self.db_path = Path(db_path)
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        self._init_db()

    @contextlib.contextmanager
    def _get_connection(self):
        conn = sqlite3.connect(self.db_path, timeout=30.0, check_same_thread=False)
        try:
            conn.execute("PRAGMA journal_mode=WAL;")
            conn.execute("PRAGMA busy_timeout=30000;")
            with conn:
                yield conn
        finally:
            conn.close()

    def _init_db(self):
        with self._get_connection() as conn:
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
                    is_ignored INTEGER DEFAULT 0,
                    flag_status INTEGER DEFAULT 0
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

            # Schema migration: add is_ignored and flag_status if existing DB lacks them
            cursor.execute("PRAGMA table_info(general_emails)")
            columns = [info[1] for info in cursor.fetchall()]
            if 'is_ignored' not in columns:
                cursor.execute("ALTER TABLE general_emails ADD COLUMN is_ignored INTEGER DEFAULT 0")
            if 'flag_status' not in columns:
                cursor.execute("ALTER TABLE general_emails ADD COLUMN flag_status INTEGER DEFAULT 0")

            cursor.execute("CREATE INDEX IF NOT EXISTS idx_gen_tdoc ON general_email_tdoc_matches(tdoc_id)")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_gen_email_id ON general_email_tdoc_matches(email_id)")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_gen_read ON general_emails(is_read)")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_gen_ignored ON general_emails(is_ignored)")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_gen_flagged ON general_emails(flag_status)")
            conn.commit()

    def save_emails_batch(self, emails_data: List[dict], matches_data: List[dict]):
        if not emails_data:
            return
        with self._get_connection() as conn:
            cursor = conn.cursor()
            email_tuples = [
                (
                    e['id'], e.get('folder_tag', ''), e.get('subject', ''),
                    e.get('sender_name', ''), e.get('sender_email', ''),
                    e.get('company', ''), e.get('date_received', ''),
                    e.get('body_text', ''), e.get('is_read', 0),
                    e.get('is_ignored', 0), 2 if int(e.get('flag_status', 0) or 0) == 2 else 0
                )
                for e in emails_data
            ]
            # Preserve existing is_ignored and local read marks while syncing read/flag updates from Outlook
            cursor.executemany("""
                INSERT INTO general_emails (id, folder_tag, subject, sender_name, sender_email, company, date_received, body_text, is_read, is_ignored, flag_status)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(id) DO UPDATE SET
                    folder_tag = excluded.folder_tag,
                    subject = excluded.subject,
                    sender_name = excluded.sender_name,
                    sender_email = excluded.sender_email,
                    company = excluded.company,
                    date_received = excluded.date_received,
                    body_text = excluded.body_text,
                    is_read = CASE WHEN excluded.is_read = 1 THEN 1 ELSE general_emails.is_read END,
                    flag_status = CASE WHEN excluded.flag_status != 0 THEN excluded.flag_status ELSE general_emails.flag_status END
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

    def get_email_counts_per_tdoc(self) -> Dict[str, Dict]:
        """
        Returns a dictionary mapping tdoc_id to email statistics excluding ignored emails.
        Stores both scalar counts and sets of unique email IDs for exact family deduplication.
        """
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                           SELECT m.tdoc_id, e.id, e.is_read, COALESCE(e.flag_status, 0)
                           FROM general_email_tdoc_matches m
                                    JOIN general_emails e ON m.email_id = e.id
                           WHERE e.is_ignored = 0
                           """)
            raw_stats = {}
            for tdoc_id, email_id, is_read, flag_status in cursor.fetchall():
                tdoc_key = tdoc_id.upper()
                if tdoc_key not in raw_stats:
                    raw_stats[tdoc_key] = {"all_ids": set(), "unread_ids": set(), "flagged_ids": set()}
                raw_stats[tdoc_key]["all_ids"].add(email_id)
                if is_read == 0:
                    raw_stats[tdoc_key]["unread_ids"].add(email_id)
                if int(flag_status or 0) == 2:
                    raw_stats[tdoc_key]["flagged_ids"].add(email_id)

            return {
                tdoc: {
                    "total": len(data["all_ids"]),
                    "unread": len(data["unread_ids"]),
                    "flagged": len(data["flagged_ids"]),
                    "all_ids": data["all_ids"],
                    "unread_ids": data["unread_ids"],
                    "flagged_ids": data["flagged_ids"],
                }
                for tdoc, data in raw_stats.items()
            }

    def get_emails_for_tdocs(self, tdoc_ids: Set[str], show_ignored: bool = False) -> List[dict]:
        """
        Retrieves unique emails matching any of the specified TDocs.
        Aggregates across multiple family matches to prevent duplicate rows in dialogs.
        """
        if not tdoc_ids:
            return []
        placeholders = ",".join(["?"] * len(tdoc_ids))
        ignored_clause = "" if show_ignored else "AND e.is_ignored = 0"
        with self._get_connection() as conn:
            conn.row_factory = sqlite3.Row
            cursor = conn.cursor()
            cursor.execute(f"""
                SELECT 
                    e.*,
                    GROUP_CONCAT(DISTINCT m.tdoc_id) AS matched_tdoc,
                    MAX(m.rev_matched) AS rev_matched,
                    CASE 
                        WHEN SUM(CASE WHEN m.match_location = 'Subject' THEN 1 ELSE 0 END) > 0 THEN 'Subject'
                        WHEN SUM(CASE WHEN m.match_location = 'Body' THEN 1 ELSE 0 END) > 0 THEN 'Body'
                        ELSE 'Quoted'
                    END AS match_location
                FROM general_emails e
                JOIN general_email_tdoc_matches m ON e.id = m.email_id
                WHERE m.tdoc_id IN ({placeholders}) {ignored_clause}
                GROUP BY e.id
                ORDER BY e.date_received DESC
            """, [t.upper() for t in tdoc_ids])
            return [dict(row) for row in cursor.fetchall()]

    def get_all_emails_with_tdocs(self, show_ignored: bool = False) -> List[dict]:
        """Retrieves all indexed meeting emails annotated with their matched TDoc(s) for the unified inbox."""
        ignored_clause = "" if show_ignored else "WHERE e.is_ignored = 0"
        with self._get_connection() as conn:
            conn.row_factory = sqlite3.Row
            return [
                dict(row)
                for row in conn.execute(
                    f"""
                    SELECT
                        e.*,
                        COALESCE(GROUP_CONCAT(DISTINCT m.tdoc_id), '') AS matched_tdocs
                    FROM general_emails e
                    LEFT JOIN general_email_tdoc_matches m ON e.id = m.email_id
                    {ignored_clause}
                    GROUP BY e.id
                    ORDER BY e.date_received DESC
                    """
                ).fetchall()
            ]

    def get_tdoc_match_locations_by_email_id(self) -> Dict[str, Set[str]]:
        """Returns {email_id: {match_location, ...}} for non-ignored indexed emails."""
        with self._get_connection() as conn:
            out: Dict[str, Set[str]] = {}
            for eid, loc in conn.execute(
                """
                SELECT m.email_id, m.match_location
                FROM general_email_tdoc_matches m
                JOIN general_emails e ON e.id = m.email_id
                WHERE e.is_ignored = 0
                """
            ).fetchall():
                if eid:
                    out.setdefault(str(eid), set()).add(str(loc or "Body"))
            return out

    @staticmethod
    def _sync_outlook_read_status(to_sync: List[str], is_read: bool) -> int:
        if not to_sync:
            return 0
        from modules.emails.core.outlook_client import OutlookClient
        return OutlookClient.set_messages_read_status(to_sync, is_read=bool(is_read))

    @staticmethod
    def _sync_outlook_flag_status(to_sync: List[str], is_flagged: bool) -> int:
        if not to_sync:
            return 0
        from modules.emails.core.outlook_client import OutlookClient
        return OutlookClient.set_messages_flag_status(to_sync, is_flagged=bool(is_flagged))

    def set_emails_flag_status(self, email_ids: List[str], is_flagged: bool, sync_outlook: bool = True) -> int:
        if not email_ids:
            return 0
        target_val = 2 if is_flagged else 0
        unique_ids = list(dict.fromkeys(str(eid) for eid in email_ids if eid))
        if not unique_ids:
            return 0
        placeholders = ",".join(["?"] * len(unique_ids))
        with self._get_connection() as conn:
            to_sync = [
                row[0]
                for row in conn.execute(
                    f"SELECT id FROM general_emails WHERE id IN ({placeholders}) AND COALESCE(flag_status, 0) != ?",
                    unique_ids + [target_val],
                ).fetchall()
            ]
            conn.executemany(
                "UPDATE general_emails SET flag_status = ? WHERE id = ?",
                [(target_val, eid) for eid in unique_ids],
            )
            conn.commit()
        return self._sync_outlook_flag_status(to_sync, is_flagged) if sync_outlook else 0

    def set_emails_read_status(self, email_ids: List[str], is_read: bool, sync_outlook: bool = True) -> int:
        if not email_ids:
            return 0
        target_val = 1 if is_read else 0
        unique_ids = list(dict.fromkeys(str(eid) for eid in email_ids if eid))
        if not unique_ids:
            return 0
        placeholders = ",".join(["?"] * len(unique_ids))
        with self._get_connection() as conn:
            to_sync = [
                row[0]
                for row in conn.execute(
                    f"SELECT id FROM general_emails WHERE id IN ({placeholders}) AND is_read != ?",
                    unique_ids + [target_val],
                ).fetchall()
            ]
            conn.executemany(
                "UPDATE general_emails SET is_read = ? WHERE id = ?",
                [(target_val, eid) for eid in unique_ids],
            )
            conn.commit()
        return self._sync_outlook_read_status(to_sync, is_read) if sync_outlook else 0

    def set_emails_ignored_status(self, email_ids: List[str], is_ignored: bool):
        if not email_ids:
            return
        with self._get_connection() as conn:
            conn.executemany("UPDATE general_emails SET is_ignored = ? WHERE id = ?",
                             [(1 if is_ignored else 0, eid) for eid in email_ids])
            conn.commit()

    def delete_emails(self, email_ids: List[str]):
        if not email_ids:
            return
        with self._get_connection() as conn:
            conn.executemany("DELETE FROM general_email_tdoc_matches WHERE email_id = ?", [(eid,) for eid in email_ids])
            conn.executemany("DELETE FROM general_emails WHERE id = ?", [(eid,) for eid in email_ids])
            conn.commit()

    def set_tdocs_read_status(self, tdoc_ids: Set[str], is_read: bool, sync_outlook: bool = True) -> int:
        if not tdoc_ids:
            return 0
        target_val = 1 if is_read else 0
        tdoc_list = list(tdoc_ids)
        placeholders = ",".join(["?"] * len(tdoc_list))
        with self._get_connection() as conn:
            to_sync = [
                row[0]
                for row in conn.execute(
                    f"""
                    SELECT DISTINCT e.id
                    FROM general_emails e
                    JOIN general_email_tdoc_matches m ON e.id = m.email_id
                    WHERE m.tdoc_id IN ({placeholders}) AND e.is_read != ?
                    """,
                    tdoc_list + [target_val],
                ).fetchall()
            ]
            conn.execute(f"""
                UPDATE general_emails 
                SET is_read = ?
                WHERE id IN (
                    SELECT email_id FROM general_email_tdoc_matches WHERE tdoc_id IN ({placeholders})
                )
            """, [target_val] + tdoc_list)
            conn.commit()
        return self._sync_outlook_read_status(to_sync, is_read) if sync_outlook else 0

    def mark_all_read(self, sync_outlook: bool = True) -> int:
        with self._get_connection() as conn:
            to_sync = [
                row[0]
                for row in conn.execute(
                    "SELECT id FROM general_emails WHERE is_ignored = 0 AND is_read = 0"
                ).fetchall()
            ]
            conn.execute("UPDATE general_emails SET is_read = 1 WHERE is_ignored = 0")
            conn.commit()
        return self._sync_outlook_read_status(to_sync, True) if sync_outlook else 0

    def wipe_generic_emails(self):
        """High-speed purge of generic tables only; preserves SA2 eMeeting tables."""
        with self._get_connection() as conn:
            conn.execute("DELETE FROM general_email_tdoc_matches")
            conn.execute("DELETE FROM general_emails")
            conn.execute("PRAGMA optimize")
            conn.commit()