# --- File: src/modules/emails/core/general_email_db.py ---
import contextlib
import json
import sqlite3
from pathlib import Path
from typing import Any, Dict, List, Set


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
                    clean_body TEXT DEFAULT '',
                    attachments TEXT DEFAULT '[]',
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
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS sync_metadata (
                    key TEXT PRIMARY KEY,
                    value TEXT
                )
            """)

            # Schema migration: add missing columns on existing DBs
            cursor.execute("PRAGMA table_info(general_emails)")
            columns = [info[1] for info in cursor.fetchall()]
            if 'is_ignored' not in columns:
                cursor.execute("ALTER TABLE general_emails ADD COLUMN is_ignored INTEGER DEFAULT 0")
            if 'flag_status' not in columns:
                cursor.execute("ALTER TABLE general_emails ADD COLUMN flag_status INTEGER DEFAULT 0")
            if 'clean_body' not in columns:
                cursor.execute("ALTER TABLE general_emails ADD COLUMN clean_body TEXT DEFAULT ''")
            if 'attachments' not in columns:
                cursor.execute("ALTER TABLE general_emails ADD COLUMN attachments TEXT DEFAULT '[]'")

            cursor.execute("CREATE INDEX IF NOT EXISTS idx_gen_tdoc ON general_email_tdoc_matches(tdoc_id)")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_gen_email_id ON general_email_tdoc_matches(email_id)")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_gen_read ON general_emails(is_read)")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_gen_ignored ON general_emails(is_ignored)")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_gen_flagged ON general_emails(flag_status)")
            conn.commit()
            from modules.emails.core.legacy_email_migrator import migrate_legacy_emeeting_emails
            from modules.emails.core.email_deduplicator import deduplicate_general_emails
            migrate_legacy_emeeting_emails(conn)
            deduplicate_general_emails(conn)

    def get_cached_email_states(self) -> Dict[str, Dict[str, Any]]:
        """Returns {email_id: {'is_read': int, 'flag_status': int, 'has_body': bool}} for cached emails."""
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(
                "SELECT id, is_read, COALESCE(flag_status, 0), body_text FROM general_emails"
            )
            return {
                row[0]: {
                    "is_read": int(row[1] or 0),
                    "flag_status": 2 if int(row[2] or 0) == 2 else 0,
                    "has_body": bool(str(row[3] or "").strip()),
                }
                for row in cursor.fetchall()
            }

    def update_email_flags_batch(self, states: List[tuple]) -> int:
        """Applies lightweight read/flag status updates [(email_id, is_read, flag_status)] for cached emails."""
        if not states:
            return 0
        updated = 0
        with self._get_connection() as conn:
            cursor = conn.cursor()
            for eid, is_read, flag_status in states:
                if not eid or str(eid).startswith("MSG-"):
                    continue
                ir = 1 if int(is_read or 0) else 0
                fs = 2 if int(flag_status or 0) == 2 else 0
                cursor.execute(
                    """
                    UPDATE general_emails
                    SET is_read = ?, flag_status = ?
                    WHERE id = ? AND (is_read != ? OR COALESCE(flag_status, 0) != ?)
                    """,
                    (ir, fs, eid, ir, fs),
                )
                updated += cursor.rowcount
            conn.commit()
        return updated

    def save_emails_batch(self, emails_data: List[dict], matches_data: List[dict], sync_unread: bool = False):
        if not emails_data:
            return
        read_expr = (
            "CASE WHEN excluded.id LIKE 'MSG-%' THEN general_emails.is_read ELSE excluded.is_read END"
            if sync_unread
            else "CASE WHEN excluded.is_read = 1 THEN 1 ELSE general_emails.is_read END"
        )
        with self._get_connection() as conn:
            cursor = conn.cursor()
            email_tuples = [
                (
                    e['id'], e.get('folder_tag', ''), e.get('subject', ''),
                    e.get('sender_name', ''), e.get('sender_email', ''),
                    e.get('company', ''), e.get('date_received', ''),
                    e.get('body_text', ''), e.get('clean_body', ''),
                    json.dumps(e.get('attachments', []) or []),
                    e.get('is_read', 0),
                    e.get('is_ignored', 0), 2 if int(e.get('flag_status', 0) or 0) == 2 else 0
                )
                for e in emails_data
            ]
            # Preserve existing is_ignored; sync Outlook read/unread (when sync_unread=True) and flag_status both ways
            cursor.executemany(f"""
                INSERT INTO general_emails (
                    id, folder_tag, subject, sender_name, sender_email, company, date_received,
                    body_text, clean_body, attachments, is_read, is_ignored, flag_status
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(id) DO UPDATE SET
                    folder_tag = excluded.folder_tag,
                    subject = excluded.subject,
                    sender_name = excluded.sender_name,
                    sender_email = excluded.sender_email,
                    company = excluded.company,
                    date_received = excluded.date_received,
                    body_text = CASE WHEN excluded.body_text != '' THEN excluded.body_text ELSE general_emails.body_text END,
                    clean_body = CASE WHEN excluded.clean_body != '' THEN excluded.clean_body ELSE general_emails.clean_body END,
                    attachments = CASE WHEN excluded.attachments != '[]' THEN excluded.attachments ELSE general_emails.attachments END,
                    is_read = {read_expr},
                    flag_status = CASE WHEN excluded.id LIKE 'MSG-%' THEN general_emails.flag_status ELSE excluded.flag_status END
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
            from modules.emails.core.email_deduplicator import deduplicate_general_emails
            deduplicate_general_emails(conn)

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

    @staticmethod
    def _decode_row(row: sqlite3.Row) -> dict:
        d = dict(row)
        raw_atts = d.get("attachments")
        if isinstance(raw_atts, str):
            try:
                parsed = json.loads(raw_atts)
                d["attachments"] = parsed if isinstance(parsed, list) else []
            except Exception:
                d["attachments"] = []
        elif not isinstance(raw_atts, list):
            d["attachments"] = []
        d.setdefault("date_sent", d.get("date_received", ""))
        return d

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
            return [self._decode_row(row) for row in cursor.fetchall()]

    def get_all_emails_with_tdocs(self, show_ignored: bool = False) -> List[dict]:
        """Retrieves all indexed meeting emails annotated with their matched TDoc(s) for the unified inbox."""
        ignored_clause = "" if show_ignored else "WHERE e.is_ignored = 0"
        with self._get_connection() as conn:
            conn.row_factory = sqlite3.Row
            return [
                self._decode_row(row)
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

    def update_reloaded_message(self, email_id: str, parsed: dict):
        """Updates an existing general_emails record with richly parsed Outlook fields."""
        eid = str(email_id or "").strip()
        if not eid or not parsed:
            return
        with self._get_connection() as conn:
            conn.execute(
                """
                UPDATE general_emails
                SET subject = ?,
                    sender_name = COALESCE(NULLIF(?, ''), sender_name),
                    sender_email = COALESCE(NULLIF(?, ''), sender_email),
                    company = COALESCE(NULLIF(?, ''), company),
                    body_text = ?,
                    clean_body = ?,
                    attachments = ?,
                    is_read = ?,
                    flag_status = ?
                WHERE id = ?
                """,
                (
                    parsed.get("subject", ""),
                    parsed.get("sender_name", ""),
                    parsed.get("sender_email", ""),
                    parsed.get("company", ""),
                    parsed.get("body_text", ""),
                    parsed.get("clean_body", ""),
                    json.dumps(parsed.get("attachments", []) or []),
                    int(parsed.get("is_read", 0) or 0),
                    2 if int(parsed.get("flag_status", 0) or 0) == 2 else 0,
                    eid,
                ),
            )
            conn.commit()

    def get_sync_meta(self, key: str, default: str = "") -> str:
        """Returns a persisted sync metadata value by key."""
        clean_key = str(key or "").strip()
        if not clean_key:
            return default
        with self._get_connection() as conn:
            row = conn.execute("SELECT value FROM sync_metadata WHERE key = ?", (clean_key,)).fetchone()
            return str(row[0]) if row and row[0] is not None else default

    def set_sync_meta(self, key: str, value: str) -> None:
        """Persists a sync metadata key/value pair."""
        clean_key = str(key or "").strip()
        if not clean_key:
            return
        with self._get_connection() as conn:
            conn.execute(
                """
                INSERT INTO sync_metadata (key, value) VALUES (?, ?)
                ON CONFLICT(key) DO UPDATE SET value = excluded.value
                """,
                (clean_key, str(value if value is not None else "")),
            )
            conn.commit()

    def delete_sync_meta(self, key: str) -> None:
        """Removes a sync metadata entry by key."""
        clean_key = str(key or "").strip()
        if not clean_key:
            return
        with self._get_connection() as conn:
            conn.execute("DELETE FROM sync_metadata WHERE key = ?", (clean_key,))
            conn.commit()

    def wipe_generic_emails(self):
        """High-speed purge of generic tables only; preserves SA2 eMeeting tables."""
        with self._get_connection() as conn:
            conn.execute("DELETE FROM general_email_tdoc_matches")
            conn.execute("DELETE FROM general_emails")
            conn.execute("DELETE FROM sync_metadata WHERE key = 'last_sync_utc'")
            conn.execute("PRAGMA optimize")
            conn.commit()


def reload_general_email_from_outlook(db: GeneralEmailDatabase, email_id: str) -> bool:
    """Re-fetches an indexed TDoc email from Outlook COM to populate HTML tables, attachments, and read/flag status."""
    eid = str(email_id or "").strip()
    if not eid or eid.startswith("MSG-"):
        return False
    from modules.emails.core.outlook_client import OutlookClient
    from modules.emails.core.tracked_threads_service import _mail_item_to_dict

    ns = OutlookClient.get_namespace()
    if not ns:
        return False
    try:
        item = ns.GetItemFromID(eid)
        if not item:
            return False
        parsed = _mail_item_to_dict(item)
        db.update_reloaded_message(eid, parsed)
        return True
    except Exception:
        return False


def reload_tdoc_emails_from_outlook(db: GeneralEmailDatabase, email_ids: List[str]) -> int:
    """Batch reloads multiple cached TDoc emails from Outlook COM."""
    updated = 0
    for eid in (email_ids or []):
        if reload_general_email_from_outlook(db, eid):
            updated += 1
    return updated