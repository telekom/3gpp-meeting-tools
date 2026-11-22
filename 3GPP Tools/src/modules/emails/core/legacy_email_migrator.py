# --- File: src/modules/emails/core/legacy_email_migrator.py ---
"""
Idempotent migration from legacy eMeeting `emails` table (inside `<meeting_dir>/Agenda/emails.db`)
into `general_emails` and `general_email_tdoc_matches`.
"""

from __future__ import annotations

import logging
import sqlite3
from typing import List, Set

logger = logging.getLogger(__name__)


def _compose_migrated_body(short_text: str, free_text: str) -> str:
    """Reconstructs an email body preserving formal `<<START>>...<<END>>` e-Meeting comment blocks."""
    st = str(short_text or "").strip()
    ft = str(free_text or "").strip()
    if st and ft:
        return f"<<START>>\n{st}\n<<END>>\n\n{ft}"
    if st:
        return f"<<START>>\n{st}\n<<END>>"
    return ft


def migrate_legacy_emeeting_emails(conn: sqlite3.Connection) -> int:
    """
    Migrates any rows from the legacy `emails` table in `emails.db` that are not yet present
    in `general_emails` into `general_emails` + `general_email_tdoc_matches`.
    Never deletes or alters the legacy `emails` table, and never overwrites existing `general_emails` rows.
    Returns the number of migrated email rows.
    """
    try:
        cursor = conn.cursor()
        tables = {
            row[0]
            for row in cursor.execute(
                "SELECT name FROM sqlite_master WHERE type='table'"
            ).fetchall()
        }
        if "emails" not in tables or "general_emails" not in tables:
            return 0

        existing_ids: Set[str] = {
            str(r[0])
            for r in cursor.execute("SELECT id FROM general_emails").fetchall()
            if r[0]
        }

        conn.row_factory = sqlite3.Row
        legacy_rows = conn.execute("SELECT * FROM emails").fetchall()
        if not legacy_rows:
            return 0

        to_insert_emails: List[tuple] = []
        to_insert_matches: List[tuple] = []

        for row in legacy_rows:
            d = dict(row)
            eid = str(d.get("id") or "").strip()
            if not eid or eid in existing_ids:
                continue
            existing_ids.add(eid)

            tdoc_id = str(d.get("tdoc_id") or "").strip().upper()
            revs_raw = str(d.get("revisions_mentioned") or "").strip()
            rev_list = [r.strip().upper() for r in revs_raw.split(",") if r.strip()]
            first_rev = rev_list[0] if rev_list else ""

            body = _compose_migrated_body(
                d.get("short_text", ""),
                d.get("free_text", ""),
            )
            to_insert_emails.append((
                eid,
                "eMeeting",
                str(d.get("subject") or "").strip(),
                str(d.get("sender_name") or "").strip(),
                str(d.get("sender_email") or "").strip(),
                str(d.get("company") or "").strip(),
                str(d.get("date_received") or "").strip(),
                body,
                body,
                "[]",
                1,  # Historical migrated eMeeting emails default to read
                0,
                0,
            ))
            if tdoc_id:
                to_insert_matches.append((
                    eid,
                    tdoc_id,
                    first_rev,
                    "Subject",
                ))

        if not to_insert_emails:
            return 0

        cursor.executemany(
            """
            INSERT OR IGNORE INTO general_emails (
                id, folder_tag, subject, sender_name, sender_email, company, date_received,
                body_text, clean_body, attachments, is_read, is_ignored, flag_status
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            to_insert_emails,
        )
        if to_insert_matches:
            cursor.executemany(
                """
                INSERT INTO general_email_tdoc_matches (email_id, tdoc_id, rev_matched, match_location)
                VALUES (?, ?, ?, ?)
                """,
                to_insert_matches,
            )
        conn.commit()
        logger.info("Migrated %d legacy eMeeting email(s) into general_emails.", len(to_insert_emails))
        return len(to_insert_emails)
    except Exception as exc:
        logger.warning("Could not auto-migrate legacy eMeeting emails: %s", exc)
        return 0


def read_legacy_followed_ais_and_starred_tdocs(db_path) -> tuple[Set[str], Set[str]]:
    """
    Reads `followed_ais` and `starred_tdocs` from `Agenda/emails.db` if those legacy tables exist.
    Returns `(followed_ais, starred_tdocs)`.
    """
    followed_ais: Set[str] = set()
    starred_tdocs: Set[str] = set()
    try:
        conn = sqlite3.connect(db_path, timeout=10.0)
        try:
            tables = {
                r[0]
                for r in conn.execute("SELECT name FROM sqlite_master WHERE type='table'").fetchall()
            }
            if "followed_ais" in tables:
                followed_ais = {
                    str(r[0]).strip()
                    for r in conn.execute("SELECT agenda_item FROM followed_ais").fetchall()
                    if r[0] and str(r[0]).strip()
                }
            if "starred_tdocs" in tables:
                starred_tdocs = {
                    str(r[0]).strip().upper()
                    for r in conn.execute("SELECT tdoc_id FROM starred_tdocs").fetchall()
                    if r[0] and str(r[0]).strip()
                }
        finally:
            conn.close()
    except Exception:
        pass
    return followed_ais, starred_tdocs
