"""SQLite connection management, schema initialization, upserts, and wipe/delete helpers."""

from __future__ import annotations

import contextlib
import sqlite3
from typing import Any, Dict, Iterator, List


class SpecHistorySchemaAndWritesMixin:
    """Mixin providing thread-safe WAL connections, schema setup, history upserts, and deletion."""

    @contextlib.contextmanager
    def _get_connection(self) -> Iterator[sqlite3.Connection]:
        """Creates a thread-safe SQLite connection with WAL mode and foreign key constraints."""
        conn = sqlite3.connect(str(self.db_path), timeout=30.0, check_same_thread=False)
        try:
            conn.row_factory = sqlite3.Row
            conn.execute("PRAGMA foreign_keys = ON;")
            conn.execute("PRAGMA busy_timeout = 30000;")
            with conn:
                yield conn
        finally:
            conn.close()

    def _init_db(self):
        """Initializes the database schema with WAL mode and indexed lookup tables."""
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("PRAGMA journal_mode = WAL;")

            # 1. Master Change History Table
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS spec_change_history (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    spec_number TEXT NOT NULL,                  -- e.g. "23.501", "23.801-01"
                    version TEXT NOT NULL DEFAULT '',           -- Target version, e.g. "0.8.0", "16.20.0"
                    row_order INTEGER NOT NULL DEFAULT 0,       -- Preserves document order
                    date TEXT NOT NULL DEFAULT '',              -- Date (YYYY-MM or text)
                    meeting TEXT NOT NULL DEFAULT '',           -- Meeting name, e.g. "SA WG2#175-AHE"
                    source_tdoc TEXT NOT NULL DEFAULT '',       -- Explicit TDoc from table column
                    cr_number TEXT NOT NULL DEFAULT '',         -- CR number (e.g. "5357", or "" for p-CRs)
                    revision TEXT NOT NULL DEFAULT '',          -- Revision string (e.g. "2" or "")
                    category TEXT NOT NULL DEFAULT '',          -- 3GPP Category (F, B, C, A, etc.)
                    subject TEXT NOT NULL DEFAULT '',           -- Full subject / rapporteur commentary
                    old_version TEXT NOT NULL DEFAULT '',       -- Preceding version if noted
                    source_file TEXT NOT NULL DEFAULT '',       -- Cached file that supplied this entry
                    affected_clauses TEXT NOT NULL DEFAULT '',  -- Clauses modified by this CR/p-CR
                    extracted_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    UNIQUE(spec_number, version, meeting, source_tdoc, cr_number, revision, subject)
                );
            """)

            # 2. TDoc Cross-Reference Index Table
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS spec_tdoc_cross_ref (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    tdoc TEXT NOT NULL,                         -- Normalized base TDoc (e.g. "S2-2606265")
                    raw_tdoc TEXT NOT NULL,                     -- Exact token with revision (e.g. "S2-2606265r01")
                    spec_number TEXT NOT NULL,                  -- e.g. "23.801-01"
                    version TEXT NOT NULL DEFAULT '',           -- e.g. "0.8.0"
                    meeting TEXT NOT NULL DEFAULT '',           -- e.g. "SA WG2#175-AHE"
                    is_pcr INTEGER NOT NULL DEFAULT 0,          -- 1 if extracted from p-CR list; 0 if primary column TDoc
                    history_id INTEGER NOT NULL,                -- Reference to spec_change_history(id)
                    UNIQUE(history_id, raw_tdoc),
                    FOREIGN KEY(history_id) REFERENCES spec_change_history(id) ON DELETE CASCADE
                );
            """)

            # 3. Persistent Meeting Agenda TDoc Metadata Cache
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS tdoc_metadata_cache (
                    tdoc TEXT PRIMARY KEY,                      -- Uppercase token, e.g. "S2-2606363"
                    base_tdoc TEXT NOT NULL,                    -- Base token without revision
                    meeting TEXT NOT NULL DEFAULT '',
                    title TEXT NOT NULL DEFAULT '',
                    source TEXT NOT NULL DEFAULT '',
                    tdoc_type TEXT NOT NULL DEFAULT '',
                    status TEXT NOT NULL DEFAULT '',
                    agenda_item TEXT NOT NULL DEFAULT '',
                    ai_desc TEXT NOT NULL DEFAULT '',
                    abstract TEXT NOT NULL DEFAULT '',
                    work_item TEXT NOT NULL DEFAULT '',
                    revision_of TEXT NOT NULL DEFAULT '',
                    revised_to TEXT NOT NULL DEFAULT '',
                    remarks TEXT NOT NULL DEFAULT '',
                    affected_clauses TEXT NOT NULL DEFAULT '',  -- Clauses affected extracted from cover sheet
                    cached_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                );
            """)

            # Schema Migration: ensure affected_clauses exists in existing tables
            cursor.execute("PRAGMA table_info(spec_change_history);")
            cols_hist = [row[1] for row in cursor.fetchall()]
            if cols_hist and "affected_clauses" not in cols_hist:
                cursor.execute("ALTER TABLE spec_change_history ADD COLUMN affected_clauses TEXT NOT NULL DEFAULT '';")

            cursor.execute("PRAGMA table_info(tdoc_metadata_cache);")
            cols_tdoc = [row[1] for row in cursor.fetchall()]
            if cols_tdoc and "affected_clauses" not in cols_tdoc:
                cursor.execute("ALTER TABLE tdoc_metadata_cache ADD COLUMN affected_clauses TEXT NOT NULL DEFAULT '';")

            # Indexes for ultra-fast lookup
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_hist_spec_ver ON spec_change_history(spec_number, version);")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_hist_spec ON spec_change_history(spec_number);")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_tdoc_lookup ON spec_tdoc_cross_ref(tdoc);")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_raw_tdoc_lookup ON spec_tdoc_cross_ref(raw_tdoc);")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_tdoc_spec_ver ON spec_tdoc_cross_ref(spec_number, version);")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_tdoc_meta_base ON tdoc_metadata_cache(base_tdoc);")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_tdoc_meta_meeting ON tdoc_metadata_cache(meeting);")

    def upsert_history_records(self, records: List[Dict[str, Any]]) -> int:
        """
        Bulk upserts parsed change history records and populates the TDoc cross-reference table.
        Returns the number of history rows upserted.
        """
        if not records:
            return 0

        inserted_count = 0
        with self._get_connection() as conn:
            cursor = conn.cursor()

            for rec in records:
                spec_num = str(rec.get("spec_number", "")).strip()
                version = str(rec.get("version", "")).strip()
                row_order = int(rec.get("row_order", 0))
                date_val = str(rec.get("date", "")).strip()
                meeting = str(rec.get("meeting", "")).strip()
                source_tdoc = str(rec.get("source_tdoc", "")).strip()
                cr_number = str(rec.get("cr_number", "")).strip()
                revision = str(rec.get("revision", "")).strip()
                category = str(rec.get("category", "")).strip()
                subject = str(rec.get("subject", "")).strip()
                old_ver = str(rec.get("old_version", "")).strip()
                source_file = str(rec.get("source_file", "")).strip()
                affected_clauses = str(rec.get("affected_clauses", "")).strip()

                if not spec_num:
                    continue

                try:
                    cursor.execute("""
                        INSERT INTO spec_change_history (
                            spec_number, version, row_order, date, meeting,
                            source_tdoc, cr_number, revision, category,
                            subject, old_version, source_file, affected_clauses
                        )
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        ON CONFLICT(spec_number, version, meeting, source_tdoc, cr_number, revision, subject)
                        DO UPDATE SET
                            row_order=excluded.row_order,
                            date=excluded.date,
                            category=excluded.category,
                            old_version=excluded.old_version,
                            source_file=excluded.source_file,
                            affected_clauses=CASE WHEN excluded.affected_clauses != '' THEN excluded.affected_clauses ELSE spec_change_history.affected_clauses END,
                            extracted_at=CURRENT_TIMESTAMP
                        RETURNING id;
                    """, (
                        spec_num, version, row_order, date_val, meeting,
                        source_tdoc, cr_number, revision, category,
                        subject, old_ver, source_file, affected_clauses,
                    ))
                    res = cursor.fetchone()
                    history_id = res[0] if res else None
                except sqlite3.OperationalError:
                    cursor.execute("""
                        INSERT INTO spec_change_history (
                            spec_number, version, row_order, date, meeting,
                            source_tdoc, cr_number, revision, category,
                            subject, old_version, source_file, affected_clauses
                        )
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        ON CONFLICT(spec_number, version, meeting, source_tdoc, cr_number, revision, subject)
                        DO UPDATE SET
                            row_order=excluded.row_order,
                            date=excluded.date,
                            category=excluded.category,
                            old_version=excluded.old_version,
                            source_file=excluded.source_file,
                            affected_clauses=CASE WHEN excluded.affected_clauses != '' THEN excluded.affected_clauses ELSE spec_change_history.affected_clauses END,
                            extracted_at=CURRENT_TIMESTAMP;
                    """, (
                        spec_num, version, row_order, date_val, meeting,
                        source_tdoc, cr_number, revision, category,
                        subject, old_ver, source_file, affected_clauses,
                    ))
                    row = cursor.execute("""
                        SELECT id FROM spec_change_history
                        WHERE spec_number = ? AND version = ? AND meeting = ?
                          AND source_tdoc = ? AND cr_number = ? AND revision = ? AND subject = ?
                    """, (spec_num, version, meeting, source_tdoc, cr_number, revision, subject)).fetchone()
                    history_id = row[0] if row else None

                if not history_id:
                    continue

                inserted_count += 1

                if source_tdoc and source_tdoc != "-":
                    clean_base = source_tdoc.split("r")[0].upper()
                    cursor.execute("""
                        INSERT INTO spec_tdoc_cross_ref (
                            tdoc, raw_tdoc, spec_number, version, meeting, is_pcr, history_id
                        )
                        VALUES (?, ?, ?, ?, ?, 0, ?)
                        ON CONFLICT(history_id, raw_tdoc) DO UPDATE SET
                            meeting=excluded.meeting,
                            version=excluded.version
                    """, (clean_base, source_tdoc, spec_num, version, meeting, history_id))

                for pcr in rec.get("pcr_tdocs", []):
                    raw_pcr = str(pcr).strip()
                    if not raw_pcr or raw_pcr == "-":
                        continue
                    clean_pcr_base = raw_pcr.split("r")[0].split("R")[0].upper()
                    cursor.execute("""
                        INSERT INTO spec_tdoc_cross_ref (
                            tdoc, raw_tdoc, spec_number, version, meeting, is_pcr, history_id
                        )
                        VALUES (?, ?, ?, ?, ?, 1, ?)
                        ON CONFLICT(history_id, raw_tdoc) DO UPDATE SET
                            meeting=excluded.meeting,
                            version=excluded.version
                    """, (clean_pcr_base, raw_pcr, spec_num, version, meeting, history_id))

        return inserted_count

    def delete_spec_history(self, spec_number: str) -> int:
        """Deletes all history records for a specification. Cascade deletes cross-references."""
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("DELETE FROM spec_change_history WHERE spec_number = ?", (spec_number.strip(),))
            return cursor.rowcount

    def wipe_database(self) -> bool:
        """Completely drops tables and clears the history database."""
        try:
            with self._get_connection() as conn:
                conn.execute("DROP TABLE IF EXISTS tdoc_metadata_cache;")
                conn.execute("DROP TABLE IF EXISTS spec_tdoc_cross_ref;")
                conn.execute("DROP TABLE IF EXISTS spec_change_history;")
                conn.execute("PRAGMA wal_checkpoint(TRUNCATE);")
                conn.execute("VACUUM;")
            self._init_db()
            self.logger.info("Successfully wiped and reinitialized 3gpp_spec_history.db")
            return True
        except Exception as e:
            self.logger.error(f"Error wiping 3gpp_spec_history.db: {e}", exc_info=True)
            return False
