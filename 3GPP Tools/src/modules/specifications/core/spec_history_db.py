# --- File: src/modules/specifications/core/spec_history_db.py ---
"""
Specification Change History Database Engine.

Manages persistent SQLite storage (in 3gpp_spec_history.db) for 3GPP Technical
Specifications and Technical Reports Change History, formal TSG Change Requests (CRs),
Working Group Pseudo-CRs (p-CRs), and integrated TDocs.
"""

import contextlib
import logging
import sqlite3
from pathlib import Path
from typing import Any, Dict, Iterator, List, Optional, Sequence, Tuple, Union

logger = logging.getLogger(__name__)


class SpecHistoryDatabase:
    """Thread-safe SQLite database manager for 3GPP Specification Change History and TDoc cross-referencing."""

    def __init__(self, db_path: Optional[Union[Path, str]] = None):
        if db_path is None:
            from core.utils.paths import get_project_root
            self.db_path = get_project_root() / "db" / "3gpp_spec_history.db"
        else:
            self.db_path = Path(db_path)
        self.logger = logging.getLogger(__name__)
        self._init_db()

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

            # Indexes for ultra-fast lookup
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_hist_spec_ver ON spec_change_history(spec_number, version);")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_hist_spec ON spec_change_history(spec_number);")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_tdoc_lookup ON spec_tdoc_cross_ref(tdoc);")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_raw_tdoc_lookup ON spec_tdoc_cross_ref(raw_tdoc);")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_tdoc_spec_ver ON spec_tdoc_cross_ref(spec_number, version);")

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

                if not spec_num:
                    continue

                # 1. Insert master history row with RETURNING id for bulletproof FK linkage
                try:
                    cursor.execute("""
                        INSERT INTO spec_change_history (
                            spec_number, version, row_order, date, meeting,
                            source_tdoc, cr_number, revision, category,
                            subject, old_version, source_file
                        )
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        ON CONFLICT(spec_number, version, meeting, source_tdoc, cr_number, revision, subject)
                        DO UPDATE SET
                            row_order=excluded.row_order,
                            date=excluded.date,
                            category=excluded.category,
                            old_version=excluded.old_version,
                            source_file=excluded.source_file,
                            extracted_at=CURRENT_TIMESTAMP
                        RETURNING id;
                    """, (
                        spec_num, version, row_order, date_val, meeting,
                        source_tdoc, cr_number, revision, category,
                        subject, old_ver, source_file
                    ))
                    res = cursor.fetchone()
                    history_id = res[0] if res else None
                except sqlite3.OperationalError:
                    # Fallback for older SQLite engines without RETURNING clause
                    cursor.execute("""
                        INSERT INTO spec_change_history (
                            spec_number, version, row_order, date, meeting,
                            source_tdoc, cr_number, revision, category,
                            subject, old_version, source_file
                        )
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        ON CONFLICT(spec_number, version, meeting, source_tdoc, cr_number, revision, subject)
                        DO UPDATE SET
                            row_order=excluded.row_order,
                            date=excluded.date,
                            category=excluded.category,
                            old_version=excluded.old_version,
                            source_file=excluded.source_file,
                            extracted_at=CURRENT_TIMESTAMP;
                    """, (
                        spec_num, version, row_order, date_val, meeting,
                        source_tdoc, cr_number, revision, category,
                        subject, old_ver, source_file
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

                # 2. Insert primary TDoc if present
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

                # 3. Insert embedded p-CR TDocs
                pcr_tdocs = rec.get("pcr_tdocs", [])
                for pcr in pcr_tdocs:
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

    def get_spec_history(
        self,
        spec_number: str,
        version: Optional[str] = None
    ) -> List[Dict[str, Any]]:
        """
        Retrieves full change history records for a specification, optionally filtered by version.
        Each record includes an embedded 'tdocs' list of all associated TDocs.
        """
        clean_spec = spec_number.strip()
        with self._get_connection() as conn:
            query = """
                SELECT id, spec_number, version, row_order, date, meeting,
                       source_tdoc, cr_number, revision, category, subject,
                       old_version, source_file, extracted_at
                FROM spec_change_history
                WHERE spec_number = ?
            """
            params: List[Any] = [clean_spec]

            if version:
                clean_ver = version.lstrip("vV").strip()
                query += " AND version = ?"
                params.append(clean_ver)

            query += " ORDER BY row_order ASC, id ASC"

            rows = conn.execute(query, params).fetchall()
            results = [dict(r) for r in rows]

            # Fetch associated TDocs for each history record
            if results:
                history_ids = [r["id"] for r in results]
                placeholders = ",".join("?" for _ in history_ids)
                tdoc_rows = conn.execute(f"""
                    SELECT history_id, tdoc, raw_tdoc, is_pcr
                    FROM spec_tdoc_cross_ref
                    WHERE history_id IN ({placeholders})
                    ORDER BY id ASC
                """, history_ids).fetchall()

                tdoc_map: Dict[int, List[Dict[str, Any]]] = {}
                for t in tdoc_rows:
                    hid = t["history_id"]
                    tdoc_map.setdefault(hid, []).append({
                        "tdoc": t["tdoc"],
                        "raw_tdoc": t["raw_tdoc"],
                        "is_pcr": bool(t["is_pcr"]),
                    })

                for r in results:
                    r["tdocs"] = tdoc_map.get(r["id"], [])

            return results

    # Convenience alias for get_spec_history
    get_history_for_spec = get_spec_history

    def get_tdocs_for_spec_version(self, spec_number: str, version: str) -> List[Dict[str, Any]]:
        """Returns all distinct TDocs cross-referenced for a given specification and version."""
        clean_spec = spec_number.strip()
        clean_ver = version.lstrip("vV").strip()
        with self._get_connection() as conn:
            rows = conn.execute("""
                SELECT DISTINCT tdoc, raw_tdoc, is_pcr, meeting
                FROM spec_tdoc_cross_ref
                WHERE spec_number = ? AND version = ?
                ORDER BY tdoc ASC
            """, (clean_spec, clean_ver)).fetchall()
            return [dict(r) for r in rows]

    def get_version_summary(
        self,
        spec_number: Optional[Union[str, Sequence[str]]] = None
    ) -> Dict[Tuple[str, str], int]:
        """
        Fast bulk pre-fetch returning a dictionary of {(spec_number, version): record_count}.
        Supports a single spec number, a sequence/list of spec numbers, or None for the full database.
        Executes in ~2ms across the whole database to prevent N+1 query lag in UI loops.
        """
        summary: Dict[Tuple[str, str], int] = {}
        with self._get_connection() as conn:
            if spec_number:
                if isinstance(spec_number, str):
                    rows = conn.execute("""
                        SELECT spec_number, version, COUNT(id) AS cnt
                        FROM spec_change_history
                        WHERE spec_number = ?
                        GROUP BY spec_number, version
                    """, (spec_number.strip(),)).fetchall()
                else:
                    specs = [str(s).strip() for s in spec_number if str(s).strip()]
                    if not specs:
                        return summary
                    placeholders = ",".join("?" for _ in specs)
                    rows = conn.execute(f"""
                        SELECT spec_number, version, COUNT(id) AS cnt
                        FROM spec_change_history
                        WHERE spec_number IN ({placeholders})
                        GROUP BY spec_number, version
                    """, specs).fetchall()
            else:
                rows = conn.execute("""
                    SELECT spec_number, version, COUNT(id) AS cnt
                    FROM spec_change_history
                    GROUP BY spec_number, version
                """).fetchall()

            for r in rows:
                summary[(r["spec_number"], r["version"])] = r["cnt"]

        return summary

    def get_versions_for_spec(self, spec_number: str) -> List[str]:
        """Returns distinct version numbers indexed for a specification in natural order."""
        with self._get_connection() as conn:
            rows = conn.execute("""
                SELECT DISTINCT version
                FROM spec_change_history
                WHERE spec_number = ? AND version != ''
                ORDER BY row_order DESC, id DESC
            """, (spec_number.strip(),)).fetchall()
            return [r[0] for r in rows]

    def get_tdoc_references(self, tdoc: str) -> List[Dict[str, Any]]:
        """
        Reverse lookup: Finds all specifications and versions that integrated a given TDoc ID.
        Matches either normalized base ID (S2-2606265) or raw token with revision (S2-2606265r01).
        """
        clean_tdoc = tdoc.strip().upper()
        base_tdoc = clean_tdoc.split("R")[0]
        with self._get_connection() as conn:
            rows = conn.execute("""
                SELECT c.id, c.tdoc, c.raw_tdoc, c.spec_number, c.version, c.meeting, c.is_pcr,
                       h.date, h.cr_number, h.revision, h.category, h.subject
                FROM spec_tdoc_cross_ref c
                JOIN spec_change_history h ON c.history_id = h.id
                WHERE c.tdoc = ? OR c.raw_tdoc = ?
                ORDER BY h.date DESC, c.version DESC
            """, (base_tdoc, clean_tdoc)).fetchall()
            return [dict(r) for r in rows]

    def has_spec_history(self, spec_number: str) -> bool:
        """Checks if any history records exist for the specification."""
        with self._get_connection() as conn:
            row = conn.execute(
                "SELECT 1 FROM spec_change_history WHERE spec_number = ? LIMIT 1",
                (spec_number.strip(),)
            ).fetchone()
            return bool(row)

    def delete_spec_history(self, spec_number: str) -> int:
        """Deletes all history records for a specification. Cascade deletes cross-references."""
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("DELETE FROM spec_change_history WHERE spec_number = ?", (spec_number.strip(),))
            return cursor.rowcount

    def get_statistics(self) -> Dict[str, Any]:
        """Returns aggregate metrics on indexed specifications, versions, change entries, and TDocs."""
        with self._get_connection() as conn:
            row_rec = conn.execute("SELECT COUNT(*) FROM spec_change_history;").fetchone()
            row_tdoc = conn.execute("SELECT COUNT(*) FROM spec_tdoc_cross_ref;").fetchone()
            row_specs = conn.execute("SELECT COUNT(DISTINCT spec_number) FROM spec_change_history;").fetchone()
            row_vers = conn.execute("SELECT COUNT(DISTINCT spec_number || ':' || version) FROM spec_change_history;").fetchone()
            return {
                "total_records": row_rec[0] if row_rec else 0,
                "total_tdocs": row_tdoc[0] if row_tdoc else 0,
                "unique_specs": row_specs[0] if row_specs else 0,
                "unique_versions": row_vers[0] if row_vers else 0,
            }

    def wipe_database(self) -> bool:
        """Completely drops tables and clears the history database."""
        try:
            with self._get_connection() as conn:
                conn.execute("DROP TABLE IF EXISTS spec_tdoc_cross_ref;")
                conn.execute("DROP TABLE IF EXISTS spec_change_history;")
                conn.execute("PRAGMA wal_checkpoint(TRUNCATE);")
                conn.execute("VACUUM;")
            self._init_db()
            logger.info("Successfully wiped and reinitialized 3gpp_spec_history.db")
            return True
        except Exception as e:
            logger.error(f"Error wiping 3gpp_spec_history.db: {e}", exc_info=True)
            return False
