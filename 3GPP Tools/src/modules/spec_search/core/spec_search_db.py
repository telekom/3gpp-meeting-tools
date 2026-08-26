"""
Specification Search Database Engine.
Manages metadata tables, release dates, full-text trigram indices (FTS5), and chronological release diffs.
"""

import logging
from pathlib import Path
import re
import sqlite3
from typing import Any, Dict, List, Optional, Tuple
import pandas as pd


def parse_version_tuple(version_str: str) -> tuple:
    """Converts version strings (e.g. '18.4.0', 'v15.2.1') into integer tuples for natural sorting."""
    if not version_str:
        return ()
    clean = str(version_str).lstrip("vV").strip()
    parts = []
    for part in clean.split("."):
        if part.isdigit():
            parts.append(int(part))
        else:
            parts.append(part)
    return tuple(parts)


class SpecSearchDatabase:
    """SQLite database manager for 3GPP Specification Full-Text and Substring Indexing with Release Date tracking."""

    def __init__(self, db_path: Path):
        self.db_path = Path(db_path)
        self.logger = logging.getLogger(__name__)
        self._init_db()

    def _get_connection(self) -> sqlite3.Connection:
        conn = sqlite3.connect(str(self.db_path), check_same_thread=False)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA foreign_keys = ON;")
        conn.execute("PRAGMA journal_mode = WAL;")
        conn.execute("PRAGMA synchronous = NORMAL;")
        return conn

    def _init_db(self):
        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()

                # Relational metadata tables
                cursor.execute("""
                    CREATE TABLE IF NOT EXISTS indexed_versions (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        spec_number TEXT NOT NULL,
                        version TEXT NOT NULL,
                        release_date TEXT,
                        import_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                        total_clauses INTEGER DEFAULT 0,
                        total_chars INTEGER DEFAULT 0,
                        UNIQUE(spec_number, version)
                    )
                """)

                cursor.execute("""
                    CREATE TABLE IF NOT EXISTS spec_clauses (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        version_id INTEGER NOT NULL,
                        clause_number TEXT NOT NULL,
                        clause_title TEXT NOT NULL,
                        content TEXT NOT NULL,
                        content_length INTEGER DEFAULT 0,
                        order_index INTEGER NOT NULL,
                        FOREIGN KEY(version_id) REFERENCES indexed_versions(id) ON DELETE CASCADE
                    )
                """)

                # Virtual table with FTS5 Trigram Tokenizer for fast substring matching
                cursor.execute("""
                    CREATE VIRTUAL TABLE IF NOT EXISTS spec_fts USING fts5(
                        clause_pk UNINDEXED,
                        version_id UNINDEXED,
                        spec_number UNINDEXED,
                        version UNINDEXED,
                        clause_number UNINDEXED,
                        clause_title,
                        content,
                        tokenize="trigram"
                    )
                """)

                cursor.execute("CREATE INDEX IF NOT EXISTS idx_clause_ver ON spec_clauses(version_id);")
                cursor.execute("CREATE INDEX IF NOT EXISTS idx_clause_num ON spec_clauses(clause_number);")
                cursor.execute("CREATE INDEX IF NOT EXISTS idx_ver_spec ON indexed_versions(spec_number);")

                # Dynamic schema migration for older search databases
                cursor.execute("PRAGMA table_info(indexed_versions);")
                cols = [col["name"] for col in cursor.fetchall()]
                if "release_date" not in cols:
                    cursor.execute("ALTER TABLE indexed_versions ADD COLUMN release_date TEXT;")

                conn.commit()
        except Exception as e:
            self.logger.error(f"Failed to initialize Spec Search DB: {e}")

    def get_imported_versions(self) -> List[Dict[str, Any]]:
        """Returns all imported versions sorted chronologically with release dates."""
        query = "SELECT id, spec_number, version, release_date, import_date, total_clauses, total_chars FROM indexed_versions"
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(query)
            rows = [dict(row) for row in cursor.fetchall()]
        return sorted(rows, key=lambda x: (x["spec_number"], parse_version_tuple(x["version"])), reverse=True)

    def insert_parsed_spec(
        self,
        spec_number: str,
        version: str,
        clauses: List[Dict[str, Any]],
        release_date: Optional[str] = None,
    ) -> bool:
        """Batch inserts clauses, release date, and populates the FTS5 trigram index atomically."""
        if not clauses:
            return False

        conn = self._get_connection()
        try:
            with conn:
                cursor = conn.cursor()

                # 1. Clean existing version if re-importing
                cursor.execute(
                    "SELECT id FROM indexed_versions WHERE spec_number = ? AND version = ?",
                    (spec_number, version),
                )
                existing = cursor.fetchone()
                if existing:
                    v_id = existing[0]
                    cursor.execute("DELETE FROM spec_fts WHERE version_id = ?", (v_id,))
                    cursor.execute("DELETE FROM indexed_versions WHERE id = ?", (v_id,))

                # 2. Insert into indexed_versions
                total_chars = sum(len(c.get("content", "")) for c in clauses)
                cursor.execute(
                    """
                    INSERT INTO indexed_versions (spec_number, version, release_date, total_clauses, total_chars)
                    VALUES (?, ?, ?, ?, ?)
                """,
                    (spec_number, version, release_date, len(clauses), total_chars),
                )
                version_id = cursor.lastrowid

                # 3. Batch insert into spec_clauses
                clause_rows = []
                for idx, c in enumerate(clauses):
                    content = c.get("content", "")
                    clause_rows.append((
                        version_id,
                        c.get("clause_number", ""),
                        c.get("clause_title", ""),
                        content,
                        len(content),
                        idx,
                    ))

                cursor.executemany(
                    """
                    INSERT INTO spec_clauses (version_id, clause_number, clause_title, content, content_length, order_index)
                    VALUES (?, ?, ?, ?, ?, ?)
                """,
                    clause_rows,
                )

                # 4. Fetch the generated clause primary keys to mirror into FTS5
                cursor.execute(
                    "SELECT id, clause_number, clause_title, content FROM spec_clauses WHERE version_id = ? ORDER BY order_index ASC",
                    (version_id,),
                )
                inserted_clauses = cursor.fetchall()

                fts_rows = [
                    (
                        row["id"],
                        version_id,
                        spec_number,
                        version,
                        row["clause_number"],
                        row["clause_title"],
                        row["content"],
                    )
                    for row in inserted_clauses
                ]

                cursor.executemany(
                    """
                    INSERT INTO spec_fts (clause_pk, version_id, spec_number, version, clause_number, clause_title, content)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                """,
                    fts_rows,
                )

            return True
        except Exception as e:
            self.logger.error(f"Failed to insert TS {spec_number} v{version} into search DB: {e}")
            return False
        finally:
            conn.close()

    def search_substring(
        self,
        query_str: str,
        version_ids: List[int],
        clause_filter: Optional[str] = None,
    ) -> pd.DataFrame:
        """
        Executes substring matching across selected specification versions using SQLite FTS5.
        Returns a DataFrame containing clause metadata, version details, release dates, and hit snippets.
        """
        if not version_ids or not query_str.strip():
            return pd.DataFrame()

        clean_query = query_str.strip()
        if len(clean_query) < 3:
            return self._search_like_fallback(clean_query, version_ids, clause_filter)

        placeholders = ",".join("?" for _ in version_ids)
        sanitized_query = clean_query.replace('"', '""')
        escaped_query = f'"{sanitized_query}"'

        sql = f"""
            SELECT 
                f.spec_number,
                f.version,
                v.release_date,
                f.clause_number,
                f.clause_title,
                f.version_id,
                f.clause_pk,
                snippet(spec_fts, 6, '<mark>', '</mark>', '...', 20) AS snippet_text,
                c.order_index
            FROM spec_fts f
            JOIN spec_clauses c ON f.clause_pk = c.id
            JOIN indexed_versions v ON f.version_id = v.id
            WHERE f.version_id IN ({placeholders})
              AND spec_fts MATCH ?
        """
        params: List[Any] = list(version_ids) + [escaped_query]

        if clause_filter and clause_filter.strip():
            sql += " AND (f.clause_number LIKE ? OR f.clause_title LIKE ?)"
            pat = f"%{clause_filter.strip()}%"
            params.extend([pat, pat])

        sql += " ORDER BY c.order_index ASC"

        with self._get_connection() as conn:
            return pd.read_sql_query(sql, conn, params=params)

    def _search_like_fallback(
        self,
        clean_query: str,
        version_ids: List[int],
        clause_filter: Optional[str] = None,
    ) -> pd.DataFrame:
        """Fallback scanning for queries with fewer than 3 characters."""
        placeholders = ",".join("?" for _ in version_ids)
        pattern = f"%{clean_query}%"

        sql = f"""
            SELECT 
                v.spec_number,
                v.version,
                v.release_date,
                c.clause_number,
                c.clause_title,
                c.version_id,
                c.id AS clause_pk,
                substr(c.content, max(1, instr(lower(c.content), lower(?)) - 25), 80) AS snippet_text,
                c.order_index
            FROM spec_clauses c
            JOIN indexed_versions v ON c.version_id = v.id
            WHERE c.version_id IN ({placeholders})
              AND (c.content LIKE ? OR c.clause_title LIKE ?)
        """
        params: List[Any] = [clean_query] + list(version_ids) + [pattern, pattern]

        if clause_filter and clause_filter.strip():
            sql += " AND (c.clause_number LIKE ? OR c.clause_title LIKE ?)"
            pat = f"%{clause_filter.strip()}%"
            params.extend([pat, pat])

        sql += " ORDER BY c.order_index ASC"

        with self._get_connection() as conn:
            return pd.read_sql_query(sql, conn, params=params)

    def get_clause_content(self, clause_pk: int) -> Optional[Dict[str, Any]]:
        query = """
            SELECT c.*, v.spec_number, v.version, v.release_date
            FROM spec_clauses c
            JOIN indexed_versions v ON c.version_id = v.id
            WHERE c.id = ?
        """
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(query, (clause_pk,))
            row = cursor.fetchone()
            return dict(row) if row else None

    def get_clause_content_by_spec_ver(
        self, spec_number: str, version: str, clause_number: str
    ) -> Optional[Dict[str, Any]]:
        query = """
            SELECT c.*, v.spec_number, v.version, v.release_date
            FROM spec_clauses c
            JOIN indexed_versions v ON c.version_id = v.id
            WHERE v.spec_number = ? AND v.version = ? AND c.clause_number = ?
            LIMIT 1
        """
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(query, (spec_number, version, clause_number))
            row = cursor.fetchone()
            return dict(row) if row else None

    def clear_version(self, spec_number: str, version: str) -> bool:
        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute(
                    "SELECT id FROM indexed_versions WHERE spec_number = ? AND version = ?",
                    (spec_number, version),
                )
                row = cursor.fetchone()
                if row:
                    v_id = row[0]
                    cursor.execute("DELETE FROM spec_fts WHERE version_id = ?", (v_id,))
                    cursor.execute("DELETE FROM indexed_versions WHERE id = ?", (v_id,))
                conn.commit()
                return True
        except Exception as e:
            self.logger.error(f"Failed to delete TS {spec_number} v{version}: {e}")
            return False

    def wipe_database(self) -> bool:
        """
        Wipes all indexed specifications and resets the database schema.
        Attempts fast file-level unlinking, falling back to SQL drops if file locks exist.
        """
        try:
            # 1. Checkpoint and close active handles
            try:
                with self._get_connection() as conn:
                    conn.execute("PRAGMA wal_checkpoint(TRUNCATE);")
            except Exception:
                pass

            # 2. Fast file unlink (near instantaneous)
            wiped_files = False
            try:
                wal_file = Path(str(self.db_path) + "-wal")
                shm_file = Path(str(self.db_path) + "-shm")

                if self.db_path.exists():
                    self.db_path.unlink()
                if wal_file.exists():
                    wal_file.unlink()
                if shm_file.exists():
                    shm_file.unlink()
                wiped_files = True
            except (PermissionError, OSError):
                wiped_files = False

            # 3. Fallback to standard SQL table drops if file is locked
            if not wiped_files:
                with self._get_connection() as conn:
                    cursor = conn.cursor()
                    cursor.execute("DROP TABLE IF EXISTS spec_fts;")
                    cursor.execute("DROP TABLE IF EXISTS spec_clauses;")
                    cursor.execute("DROP TABLE IF EXISTS indexed_versions;")
                    conn.commit()
                    conn.execute("PRAGMA wal_checkpoint(TRUNCATE);")
                    conn.execute("VACUUM;")

            # 4. Re-initialize fresh schema
            self._init_db()
            return True
        except Exception as e:
            self.logger.error(f"Failed to wipe Spec Search DB: {e}")
            return False