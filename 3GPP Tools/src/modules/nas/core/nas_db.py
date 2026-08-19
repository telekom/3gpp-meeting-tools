import logging
from pathlib import Path
import sqlite3
from typing import Any, Dict, List, Optional
import pandas as pd


def parse_version_tuple(version_str: str) -> tuple:
    """Converts a version string into a tuple of integers for natural sorting."""
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


class NASDatabase:
    """Manages the SQLite database for 3GPP TS 24.501 NAS messages and IE definitions."""

    def __init__(self, db_path: Path):
        self.db_path = Path(db_path)
        self.logger = logging.getLogger(__name__)
        self._init_db()

    def _get_connection(self) -> sqlite3.Connection:
        conn = sqlite3.connect(str(self.db_path), check_same_thread=False)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA foreign_keys = ON;")
        conn.execute("PRAGMA journal_mode = WAL;")
        return conn

    def _init_db(self):
        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute("""
                    CREATE TABLE IF NOT EXISTS spec_versions (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        spec_number TEXT NOT NULL,
                        version TEXT NOT NULL,
                        import_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                        UNIQUE(spec_number, version)
                    )
                """)

                cursor.execute("""
                    CREATE TABLE IF NOT EXISTS nas_messages (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        version_id INTEGER NOT NULL,
                        clause TEXT NOT NULL,
                        message_name TEXT NOT NULL,
                        table_caption TEXT,
                        FOREIGN KEY(version_id) REFERENCES spec_versions(id) ON DELETE CASCADE
                    )
                """)

                cursor.execute("""
                    CREATE TABLE IF NOT EXISTS message_ies (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        message_id INTEGER NOT NULL,
                        iei TEXT,
                        ie_name TEXT NOT NULL,
                        type_reference TEXT,
                        presence TEXT,
                        format TEXT,
                        length TEXT,
                        order_index INTEGER NOT NULL,
                        FOREIGN KEY(message_id) REFERENCES nas_messages(id) ON DELETE CASCADE
                    )
                """)

                cursor.execute("""
                    CREATE TABLE IF NOT EXISTS ie_definitions (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        version_id INTEGER NOT NULL,
                        clause TEXT NOT NULL,
                        ie_name TEXT NOT NULL,
                        raw_description TEXT,
                        structure_table TEXT,
                        FOREIGN KEY(version_id) REFERENCES spec_versions(id) ON DELETE CASCADE,
                        UNIQUE(version_id, clause)
                    )
                """)

                cursor.execute("CREATE INDEX IF NOT EXISTS idx_msg_ver ON nas_messages(version_id);")
                cursor.execute("CREATE INDEX IF NOT EXISTS idx_ie_msg ON message_ies(message_id);")
                cursor.execute("CREATE INDEX IF NOT EXISTS idx_def_ver ON ie_definitions(version_id);")
                conn.commit()
        except Exception as e:
            self.logger.error(f"Error initializing NAS DB: {e}")

    def get_imported_versions(self) -> List[Dict[str, Any]]:
        """Retrieves all imported specification versions sorted numerically descending."""
        query = "SELECT id, spec_number, version, import_date FROM spec_versions"
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(query)
            rows = [dict(row) for row in cursor.fetchall()]
        return sorted(rows, key=lambda x: parse_version_tuple(x["version"]), reverse=True)

    def clear_version(self, spec_number: str, version: str) -> bool:
        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute(
                    "DELETE FROM spec_versions WHERE spec_number = ? AND version = ?",
                    (spec_number, version),
                )
                conn.commit()
                return True
        except Exception as e:
            self.logger.error(f"Failed to clear version {version}: {e}")
            return False

    def wipe_database(self) -> bool:
        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute("DROP TABLE IF EXISTS message_ies;")
                cursor.execute("DROP TABLE IF EXISTS nas_messages;")
                cursor.execute("DROP TABLE IF EXISTS ie_definitions;")
                cursor.execute("DROP TABLE IF EXISTS spec_versions;")
                conn.commit()
            self._init_db()
            return True
        except Exception as e:
            self.logger.error(f"Failed to wipe NAS DB: {e}")
            return False

    def insert_parsed_spec(
        self,
        spec_number: str,
        version: str,
        messages: List[Dict[str, Any]],
        ie_defs: List[Dict[str, Any]],
    ) -> bool:
        conn = self._get_connection()
        try:
            with conn:
                cursor = conn.cursor()
                cursor.execute(
                    "DELETE FROM spec_versions WHERE spec_number = ? AND version = ?",
                    (spec_number, version),
                )

                cursor.execute(
                    "INSERT INTO spec_versions (spec_number, version) VALUES (?, ?)",
                    (spec_number, version),
                )
                version_id = cursor.lastrowid

                for msg in messages:
                    cursor.execute(
                        """
                        INSERT INTO nas_messages (version_id, clause, message_name, table_caption)
                        VALUES (?, ?, ?, ?)
                    """,
                        (
                            version_id,
                            msg.get("clause", ""),
                            msg.get("message_name", ""),
                            msg.get("table_caption", ""),
                        ),
                    )
                    message_id = cursor.lastrowid

                    ie_rows = []
                    for idx, ie in enumerate(msg.get("ies", [])):
                        ie_rows.append((
                            message_id,
                            ie.get("iei", ""),
                            ie.get("information_element", ""),
                            ie.get("type_reference", ""),
                            ie.get("presence", ""),
                            ie.get("format", ""),
                            ie.get("length", ""),
                            idx,
                        ))

                    cursor.executemany(
                        """
                        INSERT INTO message_ies (message_id, iei, ie_name, type_reference, presence, format, length, order_index)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                        ie_rows,
                    )

                def_rows = [
                    (
                        version_id,
                        d.get("clause", ""),
                        d.get("ie_name", ""),
                        d.get("raw_description", ""),
                        d.get("structure_table", ""),
                    )
                    for d in ie_defs
                ]
                cursor.executemany(
                    """
                    INSERT OR REPLACE INTO ie_definitions (version_id, clause, ie_name, raw_description, structure_table)
                    VALUES (?, ?, ?, ?, ?)
                """,
                    def_rows,
                )
            return True
        except Exception as e:
            self.logger.error(f"Failed to insert parsed spec: {e}")
            return False
        finally:
            conn.close()

    def get_messages_list(self, version_ids: Optional[List[int]] = None) -> List[Dict[str, Any]]:
        query = "SELECT DISTINCT m.message_name, m.clause FROM nas_messages m"
        params = []
        if version_ids:
            placeholders = ",".join("?" for _ in version_ids)
            query += f" WHERE m.version_id IN ({placeholders})"
            params.extend(version_ids)
        query += " ORDER BY m.message_name ASC"

        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(query, params)
            return [dict(row) for row in cursor.fetchall()]

    def get_messages_by_ie_search(
        self, ie_query: str, version_ids: Optional[List[int]] = None
    ) -> List[Dict[str, Any]]:
        """Returns only messages that contain Information Elements matching the search query."""
        if not version_ids or not ie_query.strip():
            return self.get_messages_list(version_ids)

        placeholders = ",".join("?" for _ in version_ids)
        query = f"""
            SELECT DISTINCT m.message_name, m.clause
            FROM nas_messages m
            JOIN message_ies i ON i.message_id = m.id
            WHERE m.version_id IN ({placeholders})
              AND (
                  i.ie_name LIKE ? 
                  OR i.type_reference LIKE ? 
                  OR i.iei LIKE ?
              )
            ORDER BY m.message_name ASC
        """
        pattern = f"%{ie_query.strip()}%"
        params = list(version_ids) + [pattern, pattern, pattern]

        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(query, params)
            return [dict(row) for row in cursor.fetchall()]

    def get_messages_using_ie(
        self, clause: str, ie_name: str = "", version_ids: Optional[List[int]] = None
    ) -> List[Dict[str, Any]]:
        """
        Finds all distinct NAS messages across active versions that reference an IE by clause or name.
        """
        if not version_ids:
            return []

        placeholders = ",".join("?" for _ in version_ids)
        clause_pattern = f"%{clause.strip()}%" if clause.strip() else ""
        name_pattern = f"%{ie_name.strip()}%" if ie_name.strip() else ""

        query = f"""
            SELECT DISTINCT m.message_name, m.clause
            FROM nas_messages m
            JOIN message_ies i ON i.message_id = m.id
            WHERE m.version_id IN ({placeholders})
              AND (
                  (? != '' AND i.type_reference LIKE ?)
                  OR (? != '' AND i.ie_name LIKE ?)
              )
            ORDER BY m.message_name ASC
        """
        params = list(version_ids) + [clause_pattern, clause_pattern, name_pattern, name_pattern]

        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(query, params)
            return [dict(row) for row in cursor.fetchall()]

    def get_message_evolution_df(self, message_name: str, version_ids: List[int]) -> pd.DataFrame:
        if not version_ids:
            return pd.DataFrame()

        placeholders = ",".join("?" for _ in version_ids)
        query = f"""
            SELECT 
                sv.spec_number,
                sv.version,
                i.iei,
                i.ie_name,
                i.type_reference,
                i.presence,
                i.format,
                i.length,
                i.order_index
            FROM message_ies i
            JOIN nas_messages m ON i.message_id = m.id
            JOIN spec_versions sv ON m.version_id = sv.id
            WHERE m.message_name = ? AND sv.id IN ({placeholders})
            ORDER BY i.order_index ASC
        """
        params = [message_name] + version_ids

        with self._get_connection() as conn:
            return pd.read_sql_query(query, conn, params=params)

    def get_ie_definitions_by_clause(
        self, clause: str, version_ids: Optional[List[int]] = None
    ) -> List[Dict[str, Any]]:
        query = """
            SELECT d.*, sv.version, sv.spec_number
            FROM ie_definitions d
            JOIN spec_versions sv ON d.version_id = sv.id
            WHERE d.clause = ?
        """
        params = [clause]
        if version_ids:
            placeholders = ",".join("?" for _ in version_ids)
            query += f" AND d.version_id IN ({placeholders})"
            params.extend(version_ids)

        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(query, params)
            rows = [dict(row) for row in cursor.fetchall()]

        return sorted(rows, key=lambda x: parse_version_tuple(x["version"]), reverse=True)

    def get_ie_definition(self, clause: str, version_id: Optional[int] = None) -> Optional[Dict[str, Any]]:
        query = "SELECT * FROM ie_definitions WHERE clause = ?"
        params = [clause]
        if version_id:
            query += " AND version_id = ?"
            params.append(version_id)
        query += " ORDER BY id DESC LIMIT 1"

        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(query, params)
            row = cursor.fetchone()
            return dict(row) if row else None