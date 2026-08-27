import logging
from pathlib import Path
import re
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
    """Manages the SQLite database for 3GPP NAS (24.501, 24.301) and ASN.1 (38.331, 36.331, 38.413) protocols."""

    def __init__(self, db_path: Path):
        self.db_path = Path(db_path)
        self.logger = logging.getLogger(__name__)
        self._init_db()

    def _get_connection(self) -> sqlite3.Connection:
        conn = sqlite3.connect(str(self.db_path), check_same_thread=False)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA foreign_keys = ON;")
        conn.execute("PRAGMA journal_mode = WAL;")
        conn.create_function(
            "REGEXP",
            2,
            lambda expr, item: bool(re.search(expr, str(item))) if item is not None else False,
        )
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
                        spec_type TEXT DEFAULT 'NAS',
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
                        field_path TEXT,
                        depth INTEGER DEFAULT 0,
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
                        UNIQUE(version_id, ie_name)
                    )
                """)

                cursor.execute("CREATE INDEX IF NOT EXISTS idx_msg_ver ON nas_messages(version_id);")
                cursor.execute("CREATE INDEX IF NOT EXISTS idx_msg_name ON nas_messages(message_name);")
                cursor.execute("CREATE INDEX IF NOT EXISTS idx_ie_msg ON message_ies(message_id);")
                cursor.execute("CREATE INDEX IF NOT EXISTS idx_ie_name ON message_ies(ie_name);")
                cursor.execute("CREATE INDEX IF NOT EXISTS idx_ie_type ON message_ies(type_reference);")
                cursor.execute("CREATE INDEX IF NOT EXISTS idx_def_ver ON ie_definitions(version_id);")
                cursor.execute("CREATE INDEX IF NOT EXISTS idx_def_name ON ie_definitions(ie_name);")
                cursor.execute("CREATE INDEX IF NOT EXISTS idx_def_clause ON ie_definitions(clause);")
                conn.commit()
        except Exception as e:
            self.logger.error(f"Error initializing Protocol DB: {e}")

    def get_imported_versions(self) -> List[Dict[str, Any]]:
        query = "SELECT id, spec_number, version, spec_type, import_date FROM spec_versions"
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

            with self._get_connection() as conn:
                conn.execute("PRAGMA wal_checkpoint(TRUNCATE);")
                conn.execute("VACUUM;")

            return True
        except Exception as e:
            self.logger.error(f"Failed to wipe Protocol DB: {e}")
            return False

    def insert_parsed_spec(
        self,
        spec_number: str,
        version: str,
        messages: List[Dict[str, Any]],
        ie_defs: List[Dict[str, Any]],
        spec_type: str = "NAS",
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
                    "INSERT INTO spec_versions (spec_number, version, spec_type) VALUES (?, ?, ?)",
                    (spec_number, version, spec_type),
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
                            ie.get("field_path", ie.get("information_element", "")),
                            ie.get("depth", 0),
                            ie.get("type_reference", ""),
                            ie.get("presence", ""),
                            ie.get("format", ""),
                            ie.get("length", ""),
                            idx,
                        ))

                    cursor.executemany(
                        """
                        INSERT INTO message_ies (message_id, iei, ie_name, field_path, depth, type_reference, presence, format, length, order_index)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
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
            self.logger.error(f"Failed to insert parsed spec TS {spec_number} v{version}: {e}")
            return False
        finally:
            conn.close()

    def get_messages_list(self, version_ids: Optional[List[int]] = None) -> List[Dict[str, Any]]:
        query = """
            SELECT m.message_name, m.clause, GROUP_CONCAT(DISTINCT sv.spec_number) AS spec_number
            FROM nas_messages m
            JOIN spec_versions sv ON m.version_id = sv.id
        """
        params = []
        if version_ids:
            placeholders = ",".join("?" for _ in version_ids)
            query += f" WHERE m.version_id IN ({placeholders})"
            params.extend(version_ids)
        query += " GROUP BY m.message_name, m.clause ORDER BY m.message_name ASC"

        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(query, params)
            return [dict(row) for row in cursor.fetchall()]

    def get_messages_by_ie_search(
        self,
        ie_query: str,
        version_ids: Optional[List[int]] = None,
        search_descriptions: bool = False,
    ) -> List[Dict[str, Any]]:
        """
        Searches messages matching an IE name, field path, or clause 9 / ASN.1 description.
        Optimized with Common Table Expressions (CTE) to avoid full cross-join table scans in ASN.1.
        """
        if not version_ids or not ie_query.strip():
            return self.get_messages_list(version_ids)

        clean_query = ie_query.strip()
        placeholders = ",".join("?" for _ in version_ids)
        pattern = f"%{clean_query}%"

        # If search query is only 1 character, perform direct indexed IE/Type search to prevent full table cross-joins
        if search_descriptions and len(clean_query) >= 2:
            query = f"""
                WITH matched_defs AS (
                    SELECT version_id, ie_name, clause
                    FROM ie_definitions
                    WHERE version_id IN ({placeholders})
                      AND (raw_description LIKE ? OR ie_name LIKE ?)
                ),
                matching_messages AS (
                    -- 1. Direct Information Element / ASN.1 Field Name Matches
                    SELECT m.id AS message_id, m.version_id, m.message_name, m.clause
                    FROM nas_messages m
                    JOIN message_ies i ON i.message_id = m.id
                    WHERE m.version_id IN ({placeholders})
                      AND (
                          i.ie_name LIKE ? 
                          OR i.field_path LIKE ?
                          OR i.type_reference LIKE ? 
                          OR i.iei LIKE ?
                      )

                    UNION

                    -- 2. Matches via Associated Clause 9 or ASN.1 Field Descriptions
                    SELECT m.id AS message_id, m.version_id, m.message_name, m.clause
                    FROM nas_messages m
                    JOIN message_ies i ON i.message_id = m.id
                    JOIN matched_defs d ON d.version_id = m.version_id
                        AND (
                            i.ie_name = d.ie_name
                            OR i.type_reference = d.ie_name
                            OR i.type_reference = d.clause
                            OR (d.clause != '' AND i.type_reference = d.clause)
                            OR (length(d.ie_name) >= 4 AND instr(i.type_reference, d.ie_name) > 0)
                        )
                    WHERE m.version_id IN ({placeholders})
                )
                SELECT mm.message_name, mm.clause, GROUP_CONCAT(DISTINCT sv.spec_number) AS spec_number
                FROM matching_messages mm
                JOIN spec_versions sv ON mm.version_id = sv.id
                GROUP BY mm.message_name, mm.clause
                ORDER BY mm.message_name ASC
            """
            params = (
                list(version_ids) + [pattern, pattern]
                + list(version_ids) + [pattern, pattern, pattern, pattern]
                + list(version_ids)
            )
        else:
            query = f"""
                SELECT m.message_name, m.clause, GROUP_CONCAT(DISTINCT sv.spec_number) AS spec_number
                FROM nas_messages m
                JOIN message_ies i ON i.message_id = m.id
                JOIN spec_versions sv ON m.version_id = sv.id
                WHERE m.version_id IN ({placeholders})
                  AND (
                      i.ie_name LIKE ? 
                      OR i.field_path LIKE ?
                      OR i.type_reference LIKE ? 
                      OR i.iei LIKE ?
                  )
                GROUP BY m.message_name, m.clause
                ORDER BY m.message_name ASC
            """
            params = list(version_ids) + [pattern, pattern, pattern, pattern]

        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(query, params)
            return [dict(row) for row in cursor.fetchall()]

    def get_messages_using_ie(
        self,
        clause: Optional[str] = None,
        ie_name: Optional[str] = None,
        spec_number: Optional[str] = None,
        version_ids: Optional[List[int]] = None,
    ) -> List[Dict[str, Any]]:
        if not version_ids:
            return []

        placeholders = ",".join("?" for _ in version_ids)
        where_conditions = [f"m.version_id IN ({placeholders})"]
        params: List[Any] = list(version_ids)

        if spec_number:
            where_conditions.append("sv.spec_number = ?")
            params.append(spec_number.strip())

        clean_clause = clause.strip() if clause else ""
        clean_name = ie_name.strip() if ie_name else ""

        if clean_clause and re.match(r"^(?:9|6|D\.6)(?:\.[0-9A-Za-z]+)+$", clean_clause):
            regex_pattern = rf"(?<![0-9A-Za-z.]){re.escape(clean_clause)}(?![0-9A-Za-z.])"
            where_conditions.append("i.type_reference REGEXP ?")
            params.append(regex_pattern)
        elif clean_name or clean_clause:
            target = clean_name or clean_clause
            where_conditions.append("(LOWER(TRIM(i.ie_name)) = LOWER(?) OR i.type_reference LIKE ?)")
            params.extend([target, f"%{target}%"])
        else:
            return []

        where_clause = " AND ".join(where_conditions)
        query = f"""
            SELECT m.message_name, m.clause, GROUP_CONCAT(DISTINCT sv.spec_number) AS spec_number
            FROM nas_messages m
            JOIN message_ies i ON i.message_id = m.id
            JOIN spec_versions sv ON m.version_id = sv.id
            WHERE {where_clause}
            GROUP BY m.message_name, m.clause
            ORDER BY m.message_name ASC
        """

        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(query, params)
            return [dict(row) for row in cursor.fetchall()]

    def get_message_evolution_df(
        self,
        message_name: str,
        version_ids: List[int],
        include_descriptions: bool = False,
    ) -> pd.DataFrame:
        if not version_ids:
            return pd.DataFrame()

        placeholders = ",".join("?" for _ in version_ids)
        if include_descriptions:
            query = f"""
                SELECT 
                    sv.spec_number,
                    sv.version,
                    i.iei,
                    i.ie_name,
                    i.field_path,
                    i.depth,
                    i.type_reference,
                    i.presence,
                    i.format,
                    i.length,
                    i.order_index,
                    d.raw_description AS ie_description
                FROM message_ies i
                JOIN nas_messages m ON i.message_id = m.id
                JOIN spec_versions sv ON m.version_id = sv.id
                LEFT JOIN ie_definitions d ON d.version_id = sv.id
                    AND (
                        (d.ie_name != '' AND (d.ie_name = i.ie_name OR d.ie_name = i.type_reference))
                        OR (d.clause != '' AND d.clause = i.type_reference)
                    )
                WHERE m.message_name = ? AND sv.id IN ({placeholders})
                ORDER BY i.order_index ASC
            """
        else:
            query = f"""
                SELECT 
                    sv.spec_number,
                    sv.version,
                    i.iei,
                    i.ie_name,
                    i.field_path,
                    i.depth,
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
        self,
        clause: str,
        alt_name: str = "",
        spec_number: Optional[str] = None,
        version_ids: Optional[List[int]] = None,
    ) -> List[Dict[str, Any]]:
        params = []
        where_parts = []

        clean_c = clause.strip() if clause else ""
        clean_alt = alt_name.strip() if alt_name else ""

        if clean_c:
            where_parts.append("(LOWER(TRIM(d.clause)) = LOWER(?) OR LOWER(TRIM(d.ie_name)) = LOWER(?))")
            params.extend([clean_c, clean_c])

        if clean_alt:
            where_parts.append("(LOWER(TRIM(d.clause)) = LOWER(?) OR LOWER(TRIM(d.ie_name)) = LOWER(?))")
            params.extend([clean_alt, clean_alt])

        if not where_parts:
            return []

        where_sql = " OR ".join(where_parts)
        query = f"""
            SELECT d.*, sv.version, sv.spec_number
            FROM ie_definitions d
            JOIN spec_versions sv ON d.version_id = sv.id
            WHERE ({where_sql})
        """

        if spec_number:
            query += " AND sv.spec_number = ?"
            params.append(spec_number.strip())

        if version_ids:
            placeholders = ",".join("?" for _ in version_ids)
            query += f" AND d.version_id IN ({placeholders})"
            params.extend(version_ids)

        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(query, params)
            rows = [dict(row) for row in cursor.fetchall()]

        return sorted(rows, key=lambda x: parse_version_tuple(x["version"]), reverse=True)

    def get_latest_spec_version_by_major(
        self,
        spec_number: str,
        major_version: Optional[int] = None,
    ) -> Optional[Dict[str, Any]]:
        query = "SELECT id, spec_number, version, spec_type, import_date FROM spec_versions WHERE spec_number = ?"
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(query, (spec_number.strip(),))
            rows = [dict(r) for r in cursor.fetchall()]

        if not rows:
            return None

        if major_version is not None:
            rows = [
                r for r in rows
                if parse_version_tuple(r["version"]) and parse_version_tuple(r["version"])[0] == major_version
            ]

        if not rows:
            return None

        rows.sort(key=lambda r: parse_version_tuple(r["version"]), reverse=True)
        return rows[0]

    def get_cross_referenced_ie_definition(
        self,
        target_spec: str,
        target_clause: str,
        alt_name: str = "",
        major_version: Optional[int] = None,
    ) -> Optional[Dict[str, Any]]:
        latest_ver = self.get_latest_spec_version_by_major(target_spec, major_version)
        if not latest_ver:
            return None

        clean_c = target_clause.strip().rstrip(".") if target_clause else ""
        clean_alt = alt_name.strip() if alt_name else ""

        where_parts = []
        params: List[Any] = [latest_ver["id"]]

        if clean_c:
            where_parts.append("LOWER(TRIM(d.clause)) = LOWER(?)")
            params.append(clean_c)
        if clean_alt:
            where_parts.append("LOWER(TRIM(d.ie_name)) = LOWER(?)")
            params.append(clean_alt)

        if clean_c:
            where_parts.append("d.clause LIKE ?")
            params.append(f"%{clean_c}%")

        if not where_parts:
            return None

        where_sql = " OR ".join(where_parts)

        query = f"""
            SELECT d.*, sv.version, sv.spec_number
            FROM ie_definitions d
            JOIN spec_versions sv ON d.version_id = sv.id
            WHERE sv.id = ? AND ({where_sql})
            ORDER BY 
                CASE 
                    WHEN LOWER(TRIM(d.clause)) = LOWER('{clean_c}') THEN 1
                    WHEN LOWER(TRIM(d.ie_name)) = LOWER('{clean_alt}') THEN 2
                    ELSE 3
                END ASC,
                LENGTH(COALESCE(d.raw_description, '')) DESC
            LIMIT 1
        """
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(query, params)
            row = cursor.fetchone()
            return dict(row) if row else None