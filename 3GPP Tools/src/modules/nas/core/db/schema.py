# --- File: src/modules/nas/core/db/schema.py ---
"""Version sorting, regex constants, SQLite connection helper, and schema management."""

import contextlib
import re
import sqlite3
from typing import Any, Dict, Iterator, List

RE_NAS_IE_TYPE = re.compile(r"\btype\s+([1-6])\s+information\s+element\b", re.IGNORECASE)
RE_NAS_IE_MIN_MAX_LEN = re.compile(
    r"minimum\s+length\s+of\s+(\d+)\s+octets?(?:\s+and\s+a\s+maximum\s+length\s+of\s+(\d+|n)\s+octets?)?",
    re.IGNORECASE,
)
RE_NAS_IE_EXACT_LEN = re.compile(r"\blength\s+of\s+(\d+)\s+octets?\b", re.IGNORECASE)
RE_TRAILING_ASN1_CAPTION = re.compile(
    r"(?:\s+|^)(?:[A-Z][A-Za-z0-9]*(?:-[A-Za-z0-9]+|[A-Z0-9][A-Za-z0-9]*)[A-Za-z0-9\-]*)\s+"
    r"(?:information\s+elements?|messages?)\s*$"
)


def parse_version_tuple(version_str: str) -> tuple:
    """Converts a version string into a tuple of comparable elements for natural sorting."""
    if not version_str:
        return ()
    clean = str(version_str).lstrip("vV").strip()
    parts = []
    for part in clean.split("."):
        if part.isdigit():
            parts.append((0, int(part)))
        else:
            parts.append((1, str(part)))
    return tuple(parts)


class NasDbSchemaMixin:
    """Mixin managing SQLite connections, WAL initialization, migrations, and lifecycle cleanup."""

    @contextlib.contextmanager
    def _get_connection(self) -> Iterator[sqlite3.Connection]:
        conn = sqlite3.connect(str(self.db_path), check_same_thread=False)
        try:
            conn.row_factory = sqlite3.Row
            conn.execute("PRAGMA foreign_keys = ON;")
            conn.create_function(
                "REGEXP",
                2,
                lambda expr, item: bool(re.search(expr, str(item))) if item is not None else False,
            )
            with conn:
                yield conn
        finally:
            conn.close()

    def _init_db(self):
        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()

                # Set WAL mode once during initial schema setup
                cursor.execute('PRAGMA journal_mode=WAL;')

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
                        description TEXT DEFAULT '',
                        direction TEXT DEFAULT '',
                        FOREIGN KEY(version_id) REFERENCES spec_versions(id) ON DELETE CASCADE
                    )
                """)

                # Automatic schema migration for nas_messages
                cursor.execute("PRAGMA table_info(nas_messages);")
                existing_msg_cols = [row[1] for row in cursor.fetchall()]
                if "description" not in existing_msg_cols:
                    cursor.execute("ALTER TABLE nas_messages ADD COLUMN description TEXT DEFAULT '';")
                if "direction" not in existing_msg_cols:
                    cursor.execute("ALTER TABLE nas_messages ADD COLUMN direction TEXT DEFAULT '';")
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
                        applicability TEXT DEFAULT '',
                        FOREIGN KEY(message_id) REFERENCES nas_messages(id) ON DELETE CASCADE
                    )
                """)

                # Automatic schema migration for existing databases
                cursor.execute("PRAGMA table_info(message_ies);")
                existing_cols = [row[1] for row in cursor.fetchall()]
                if "applicability" not in existing_cols:
                    cursor.execute("ALTER TABLE message_ies ADD COLUMN applicability TEXT DEFAULT '';")
                if "field_description" not in existing_cols:
                    cursor.execute("ALTER TABLE message_ies ADD COLUMN field_description TEXT DEFAULT '';")

                cursor.execute("CREATE INDEX IF NOT EXISTS idx_ie_appl ON message_ies(applicability);")

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

                # Optimized Indices
                cursor.execute("CREATE INDEX IF NOT EXISTS idx_msg_ver ON nas_messages(version_id);")
                cursor.execute("CREATE INDEX IF NOT EXISTS idx_msg_name ON nas_messages(message_name);")
                cursor.execute("CREATE INDEX IF NOT EXISTS idx_msg_name_ver ON nas_messages(message_name, version_id);")
                cursor.execute("CREATE INDEX IF NOT EXISTS idx_ie_msg ON message_ies(message_id);")
                cursor.execute("CREATE INDEX IF NOT EXISTS idx_ie_msg_order ON message_ies(message_id, order_index);")
                cursor.execute("CREATE INDEX IF NOT EXISTS idx_ie_name ON message_ies(ie_name);")
                cursor.execute("CREATE INDEX IF NOT EXISTS idx_ie_type ON message_ies(type_reference);")
                cursor.execute("CREATE INDEX IF NOT EXISTS idx_def_ver ON ie_definitions(version_id);")
                cursor.execute("CREATE INDEX IF NOT EXISTS idx_def_name ON ie_definitions(ie_name);")
                cursor.execute("CREATE INDEX IF NOT EXISTS idx_def_clause ON ie_definitions(clause);")
                cursor.execute("CREATE INDEX IF NOT EXISTS idx_def_ver_name ON ie_definitions(version_id, ie_name);")
                cursor.execute("CREATE INDEX IF NOT EXISTS idx_def_ver_clause ON ie_definitions(version_id, clause);")
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
                conn.isolation_level = None
                conn.execute("PRAGMA wal_checkpoint(TRUNCATE);")
                conn.execute("VACUUM;")

            return True
        except Exception as e:
            self.logger.error(f"Failed to wipe Protocol DB: {e}")
            return False
