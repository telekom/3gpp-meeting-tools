"""SQLite connection management, schema initialization, orphan cleanup, and maintenance."""

from __future__ import annotations

import contextlib
import sqlite3
from typing import Iterator


class SpecsSchemaMixin:
    """Handles SpecsDatabase SQLite connections, DDL, garbage collection, and filter lists."""

    @contextlib.contextmanager
    def _get_connection(self) -> Iterator[sqlite3.Connection]:
        """
        Creates a thread-safe connection with a 30s busy retry handler.
        PRAGMA journal_mode is intentionally omitted here because it requires a write lock
        and is already permanently established in _init_db().
        """
        conn = sqlite3.connect(self.db_path, timeout=30.0, check_same_thread=False)
        try:
            conn.execute("PRAGMA busy_timeout=30000;")
            with conn:
                yield conn
        finally:
            conn.close()

    def _init_db(self):
        with self._get_connection() as conn:
            cursor = conn.cursor()
            # Set WAL mode once during initial schema setup
            cursor.execute('PRAGMA journal_mode=WAL;')

            cursor.execute('''
                CREATE TABLE IF NOT EXISTS series (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    name TEXT UNIQUE,
                    url TEXT
                )
            ''')
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS working_groups (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    name TEXT UNIQUE
                )
            ''')
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS specifications (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    series_id INTEGER,
                    number TEXT,
                    url TEXT,
                    title TEXT,
                    type TEXT,
                    initial_release TEXT,
                    radio_technology TEXT,  
                    primary_group_id INTEGER,
                    secondary_groups TEXT,    
                    UNIQUE(series_id, number),
                    FOREIGN KEY(series_id) REFERENCES series(id),
                    FOREIGN KEY(primary_group_id) REFERENCES working_groups(id)
                )
            ''')
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS files (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    spec_id INTEGER,
                    filename TEXT,
                    version TEXT,
                    url TEXT,
                    upload_date TEXT,
                    UNIQUE(spec_id, version),
                    FOREIGN KEY(spec_id) REFERENCES specifications(id)
                )
            ''')
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS radio_technologies (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    name TEXT UNIQUE
                )
            ''')
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS spec_radio_tech_map (
                    spec_id INTEGER,
                    tech_id INTEGER,
                    UNIQUE(spec_id, tech_id),
                    FOREIGN KEY(spec_id) REFERENCES specifications(id),
                    FOREIGN KEY(tech_id) REFERENCES radio_technologies(id)
                )
            ''')
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS spec_secondary_group_map (
                    spec_id INTEGER,
                    group_id INTEGER,
                    UNIQUE(spec_id, group_id),
                    FOREIGN KEY(spec_id) REFERENCES specifications(id),
                    FOREIGN KEY(group_id) REFERENCES working_groups(id)
                )
            ''')

            # Relation table mapping specifications to Work Items
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS spec_wi_map (
                    spec_id INTEGER,
                    wi_code TEXT,
                    is_primary BOOLEAN DEFAULT 0,
                    UNIQUE(spec_id, wi_code),
                    FOREIGN KEY(spec_id) REFERENCES specifications(id)
                )
            ''')
            cursor.execute('CREATE INDEX IF NOT EXISTS idx_spec_wi_code ON spec_wi_map(wi_code);')

            # Performance indexes for fast spec lookup, joins, and file resolution
            cursor.execute('CREATE INDEX IF NOT EXISTS idx_specs_number ON specifications(number);')
            cursor.execute('CREATE INDEX IF NOT EXISTS idx_specs_series_id ON specifications(series_id);')
            cursor.execute('CREATE INDEX IF NOT EXISTS idx_specs_primary_group_id ON specifications(primary_group_id);')
            cursor.execute('CREATE INDEX IF NOT EXISTS idx_files_spec_id ON files(spec_id);')
            cursor.execute('CREATE INDEX IF NOT EXISTS idx_files_version ON files(version);')
            cursor.execute('CREATE INDEX IF NOT EXISTS idx_spec_radio_tech_spec ON spec_radio_tech_map(spec_id);')
            cursor.execute('CREATE INDEX IF NOT EXISTS idx_spec_radio_tech_tech ON spec_radio_tech_map(tech_id);')
            cursor.execute('CREATE INDEX IF NOT EXISTS idx_spec_sec_group_spec ON spec_secondary_group_map(spec_id);')
            cursor.execute('CREATE INDEX IF NOT EXISTS idx_spec_sec_group_group ON spec_secondary_group_map(group_id);')

            # User annotations (tags and comments) per specification
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS spec_user_annotations (
                    spec_number TEXT PRIMARY KEY,
                    tags TEXT,
                    comments TEXT,
                    updated_at TEXT
                )
            ''')
            cursor.execute('CREATE INDEX IF NOT EXISTS idx_spec_annotations_tags ON spec_user_annotations(tags);')

            # Dynamic migration: Ensure upload_date exists on older local databases
            cursor.execute("PRAGMA table_info(files)")
            file_cols = [col[1] for col in cursor.fetchall()]
            if 'upload_date' not in file_cols:
                cursor.execute("ALTER TABLE files ADD COLUMN upload_date TEXT")

    def _cleanup_orphans(self):
        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute('''
                    DELETE FROM radio_technologies 
                    WHERE id NOT IN (SELECT DISTINCT tech_id FROM spec_radio_tech_map WHERE tech_id IS NOT NULL)
                ''')
                cursor.execute('''
                    DELETE FROM working_groups 
                    WHERE id NOT IN (SELECT DISTINCT primary_group_id FROM specifications WHERE primary_group_id IS NOT NULL)
                      AND id NOT IN (SELECT DISTINCT group_id FROM spec_secondary_group_map WHERE group_id IS NOT NULL)
                ''')
                cursor.execute('''
                    DELETE FROM series 
                    WHERE id NOT IN (SELECT DISTINCT series_id FROM specifications WHERE series_id IS NOT NULL)
                ''')
        except Exception as e:
            self.logger.error(f"Error during specifications garbage collection: {e}")

    def vacuum(self) -> bool:
        try:
            with self._get_connection() as conn:
                conn.isolation_level = None
                conn.execute("PRAGMA wal_checkpoint(TRUNCATE);")
                conn.execute("VACUUM;")
                conn.execute("PRAGMA optimize;")
            return True
        except Exception as e:
            self.logger.error(f"Failed to vacuum Specs DB: {e}")
            return False

    def get_filter_options(self) -> dict:
        options = {'series': [], 'techs': [], 'groups': [], 'types': []}
        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute("SELECT name FROM series ORDER BY CAST(name AS INTEGER)")
                options['series'] = [r[0] for r in cursor.fetchall() if r[0]]

                cursor.execute("SELECT name FROM radio_technologies ORDER BY name")
                options['techs'] = [r[0] for r in cursor.fetchall() if r[0]]

                cursor.execute("SELECT name FROM working_groups ORDER BY name")
                options['groups'] = [r[0] for r in cursor.fetchall() if r[0]]

                cursor.execute(
                    "SELECT DISTINCT type FROM specifications WHERE type IS NOT NULL AND type != '' ORDER BY type")
                options['types'] = [r[0] for r in cursor.fetchall() if r[0]]
        except Exception as e:
            self.logger.error(f"Error fetching filter options: {e}")
        return options
