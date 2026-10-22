# --- File: src/modules/specifications/core/database.py ---
import contextlib
import logging
import sqlite3
from datetime import datetime
from pathlib import Path
from typing import Dict, Iterator, List, Optional, Set, Union

from modules.specifications.utils.utils import build_dynareport_url


class SpecsDatabase:
    def __init__(self, db_path: Path):
        self.db_path = db_path
        self.logger = logging.getLogger(__name__)
        self._init_db()
        self._cleanup_orphans()

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

    def insert_or_update_file(self, series_name: str, series_url: str, spec_number: str,
                              spec_url: str, filename: str, version: str, file_url: str,
                              upload_date: Optional[str] = None):
        # Archive directories contain a small number of historical/non-standard ZIPs
        # whose filenames do not encode a valid 3GPP version. They are not selectable
        # specification versions and must not enter the version table.
        clean_version = (version or "").strip()
        if not clean_version:
            self.logger.debug(
                "Skipping specification archive file without a decoded version: %s (%s)",
                filename, spec_number
            )
            return

        version = clean_version

        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute('INSERT OR IGNORE INTO series (name, url) VALUES (?, ?)', (series_name, series_url))
            cursor.execute('SELECT id FROM series WHERE name = ?', (series_name,))
            series_id = cursor.fetchone()[0]

            cursor.execute('''
                INSERT OR IGNORE INTO specifications (series_id, number, url) 
                VALUES (?, ?, ?)
            ''', (series_id, spec_number, spec_url))
            cursor.execute('SELECT id FROM specifications WHERE number = ?', (spec_number,))
            spec_id = cursor.fetchone()[0]

            cursor.execute('''
                INSERT INTO files (spec_id, filename, version, url, upload_date)
                VALUES (?, ?, ?, ?, ?)
                ON CONFLICT(spec_id, version) DO UPDATE SET
                    filename = excluded.filename,
                    url = excluded.url,
                    upload_date = COALESCE(excluded.upload_date, files.upload_date)
            ''', (spec_id, filename, version, file_url, upload_date))

    def update_file_dates(self, spec_number: str, version_date_map: Dict[str, str]):
        """Batch-updates the portal upload dates for all matched versions of a specification."""
        if not version_date_map:
            return
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute('SELECT id FROM specifications WHERE number = ?', (spec_number,))
            spec_row = cursor.fetchone()
            if not spec_row:
                return
            spec_id = spec_row[0]

            for version, date_str in version_date_map.items():
                clean_ver = version.lstrip('v').strip()
                cursor.execute('''
                    UPDATE files 
                    SET upload_date = ?
                    WHERE spec_id = ? AND (version = ? OR version = ?)
                ''', (date_str, spec_id, clean_ver, f"v{clean_ver}"))

    def update_spec_wis(self, spec_number: str, wis_list: List[Dict]):
        """
        Maps related Work Items to a specification and inserts stubs into work_items
        if the WI has not been fully synchronized yet.
        """
        if wis_list is None:
            return

        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute('SELECT id FROM specifications WHERE number = ?', (spec_number,))
            spec_row = cursor.fetchone()
            if not spec_row:
                return
            spec_id = spec_row[0]

            # Clear existing mappings for this spec to prevent stale links
            cursor.execute('DELETE FROM spec_wi_map WHERE spec_id = ?', (spec_id,))

            for wi in wis_list:
                code = str(wi.get('code', '')).strip()
                if not code:
                    continue

                acronym = wi.get('acronym', '')
                title = wi.get('title', '')
                is_primary = 1 if wi.get('is_primary') else 0

                cursor.execute('''
                    INSERT INTO work_items (code, acronym, name, latest_wid, release, start_date, end_date)
                    VALUES (?, ?, ?, '', '', '', '')
                    ON CONFLICT(code) DO UPDATE SET
                        acronym = CASE WHEN (acronym IS NULL OR acronym = '') THEN excluded.acronym ELSE acronym END,
                        name = CASE WHEN (name IS NULL OR name = '') THEN excluded.name ELSE name END
                ''', (code, acronym, title))

                cursor.execute('''
                    INSERT OR REPLACE INTO spec_wi_map (spec_id, wi_code, is_primary)
                    VALUES (?, ?, ?)
                ''', (spec_id, code, is_primary))

    def get_spec_wis(self, spec_number: str) -> List[Dict]:
        """Fetches all Work Items mapped to a given specification."""
        query = """
            SELECT m.wi_code, m.is_primary, 
                   COALESCE(w.acronym, '') AS acronym, 
                   COALESCE(w.name, '') AS name, 
                   COALESCE(w.release, '') AS release
            FROM spec_wi_map m
            JOIN specifications s ON m.spec_id = s.id
            LEFT JOIN work_items w ON m.wi_code = w.code
            WHERE s.number = ?
            ORDER BY m.is_primary DESC, m.wi_code ASC
        """
        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute(query, (spec_number,))
                cols = [c[0] for c in cursor.description]
                return [dict(zip(cols, row)) for row in cursor.fetchall()]
        except sqlite3.OperationalError:
            return []

    def update_spec_metadata(self, spec_number: str, metadata: dict):
        """Apply successfully parsed metadata without erasing fields that failed to parse."""
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute('SELECT id FROM specifications WHERE number = ?', (spec_number,))
            spec_row = cursor.fetchone()
            if not spec_row:
                return
            spec_id = spec_row[0]

            scalar_updates = []
            scalar_values = []
            for column, key in (
                ('title', 'title'), ('type', 'type'),
                ('initial_release', 'initial_release'),
                ('radio_technology', 'radio_technology'),
                ('secondary_groups', 'secondary_groups_raw'),
            ):
                value = metadata.get(key)
                if value is not None:
                    scalar_updates.append(f'{column} = ?')
                    scalar_values.append(value)

            if metadata.get('primary_group') is not None:
                primary_group_id = None
                p_group = metadata.get('primary_group')
                if p_group:
                    cursor.execute('INSERT OR IGNORE INTO working_groups (name) VALUES (?)', (p_group,))
                    cursor.execute('SELECT id FROM working_groups WHERE name = ?', (p_group,))
                    primary_group_id = cursor.fetchone()[0]
                scalar_updates.append('primary_group_id = ?')
                scalar_values.append(primary_group_id)

            if scalar_updates:
                cursor.execute(
                    f"UPDATE specifications SET {', '.join(scalar_updates)} WHERE number = ?",
                    (*scalar_values, spec_number)
                )

            techs = metadata.get('radio_technologies_list')
            if techs is not None:
                cursor.execute('DELETE FROM spec_radio_tech_map WHERE spec_id = ?', (spec_id,))
                for tech in techs:
                    cursor.execute('INSERT OR IGNORE INTO radio_technologies (name) VALUES (?)', (tech,))
                    cursor.execute('SELECT id FROM radio_technologies WHERE name = ?', (tech,))
                    tech_id = cursor.fetchone()[0]
                    cursor.execute('INSERT OR IGNORE INTO spec_radio_tech_map (spec_id, tech_id) VALUES (?, ?)',
                                   (spec_id, tech_id))

            sec_groups = metadata.get('secondary_groups_list')
            if sec_groups is not None:
                cursor.execute('DELETE FROM spec_secondary_group_map WHERE spec_id = ?', (spec_id,))
                for sg in sec_groups:
                    cursor.execute('INSERT OR IGNORE INTO working_groups (name) VALUES (?)', (sg,))
                    cursor.execute('SELECT id FROM working_groups WHERE name = ?', (sg,))
                    sg_id = cursor.fetchone()[0]
                    cursor.execute('INSERT OR IGNORE INTO spec_secondary_group_map (spec_id, group_id) VALUES (?, ?)',
                                   (spec_id, sg_id))

    def needs_metadata(self, spec_number: str) -> bool:
        """Return True when core DynaReport metadata is absent or clearly incomplete."""
        query = """
            SELECT title, type, primary_group_id
            FROM specifications
            WHERE number = ?
        """
        with self._get_connection() as conn:
            result = conn.cursor().execute(query, (spec_number,)).fetchone()
            if not result:
                return True
            title, spec_type, primary_group_id = result
            return not title or not spec_type or primary_group_id is None

    def search_files(self, spec_number: str = None, release_version: str = None,
                     series: str = None, tech: str = None, group: str = None, spec_type: str = None,
                     tag: str = None) -> list:
        join_tech = bool(tech and tech != "Any")
        join_group = bool(group and group != "Any")
        use_distinct = join_tech or join_group

        select_clause = "SELECT DISTINCT" if use_distinct else "SELECT"
        query = f"""
            {select_clause} s.name, sp.number, sp.title, sp.type, f.filename, f.version, f.url, f.upload_date,
                            COALESCE(ua.tags, ''), COALESCE(ua.comments, ''), COALESCE(p_grp.name, '')
            FROM files f
            JOIN specifications sp ON f.spec_id = sp.id
            JOIN series s ON sp.series_id = s.id
            LEFT JOIN working_groups p_grp ON sp.primary_group_id = p_grp.id
            LEFT JOIN spec_user_annotations ua ON sp.number = ua.spec_number
        """
        if join_tech:
            query += """
            LEFT JOIN spec_radio_tech_map r_map ON sp.id = r_map.spec_id
            LEFT JOIN radio_technologies r ON r_map.tech_id = r.id
            """
        if join_group:
            query += """
            LEFT JOIN spec_secondary_group_map sg_map ON sp.id = sg_map.spec_id
            LEFT JOIN working_groups s_grp ON sg_map.group_id = s_grp.id
            """
        query += """
            WHERE 1=1
              AND f.version IS NOT NULL
              AND TRIM(f.version) != ''
        """
        params = []

        if spec_number:
            query += " AND (sp.number LIKE ? OR sp.type LIKE ? OR (sp.type || ' ' || sp.number) LIKE ? OR sp.title LIKE ? OR ua.tags LIKE ? OR ua.comments LIKE ?)"
            search_term = f"%{spec_number}%"
            params.extend([search_term, search_term, search_term, search_term, search_term, search_term])

        if tag and tag != "Any":
            query += " AND (ua.tags LIKE ? OR ua.tags = ?)"
            params.extend([f"%{tag}%", tag])

        if release_version:
            query += " AND f.version LIKE ?"
            params.append(f"%{release_version}%")

        if series:
            series_list = [s.strip() for s in series.split(',') if s.strip()]
            if series_list:
                clauses = ["sp.number LIKE ?" for _ in series_list]
                params.extend([f"{s}.%" for s in series_list])
                query += f" AND ({' OR '.join(clauses)})"

        if tech and tech != "Any":
            query += " AND (r.name = ? OR sp.radio_technology LIKE ?)"
            params.extend([tech, f"%{tech}%"])

        # Non-intrusive prefix matching for TSG groups (e.g., 'RAN' matches 'RAN', 'RAN3', etc.)
        if group and group != "Any":
            group_list = [g.strip() for g in group.split(',') if g.strip()]
            if group_list:
                group_clauses = []
                for g in group_list:
                    group_clauses.append("(p_grp.name = ? OR p_grp.name LIKE ? OR s_grp.name = ? OR s_grp.name LIKE ?)")
                    params.extend([g, f"{g}%", g, f"{g}%"])
                query += f" AND ({' OR '.join(group_clauses)})"

        if spec_type and spec_type != "Any":
            query += " AND sp.type = ?"
            params.append(spec_type)

        query += " ORDER BY sp.number ASC, f.version DESC"

        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(query, params)
            return cursor.fetchall()

    def get_filtered_specs(self, series: str, tech: str, group: str, spec_type: str, tag: str = None) -> list:
        join_tech = bool(tech and tech != "Any")
        join_group = bool(group and group != "Any")
        query = """
            SELECT DISTINCT sp.number 
            FROM specifications sp
            LEFT JOIN spec_user_annotations ua ON sp.number = ua.spec_number
        """
        if join_tech:
            query += """
            LEFT JOIN spec_radio_tech_map r_map ON sp.id = r_map.spec_id
            LEFT JOIN radio_technologies r ON r_map.tech_id = r.id
            """
        if join_group:
            query += """
            LEFT JOIN working_groups p_grp ON sp.primary_group_id = p_grp.id
            LEFT JOIN spec_secondary_group_map sg_map ON sp.id = sg_map.spec_id
            LEFT JOIN working_groups s_grp ON sg_map.group_id = s_grp.id
            """
        query += " WHERE 1=1"
        params = []

        if series and series != "Any":
            series_list = [s.strip() for s in series.split(',') if s.strip()]
            if series_list:
                clauses = [f"sp.number LIKE ?" for _ in series_list]
                params.extend([f"{s}.%" for s in series_list])
                query += f" AND ({' OR '.join(clauses)})"

        if tech and tech != "Any":
            query += " AND (r.name = ? OR sp.radio_technology LIKE ?)"
            params.extend([tech, f"%{tech}%"])

        # Non-intrusive prefix matching for TSG groups
        if group and group != "Any":
            group_list = [g.strip() for g in group.split(',') if g.strip()]
            if group_list:
                group_clauses = []
                for g in group_list:
                    group_clauses.append("(p_grp.name = ? OR p_grp.name LIKE ? OR s_grp.name = ? OR s_grp.name LIKE ?)")
                    params.extend([g, f"{g}%", g, f"{g}%"])
                query += f" AND ({' OR '.join(group_clauses)})"

        if spec_type and spec_type != "Any":
            query += " AND sp.type = ?"
            params.append(spec_type)

        if tag and tag != "Any":
            query += " AND (ua.tags LIKE ? OR ua.tags = ?)"
            params.extend([f"%{tag}%", tag])

        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(query, params)
            return [row[0] for row in cursor.fetchall()]

    def get_spec_details(self, spec_number: str) -> dict:
        query = "SELECT * FROM specifications WHERE number = ?"
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(query, (spec_number,))
            row = cursor.fetchone()
            if not row:
                return {}

            columns = [description[0] for description in cursor.description]
            details = dict(zip(columns, row))

            if details.get('primary_group_id'):
                cursor.execute('SELECT name FROM working_groups WHERE id = ?', (details['primary_group_id'],))
                p_row = cursor.fetchone()
                if p_row:
                    details['primary_group'] = p_row[0]
            details.pop('primary_group_id', None)

            cursor.execute('''
                SELECT r.name FROM radio_technologies r
                JOIN spec_radio_tech_map m ON r.id = m.tech_id
                JOIN specifications s ON s.id = m.spec_id
                WHERE s.number = ?
            ''', (spec_number,))
            techs = [r[0] for r in cursor.fetchall()]
            if techs:
                details['radio_technology'] = ", ".join(techs)

            cursor.execute('''
                SELECT w.name FROM working_groups w
                JOIN spec_secondary_group_map m ON w.id = m.group_id
                JOIN specifications s ON s.id = m.spec_id
                WHERE s.number = ?
            ''', (spec_number,))
            sec_groups = [r[0] for r in cursor.fetchall()]
            if sec_groups:
                details['secondary_groups'] = ", ".join(sec_groups)

            details['related_wis'] = self.get_spec_wis(spec_number)
            details['user_annotations'] = self.get_spec_annotations(spec_number)

            return details

    def get_spec_files(self, spec_number: str) -> List[Dict]:
        """Returns all archive file entries for a specification, sorted by semantic version descending."""
        query = """
            SELECT f.version, f.filename, f.url, f.upload_date
            FROM files f
            JOIN specifications s ON f.spec_id = s.id
            WHERE s.number = ?
        """
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(query, (spec_number,))
            rows = cursor.fetchall()
            if not rows and "-" in spec_number:
                cursor.execute(query, (spec_number.split("-")[0],))
                rows = cursor.fetchall()

            def parse_ver(v_str):
                return [(0, int(x)) if x.isdigit() else (1, str(x)) for x in str(v_str).split(".")]

            sorted_rows = sorted(rows, key=lambda x: parse_ver(x[0]), reverse=True)
            return [
                {"version": r[0], "filename": r[1], "url": r[2], "upload_date": r[3]}
                for r in sorted_rows
            ]

    def delete_specification(self, spec_number: str) -> bool:
        """Deletes a specification, its file entries, and relational mappings, then removes orphans."""
        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute("SELECT id FROM specifications WHERE number = ?", (spec_number,))
                row = cursor.fetchone()
                if not row:
                    return False
                spec_id = row[0]

                cursor.execute("DELETE FROM files WHERE spec_id = ?", (spec_id,))
                cursor.execute("DELETE FROM spec_radio_tech_map WHERE spec_id = ?", (spec_id,))
                cursor.execute("DELETE FROM spec_secondary_group_map WHERE spec_id = ?", (spec_id,))
                cursor.execute("DELETE FROM spec_wi_map WHERE spec_id = ?", (spec_id,))
                cursor.execute("DELETE FROM specifications WHERE id = ?", (spec_id,))

            self._cleanup_orphans()
            return True
        except Exception as e:
            self.logger.error(f"Failed to delete specification {spec_number}: {e}")
            return False

    def wipe_database(self) -> bool:
        """Completely purges all specifications, files, and relational mapping tables."""
        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute("DELETE FROM files;")
                cursor.execute("DELETE FROM spec_wi_map;")
                cursor.execute("DELETE FROM spec_secondary_group_map;")
                cursor.execute("DELETE FROM spec_radio_tech_map;")
                cursor.execute("DELETE FROM specifications;")
                cursor.execute("DELETE FROM working_groups;")
                cursor.execute("DELETE FROM radio_technologies;")
                cursor.execute("DELETE FROM series;")
            self.vacuum()
            return True
        except Exception as e:
            self.logger.error(f"Failed to wipe specifications database: {e}")
            return False

    def upsert_manual_spec(self, spec_data: dict) -> bool:
        """Inserts or updates a specification record from the manual addition dialog."""
        spec_number = str(spec_data.get("number", "")).strip()
        if not spec_number:
            return False

        series_name = spec_data.get("series")
        if not series_name:
            series_name = spec_number.split(".")[0].strip()

        title = spec_data.get("title", "").strip()
        spec_type = spec_data.get("type", "TS").strip().upper()
        primary_group = spec_data.get("primary_group", "").strip()
        initial_release = spec_data.get("initial_release", "").strip()
        radio_tech = spec_data.get("radio_technology", "").strip()

        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()
                # Ensure series exists
                cursor.execute(
                    "INSERT OR IGNORE INTO series (name, url) VALUES (?, ?)",
                    (series_name, f"https://www.3gpp.org/ftp/Specs/archive/{series_name}_series/")
                )
                cursor.execute("SELECT id FROM series WHERE name = ?", (series_name,))
                series_id = cursor.fetchone()[0]

                # Ensure primary group exists
                primary_group_id = None
                if primary_group:
                    cursor.execute("INSERT OR IGNORE INTO working_groups (name) VALUES (?)", (primary_group,))
                    cursor.execute("SELECT id FROM working_groups WHERE name = ?", (primary_group,))
                    primary_group_id = cursor.fetchone()[0]

                # Insert or update specification
                dyna_url = build_dynareport_url(spec_number)
                cursor.execute("""
                    INSERT INTO specifications (series_id, number, url, title, type, initial_release, radio_technology, primary_group_id)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                    ON CONFLICT(series_id, number) DO UPDATE SET
                        title = CASE WHEN excluded.title != '' THEN excluded.title ELSE specifications.title END,
                        type = excluded.type,
                        initial_release = CASE WHEN excluded.initial_release != '' THEN excluded.initial_release ELSE specifications.initial_release END,
                        radio_technology = CASE WHEN excluded.radio_technology != '' THEN excluded.radio_technology ELSE specifications.radio_technology END,
                        primary_group_id = COALESCE(excluded.primary_group_id, specifications.primary_group_id)
                """, (
                series_id, spec_number, dyna_url, title, spec_type, initial_release, radio_tech, primary_group_id))

                cursor.execute("SELECT id FROM specifications WHERE number = ?", (spec_number,))
                spec_id = cursor.fetchone()[0]

                # Link radio tech
                if radio_tech:
                    cursor.execute("INSERT OR IGNORE INTO radio_technologies (name) VALUES (?)", (radio_tech,))
                    cursor.execute("SELECT id FROM radio_technologies WHERE name = ?", (radio_tech,))
                    tech_id = cursor.fetchone()[0]
                    cursor.execute("INSERT OR IGNORE INTO spec_radio_tech_map (spec_id, tech_id) VALUES (?, ?)",
                                   (spec_id, tech_id))

            return True
        except Exception as e:
            self.logger.error(f"Failed to upsert manual specification: {e}")
            return False

    # -------------------------------------------------------------------------
    # User Annotations (Tags & Comments)
    # -------------------------------------------------------------------------
    def get_spec_annotations(self, spec_number: str) -> dict:
        """Returns user tags and comments for a specification."""
        clean_spec = (spec_number or "").strip()
        if not clean_spec:
            return {"tags": [], "comments": "", "updated_at": ""}
        query = "SELECT tags, comments, updated_at FROM spec_user_annotations WHERE spec_number = ?"
        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute(query, (clean_spec,))
                row = cursor.fetchone()
                if not row:
                    return {"tags": [], "comments": "", "updated_at": ""}
                tags_raw = row[0] or ""
                tags_list = [t.strip() for t in tags_raw.split(",") if t.strip()]
                return {
                    "tags": tags_list,
                    "comments": row[1] or "",
                    "updated_at": row[2] or "",
                }
        except Exception as e:
            self.logger.error(f"Error fetching annotations for {clean_spec}: {e}")
            return {"tags": [], "comments": "", "updated_at": ""}

    def get_all_annotations(self) -> Dict[str, dict]:
        """Returns all annotations keyed by spec_number."""
        query = "SELECT spec_number, tags, comments, updated_at FROM spec_user_annotations"
        result = {}
        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute(query)
                for s_num, tags_raw, comments, updated_at in cursor.fetchall():
                    tags_list = [t.strip() for t in (tags_raw or "").split(",") if t.strip()]
                    result[s_num] = {
                        "tags": tags_list,
                        "comments": comments or "",
                        "updated_at": updated_at or "",
                    }
        except Exception as e:
            self.logger.error(f"Error fetching all annotations: {e}")
        return result

    def get_all_unique_tags(self) -> List[str]:
        """Returns a sorted list of all unique tags used across specifications."""
        query = "SELECT tags FROM spec_user_annotations WHERE tags IS NOT NULL AND tags != ''"
        all_tags: Set[str] = set()
        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute(query)
                for (tags_str,) in cursor.fetchall():
                    for t in (tags_str or "").split(","):
                        t_clean = t.strip()
                        if t_clean:
                            all_tags.add(t_clean)
            return sorted(list(all_tags), key=lambda x: x.lower())
        except Exception as e:
            self.logger.error(f"Error fetching unique tags: {e}")
            return []

    def set_spec_annotations(self, spec_number: str, tags: Union[List[str], str], comments: str = "") -> bool:
        """Upserts user tags and comments for a specification."""
        spec_clean = (spec_number or "").strip()
        if not spec_clean:
            return False

        if isinstance(tags, list):
            clean_tags_list = sorted(list({t.strip() for t in tags if t.strip()}))
        else:
            clean_tags_list = sorted(list({t.strip() for t in (tags or "").split(",") if t.strip()}))
        clean_tags = ", ".join(clean_tags_list)

        clean_comments = (comments or "").strip()
        now_iso = datetime.now().isoformat(timespec="seconds")

        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()
                if not clean_tags and not clean_comments:
                    cursor.execute("DELETE FROM spec_user_annotations WHERE spec_number = ?", (spec_clean,))
                else:
                    cursor.execute("""
                        INSERT INTO spec_user_annotations (spec_number, tags, comments, updated_at)
                        VALUES (?, ?, ?, ?)
                        ON CONFLICT(spec_number) DO UPDATE SET
                            tags = excluded.tags,
                            comments = excluded.comments,
                            updated_at = excluded.updated_at
                    """, (spec_clean, clean_tags, clean_comments, now_iso))
            return True
        except Exception as e:
            self.logger.error(f"Failed to set annotations for {spec_clean}: {e}")
            return False

    def add_tag_to_specs(self, spec_numbers: List[str], tag: str) -> int:
        """Adds a tag to multiple specifications, preserving existing tags and comments."""
        clean_tag = (tag or "").strip()
        if not clean_tag:
            return 0

        count = 0
        for s_num in spec_numbers:
            ann = self.get_spec_annotations(s_num)
            tags = set(ann.get("tags", []))
            if clean_tag not in tags:
                tags.add(clean_tag)
                if self.set_spec_annotations(s_num, list(tags), ann.get("comments", "")):
                    count += 1
        return count

    def remove_tag_from_specs(self, spec_numbers: List[str], tag: str) -> int:
        """Removes a tag from multiple specifications."""
        clean_tag = (tag or "").strip().lower()
        if not clean_tag:
            return 0

        count = 0
        for s_num in spec_numbers:
            ann = self.get_spec_annotations(s_num)
            current_tags = ann.get("tags", [])
            filtered_tags = [t for t in current_tags if t.lower() != clean_tag]
            if len(filtered_tags) != len(current_tags):
                if self.set_spec_annotations(s_num, filtered_tags, ann.get("comments", "")):
                    count += 1
        return count

    def delete_spec_annotations(self, spec_number: str) -> bool:
        """Deletes user annotations for a specification."""
        spec_clean = (spec_number or "").strip()
        if not spec_clean:
            return False
        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute("DELETE FROM spec_user_annotations WHERE spec_number = ?", (spec_clean,))
            return True
        except Exception as e:
            self.logger.error(f"Failed to delete annotations for {spec_clean}: {e}")
            return False
