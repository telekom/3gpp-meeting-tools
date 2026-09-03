# --- File: src/modules/specifications/core/database.py ---
import logging
import sqlite3
from pathlib import Path
from typing import Dict, List, Optional


class SpecsDatabase:
    def __init__(self, db_path: Path):
        self.db_path = db_path
        self.logger = logging.getLogger(__name__)
        self._init_db()
        self._cleanup_orphans()

    def _get_connection(self) -> sqlite3.Connection:
        """
        Creates a thread-safe connection with a 30s busy retry handler.
        PRAGMA journal_mode is intentionally omitted here because it requires a write lock
        and is already permanently established in _init_db().
        """
        conn = sqlite3.connect(self.db_path, timeout=30.0, check_same_thread=False)
        conn.execute("PRAGMA busy_timeout=30000;")
        return conn

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
        if not wis_list:
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
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(query, (spec_number,))
            cols = [c[0] for c in cursor.description]
            return [dict(zip(cols, row)) for row in cursor.fetchall()]

    def update_spec_metadata(self, spec_number: str, metadata: dict):
        with self._get_connection() as conn:
            cursor = conn.cursor()
            primary_group_id = None
            p_group = metadata.get('primary_group')
            if p_group:
                cursor.execute('INSERT OR IGNORE INTO working_groups (name) VALUES (?)', (p_group,))
                cursor.execute('SELECT id FROM working_groups WHERE name = ?', (p_group,))
                primary_group_id = cursor.fetchone()[0]

            cursor.execute('''
                UPDATE specifications 
                SET title = ?, type = ?, initial_release = ?, radio_technology = ?, 
                    primary_group_id = ?, secondary_groups = ?
                WHERE number = ?
            ''', (
                metadata.get('title'), metadata.get('type'), metadata.get('initial_release'),
                metadata.get('radio_technology'), primary_group_id, metadata.get('secondary_groups_raw'),
                spec_number
            ))

            cursor.execute('SELECT id FROM specifications WHERE number = ?', (spec_number,))
            spec_row = cursor.fetchone()
            if not spec_row:
                return
            spec_id = spec_row[0]

            cursor.execute('DELETE FROM spec_radio_tech_map WHERE spec_id = ?', (spec_id,))
            cursor.execute('DELETE FROM spec_secondary_group_map WHERE spec_id = ?', (spec_id,))

            techs = metadata.get('radio_technologies_list', [])
            for tech in techs:
                cursor.execute('INSERT OR IGNORE INTO radio_technologies (name) VALUES (?)', (tech,))
                cursor.execute('SELECT id FROM radio_technologies WHERE name = ?', (tech,))
                tech_id = cursor.fetchone()[0]
                cursor.execute('INSERT OR IGNORE INTO spec_radio_tech_map (spec_id, tech_id) VALUES (?, ?)',
                               (spec_id, tech_id))

            sec_groups = metadata.get('secondary_groups_list', [])
            for sg in sec_groups:
                cursor.execute('INSERT OR IGNORE INTO working_groups (name) VALUES (?)', (sg,))
                cursor.execute('SELECT id FROM working_groups WHERE name = ?', (sg,))
                sg_id = cursor.fetchone()[0]
                cursor.execute('INSERT OR IGNORE INTO spec_secondary_group_map (spec_id, group_id) VALUES (?, ?)',
                               (spec_id, sg_id))

    def needs_metadata(self, spec_number: str) -> bool:
        query = "SELECT title FROM specifications WHERE number = ?"
        with self._get_connection() as conn:
            result = conn.cursor().execute(query, (spec_number,)).fetchone()
            return not result or not result[0]

    def search_files(self, spec_number: str = None, release_version: str = None,
                     series: str = None, tech: str = None, group: str = None, spec_type: str = None) -> list:
        query = """
            SELECT DISTINCT s.name, sp.number, sp.title, sp.type, f.filename, f.version, f.url, f.upload_date
            FROM files f
            JOIN specifications sp ON f.spec_id = sp.id
            JOIN series s ON sp.series_id = s.id
            LEFT JOIN spec_radio_tech_map r_map ON sp.id = r_map.spec_id
            LEFT JOIN radio_technologies r ON r_map.tech_id = r.id
            LEFT JOIN working_groups p_grp ON sp.primary_group_id = p_grp.id
            LEFT JOIN spec_secondary_group_map sg_map ON sp.id = sg_map.spec_id
            LEFT JOIN working_groups s_grp ON sg_map.group_id = s_grp.id
            WHERE 1=1
        """
        params = []

        if spec_number:
            query += " AND (sp.number LIKE ? OR sp.type LIKE ? OR (sp.type || ' ' || sp.number) LIKE ? OR sp.title LIKE ?)"
            search_term = f"%{spec_number}%"
            params.extend([search_term, search_term, search_term, search_term])

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
            query += " AND (p_grp.name = ? OR p_grp.name LIKE ? OR s_grp.name = ? OR s_grp.name LIKE ?)"
            params.extend([group, f"{group}%", group, f"{group}%"])

        if spec_type and spec_type != "Any":
            query += " AND sp.type = ?"
            params.append(spec_type)

        query += " ORDER BY sp.number ASC, f.version DESC"

        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(query, params)
            return cursor.fetchall()

    def get_filtered_specs(self, series: str, tech: str, group: str, spec_type: str) -> list:
        query = """
            SELECT DISTINCT sp.number 
            FROM specifications sp
            LEFT JOIN spec_radio_tech_map r_map ON sp.id = r_map.spec_id
            LEFT JOIN radio_technologies r ON r_map.tech_id = r.id
            LEFT JOIN working_groups p_grp ON sp.primary_group_id = p_grp.id
            LEFT JOIN spec_secondary_group_map sg_map ON sp.id = sg_map.spec_id
            LEFT JOIN working_groups s_grp ON sg_map.group_id = s_grp.id
            WHERE 1=1
        """
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
            query += " AND (p_grp.name = ? OR p_grp.name LIKE ? OR s_grp.name = ? OR s_grp.name LIKE ?)"
            params.extend([group, f"{group}%", group, f"{group}%"])

        if spec_type and spec_type != "Any":
            query += " AND sp.type = ?"
            params.append(spec_type)

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

            return details

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
                clean_num = spec_number.replace(".", "")
                dyna_url = f"https://www.3gpp.org/DynaReport/{clean_num}.htm"
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