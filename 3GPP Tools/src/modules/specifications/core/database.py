# --- File: modules/specifications/core/database.py ---
import sqlite3
from pathlib import Path


class SpecsDatabase:
    def __init__(self, db_path: Path):
        self.db_path = db_path
        self._init_db()

    def _get_connection(self):
        return sqlite3.connect(self.db_path, check_same_thread=False)

    def _init_db(self):
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS series (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    name TEXT UNIQUE,
                    url TEXT
                )
            ''')
            # ---> NEW: Centralized Working Groups Table
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
                    primary_group_id INTEGER, -- ---> NEW: Replaced text with Foreign Key
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
            # ---> NEW: Many-to-Many Map for Secondary Groups
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS spec_secondary_group_map (
                    spec_id INTEGER,
                    group_id INTEGER,
                    UNIQUE(spec_id, group_id),
                    FOREIGN KEY(spec_id) REFERENCES specifications(id),
                    FOREIGN KEY(group_id) REFERENCES working_groups(id)
                )
            ''')

    def insert_or_update_file(self, series_name, series_url, spec_number, spec_url, filename, version, file_url):
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
                INSERT OR REPLACE INTO files (spec_id, filename, version, url)
                VALUES (?, ?, ?, ?)
            ''', (spec_id, filename, version, file_url))

    def update_spec_metadata(self, spec_number, metadata):
        with self._get_connection() as conn:
            cursor = conn.cursor()

            # 1. Handle Primary Group Foreign Key Resolution
            primary_group_id = None
            p_group = metadata.get('primary_group')
            if p_group:
                cursor.execute('INSERT OR IGNORE INTO working_groups (name) VALUES (?)', (p_group,))
                cursor.execute('SELECT id FROM working_groups WHERE name = ?', (p_group,))
                primary_group_id = cursor.fetchone()[0]

            # 2. Update standard fields
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
            if not spec_row: return
            spec_id = spec_row[0]

            # 3. Update Radio Technologies mapping
            techs = metadata.get('radio_technologies_list', [])
            for tech in techs:
                cursor.execute('INSERT OR IGNORE INTO radio_technologies (name) VALUES (?)', (tech,))
                cursor.execute('SELECT id FROM radio_technologies WHERE name = ?', (tech,))
                tech_id = cursor.fetchone()[0]
                cursor.execute('INSERT OR IGNORE INTO spec_radio_tech_map (spec_id, tech_id) VALUES (?, ?)',
                               (spec_id, tech_id))

            # 4. Update Secondary Groups mapping
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
                     series: str = None, tech: str = None, group: str = None, types: list = None) -> list:
        # Use DISTINCT to prevent duplicate rows if a spec has multiple secondary groups
        query = """
            SELECT DISTINCT s.name, sp.number, sp.title, sp.type, f.filename, f.version, f.url
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

        # 1. Standard Search
        if spec_number:
            query += " AND (sp.number LIKE ? OR sp.type LIKE ? OR (sp.type || ' ' || sp.number) LIKE ? OR sp.title LIKE ?)"
            search_term = f"%{spec_number}%"
            params.extend([search_term, search_term, search_term, search_term])

        if release_version:
            query += " AND f.version LIKE ?"
            params.append(f"%{release_version}%")

        # 2. Advanced Table Filters
        if series:
            series_list = [s.strip() for s in series.split(',') if s.strip()]
            if series_list:
                clauses = ["sp.number LIKE ?" for _ in series_list]
                params.extend([f"{s}.%" for s in series_list])
                query += f" AND ({' OR '.join(clauses)})"

        if tech:
            query += " AND (r.name LIKE ? OR sp.radio_technology LIKE ?)"
            params.extend([f"%{tech}%", f"%{tech}%"])

        if group:
            query += " AND (p_grp.name LIKE ? OR s_grp.name LIKE ?)"
            params.extend([f"%{group}%", f"%{group}%"])

        if types:
            clauses = ["sp.type = ?" for _ in types]
            params.extend(types)
            query += f" AND ({' OR '.join(clauses)})"

        query += " ORDER BY sp.number ASC, f.version DESC"

        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(query, params)
            return cursor.fetchall()

    def get_filtered_specs(self, series: str, tech: str, group: str, types: list) -> list:
        # ---> UPGRADED: Massive join across 6 tables for absolute filtering precision
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

        if series:
            series_list = [s.strip() for s in series.split(',') if s.strip()]
            if series_list:
                clauses = []
                for s in series_list:
                    clauses.append("sp.number LIKE ?")
                    params.append(f"{s}.%")
                query += f" AND ({' OR '.join(clauses)})"

        if tech:
            query += " AND (r.name LIKE ? OR sp.radio_technology LIKE ?)"
            params.extend([f"%{tech}%", f"%{tech}%"])

        if group:
            # Scans both Primary and Secondary group tables
            query += " AND (p_grp.name LIKE ? OR s_grp.name LIKE ?)"
            params.extend([f"%{group}%", f"%{group}%"])

        if types:
            clauses = []
            for t in types:
                clauses.append("sp.type = ?")
                params.append(t)
            query += f" AND ({' OR '.join(clauses)})"

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

            # 1. Resolve Primary Group ID to its Name
            if details.get('primary_group_id'):
                cursor.execute('SELECT name FROM working_groups WHERE id = ?', (details['primary_group_id'],))
                p_row = cursor.fetchone()
                if p_row:
                    details['primary_group'] = p_row[0]

            # Hide the raw ID from the UI Popup
            details.pop('primary_group_id', None)

            # 2. Extract Normalized Radio Techs
            cursor.execute('''
                SELECT r.name FROM radio_technologies r
                JOIN spec_radio_tech_map m ON r.id = m.tech_id
                JOIN specifications s ON s.id = m.spec_id
                WHERE s.number = ?
            ''', (spec_number,))
            techs = [r[0] for r in cursor.fetchall()]
            if techs:
                details['radio_technology'] = ", ".join(techs)

            # 3. Extract Normalized Secondary Groups
            cursor.execute('''
                SELECT w.name FROM working_groups w
                JOIN spec_secondary_group_map m ON w.id = m.group_id
                JOIN specifications s ON s.id = m.spec_id
                WHERE s.number = ?
            ''', (spec_number,))
            sec_groups = [r[0] for r in cursor.fetchall()]
            if sec_groups:
                details['secondary_groups'] = ", ".join(sec_groups)

            return details