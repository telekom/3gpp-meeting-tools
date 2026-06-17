# --- File: modules/specs_db/database.py ---
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
                    primary_group TEXT,
                    secondary_groups TEXT,
                    UNIQUE(series_id, number),
                    FOREIGN KEY(series_id) REFERENCES series(id)
                )
            ''')
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS files (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    spec_id INTEGER,
                    filename TEXT,
                    version TEXT,
                    url TEXT,
                    UNIQUE(spec_id, filename),
                    FOREIGN KEY(spec_id) REFERENCES specifications(id)
                )
            ''')
            conn.commit()

    def insert_or_update_file(self, series_name, series_url, spec_number, spec_url, filename, version, file_url):
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute('INSERT OR IGNORE INTO series (name, url) VALUES (?, ?)', (series_name, series_url))
            cursor.execute('SELECT id FROM series WHERE name = ?', (series_name,))
            series_id = cursor.fetchone()[0]

            cursor.execute('INSERT OR IGNORE INTO specifications (series_id, number, url) VALUES (?, ?, ?)',
                           (series_id, spec_number, spec_url))
            cursor.execute('SELECT id FROM specifications WHERE series_id = ? AND number = ?', (series_id, spec_number))
            spec_id = cursor.fetchone()[0]

            cursor.execute('INSERT OR IGNORE INTO files (spec_id, filename, version, url) VALUES (?, ?, ?, ?)',
                           (spec_id, filename, version, file_url))
            conn.commit()

    def needs_metadata(self, spec_number: str) -> bool:
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute('SELECT title FROM specifications WHERE number = ?', (spec_number,))
            row = cursor.fetchone()
            return not row or not row[0]

    def update_spec_metadata(self, spec_number: str, metadata: dict):
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute('''
                UPDATE specifications 
                SET title = ?, type = ?, initial_release = ?, radio_technology = ?, primary_group = ?, secondary_groups = ?
                WHERE number = ?
            ''', (
                metadata.get('title', ''), metadata.get('type', ''),
                metadata.get('initial_release', ''), metadata.get('radio_technology', ''),
                metadata.get('primary_group', ''), metadata.get('secondary_groups', ''),
                spec_number
            ))
            conn.commit()

    def search_files(self, spec_number: str = None, release_version: str = None) -> list:
        query = """
            SELECT s.name, sp.number, sp.title, f.filename, f.version, f.url
            FROM files f
            JOIN specifications sp ON f.spec_id = sp.id
            JOIN series s ON sp.series_id = s.id
            WHERE 1=1
        """
        params = []
        if spec_number:
            query += " AND sp.number LIKE ?"
            params.append(f"%{spec_number}%")
        if release_version:
            query += " AND f.version LIKE ?"
            params.append(f"{release_version}%")

        query += " ORDER BY sp.number ASC, f.version DESC"
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(query, params)
            return cursor.fetchall()

    def get_all_specifications(self) -> list:
        """Returns a list of all unique specifications: (number, title, type)."""
        query = """
            SELECT DISTINCT sp.number, sp.title, sp.type
            FROM specifications sp
            ORDER BY sp.number ASC
        """
        with self._get_connection() as conn:
            return conn.cursor().execute(query).fetchall()

    def get_versions_for_spec(self, spec_number: str) -> list:
        """Returns all file versions and URLs for a specific spec number."""
        query = """
            SELECT f.version, f.url, f.filename
            FROM files f
            JOIN specifications sp ON f.spec_id = sp.id
            WHERE sp.number = ?
            ORDER BY f.version DESC
        """
        with self._get_connection() as conn:
            return conn.cursor().execute(query, (spec_number,)).fetchall()