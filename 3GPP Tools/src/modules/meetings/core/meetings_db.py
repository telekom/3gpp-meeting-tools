# --- File: modules/meetings/core/meetings_db.py ---
import sqlite3
from pathlib import Path
import logging


class MeetingsDatabase:
    def __init__(self, db_path: Path):
        self.db_path = db_path
        self._init_db()

    def _get_connection(self):
        return sqlite3.connect(self.db_path, check_same_thread=False)

    def _init_db(self):
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute('PRAGMA journal_mode=WAL;')

            # Ensure working_groups exists (it should, from specs, but we make sure)
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS working_groups (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    name TEXT UNIQUE
                )
            ''')

            # Create the new meetings table
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS meetings (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    wg_id INTEGER,
                    folder_name TEXT,
                    meeting_number TEXT,
                    url_key TEXT UNIQUE,
                    name TEXT,
                    location TEXT,
                    start_date TEXT,
                    end_date TEXT,
                    first_tdoc TEXT,
                    last_tdoc TEXT,
                    docs_folder_url TEXT,
                    FOREIGN KEY(wg_id) REFERENCES working_groups(id)
                )
            ''')

    def get_or_create_wg(self, wg_name: str) -> int:
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute('INSERT OR IGNORE INTO working_groups (name) VALUES (?)', (wg_name,))
            cursor.execute('SELECT id FROM working_groups WHERE name = ?', (wg_name,))
            return cursor.fetchone()[0]

    def insert_or_update_meeting_pass1(self, wg_name: str, folder_name: str, meeting_number: str,
                                       url_key: str, docs_url: str, first_tdoc: str, last_tdoc: str):
        wg_id = self.get_or_create_wg(wg_name)
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute('''
                INSERT INTO meetings (wg_id, folder_name, meeting_number, url_key, docs_folder_url, first_tdoc, last_tdoc)
                VALUES (?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(url_key) DO UPDATE SET
                    folder_name=excluded.folder_name,
                    meeting_number=excluded.meeting_number,
                    docs_folder_url=excluded.docs_folder_url,
                    first_tdoc=excluded.first_tdoc,
                    last_tdoc=excluded.last_tdoc
            ''', (wg_id, folder_name, meeting_number, url_key, docs_url, first_tdoc, last_tdoc))

    def update_meeting_metadata_pass2(self, wg_name: str, meeting_number: str, name: str,
                                      location: str, start_date: str, end_date: str):
        wg_id = self.get_or_create_wg(wg_name)
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute('''
                UPDATE meetings 
                SET name = ?, location = ?, start_date = ?, end_date = ?
                WHERE wg_id = ? AND meeting_number = ?
            ''', (name, location, start_date, end_date, wg_id, meeting_number))

    def search_meetings(self, wg_name: str = None, search_term: str = None,
                        location: str = None, date_from: str = None, date_to: str = None) -> list:
        query = """
            SELECT m.id, wg.name as wg_name, m.meeting_number, m.name, m.location, 
                   m.start_date, m.end_date, m.first_tdoc, m.last_tdoc, m.url_key, m.docs_folder_url
            FROM meetings m
            JOIN working_groups wg ON m.wg_id = wg.id
            WHERE 1=1
        """
        params = []

        if wg_name and wg_name != "All WGs":
            query += " AND wg.name = ?"
            params.append(wg_name)

        if search_term:
            query += " AND (m.meeting_number LIKE ? OR m.name LIKE ?)"
            term = f"%{search_term}%"
            params.extend([term, term])

        if location:
            query += " AND m.location LIKE ?"
            params.append(f"%{location}%")

        if date_from:
            query += " AND m.start_date >= ?"
            params.append(date_from)

        if date_to:
            query += " AND m.end_date <= ?"
            params.append(date_to)

        query += " ORDER BY m.start_date DESC"

        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(query, params)

            columns = [col[0] for col in cursor.description]
            return [dict(zip(columns, row)) for row in cursor.fetchall()]

    def get_working_groups(self) -> list:
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute('''
                SELECT DISTINCT wg.name 
                FROM working_groups wg
                JOIN meetings m ON wg.id = m.wg_id
                ORDER BY wg.name
            ''')
            return [r[0] for r in cursor.fetchall()]

    def delete_all_meetings(self):
        """Wipes all entries from the meetings table."""
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute('DELETE FROM meetings')

    def delete_specific_meetings(self, targets: list):
        """Deletes specific meetings based on WG and meeting number."""
        with self._get_connection() as conn:
            cursor = conn.cursor()
            for target in targets:
                cursor.execute('''
                    DELETE FROM meetings 
                    WHERE wg_id = (SELECT id FROM working_groups WHERE name = ?) 
                    AND meeting_number = ?
                ''', (target['wg'], target['meeting']))

    def insert_meeting_basic(self, wg_name: str, folder_name: str, meeting_number: str, url_key: str):
        """Phase 1: Safely inserts directory info without overwriting existing TDoc data."""
        wg_id = self.get_or_create_wg(wg_name)
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute('''
                INSERT INTO meetings (wg_id, folder_name, meeting_number, url_key)
                VALUES (?, ?, ?, ?)
                ON CONFLICT(url_key) DO UPDATE SET
                    folder_name=excluded.folder_name,
                    meeting_number=excluded.meeting_number
            ''', (wg_id, folder_name, meeting_number, url_key))

    def update_meeting_docs(self, url_key: str, docs_url: str, first_tdoc: str, last_tdoc: str):
        """Phase 2: Updates only the TDoc references."""
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute('''
                UPDATE meetings 
                SET docs_folder_url = ?, first_tdoc = ?, last_tdoc = ?
                WHERE url_key = ?
            ''', (docs_url, first_tdoc, last_tdoc, url_key))

        # --- Add this to modules/meetings/core/meetings_db.py ---

        def insert_meetings_bulk(self, meetings_data: list):
            """Phase 1: Safely bulk-inserts directory info to avoid disk I/O bottlenecks."""
            if not meetings_data:
                return

            # 1. Pre-fetch or create all WG IDs to minimize DB queries
            wg_map = {}
            for task in meetings_data:
                wg = task['wg_name']
                if wg not in wg_map:
                    wg_map[wg] = self.get_or_create_wg(wg)

            # 2. Prepare the flat tuple data for the bulk execution
            insert_data = [
                (wg_map[task['wg_name']], task['folder_name'], task['meeting_num'], task['url_key'])
                for task in meetings_data
            ]

            # 3. Execute all ~2400 inserts in ONE single transaction
            with self._get_connection() as conn:
                cursor = conn.cursor()
                cursor.executemany('''
                    INSERT INTO meetings (wg_id, folder_name, meeting_number, url_key)
                    VALUES (?, ?, ?, ?)
                    ON CONFLICT(url_key) DO UPDATE SET
                        folder_name=excluded.folder_name,
                        meeting_number=excluded.meeting_number
                ''', insert_data)

    def insert_meetings_bulk(self, meetings_data: list):
        """Phase 1: Safely bulk-inserts directory info to avoid disk I/O bottlenecks."""
        if not meetings_data:
            return

        # 1. Pre-fetch or create all WG IDs to minimize DB queries
        wg_map = {}
        for task in meetings_data:
            wg = task['wg_name']
            if wg not in wg_map:
                wg_map[wg] = self.get_or_create_wg(wg)

        # 2. Prepare the flat tuple data for the bulk execution
        insert_data = [
            (wg_map[task['wg_name']], task['folder_name'], task['meeting_num'], task['url_key'])
            for task in meetings_data
        ]

        # 3. Execute all ~2400 inserts in ONE single transaction
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.executemany('''
                INSERT INTO meetings (wg_id, folder_name, meeting_number, url_key)
                VALUES (?, ?, ?, ?)
                ON CONFLICT(url_key) DO UPDATE SET
                    folder_name=excluded.folder_name,
                    meeting_number=excluded.meeting_number
            ''', insert_data)
            conn.commit()  # Ensure changes are locked into the hard drive