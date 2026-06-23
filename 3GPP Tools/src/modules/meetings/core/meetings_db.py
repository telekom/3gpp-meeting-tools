# --- File: modules/meetings/core/meetings_db.py ---
import sqlite3
import logging
import re
from pathlib import Path


class MeetingsDatabase:
    def __init__(self, db_path: Path):
        self.db_path = db_path
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        self._create_tables()

    def _get_connection(self):
        return sqlite3.connect(self.db_path)

    def _create_tables(self):
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS working_groups (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    name TEXT UNIQUE NOT NULL
                )
            ''')
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS meetings (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    wg_id INTEGER,
                    folder_name TEXT,
                    meeting_number TEXT,
                    name TEXT,
                    location TEXT,
                    start_date TEXT,
                    end_date TEXT,
                    url_key TEXT UNIQUE,
                    docs_folder_url TEXT,
                    first_tdoc TEXT,
                    last_tdoc TEXT,
                    FOREIGN KEY (wg_id) REFERENCES working_groups (id)
                )
            ''')

            # Graceful Schema Migration
            try:
                cursor.execute("ALTER TABLE meetings ADD COLUMN sort_number INTEGER DEFAULT 0")
            except sqlite3.OperationalError:
                pass

            conn.commit()

    def _extract_sort_num(self, m_str: str) -> int:
        match = re.search(r'\d+', m_str or "")
        return int(match.group()) if match else 0

    def get_or_create_wg(self, wg_name: str) -> int:
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute('INSERT OR IGNORE INTO working_groups (name) VALUES (?)', (wg_name,))
            cursor.execute('SELECT id FROM working_groups WHERE name = ?', (wg_name,))
            return cursor.fetchone()[0]

    def insert_meeting_basic(self, wg_name: str, folder_name: str, meeting_number: str, url_key: str):
        wg_id = self.get_or_create_wg(wg_name)
        sort_num = self._extract_sort_num(meeting_number)

        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute('''
                INSERT INTO meetings (wg_id, folder_name, meeting_number, sort_number, url_key)
                VALUES (?, ?, ?, ?, ?)
                ON CONFLICT(url_key) DO UPDATE SET
                    folder_name=excluded.folder_name,
                    meeting_number=excluded.meeting_number,
                    sort_number=excluded.sort_number
            ''', (wg_id, folder_name, meeting_number, sort_num, url_key))
            conn.commit()

    # --- PHASE 1 BULK ---
    def insert_meetings_bulk(self, meetings_data: list):
        if not meetings_data: return

        wg_map = {}
        for task in meetings_data:
            wg = task['wg_name']
            if wg not in wg_map:
                wg_map[wg] = self.get_or_create_wg(wg)

        insert_data = [
            (wg_map[task['wg_name']], task['folder_name'], task['meeting_num'],
             self._extract_sort_num(task['meeting_num']), task['url_key'])
            for task in meetings_data
        ]

        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.executemany('''
                INSERT INTO meetings (wg_id, folder_name, meeting_number, sort_number, url_key)
                VALUES (?, ?, ?, ?, ?)
                ON CONFLICT(url_key) DO UPDATE SET
                    folder_name=excluded.folder_name,
                    meeting_number=excluded.meeting_number,
                    sort_number=excluded.sort_number
            ''', insert_data)
            conn.commit()

    def update_meeting_docs(self, url_key: str, docs_url: str, first_tdoc: str, last_tdoc: str):
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute('''
                UPDATE meetings 
                SET docs_folder_url = ?, first_tdoc = ?, last_tdoc = ?
                WHERE url_key = ?
            ''', (docs_url, first_tdoc, last_tdoc, url_key))
            conn.commit()

    # --- PHASE 2 BULK ---
    def update_meeting_docs_bulk(self, docs_data: list):
        """Bulk updates the documents info. Expects a list of tuples: (docs_url, first_tdoc, last_tdoc, url_key)"""
        if not docs_data: return
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.executemany('''
                UPDATE meetings 
                SET docs_folder_url = ?, first_tdoc = ?, last_tdoc = ?
                WHERE url_key = ?
            ''', docs_data)
            conn.commit()

    def update_meeting_metadata_pass2(self, wg_name: str, meeting_number: str, name: str, location: str,
                                      start_date: str, end_date: str):
        wg_id = self.get_or_create_wg(wg_name)
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute('''
                UPDATE meetings 
                SET name = ?, location = ?, start_date = ?, end_date = ?
                WHERE wg_id = ? AND meeting_number = ?
            ''', (name, location, start_date, end_date, wg_id, meeting_number))
            conn.commit()

    # --- PHASE 3 BULK ---
    def update_meeting_metadata_bulk(self, metadata_data: list):
        """Bulk updates DynaReport metadata. Expects a list of tuples: (wg_name, meeting_number, name, location, start, end)"""
        if not metadata_data: return

        wg_map = {}
        for item in metadata_data:
            wg = item[0]
            if wg not in wg_map:
                wg_map[wg] = self.get_or_create_wg(wg)

        # Reformat the tuple to match the SQL parameters layout
        update_data = [
            (item[2], item[3], item[4], item[5], wg_map[item[0]], item[1])
            for item in metadata_data
        ]

        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.executemany('''
                UPDATE meetings 
                SET name = ?, location = ?, start_date = ?, end_date = ?
                WHERE wg_id = ? AND meeting_number = ?
            ''', update_data)
            conn.commit()

    def search_meetings(self, wg_name=None, search_term=None, location=None, date_from=None, date_to=None):
        query = '''
            SELECT m.*, w.name as wg_name 
            FROM meetings m
            JOIN working_groups w ON m.wg_id = w.id
            WHERE 1=1
        '''
        params = []
        if wg_name and wg_name != "All WGs":
            query += " AND w.name = ?"
            params.append(wg_name)
        if search_term:
            query += " AND (m.meeting_number LIKE ? OR m.name LIKE ?)"
            params.extend([f"%{search_term}%", f"%{search_term}%"])
        if location:
            query += " AND m.location LIKE ?"
            params.append(f"%{location}%")
        if date_from:
            query += " AND m.start_date >= ?"
            params.append(date_from)
        if date_to:
            query += " AND m.end_date <= ?"
            params.append(date_to)

        query += " ORDER BY m.sort_number DESC, m.meeting_number DESC"

        with self._get_connection() as conn:
            conn.row_factory = sqlite3.Row
            cursor = conn.cursor()
            cursor.execute(query, params)
            return [dict(row) for row in cursor.fetchall()]

    def get_working_groups(self):
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute('SELECT name FROM working_groups ORDER BY name')
            return [row[0] for row in cursor.fetchall()]

    def delete_all_meetings(self):
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute('DELETE FROM meetings')
            conn.commit()

    def delete_specific_meetings(self, targets: list):
        with self._get_connection() as conn:
            cursor = conn.cursor()
            for t in targets:
                cursor.execute('''
                    DELETE FROM meetings 
                    WHERE wg_id = (SELECT id FROM working_groups WHERE name = ?) AND meeting_number = ?
                ''', (t["wg"], t["meeting"]))
            conn.commit()