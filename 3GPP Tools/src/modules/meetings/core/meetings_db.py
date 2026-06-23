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

            # --- Graceful Schema Migrations ---
            try:
                cursor.execute("ALTER TABLE meetings ADD COLUMN sort_number INTEGER DEFAULT 0")
            except sqlite3.OperationalError:
                pass
            try:
                cursor.execute("ALTER TABLE meetings ADD COLUMN is_ad_hoc INTEGER DEFAULT 0")
            except sqlite3.OperationalError:
                pass
            try:
                cursor.execute("ALTER TABLE meetings ADD COLUMN is_electronic INTEGER DEFAULT 0")
            except sqlite3.OperationalError:
                pass

            try:
                cursor.execute("ALTER TABLE meetings ADD COLUMN first_tdoc_prefix TEXT")
            except sqlite3.OperationalError:
                pass
            try:
                cursor.execute("ALTER TABLE meetings ADD COLUMN first_tdoc_num INTEGER DEFAULT 0")
            except sqlite3.OperationalError:
                pass
            try:
                cursor.execute("ALTER TABLE meetings ADD COLUMN last_tdoc_prefix TEXT")
            except sqlite3.OperationalError:
                pass
            try:
                cursor.execute("ALTER TABLE meetings ADD COLUMN last_tdoc_num INTEGER DEFAULT 0")
            except sqlite3.OperationalError:
                pass

            try:
                cursor.execute("ALTER TABLE meetings ADD COLUMN mtg_id TEXT")
            except sqlite3.OperationalError:
                pass

            conn.commit()

    def _extract_sort_num(self, m_str: str) -> int:
        match = re.search(r'\d+', m_str or "")
        return int(match.group()) if match else 0

    def _get_meeting_flags(self, m_str: str):
        num_upper = (m_str or "").upper()
        is_ad_hoc = 1 if ("A" in num_upper or "BIS" in num_upper) else 0
        is_electronic = 1 if ("E" in num_upper) else 0
        return is_ad_hoc, is_electronic

    def get_or_create_wg(self, wg_name: str) -> int:
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute('INSERT OR IGNORE INTO working_groups (name) VALUES (?)', (wg_name,))
            cursor.execute('SELECT id FROM working_groups WHERE name = ?', (wg_name,))
            return cursor.fetchone()[0]

    def insert_meeting_basic(self, wg_name: str, folder_name: str, meeting_number: str, url_key: str):
        wg_id = self.get_or_create_wg(wg_name)
        sort_num = self._extract_sort_num(meeting_number)
        is_ad_hoc, is_electronic = self._get_meeting_flags(meeting_number)

        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute('''
                INSERT INTO meetings (wg_id, folder_name, meeting_number, sort_number, is_ad_hoc, is_electronic, url_key)
                VALUES (?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(url_key) DO UPDATE SET
                    folder_name=excluded.folder_name,
                    meeting_number=excluded.meeting_number,
                    sort_number=excluded.sort_number,
                    is_ad_hoc=excluded.is_ad_hoc,
                    is_electronic=excluded.is_electronic
            ''', (wg_id, folder_name, meeting_number, sort_num, is_ad_hoc, is_electronic, url_key))
            conn.commit()

    def insert_meetings_bulk(self, meetings_data: list):
        if not meetings_data: return

        wg_map = {}
        for task in meetings_data:
            wg = task['wg_name']
            if wg not in wg_map:
                wg_map[wg] = self.get_or_create_wg(wg)

        insert_data = []
        for task in meetings_data:
            m_num = task['meeting_num']
            is_ah, is_e = self._get_meeting_flags(m_num)
            final_ah = 1 if (is_ah or task.get('is_ad_hoc')) else 0

            insert_data.append((
                wg_map[task['wg_name']], task['folder_name'], m_num,
                self._extract_sort_num(m_num), final_ah, is_e,
                task['url_key'], task.get('docs_url', '')
            ))

        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.executemany('''
                INSERT INTO meetings (wg_id, folder_name, meeting_number, sort_number, is_ad_hoc, is_electronic, url_key, docs_folder_url)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(url_key) DO UPDATE SET
                    folder_name=excluded.folder_name,
                    meeting_number=excluded.meeting_number,
                    sort_number=excluded.sort_number,
                    is_ad_hoc=excluded.is_ad_hoc,
                    is_electronic=excluded.is_electronic,
                    docs_folder_url=excluded.docs_folder_url
            ''', insert_data)
            conn.commit()

    def update_meeting_docs_bulk(self, docs_data: list):
        if not docs_data: return

        formatted_data = [
            (
                d[0],
                d[1], d[1], d[1], d[2], d[1], d[3],
                d[4], d[4], d[4], d[5], d[4], d[6],
                d[7]
            ) for d in docs_data
        ]

        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.executemany('''
                UPDATE meetings 
                SET docs_folder_url = ?, 
                    first_tdoc = CASE WHEN ? != '' THEN ? ELSE first_tdoc END,
                    first_tdoc_prefix = CASE WHEN ? != '' THEN ? ELSE first_tdoc_prefix END,
                    first_tdoc_num = CASE WHEN ? != '' THEN ? ELSE first_tdoc_num END,
                    last_tdoc = CASE WHEN ? != '' THEN ? ELSE last_tdoc END,
                    last_tdoc_prefix = CASE WHEN ? != '' THEN ? ELSE last_tdoc_prefix END,
                    last_tdoc_num = CASE WHEN ? != '' THEN ? ELSE last_tdoc_num END
                WHERE url_key = ?
            ''', formatted_data)
            conn.commit()

    def update_meeting_metadata_bulk(self, metadata_data: list):
        if not metadata_data: return

        wg_map = {}
        for item in metadata_data:
            wg = item[0]
            if wg not in wg_map:
                wg_map[wg] = self.get_or_create_wg(wg)

        formatted_data = []
        for item in metadata_data:
            wg_name, m_num, url_key, mtg_id, m_name, town, start_d, end_d = item
            wg_id = wg_map[wg_name]

            # Ensure None falls back to a clean string
            mtg_id = mtg_id or ""
            m_name = m_name or ""
            town = town or ""
            start_d = start_d or ""
            end_d = end_d or ""

            # Prevent accidental empty string matches across the DB
            url_key = url_key or "NO_MATCH_URL"
            m_num = m_num or "NO_MATCH_NUM"

            formatted_data.append((
                mtg_id, mtg_id,
                m_name, m_name,
                town, town,
                start_d, start_d,
                end_d, end_d,
                wg_id, url_key, m_num
            ))

        with self._get_connection() as conn:
            cursor = conn.cursor()
            # --- FIXED: Dual-Match (Matches URL OR Meeting Number!) ---
            cursor.executemany('''
                UPDATE meetings 
                SET mtg_id = CASE WHEN ? != '' THEN ? ELSE mtg_id END,
                    name = CASE WHEN ? != '' THEN ? ELSE name END,
                    location = CASE WHEN ? != '' THEN ? ELSE location END,
                    start_date = CASE WHEN ? != '' THEN ? ELSE start_date END,
                    end_date = CASE WHEN ? != '' THEN ? ELSE end_date END
                WHERE wg_id = ? 
                  AND (
                      LOWER(RTRIM(url_key, '/')) = LOWER(RTRIM(?, '/'))
                      OR UPPER(meeting_number) = UPPER(?)
                  )
            ''', formatted_data)
            conn.commit()

    def search_meetings(self, wg_name=None, search_term=None, location=None, date_from=None, date_to=None,
                        adhoc_filter=None, type_filter=None):
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

        if adhoc_filter == "Ad-Hoc / BIS":
            query += " AND m.is_ad_hoc = 1"
        elif adhoc_filter == "Regular":
            query += " AND m.is_ad_hoc = 0"

        if type_filter == "Electronic":
            query += " AND m.is_electronic = 1"
        elif type_filter == "In-Person":
            query += " AND m.is_electronic = 0"

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