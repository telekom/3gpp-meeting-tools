# --- File: modules/meetings/core/meetings_db.py ---
import datetime
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
        # 30-second timeout allows background writers to finish without crashing readers
        conn = sqlite3.connect(self.db_path, timeout=30.0)
        # Enable WAL mode for concurrent multi-thread read/write support
        conn.execute("PRAGMA journal_mode=WAL;")
        conn.execute("PRAGMA busy_timeout=30000;")
        return conn

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
            (d[0], d[1], d[1], d[1], d[2], d[1], d[3], d[4], d[4], d[4], d[5], d[4], d[6], d[7])
            for d in docs_data
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

        with self._get_connection() as conn:
            cursor = conn.cursor()
            for item in metadata_data:
                wg_name, m_num, url_key, mtg_id, m_name, town, start_d, end_d, new_m_num = item
                wg_id = wg_map[wg_name]

                mtg_id = mtg_id or ""
                m_name = m_name or ""
                town = town or ""
                start_d = start_d or ""
                end_d = end_d or ""
                m_num = m_num or ""
                new_m_num = new_m_num or ""

                # Safe SQLite Null Handling
                db_url_key = url_key.strip() if url_key and url_key.strip() else None

                final_m_num = new_m_num if new_m_num else m_num
                sort_n = self._extract_sort_num(final_m_num)
                is_ah, is_e = self._get_meeting_flags(final_m_num)

                if wg_name.startswith("RAN") and (
                        re.search(r'AH|Ad\s*Hoc', final_m_num, re.IGNORECASE) or re.search(r'Ad\s*Hoc', m_name,
                                                                                           re.IGNORECASE)):
                    is_ah = 1

                row_id = None

                # 1. Attempt strict match via FTP URL
                if db_url_key:
                    cursor.execute("SELECT id FROM meetings WHERE LOWER(RTRIM(url_key, '/')) = LOWER(RTRIM(?, '/'))",
                                   (db_url_key,))
                    res = cursor.fetchone()
                    if res: row_id = res[0]

                # 2. Attempt aggressive fallback match via the calculated Meeting Number
                if not row_id and m_num:
                    cursor.execute('''
                        SELECT id FROM meetings 
                        WHERE wg_id = ? AND (
                            UPPER(meeting_number) = UPPER(?) 
                            OR UPPER(meeting_number) = UPPER('AH' || ?) 
                            OR UPPER(meeting_number) = UPPER(REPLACE(?, ' ', ''))
                            OR UPPER(REPLACE(meeting_number, '-', '')) = UPPER(REPLACE(?, '-', ''))
                        )
                    ''', (wg_id, m_num, m_num, m_name, m_name))
                    res = cursor.fetchone()
                    if res: row_id = res[0]

                # 3. Safe UPSERT
                if row_id:
                    try:
                        cursor.execute('''
                                            UPDATE meetings 
                                            SET mtg_id = CASE WHEN ? != '' THEN ? ELSE mtg_id END,
                                                name = CASE WHEN ? != '' THEN ? ELSE name END,
                                                location = CASE WHEN ? != '' THEN ? ELSE location END,
                                                start_date = CASE WHEN ? != '' THEN ? ELSE start_date END,
                                                end_date = CASE WHEN ? != '' THEN ? ELSE end_date END,
                                                meeting_number = CASE WHEN ? != '' THEN ? ELSE meeting_number END,
                                                sort_number = CASE WHEN ? != 0 THEN ? ELSE sort_number END,
                                                is_ad_hoc = CASE WHEN ? = 1 THEN 1 ELSE is_ad_hoc END,
                                                is_electronic = CASE WHEN ? = 1 THEN 1 ELSE is_electronic END,
                                                url_key = CASE WHEN (url_key IS NULL OR url_key = '') AND ? IS NOT NULL THEN ? ELSE url_key END
                                            WHERE id = ?
                                        ''',
                                       (mtg_id, mtg_id, m_name, m_name, town, town, start_d, start_d, end_d, end_d,
                                        new_m_num, new_m_num, sort_n, sort_n, is_ah, is_e, db_url_key, db_url_key,
                                        row_id))
                    except sqlite3.IntegrityError:
                        cursor.execute('''
                                            UPDATE meetings 
                                            SET mtg_id = CASE WHEN ? != '' THEN ? ELSE mtg_id END,
                                                name = CASE WHEN ? != '' THEN ? ELSE name END,
                                                location = CASE WHEN ? != '' THEN ? ELSE location END,
                                                start_date = CASE WHEN ? != '' THEN ? ELSE start_date END,
                                                end_date = CASE WHEN ? != '' THEN ? ELSE end_date END,
                                                meeting_number = CASE WHEN ? != '' THEN ? ELSE meeting_number END,
                                                sort_number = CASE WHEN ? != 0 THEN ? ELSE sort_number END,
                                                is_ad_hoc = CASE WHEN ? = 1 THEN 1 ELSE is_ad_hoc END,
                                                is_electronic = CASE WHEN ? = 1 THEN 1 ELSE is_electronic END
                                            WHERE id = ?
                                        ''',
                                       (mtg_id, mtg_id, m_name, m_name, town, town, start_d, start_d, end_d, end_d,
                                        new_m_num, new_m_num, sort_n, sort_n, is_ah, is_e, row_id))

            conn.commit()

    def upsert_single_meeting(self, data: dict) -> bool:
        """Inserts or completely updates a single meeting entry in the database."""
        wg_name = data.get("wg_name", "").strip()
        if not wg_name:
            return False

        wg_id = self.get_or_create_wg(wg_name)
        meeting_number = data.get("meeting_number", "").strip()
        folder_name = data.get("folder_name", "").strip() or meeting_number
        name = data.get("name", "").strip()
        location = data.get("location", "").strip()
        start_date = data.get("start_date", "").strip()
        end_date = data.get("end_date", "").strip()
        mtg_id = str(data.get("mtg_id", "")).strip()

        # Clean and normalize url_key
        url_key = data.get("url_key", "").strip()
        if url_key.startswith("https://www.3gpp.org/ftp/"):
            url_key = url_key.replace("https://www.3gpp.org/ftp/", "")
        elif url_key.startswith("http://www.3gpp.org/ftp/"):
            url_key = url_key.replace("http://www.3gpp.org/ftp/", "")
        url_key = url_key.strip('/')
        db_url_key = url_key if url_key else None

        docs_folder_url = data.get("docs_folder_url", "").strip()
        if not docs_folder_url and db_url_key:
            docs_folder_url = f"https://www.3gpp.org/ftp/{db_url_key}/Docs/"

        first_tdoc = data.get("first_tdoc", "").strip()
        first_tdoc_prefix = data.get("first_tdoc_prefix", "").strip()
        first_tdoc_num = int(data.get("first_tdoc_num") or 0)
        if first_tdoc and not first_tdoc_prefix:
            m = re.match(r'^([A-Za-z0-9]+)-?(\d+)', first_tdoc)
            if m:
                first_tdoc_prefix = m.group(1).upper()
                first_tdoc_num = int(m.group(2))

        last_tdoc = data.get("last_tdoc", "").strip()
        last_tdoc_prefix = data.get("last_tdoc_prefix", "").strip()
        last_tdoc_num = int(data.get("last_tdoc_num") or 0)
        if last_tdoc and not last_tdoc_prefix:
            m = re.match(r'^([A-Za-z0-9]+)-?(\d+)', last_tdoc)
            if m:
                last_tdoc_prefix = m.group(1).upper()
                last_tdoc_num = int(m.group(2))

        sort_num = self._extract_sort_num(meeting_number)
        is_ad_hoc, is_electronic = self._get_meeting_flags(meeting_number)
        if data.get("is_ad_hoc") is not None:
            is_ad_hoc = int(bool(data.get("is_ad_hoc")))
        if data.get("is_electronic") is not None:
            is_electronic = int(bool(data.get("is_electronic")))

        with self._get_connection() as conn:
            cursor = conn.cursor()
            row_id = None

            # 1. Match by url_key
            if db_url_key:
                cursor.execute("SELECT id FROM meetings WHERE LOWER(RTRIM(url_key, '/')) = LOWER(RTRIM(?, '/'))",
                               (db_url_key,))
                res = cursor.fetchone()
                if res: row_id = res[0]

            # 2. Match by mtg_id
            if not row_id and mtg_id:
                cursor.execute("SELECT id FROM meetings WHERE mtg_id = ?", (mtg_id,))
                res = cursor.fetchone()
                if res: row_id = res[0]

            # 3. Match by WG + Meeting number
            if not row_id and meeting_number:
                cursor.execute("SELECT id FROM meetings WHERE wg_id = ? AND UPPER(meeting_number) = UPPER(?)",
                               (wg_id, meeting_number))
                res = cursor.fetchone()
                if res: row_id = res[0]

            if row_id:
                cursor.execute('''
                    UPDATE meetings
                    SET wg_id = ?, folder_name = ?, meeting_number = ?, name = ?, location = ?,
                        start_date = ?, end_date = ?, url_key = ?, docs_folder_url = ?,
                        first_tdoc = ?, first_tdoc_prefix = ?, first_tdoc_num = ?,
                        last_tdoc = ?, last_tdoc_prefix = ?, last_tdoc_num = ?,
                        sort_number = ?, is_ad_hoc = ?, is_electronic = ?, mtg_id = ?
                    WHERE id = ?
                ''', (
                    wg_id, folder_name, meeting_number, name, location,
                    start_date, end_date, db_url_key, docs_folder_url,
                    first_tdoc, first_tdoc_prefix, first_tdoc_num,
                    last_tdoc, last_tdoc_prefix, last_tdoc_num,
                    sort_num, is_ad_hoc, is_electronic, mtg_id,
                    row_id
                ))
            else:
                cursor.execute('''
                    INSERT INTO meetings (
                        wg_id, folder_name, meeting_number, name, location,
                        start_date, end_date, url_key, docs_folder_url,
                        first_tdoc, first_tdoc_prefix, first_tdoc_num,
                        last_tdoc, last_tdoc_prefix, last_tdoc_num,
                        sort_number, is_ad_hoc, is_electronic, mtg_id
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', (
                    wg_id, folder_name, meeting_number, name, location,
                    start_date, end_date, db_url_key, docs_folder_url,
                    first_tdoc, first_tdoc_prefix, first_tdoc_num,
                    last_tdoc, last_tdoc_prefix, last_tdoc_num,
                    sort_num, is_ad_hoc, is_electronic, mtg_id
                ))
            conn.commit()
        return True

    def search_meetings(self, wg_name=None, search_term=None, location=None, date_from=None, date_to=None,
                        adhoc_filter=None, type_filter=None):
        query = '''
            SELECT m.*, w.name as wg_name 
            FROM meetings m
            JOIN working_groups w ON m.wg_id = w.id
            WHERE 1=1
        '''
        params = []

        if wg_name is not None:
            if isinstance(wg_name, str):
                wg_name = [wg_name]
            if len(wg_name) == 0:
                return []
            valid_wgs = [wg for wg in wg_name if wg != "All WGs"]
            if valid_wgs:
                placeholders = ','.join('?' * len(valid_wgs))
                query += f" AND w.name IN ({placeholders})"
                params.extend(valid_wgs)

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

        query += " ORDER BY m.start_date DESC, w.name ASC"

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

    def is_active_sync_meeting(self, wg_name: str, start_date: str, end_date: str, is_electronic: int) -> bool:
        if is_electronic == 1 or not start_date or not end_date:
            return False

        today = datetime.date.today().strftime("%Y-%m-%d")

        if start_date <= today <= end_date:
            return True

        if today > end_date:
            query = '''
                SELECT m.id 
                FROM meetings m
                JOIN working_groups w ON m.wg_id = w.id
                WHERE w.name = ? 
                  AND m.start_date > ? 
                  AND m.start_date <= ?
                LIMIT 1
            '''
            with self._get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute(query, (wg_name, start_date, today))
                if not cursor.fetchone():
                    return True

        return False

    def find_meeting_by_tdoc(self, tdoc_str: str) -> dict:
        match = re.match(r'^([A-Za-z0-9]+)-?(\d+)', tdoc_str.strip(), re.IGNORECASE)
        if not match:
            return {}

        prefix = match.group(1).upper()
        num = int(match.group(2))

        query = '''
            SELECT m.*, w.name as wg_name 
            FROM meetings m
            JOIN working_groups w ON m.wg_id = w.id
            WHERE m.first_tdoc_num <= ? AND m.last_tdoc_num >= ?
              AND (UPPER(m.first_tdoc_prefix) = ? OR UPPER(m.last_tdoc_prefix) = ?)
        '''

        with self._get_connection() as conn:
            conn.row_factory = sqlite3.Row
            cursor = conn.cursor()
            cursor.execute(query, (num, num, prefix, prefix))
            row = cursor.fetchone()

            if row:
                return dict(row)
        return {}