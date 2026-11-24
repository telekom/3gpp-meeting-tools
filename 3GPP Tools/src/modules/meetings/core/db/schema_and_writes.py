"""SQLite connection management, schema migrations, and meeting write/upsert operations."""

from __future__ import annotations

import contextlib
import re
import sqlite3


class MeetingsSchemaAndWritesMixin:
    """Mixin providing connection setup, schema migrations, and bulk/single writes."""

    @contextlib.contextmanager
    def _get_connection(self):
        # 30-second timeout allows background writers to finish without crashing readers
        conn = sqlite3.connect(self.db_path, timeout=30.0)
        try:
            # Enable WAL mode for concurrent multi-thread read/write support
            conn.execute("PRAGMA journal_mode=WAL;")
            conn.execute("PRAGMA busy_timeout=30000;")
            with conn:
                yield conn
        finally:
            conn.close()

    def _create_tables(self):
        with self._get_connection() as conn:
            cursor = conn.cursor()

            # Set WAL mode once during initial schema setup
            cursor.execute("PRAGMA journal_mode=WAL;")

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
            for alter_sql in (
                "ALTER TABLE meetings ADD COLUMN sort_number INTEGER DEFAULT 0",
                "ALTER TABLE meetings ADD COLUMN is_ad_hoc INTEGER DEFAULT 0",
                "ALTER TABLE meetings ADD COLUMN is_electronic INTEGER DEFAULT 0",
                "ALTER TABLE meetings ADD COLUMN first_tdoc_prefix TEXT",
                "ALTER TABLE meetings ADD COLUMN first_tdoc_num INTEGER DEFAULT 0",
                "ALTER TABLE meetings ADD COLUMN last_tdoc_prefix TEXT",
                "ALTER TABLE meetings ADD COLUMN last_tdoc_num INTEGER DEFAULT 0",
                "ALTER TABLE meetings ADD COLUMN mtg_id TEXT",
            ):
                try:
                    cursor.execute(alter_sql)
                except sqlite3.OperationalError:
                    pass

            conn.commit()

    def _extract_sort_num(self, m_str: str) -> int:
        match = re.search(r"\d+", m_str or "")
        return int(match.group()) if match else 0

    def _get_meeting_flags(self, m_str: str):
        num_upper = (m_str or "").upper()
        is_ad_hoc = 1 if ("A" in num_upper or "BIS" in num_upper) else 0
        is_electronic = 1 if ("E" in num_upper) else 0
        return is_ad_hoc, is_electronic

    def get_or_create_wg(self, wg_name: str) -> int:
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("INSERT OR IGNORE INTO working_groups (name) VALUES (?)", (wg_name,))
            cursor.execute("SELECT id FROM working_groups WHERE name = ?", (wg_name,))
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
        if not meetings_data:
            return
        wg_map = {}
        for task in meetings_data:
            wg = task["wg_name"]
            if wg not in wg_map:
                wg_map[wg] = self.get_or_create_wg(wg)

        insert_data = []
        for task in meetings_data:
            m_num = task["meeting_num"]
            is_ah, is_e = self._get_meeting_flags(m_num)
            final_ah = 1 if (is_ah or task.get("is_ad_hoc")) else 0

            insert_data.append((
                wg_map[task["wg_name"]], task["folder_name"], m_num,
                self._extract_sort_num(m_num), final_ah, is_e,
                task["url_key"], task.get("docs_url", ""),
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
        if not docs_data:
            return
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
        if not metadata_data:
            return
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
                    re.search(r"AH|Ad\s*Hoc", final_m_num, re.IGNORECASE)
                    or re.search(r"Ad\s*Hoc", m_name, re.IGNORECASE)
                ):
                    is_ah = 1

                row_id = None

                # 1. Attempt strict match via FTP URL
                if db_url_key:
                    cursor.execute(
                        "SELECT id FROM meetings WHERE LOWER(RTRIM(url_key, '/')) = LOWER(RTRIM(?, '/'))",
                        (db_url_key,),
                    )
                    res = cursor.fetchone()
                    if res:
                        row_id = res[0]

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
                    if res:
                        row_id = res[0]

                # 3. Safe UPSERT
                if row_id:
                    try:
                        cursor.execute(
                            '''
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
                            (
                                mtg_id, mtg_id, m_name, m_name, town, town, start_d, start_d, end_d, end_d,
                                new_m_num, new_m_num, sort_n, sort_n, is_ah, is_e, db_url_key, db_url_key,
                                row_id,
                            ),
                        )
                    except sqlite3.IntegrityError:
                        cursor.execute(
                            '''
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
                            (
                                mtg_id, mtg_id, m_name, m_name, town, town, start_d, start_d, end_d, end_d,
                                new_m_num, new_m_num, sort_n, sort_n, is_ah, is_e, row_id,
                            ),
                        )

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
        url_key = url_key.strip("/")
        db_url_key = url_key if url_key else None

        docs_folder_url = data.get("docs_folder_url", "").strip()
        if not docs_folder_url and db_url_key:
            docs_folder_url = f"https://www.3gpp.org/ftp/{db_url_key}/Docs/"

        first_tdoc = data.get("first_tdoc", "").strip()
        first_tdoc_prefix = data.get("first_tdoc_prefix", "").strip()
        first_tdoc_num = int(data.get("first_tdoc_num") or 0)
        if first_tdoc and not first_tdoc_prefix:
            m = re.match(r"^([A-Za-z0-9]+)-?(\d+)", first_tdoc)
            if m:
                first_tdoc_prefix = m.group(1).upper()
                first_tdoc_num = int(m.group(2))

        last_tdoc = data.get("last_tdoc", "").strip()
        last_tdoc_prefix = data.get("last_tdoc_prefix", "").strip()
        last_tdoc_num = int(data.get("last_tdoc_num") or 0)
        if last_tdoc and not last_tdoc_prefix:
            m = re.match(r"^([A-Za-z0-9]+)-?(\d+)", last_tdoc)
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
                cursor.execute(
                    "SELECT id FROM meetings WHERE LOWER(RTRIM(url_key, '/')) = LOWER(RTRIM(?, '/'))",
                    (db_url_key,),
                )
                res = cursor.fetchone()
                if res:
                    row_id = res[0]

            # 2. Match by mtg_id
            if not row_id and mtg_id:
                cursor.execute("SELECT id FROM meetings WHERE mtg_id = ?", (mtg_id,))
                res = cursor.fetchone()
                if res:
                    row_id = res[0]

            # 3. Match by WG + Meeting number
            if not row_id and meeting_number:
                cursor.execute(
                    "SELECT id FROM meetings WHERE wg_id = ? AND UPPER(meeting_number) = UPPER(?)",
                    (wg_id, meeting_number),
                )
                res = cursor.fetchone()
                if res:
                    row_id = res[0]

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
                    row_id,
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
                    sort_num, is_ad_hoc, is_electronic, mtg_id,
                ))
            conn.commit()
        return True
