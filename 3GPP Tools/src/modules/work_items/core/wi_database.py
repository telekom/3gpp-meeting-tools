# --- File: src/modules/work_items/core/wi_database.py ---
import re
import sqlite3
from pathlib import Path


class WorkItemsDatabase:
    """
    Handles all database operations for 3GPP Work Items.
    Connects to the shared 3gpp_data.db file to maintain a single source of truth.
    """

    def __init__(self, db_path: Path):
        self.db_path = db_path
        self._init_db()

    def _get_connection(self):
        return sqlite3.connect(self.db_path, check_same_thread=False)

    def _init_db(self):
        with self._get_connection() as conn:
            cursor = conn.cursor()

            # Ensure Write-Ahead Logging is enabled for concurrent access across modules
            cursor.execute('PRAGMA journal_mode=WAL;')

            cursor.execute('''
                CREATE TABLE IF NOT EXISTS work_items (
                    code TEXT PRIMARY KEY,
                    acronym TEXT,
                    name TEXT,
                    latest_wid TEXT,
                    release TEXT,
                    start_date TEXT,
                    end_date TEXT
                )
            ''')

            cursor.execute('''
                CREATE TABLE IF NOT EXISTS wi_group_map (
                    wi_code TEXT,
                    group_id INTEGER,
                    UNIQUE(wi_code, group_id),
                    FOREIGN KEY(wi_code) REFERENCES work_items(code),
                    FOREIGN KEY(group_id) REFERENCES working_groups(id)
                )
            ''')

            cursor.execute('''
                CREATE TABLE IF NOT EXISTS wi_remarks (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    wi_code TEXT,
                    creation_date TEXT,
                    remark TEXT,
                    FOREIGN KEY(wi_code) REFERENCES work_items(code)
                )
            ''')

    def get_all_work_items(self) -> list:
        """Fetches all work items to populate the UI table, ordered by code descending numerically."""
        query = """
            SELECT wi.code, wi.acronym, wi.name, wi.latest_wid, wi.release, wi.start_date, wi.end_date,
                   rem.remarks,
                   grp.wg_names,
                   COALESCE(sp.spec_count, 0) AS spec_count
            FROM work_items wi
            LEFT JOIN (
                SELECT wi_code, GROUP_CONCAT(creation_date || ':::' || remark, '|||') AS remarks
                FROM wi_remarks GROUP BY wi_code
            ) rem ON wi.code = rem.wi_code
            LEFT JOIN (
                SELECT m.wi_code, GROUP_CONCAT(w.name, ', ') AS wg_names
                FROM wi_group_map m
                JOIN working_groups w ON m.group_id = w.id
                GROUP BY m.wi_code
            ) grp ON wi.code = grp.wi_code
            LEFT JOIN (
                SELECT wi_code, COUNT(spec_id) AS spec_count
                FROM spec_wi_map
                GROUP BY wi_code
            ) sp ON wi.code = sp.wi_code
            ORDER BY CAST(wi.code AS INTEGER) DESC
        """
        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute(query)
                columns = [col[0] for col in cursor.description]
                return [dict(zip(columns, row)) for row in cursor.fetchall()]
        except Exception as e:
            import logging
            logging.error(f"Failed to fetch Work Items: {e}")
            return []

    def get_linked_specs_for_wi(self, wi_code: str) -> list:
        """Retrieves all specifications impacted by a given Work Item code."""
        query = """
            SELECT sp.number, sp.title, sp.type, s.name AS series_name, m.is_primary
            FROM spec_wi_map m
            JOIN specifications sp ON m.spec_id = sp.id
            JOIN series s ON sp.series_id = s.id
            WHERE m.wi_code = ?
            ORDER BY m.is_primary DESC, CAST(s.name AS INTEGER) ASC, sp.number ASC
        """
        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute(query, (wi_code,))
                cols = [col[0] for col in cursor.description]
                return [dict(zip(cols, row)) for row in cursor.fetchall()]
        except Exception as e:
            import logging
            logging.error(f"Failed to fetch linked specs for WI {wi_code}: {e}")
            return []

    def upsert_work_items(self, wg_name: str, items: list):
        """
        Bulk inserts or updates Work Items and maps them to their Working Group.
        """
        if not items:
            return

        with self._get_connection() as conn:
            cursor = conn.cursor()

            # 1. Ensure the Working Group exists in the shared table and grab its ID
            cursor.execute('INSERT OR IGNORE INTO working_groups (name) VALUES (?)', (wg_name,))
            cursor.execute('SELECT id FROM working_groups WHERE name = ?', (wg_name,))
            wg_row = cursor.fetchone()
            if not wg_row:
                return
            wg_id = wg_row[0]

            wi_data = []
            map_data = []
            for item in items:
                wi_data.append((
                    item['code'], item['acronym'], item['name'], '', item['release'], '', ''
                ))
                map_data.append((item['code'], wg_id))

            # 2. Bulk UPSERT the Work Items
            cursor.executemany('''
                INSERT INTO work_items (code, acronym, name, latest_wid, release, start_date, end_date)
                VALUES (?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(code) DO UPDATE SET
                    acronym = CASE WHEN (excluded.acronym IS NOT NULL AND excluded.acronym != '') THEN excluded.acronym ELSE work_items.acronym END,
                    name = CASE WHEN (excluded.name IS NOT NULL AND excluded.name != '') THEN excluded.name ELSE work_items.name END,
                    release = CASE WHEN (excluded.release IS NOT NULL AND excluded.release != '') THEN excluded.release ELSE work_items.release END
            ''', wi_data)

            # 3. Bulk UPSERT the mapping
            cursor.executemany('''
                INSERT OR IGNORE INTO wi_group_map (wi_code, group_id)
                VALUES (?, ?)
            ''', map_data)

            conn.commit()

    def get_filter_options(self) -> dict:
        """Fetches unique Release versions and mapped Working Groups for the UI dropdowns."""
        options = {'releases': [], 'groups': []}
        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()

                # Fetch unique releases
                cursor.execute("SELECT DISTINCT release FROM work_items WHERE release IS NOT NULL AND release != ''")
                raw_releases = [str(r[0]).strip() for r in cursor.fetchall()]

                def release_sort_key(rel):
                    if rel.upper() == 'R99':
                        return -1
                    match = re.search(r'\d+', rel)
                    if match:
                        return int(match.group())
                    return 0

                raw_releases.sort(key=release_sort_key, reverse=True)
                options['releases'] = raw_releases

                # Fetch only WGs that are actually mapped to work items
                cursor.execute("""
                    SELECT DISTINCT w.name 
                    FROM working_groups w
                    JOIN wi_group_map m ON w.id = m.group_id
                    ORDER BY w.name
                """)
                options['groups'] = [str(r[0]).strip() for r in cursor.fetchall()]

        except Exception as e:
            import logging
            logging.error(f"Error fetching WI filter options: {e}")
        return options

    def search_work_items(self, search_term: str = None, releases: list = None,
                          wg_names: list = None, status: str = "all") -> list:
        """Searches Work Items by text, multiple releases, multiple working groups, and completion status."""
        query = """
            SELECT wi.code, wi.acronym, wi.name, wi.latest_wid, wi.release, wi.start_date, wi.end_date,
                   rem.remarks,
                   grp.wg_names,
                   COALESCE(sp.spec_count, 0) AS spec_count
            FROM work_items wi
            LEFT JOIN (
                SELECT wi_code, GROUP_CONCAT(creation_date || ':::' || remark, '|||') AS remarks
                FROM wi_remarks GROUP BY wi_code
            ) rem ON wi.code = rem.wi_code
            LEFT JOIN (
                SELECT m.wi_code, GROUP_CONCAT(w.name, ', ') AS wg_names
                FROM wi_group_map m
                JOIN working_groups w ON m.group_id = w.id
                GROUP BY m.wi_code
            ) grp ON wi.code = grp.wi_code
            LEFT JOIN (
                SELECT wi_code, COUNT(spec_id) AS spec_count
                FROM spec_wi_map
                GROUP BY wi_code
            ) sp ON wi.code = sp.wi_code
        """
        params = []

        filter_by_wg = wg_names and 'ALL' not in wg_names and len(wg_names) > 0

        if filter_by_wg:
            query += """
                JOIN wi_group_map filter_m ON wi.code = filter_m.wi_code
                JOIN working_groups filter_w ON filter_m.group_id = filter_w.id
            """

        query += " WHERE 1=1"

        if filter_by_wg:
            placeholders = ','.join(['?'] * len(wg_names))
            query += f" AND filter_w.name IN ({placeholders})"
            params.extend(wg_names)

        if releases and 'ALL' not in releases and len(releases) > 0:
            placeholders = ','.join(['?'] * len(releases))
            query += f" AND wi.release IN ({placeholders})"
            params.extend(releases)

        if search_term:
            query += " AND (wi.acronym LIKE ? OR wi.name LIKE ? OR wi.code LIKE ?)"
            term = f"%{search_term}%"
            params.extend([term, term, term])

        # Filter by completion status based on end_date
        if status == "finished":
            query += " AND (wi.end_date IS NOT NULL AND TRIM(wi.end_date) != '' AND date(wi.end_date) < date('now'))"
        elif status == "active":
            query += " AND (wi.end_date IS NULL OR TRIM(wi.end_date) == '' OR date(wi.end_date) >= date('now'))"

        query += " GROUP BY wi.code ORDER BY CAST(wi.code AS INTEGER) DESC"

        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute(query, params)
                columns = [col[0] for col in cursor.description]
                return [dict(zip(columns, row)) for row in cursor.fetchall()]
        except Exception as e:
            import logging
            logging.error(f"Failed to search Work Items: {e}")
            return []

    def delete_work_item(self, code: str):
        """Deletes a Work Item and its associated group mappings and remarks from the database."""
        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute("DELETE FROM wi_group_map WHERE wi_code = ?", (code,))
                cursor.execute("DELETE FROM wi_remarks WHERE wi_code = ?", (code,))
                cursor.execute("DELETE FROM spec_wi_map WHERE wi_code = ?", (code,))
                cursor.execute("DELETE FROM work_items WHERE code = ?", (code,))
                conn.commit()
        except Exception as e:
            import logging
            logging.error(f"Failed to delete Work Item {code}: {e}")

    def delete_work_items(self, code_list: list):
        """Deletes multiple Work Items and their associated group mappings and remarks in a single transaction."""
        if not code_list:
            return
        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()
                placeholders = ','.join(['?'] * len(code_list))

                cursor.execute(f"DELETE FROM wi_group_map WHERE wi_code IN ({placeholders})", code_list)
                cursor.execute(f"DELETE FROM wi_remarks WHERE wi_code IN ({placeholders})", code_list)
                cursor.execute(f"DELETE FROM spec_wi_map WHERE wi_code IN ({placeholders})", code_list)
                cursor.execute(f"DELETE FROM work_items WHERE code IN ({placeholders})", code_list)

                conn.commit()
        except Exception as e:
            import logging
            logging.error(f"Failed to batch delete Work Items: {e}")

    def update_work_items_metadata(self, metadata_list: list):
        """
        Batch updates multiple Work Items with scraped metadata using a single transaction,
        including clearing and re-inserting their associated remarks.
        """
        import logging

        if not metadata_list:
            logging.warning("update_work_items_metadata was called with an empty list.")
            return

        logging.info(f"Preparing database batch update for {len(metadata_list)} Work Items...")

        update_tuples = []
        remark_tuples = []
        wi_codes_to_clear = []

        for meta in metadata_list:
            wi_code = meta.get('code')
            wi_codes_to_clear.append((wi_code,))

            update_tuples.append((
                meta.get('start_date', ''), meta.get('start_date', ''),
                meta.get('end_date', ''), meta.get('end_date', ''),
                meta.get('latest_wid', ''), meta.get('latest_wid', ''),
                wi_code
            ))

            for remark in meta.get('remarks', []):
                remark_tuples.append((
                    wi_code,
                    remark['date'],
                    remark['text']
                ))

        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()

                # 1. Update the main Work Item metadata
                cursor.executemany('''
                    UPDATE work_items 
                    SET start_date = CASE WHEN ? != '' THEN ? ELSE start_date END,
                        end_date = CASE WHEN ? != '' THEN ? ELSE end_date END,
                        latest_wid = CASE WHEN ? != '' THEN ? ELSE latest_wid END
                    WHERE code = ?
                ''', update_tuples)

                logging.info(f"Database UPDATE for work_items affected {cursor.rowcount} row(s).")

                # 2. Delete existing remarks for these specific WIs
                cursor.executemany('''
                    DELETE FROM wi_remarks WHERE wi_code = ?
                ''', wi_codes_to_clear)

                logging.info(f"Database DELETE for old wi_remarks affected {cursor.rowcount} row(s).")

                # 3. Insert the newly scraped remarks
                if remark_tuples:
                    cursor.executemany('''
                        INSERT INTO wi_remarks (wi_code, creation_date, remark)
                        VALUES (?, ?, ?)
                    ''', remark_tuples)
                    logging.info(f"Database INSERT for new wi_remarks added {cursor.rowcount} row(s).")

        except Exception as e:
            logging.error(f"Failed to batch update Work Items metadata: {e}", exc_info=True)