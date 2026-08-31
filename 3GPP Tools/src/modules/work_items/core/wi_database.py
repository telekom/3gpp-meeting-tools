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

    def _build_filter_clause(self, search_term: str = None, releases: list = None,
                             wg_names: list = None, status: str = "all"):
        """Shared query builder for search and count operations."""
        query_joins = ""
        where_clauses = ["1=1"]
        params = []

        filter_by_wg = wg_names and 'ALL' not in wg_names and len(wg_names) > 0
        if filter_by_wg:
            query_joins += """
                JOIN wi_group_map filter_m ON wi.code = filter_m.wi_code
                JOIN working_groups filter_w ON filter_m.group_id = filter_w.id
            """
            placeholders = ','.join(['?'] * len(wg_names))
            where_clauses.append(f"filter_w.name IN ({placeholders})")
            params.extend(wg_names)

        if releases and 'ALL' not in releases and len(releases) > 0:
            placeholders = ','.join(['?'] * len(releases))
            where_clauses.append(f"wi.release IN ({placeholders})")
            params.extend(releases)

        if search_term:
            where_clauses.append("(wi.acronym LIKE ? OR wi.name LIKE ? OR wi.code LIKE ?)")
            term = f"%{search_term}%"
            params.extend([term, term, term])

        if status == "finished":
            where_clauses.append("(wi.end_date IS NOT NULL AND TRIM(wi.end_date) != '' AND date(wi.end_date) < date('now'))")
        elif status == "active":
            where_clauses.append("(wi.end_date IS NULL OR TRIM(wi.end_date) == '' OR date(wi.end_date) >= date('now'))")

        where_sql = " WHERE " + " AND ".join(where_clauses)
        return query_joins, where_sql, params

    def count_work_items(self, search_term: str = None, releases: list = None,
                         wg_names: list = None, status: str = "all") -> int:
        """Returns the total number of matching work items without retrieving row payloads."""
        joins, where_sql, params = self._build_filter_clause(search_term, releases, wg_names, status)
        query = f"SELECT COUNT(DISTINCT wi.code) FROM work_items wi {joins} {where_sql}"
        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute(query, params)
                row = cursor.fetchone()
                return row[0] if row else 0
        except Exception as e:
            import logging
            logging.error(f"Failed to count Work Items: {e}")
            return 0

    def search_work_items(self, search_term: str = None, releases: list = None,
                          wg_names: list = None, status: str = "all",
                          limit: int = None, offset: int = 0) -> list:
        """Searches Work Items with optional pagination (LIMIT and OFFSET)."""
        joins, where_sql, params = self._build_filter_clause(search_term, releases, wg_names, status)

        query = f"""
            SELECT wi.code, wi.acronym, wi.name, wi.latest_wid, wi.release, wi.start_date, wi.end_date,
                   rem.remarks,
                   grp.wg_names,
                   COALESCE(sp.spec_count, 0) AS spec_count
            FROM work_items wi
            {joins}
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
            {where_sql}
            GROUP BY wi.code
            ORDER BY CAST(wi.code AS INTEGER) DESC
        """

        if limit is not None:
            query += " LIMIT ? OFFSET ?"
            params.extend([limit, offset])

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

    def get_work_item_details(self, wi_code: str) -> dict:
        """Retrieves comprehensive details for a single Work Item."""
        query = """
            SELECT wi.code, wi.acronym, wi.name, wi.latest_wid, wi.release, wi.start_date, wi.end_date,
                   grp.wg_names
            FROM work_items wi
            LEFT JOIN (
                SELECT m.wi_code, GROUP_CONCAT(w.name, ', ') AS wg_names
                FROM wi_group_map m
                JOIN working_groups w ON m.group_id = w.id
                GROUP BY m.wi_code
            ) grp ON wi.code = grp.wi_code
            WHERE wi.code = ?
        """
        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute(query, (wi_code,))
                row = cursor.fetchone()
                if not row:
                    return {}

                cols = [col[0] for col in cursor.description]
                details = dict(zip(cols, row))

                # Fetch remarks
                cursor.execute(
                    "SELECT creation_date, remark FROM wi_remarks WHERE wi_code = ? ORDER BY id DESC",
                    (wi_code,)
                )
                details['remarks'] = [
                    {'date': r[0], 'text': r[1]} for r in cursor.fetchall()
                ]

                # Fetch linked specs
                details['linked_specs'] = self.get_linked_specs_for_wi(wi_code)
                return details
        except Exception as e:
            import logging
            logging.error(f"Failed to fetch details for WI {wi_code}: {e}")
            return {}

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

    def get_filter_options(self) -> dict:
        """Fetches unique Release versions and mapped Working Groups for the UI dropdowns."""
        options = {'releases': [], 'groups': []}
        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()

                cursor.execute("SELECT DISTINCT release FROM work_items WHERE release IS NOT NULL AND release != ''")
                raw_releases = [str(r[0]).strip() for r in cursor.fetchall()]

                def release_sort_key(rel):
                    if rel.upper() == 'R99':
                        return -1
                    match = re.search(r'\d+', rel)
                    return int(match.group()) if match else 0

                raw_releases.sort(key=release_sort_key, reverse=True)
                options['releases'] = raw_releases

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
        """Deletes multiple Work Items and their associated mappings in a single transaction."""
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

    def upsert_work_items(self, wg_name: str, items: list):
        """Bulk inserts or updates Work Items and maps them to their Working Group."""
        if not items:
            return

        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute('INSERT OR IGNORE INTO working_groups (name) VALUES (?)', (wg_name,))
            cursor.execute('SELECT id FROM working_groups WHERE name = ?', (wg_name,))
            wg_row = cursor.fetchone()
            if not wg_row:
                return
            wg_id = wg_row[0]

            wi_data = []
            map_data = []
            for item in items:
                wi_data.append((item['code'], item['acronym'], item['name'], '', item['release'], '', ''))
                map_data.append((item['code'], wg_id))

            cursor.executemany('''
                INSERT INTO work_items (code, acronym, name, latest_wid, release, start_date, end_date)
                VALUES (?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(code) DO UPDATE SET
                    acronym = CASE WHEN (excluded.acronym IS NOT NULL AND excluded.acronym != '') THEN excluded.acronym ELSE work_items.acronym END,
                    name = CASE WHEN (excluded.name IS NOT NULL AND excluded.name != '') THEN excluded.name ELSE work_items.name END,
                    release = CASE WHEN (excluded.release IS NOT NULL AND excluded.release != '') THEN excluded.release ELSE work_items.release END
            ''', wi_data)

            cursor.executemany('''
                INSERT OR IGNORE INTO wi_group_map (wi_code, group_id)
                VALUES (?, ?)
            ''', map_data)
            conn.commit()

    def update_work_items_metadata(self, metadata_list: list):
        """Batch updates multiple Work Items with scraped metadata and remarks."""
        if not metadata_list:
            return

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
                remark_tuples.append((wi_code, remark['date'], remark['text']))

        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()
                cursor.executemany('''
                    UPDATE work_items 
                    SET start_date = CASE WHEN ? != '' THEN ? ELSE start_date END,
                        end_date = CASE WHEN ? != '' THEN ? ELSE end_date END,
                        latest_wid = CASE WHEN ? != '' THEN ? ELSE latest_wid END
                    WHERE code = ?
                ''', update_tuples)

                cursor.executemany('DELETE FROM wi_remarks WHERE wi_code = ?', wi_codes_to_clear)

                if remark_tuples:
                    cursor.executemany('''
                        INSERT INTO wi_remarks (wi_code, creation_date, remark)
                        VALUES (?, ?, ?)
                    ''', remark_tuples)
                conn.commit()
        except Exception as e:
            import logging
            logging.error(f"Failed to batch update Work Items metadata: {e}", exc_info=True)