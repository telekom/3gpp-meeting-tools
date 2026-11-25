"""Meeting search, deletion, TDoc-to-meeting resolution, and active/upcoming queries."""

from __future__ import annotations

import datetime
import re
import sqlite3
from typing import Optional


class MeetingsQueriesMixin:
    """Mixin providing read queries, deletion helpers, TDoc lookup, and date-cohort queries."""

    _NON_PHYSICAL_LOCATIONS = (
        "ONLINE",
        "ELECTRONIC",
        "ELECTRONIC MEETING",
        "ELECTRONIC MTG",
        "TELECONFERENCE MTG",
        "TBD",
    )

    def _build_meeting_filter_clause(
        self,
        wg_name=None,
        search_term=None,
        location=None,
        date_from=None,
        date_to=None,
        adhoc_filter=None,
        type_filter=None,
        date_mode: str = "all",
    ):
        """Build parameterized WHERE conditions for meeting search and count queries."""
        clauses = ["1=1"]
        params = []

        if wg_name is not None:
            if isinstance(wg_name, str):
                wg_name = [wg_name]
            if len(wg_name) == 0:
                return None, []
            valid_wgs = [wg for wg in wg_name if wg != "All WGs"]
            if valid_wgs:
                placeholders = ",".join("?" * len(valid_wgs))
                clauses.append(f"w.name IN ({placeholders})")
                params.extend(valid_wgs)

        if search_term:
            clauses.append("(m.meeting_number LIKE ? OR m.name LIKE ? OR m.location LIKE ? OR m.folder_name LIKE ?)")
            like_val = f"%{search_term}%"
            params.extend([like_val, like_val, like_val, like_val])
        if location:
            clauses.append("m.location LIKE ?")
            params.append(f"%{location}%")

        if date_mode == "unknown":
            clauses.append("COALESCE(TRIM(m.start_date), '') = ''")
        else:
            if date_mode == "dated":
                clauses.append("COALESCE(TRIM(m.start_date), '') != ''")
            if date_from:
                clauses.append(
                    "COALESCE(TRIM(m.start_date), '') != '' "
                    "AND COALESCE(NULLIF(TRIM(m.end_date), ''), m.start_date) >= ?"
                )
                params.append(date_from)
            if date_to:
                clauses.append("COALESCE(TRIM(m.start_date), '') != '' AND m.start_date <= ?")
                params.append(date_to)

        if adhoc_filter == "Ad-Hoc / BIS":
            clauses.append("m.is_ad_hoc = 1")
        elif adhoc_filter == "Regular":
            clauses.append("m.is_ad_hoc = 0")

        if type_filter == "Electronic":
            clauses.append("m.is_electronic = 1")
        elif type_filter == "In-Person":
            clauses.append("m.is_electronic = 0")

        return " AND ".join(clauses), params

    def search_meetings(
        self,
        wg_name=None,
        search_term=None,
        location=None,
        date_from=None,
        date_to=None,
        adhoc_filter=None,
        type_filter=None,
        date_mode: str = "all",
    ):
        where_sql, params = self._build_meeting_filter_clause(
            wg_name=wg_name,
            search_term=search_term,
            location=location,
            date_from=date_from,
            date_to=date_to,
            adhoc_filter=adhoc_filter,
            type_filter=type_filter,
            date_mode=date_mode,
        )
        if where_sql is None:
            return []

        order_sql = (
            "w.name ASC, m.sort_number DESC, m.folder_name ASC"
            if date_mode == "unknown"
            else "m.start_date DESC, w.name ASC, m.sort_number DESC"
        )
        query = f'''
            SELECT m.*, w.name as wg_name
            FROM meetings m
            JOIN working_groups w ON m.wg_id = w.id
            WHERE {where_sql}
            ORDER BY {order_sql}
        '''

        with self._get_connection() as conn:
            conn.row_factory = sqlite3.Row
            cursor = conn.cursor()
            cursor.execute(query, params)
            return [dict(row) for row in cursor.fetchall()]

    def count_unknown_date_meetings(
        self,
        wg_name=None,
        search_term=None,
        adhoc_filter=None,
        type_filter=None,
    ) -> int:
        """Count undated meetings matching the active WG/search/nature/format filters."""
        where_sql, params = self._build_meeting_filter_clause(
            wg_name=wg_name,
            search_term=search_term,
            adhoc_filter=adhoc_filter,
            type_filter=type_filter,
            date_mode="unknown",
        )
        if where_sql is None:
            return 0
        query = f'''
            SELECT COUNT(*)
            FROM meetings m
            JOIN working_groups w ON m.wg_id = w.id
            WHERE {where_sql}
        '''
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(query, params)
            row = cursor.fetchone()
            return int(row[0]) if row else 0

    def get_smart_docs_candidates(
        self,
        wg_names: Optional[list] = None,
        recent_days: int = 365,
        missing_lookback_days: int = 730,
        force_url_keys: Optional[set] = None,
    ) -> list:
        """Return FTP Docs/ scan tasks for recent, upcoming, newly added, or incomplete meetings."""
        today = datetime.date.today()
        recent_cutoff = (today - datetime.timedelta(days=max(0, int(recent_days)))).isoformat()
        missing_cutoff = (today - datetime.timedelta(days=max(0, int(missing_lookback_days)))).isoformat()

        clauses = ["COALESCE(TRIM(m.url_key), '') != ''"]
        params: list = []

        if wg_names:
            clean_wgs = [str(w).strip().upper() for w in wg_names if str(w).strip()]
            if not clean_wgs:
                return []
            placeholders = ",".join("?" * len(clean_wgs))
            clauses.append(f"UPPER(w.name) IN ({placeholders})")
            params.extend(clean_wgs)

        scope_or = [
            "m.start_date >= ?",
            "(m.start_date >= ? AND (COALESCE(m.first_tdoc_num, 0) <= 0 OR COALESCE(m.last_tdoc_num, 0) <= 0))",
        ]
        params.extend([recent_cutoff, missing_cutoff])

        norm_force = [str(k).strip().strip("/").lower() for k in (force_url_keys or []) if str(k).strip()]
        if norm_force:
            placeholders = ",".join("?" * len(norm_force))
            scope_or.append(f"LOWER(RTRIM(m.url_key, '/')) IN ({placeholders})")
            params.extend(norm_force)

        clauses.append("(" + " OR ".join(scope_or) + ")")

        query = f'''
            SELECT m.folder_name, m.meeting_number, m.url_key, w.name as wg_name
            FROM meetings m
            JOIN working_groups w ON m.wg_id = w.id
            WHERE {' AND '.join(clauses)}
            ORDER BY m.start_date DESC, m.sort_number DESC, m.id DESC
        '''
        with self._get_connection() as conn:
            conn.row_factory = sqlite3.Row
            cursor = conn.cursor()
            cursor.execute(query, params)
            tasks = []
            for row in cursor.fetchall():
                url_key = str(row["url_key"]).strip("/")
                absolute_url = f"https://www.3gpp.org/ftp/{url_key}"
                tasks.append(
                    {
                        "wg_name": row["wg_name"],
                        "folder_name": row["folder_name"] or row["meeting_number"],
                        "meeting_num": row["meeting_number"],
                        "url_key": url_key,
                        "absolute_url": absolute_url,
                        "docs_url": absolute_url.rstrip("/") + "/Docs/",
                    }
                )
            return tasks

    def get_working_groups(self):
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT name FROM working_groups ORDER BY name")
            return [row[0] for row in cursor.fetchall()]

    def delete_all_meetings(self):
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("DELETE FROM meetings")
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

    @staticmethod
    def _parse_tdoc_identity(tdoc_str: str):
        """
        Parse a TDoc into its prefix, numeric value, encoded year, and owning WG.

        3GPP TDoc numbers encode the two-digit meeting year at the start of the
        numeric portion, e.g. S2-2605740 -> 2026.
        """
        match = re.match(r"^([A-Za-z0-9]+)-?(\d+)", (tdoc_str or "").strip(), re.IGNORECASE)
        if not match:
            return None

        prefix = match.group(1).upper()
        numeric_str = match.group(2)
        if len(numeric_str) < 2:
            return None

        wg_map = {
            "SP": "SA",
            "S1": "SA1", "S2": "SA2", "S3": "SA3", "S4": "SA4", "S5": "SA5", "S6": "SA6",
            "RP": "RAN",
            "R1": "RAN1", "R2": "RAN2", "R3": "RAN3", "R4": "RAN4",
            "CP": "CT",
            "C1": "CT1", "C2": "CT2", "C3": "CT3", "C4": "CT4", "C5": "CT5", "C6": "CT6",
        }
        wg_name = wg_map.get(prefix)
        if not wg_name:
            return None

        return {
            "prefix": prefix,
            "number": int(numeric_str),
            "year": 2000 + int(numeric_str[:2]),
            "wg_name": wg_name,
        }

    def find_meeting_by_tdoc(self, tdoc_str: str) -> dict:
        """
        Find the meeting owning a TDoc using a strict WG + year + range match.

        Deliberately does not fall back to a different year: opening a document
        from the wrong meeting is worse than reporting incomplete local metadata.
        """
        identity = self._parse_tdoc_identity(tdoc_str)
        if not identity:
            return {}

        query = """
            SELECT m.*, w.name as wg_name
            FROM meetings m
            JOIN working_groups w ON m.wg_id = w.id
            WHERE UPPER(w.name) = ?
              AND SUBSTR(m.start_date, 1, 4) = ?
              AND m.first_tdoc_num > 0
              AND m.last_tdoc_num > 0
              AND m.first_tdoc_num <= ?
              AND m.last_tdoc_num >= ?
              AND (
                    UPPER(m.first_tdoc_prefix) = ?
                 OR UPPER(m.last_tdoc_prefix) = ?
              )
            ORDER BY m.start_date DESC, m.id DESC
            LIMIT 1
        """

        params = (
            identity["wg_name"].upper(),
            str(identity["year"]),
            identity["number"],
            identity["number"],
            identity["prefix"],
            identity["prefix"],
        )

        with self._get_connection() as conn:
            conn.row_factory = sqlite3.Row
            cursor = conn.cursor()
            cursor.execute(query, params)
            row = cursor.fetchone()
            return dict(row) if row else {}

    def find_incomplete_tdoc_meetings(self, tdoc_str: str) -> list:
        """
        Return refreshable same-WG/same-year candidate meetings for an unresolved TDoc.

        Historically this method returned only rows whose first/last TDoc metadata
        was empty. That misses the equally important case where a meeting has a
        populated but stale range. Quick TDoc Jump calls this method only after the
        strict range lookup failed, so at that point every same-WG/same-year meeting
        with an FTP path is a legitimate Phase-2 repair candidate.

        Candidates with incomplete ranges are ordered first, followed by populated
        ranges ordered chronologically. The existing method name is retained to keep
        the UI/controller API backward compatible.
        """
        identity = self._parse_tdoc_identity(tdoc_str)
        if not identity:
            return []

        query = """
            SELECT m.*, w.name as wg_name,
                   CASE WHEN (
                        COALESCE(m.first_tdoc_num, 0) <= 0
                     OR COALESCE(m.last_tdoc_num, 0) <= 0
                     OR COALESCE(TRIM(m.first_tdoc_prefix), '') = ''
                     OR COALESCE(TRIM(m.last_tdoc_prefix), '') = ''
                   ) THEN 0 ELSE 1 END AS range_complete
            FROM meetings m
            JOIN working_groups w ON m.wg_id = w.id
            WHERE UPPER(w.name) = ?
              AND SUBSTR(m.start_date, 1, 4) = ?
              AND COALESCE(TRIM(m.url_key), '') != ''
            ORDER BY range_complete ASC, m.start_date ASC, m.sort_number ASC, m.id ASC
        """

        with self._get_connection() as conn:
            conn.row_factory = sqlite3.Row
            cursor = conn.cursor()
            cursor.execute(query, (identity["wg_name"].upper(), str(identity["year"])))
            rows = [dict(row) for row in cursor.fetchall()]

        # range_complete is an internal ranking helper, not part of the meeting model.
        for row in rows:
            row.pop("range_complete", None)
        return rows

    def get_active_meetings_for_date(
        self,
        ref_date_str: str,
        buffer_days: int = 2,
        wg_filter: Optional[list] = None,
    ) -> list:
        """
        Returns all in-person (non-electronic) meetings with a physical location
        whose date range overlaps `[ref_date - buffer_days, ref_date + buffer_days]`.
        """
        try:
            ref_d = datetime.date.fromisoformat(ref_date_str[:10])
        except Exception:
            ref_d = datetime.date.today()

        buf = max(0, int(buffer_days))
        window_start = (ref_d - datetime.timedelta(days=buf)).isoformat()
        window_end = (ref_d + datetime.timedelta(days=buf)).isoformat()

        excl_placeholders = ",".join("?" * len(self._NON_PHYSICAL_LOCATIONS))
        query = f"""
            SELECT m.*, w.name as wg_name
            FROM meetings m
            JOIN working_groups w ON m.wg_id = w.id
            WHERE COALESCE(m.is_electronic, 0) = 0
              AND COALESCE(TRIM(m.location), '') != ''
              AND UPPER(TRIM(m.location)) NOT IN ({excl_placeholders})
              AND COALESCE(TRIM(m.start_date), '') != ''
              AND COALESCE(TRIM(m.end_date), '') != ''
              AND m.start_date <= ?
              AND m.end_date >= ?
        """
        params: list = list(self._NON_PHYSICAL_LOCATIONS) + [window_end, window_start]

        if wg_filter:
            clean_wgs = [str(w).strip().upper() for w in wg_filter if str(w).strip()]
            if clean_wgs:
                wg_placeholders = ",".join("?" * len(clean_wgs))
                query += f" AND UPPER(w.name) IN ({wg_placeholders})"
                params.extend(clean_wgs)

        query += " ORDER BY m.start_date ASC, w.name ASC, m.id ASC"

        with self._get_connection() as conn:
            conn.row_factory = sqlite3.Row
            cursor = conn.cursor()
            cursor.execute(query, params)
            return [dict(row) for row in cursor.fetchall()]

    def get_next_upcoming_meetings(
        self,
        ref_date_str: str,
        wg_filter: Optional[list] = None,
    ) -> list:
        """
        Returns the next upcoming cohort of in-person meetings starting after `ref_date_str`
        (all meetings starting within 7 days of the earliest upcoming meeting date).
        """
        try:
            ref_d = datetime.date.fromisoformat(ref_date_str[:10]).isoformat()
        except Exception:
            ref_d = datetime.date.today().isoformat()

        excl_placeholders = ",".join("?" * len(self._NON_PHYSICAL_LOCATIONS))
        base_where = f"""
            COALESCE(m.is_electronic, 0) = 0
            AND COALESCE(TRIM(m.location), '') != ''
            AND UPPER(TRIM(m.location)) NOT IN ({excl_placeholders})
            AND COALESCE(TRIM(m.start_date), '') != ''
            AND m.start_date > ?
        """
        base_params: list = list(self._NON_PHYSICAL_LOCATIONS) + [ref_d]

        wg_clause = ""
        wg_params: list = []
        if wg_filter:
            clean_wgs = [str(w).strip().upper() for w in wg_filter if str(w).strip()]
            if clean_wgs:
                wg_placeholders = ",".join("?" * len(clean_wgs))
                wg_clause = f" AND UPPER(w.name) IN ({wg_placeholders})"
                wg_params.extend(clean_wgs)

        with self._get_connection() as conn:
            conn.row_factory = sqlite3.Row
            cursor = conn.cursor()
            cursor.execute(
                f"""
                SELECT MIN(m.start_date) as next_start
                FROM meetings m
                JOIN working_groups w ON m.wg_id = w.id
                WHERE {base_where} {wg_clause}
                """,
                base_params + wg_params,
            )
            row = cursor.fetchone()
            if not row or not row["next_start"]:
                return []
            earliest_str = str(row["next_start"])[:10]
            try:
                cohort_end = (datetime.date.fromisoformat(earliest_str) + datetime.timedelta(days=7)).isoformat()
            except Exception:
                cohort_end = earliest_str

            cursor.execute(
                f"""
                SELECT m.*, w.name as wg_name
                FROM meetings m
                JOIN working_groups w ON m.wg_id = w.id
                WHERE {base_where} {wg_clause}
                  AND m.start_date <= ?
                ORDER BY m.start_date ASC, w.name ASC, m.id ASC
                """,
                base_params + wg_params + [cohort_end],
            )
            return [dict(r) for r in cursor.fetchall()]
