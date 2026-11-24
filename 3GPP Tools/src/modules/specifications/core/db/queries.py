"""Read and filter queries for specifications, archive files, titles, and WIs."""

from __future__ import annotations

import sqlite3
from typing import Dict, List


class SpecsQueryMixin:
    """Query methods for searching specifications, file versions, and related WIs."""

    def get_spec_wis(self, spec_number: str) -> List[Dict]:
        """Fetches all Work Items mapped to a given specification."""
        query = """
            SELECT m.wi_code, m.is_primary, 
                   COALESCE(w.acronym, '') AS acronym, 
                   COALESCE(w.name, '') AS name, 
                   COALESCE(w.release, '') AS release
            FROM spec_wi_map m
            JOIN specifications s ON m.spec_id = s.id
            LEFT JOIN work_items w ON m.wi_code = w.code
            WHERE s.number = ?
            ORDER BY m.is_primary DESC, m.wi_code ASC
        """
        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute(query, (spec_number,))
                cols = [c[0] for c in cursor.description]
                return [dict(zip(cols, row)) for row in cursor.fetchall()]
        except sqlite3.OperationalError:
            return []

    def search_files(
        self,
        spec_number: str = None,
        release_version: str = None,
        series: str = None,
        tech: str = None,
        group: str = None,
        spec_type: str = None,
        tag: str = None,
        tsg: str = None,
    ) -> list:
        join_tech = bool(tech and tech != "Any")
        join_group = bool((group and group != "Any") or (tsg and tsg != "Any"))
        use_distinct = join_tech or join_group

        select_clause = "SELECT DISTINCT" if use_distinct else "SELECT"
        query = f"""
            {select_clause} s.name, sp.number, sp.title, sp.type, f.filename, f.version, f.url, f.upload_date,
                            COALESCE(ua.tags, ''), COALESCE(ua.comments, ''), COALESCE(p_grp.name, '')
            FROM files f
            JOIN specifications sp ON f.spec_id = sp.id
            JOIN series s ON sp.series_id = s.id
            LEFT JOIN working_groups p_grp ON sp.primary_group_id = p_grp.id
            LEFT JOIN spec_user_annotations ua ON sp.number = ua.spec_number
        """
        if join_tech:
            query += """
            LEFT JOIN spec_radio_tech_map r_map ON sp.id = r_map.spec_id
            LEFT JOIN radio_technologies r ON r_map.tech_id = r.id
            """
        if join_group:
            query += """
            LEFT JOIN spec_secondary_group_map sg_map ON sp.id = sg_map.spec_id
            LEFT JOIN working_groups s_grp ON sg_map.group_id = s_grp.id
            """
        query += """
            WHERE 1=1
              AND f.version IS NOT NULL
              AND TRIM(f.version) != ''
        """
        params = []

        if spec_number:
            query += " AND (sp.number LIKE ? OR sp.type LIKE ? OR (sp.type || ' ' || sp.number) LIKE ? OR sp.title LIKE ? OR ua.tags LIKE ? OR ua.comments LIKE ?)"
            search_term = f"%{spec_number}%"
            params.extend([search_term, search_term, search_term, search_term, search_term, search_term])

        if tag and tag != "Any":
            query += " AND (ua.tags LIKE ? OR ua.tags = ?)"
            params.extend([f"%{tag}%", tag])

        if release_version:
            query += " AND f.version LIKE ?"
            params.append(f"%{release_version}%")

        if series:
            series_list = [s.strip() for s in series.split(',') if s.strip()]
            if series_list:
                clauses = ["sp.number LIKE ?" for _ in series_list]
                params.extend([f"{s}.%" for s in series_list])
                query += f" AND ({' OR '.join(clauses)})"

        if tech and tech != "Any":
            tech_list = [t.strip() for t in tech.split(',') if t.strip()]
            if tech_list:
                clauses = ["(r.name = ? OR sp.radio_technology LIKE ?)" for _ in tech_list]
                for t in tech_list:
                    params.extend([t, f"%{t}%"])
                query += f" AND ({' OR '.join(clauses)})"

        if tsg and tsg != "Any":
            if tsg == "SA":
                query += " AND (p_grp.name LIKE 'SA%' OR s_grp.name LIKE 'SA%')"
            elif tsg == "RAN":
                query += " AND (p_grp.name LIKE 'RAN%' OR s_grp.name LIKE 'RAN%' OR p_grp.name LIKE 'GERAN%' OR s_grp.name LIKE 'GERAN%')"
            elif tsg == "CT":
                query += " AND (p_grp.name LIKE 'CT%' OR s_grp.name LIKE 'CT%' OR p_grp.name LIKE 'CN%' OR s_grp.name LIKE 'CN%')"

        # Working Group matching: supports comma-separated WG names
        if group and group != "Any":
            group_list = [g.strip() for g in group.split(',') if g.strip()]
            if group_list:
                if not tsg and len(group_list) == 1 and group_list[0] in ("SA", "RAN", "CT"):
                    g = group_list[0]
                    if g == "SA":
                        query += " AND (p_grp.name LIKE 'SA%' OR s_grp.name LIKE 'SA%')"
                    elif g == "RAN":
                        query += " AND (p_grp.name LIKE 'RAN%' OR s_grp.name LIKE 'RAN%' OR p_grp.name LIKE 'GERAN%' OR s_grp.name LIKE 'GERAN%')"
                    elif g == "CT":
                        query += " AND (p_grp.name LIKE 'CT%' OR s_grp.name LIKE 'CT%' OR p_grp.name LIKE 'CN%' OR s_grp.name LIKE 'CN%')"
                else:
                    placeholders = ",".join(["?"] * len(group_list))
                    query += f" AND (p_grp.name IN ({placeholders}) OR s_grp.name IN ({placeholders}))"
                    params.extend(group_list)
                    params.extend(group_list)

        if spec_type and spec_type != "Any":
            query += " AND sp.type = ?"
            params.append(spec_type)

        query += " ORDER BY sp.number ASC, f.version DESC"

        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(query, params)
            return cursor.fetchall()

    def get_filtered_specs(
        self,
        series: str,
        tech: str,
        group: str,
        spec_type: str,
        tag: str = None,
        tsg: str = None,
    ) -> list:
        join_tech = bool(tech and tech != "Any")
        join_group = bool((group and group != "Any") or (tsg and tsg != "Any"))
        query = """
            SELECT DISTINCT sp.number 
            FROM specifications sp
            LEFT JOIN spec_user_annotations ua ON sp.number = ua.spec_number
        """
        if join_tech:
            query += """
            LEFT JOIN spec_radio_tech_map r_map ON sp.id = r_map.spec_id
            LEFT JOIN radio_technologies r ON r_map.tech_id = r.id
            """
        if join_group:
            query += """
            LEFT JOIN working_groups p_grp ON sp.primary_group_id = p_grp.id
            LEFT JOIN spec_secondary_group_map sg_map ON sp.id = sg_map.spec_id
            LEFT JOIN working_groups s_grp ON sg_map.group_id = s_grp.id
            """
        query += " WHERE 1=1"
        params = []

        if series and series != "Any":
            series_list = [s.strip() for s in series.split(',') if s.strip()]
            if series_list:
                clauses = ["sp.number LIKE ?" for _ in series_list]
                params.extend([f"{s}.%" for s in series_list])
                query += f" AND ({' OR '.join(clauses)})"

        if tech and tech != "Any":
            tech_list = [t.strip() for t in tech.split(',') if t.strip()]
            if tech_list:
                clauses = ["(r.name = ? OR sp.radio_technology LIKE ?)" for _ in tech_list]
                for t in tech_list:
                    params.extend([t, f"%{t}%"])
                query += f" AND ({' OR '.join(clauses)})"

        if tsg and tsg != "Any":
            if tsg == "SA":
                query += " AND (p_grp.name LIKE 'SA%' OR s_grp.name LIKE 'SA%')"
            elif tsg == "RAN":
                query += " AND (p_grp.name LIKE 'RAN%' OR s_grp.name LIKE 'RAN%' OR p_grp.name LIKE 'GERAN%' OR s_grp.name LIKE 'GERAN%')"
            elif tsg == "CT":
                query += " AND (p_grp.name LIKE 'CT%' OR s_grp.name LIKE 'CT%' OR p_grp.name LIKE 'CN%' OR s_grp.name LIKE 'CN%')"

        # Working Group matching: supports comma-separated WG names
        if group and group != "Any":
            group_list = [g.strip() for g in group.split(',') if g.strip()]
            if group_list:
                if not tsg and len(group_list) == 1 and group_list[0] in ("SA", "RAN", "CT"):
                    g = group_list[0]
                    if g == "SA":
                        query += " AND (p_grp.name LIKE 'SA%' OR s_grp.name LIKE 'SA%')"
                    elif g == "RAN":
                        query += " AND (p_grp.name LIKE 'RAN%' OR s_grp.name LIKE 'RAN%' OR p_grp.name LIKE 'GERAN%' OR s_grp.name LIKE 'GERAN%')"
                    elif g == "CT":
                        query += " AND (p_grp.name LIKE 'CT%' OR s_grp.name LIKE 'CT%' OR p_grp.name LIKE 'CN%' OR s_grp.name LIKE 'CN%')"
                else:
                    placeholders = ",".join(["?"] * len(group_list))
                    query += f" AND (p_grp.name IN ({placeholders}) OR s_grp.name IN ({placeholders}))"
                    params.extend(group_list)
                    params.extend(group_list)

        if spec_type and spec_type != "Any":
            query += " AND sp.type = ?"
            params.append(spec_type)

        if tag and tag != "Any":
            query += " AND (ua.tags LIKE ? OR ua.tags = ?)"
            params.extend([f"%{tag}%", tag])

        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(query, params)
            return [row[0] for row in cursor.fetchall()]

    def get_all_spec_titles(self) -> Dict[str, str]:
        """Return a mapping of specification number (e.g. '23.501') to its formal title."""
        query = "SELECT number, title FROM specifications WHERE title IS NOT NULL AND title != ''"
        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute(query)
                return {row[0]: row[1] for row in cursor.fetchall()}
        except Exception as e:
            self.logger.error(f"Error fetching spec titles: {e}")
            return {}

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

            if details.get('primary_group_id'):
                cursor.execute('SELECT name FROM working_groups WHERE id = ?', (details['primary_group_id'],))
                p_row = cursor.fetchone()
                if p_row:
                    details['primary_group'] = p_row[0]
            details.pop('primary_group_id', None)

            cursor.execute('''
                SELECT r.name FROM radio_technologies r
                JOIN spec_radio_tech_map m ON r.id = m.tech_id
                JOIN specifications s ON s.id = m.spec_id
                WHERE s.number = ?
            ''', (spec_number,))
            techs = [r[0] for r in cursor.fetchall()]
            if techs:
                details['radio_technology'] = ", ".join(techs)

            cursor.execute('''
                SELECT w.name FROM working_groups w
                JOIN spec_secondary_group_map m ON w.id = m.group_id
                JOIN specifications s ON s.id = m.spec_id
                WHERE s.number = ?
            ''', (spec_number,))
            sec_groups = [r[0] for r in cursor.fetchall()]
            if sec_groups:
                details['secondary_groups'] = ", ".join(sec_groups)

            details['related_wis'] = self.get_spec_wis(spec_number)
            details['user_annotations'] = self.get_spec_annotations(spec_number)

            return details

    def get_spec_files(self, spec_number: str) -> List[Dict]:
        """Returns all archive file entries for a specification, sorted by semantic version descending."""
        query = """
            SELECT f.version, f.filename, f.url, f.upload_date
            FROM files f
            JOIN specifications s ON f.spec_id = s.id
            WHERE s.number = ?
        """
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(query, (spec_number,))
            rows = cursor.fetchall()
            if not rows and "-" in spec_number:
                cursor.execute(query, (spec_number.split("-")[0],))
                rows = cursor.fetchall()

            def parse_ver(v_str):
                return [(0, int(x)) if x.isdigit() else (1, str(x)) for x in str(v_str).split(".")]

            sorted_rows = sorted(rows, key=lambda x: parse_ver(x[0]), reverse=True)
            return [
                {"version": r[0], "filename": r[1], "url": r[2], "upload_date": r[3]}
                for r in sorted_rows
            ]
