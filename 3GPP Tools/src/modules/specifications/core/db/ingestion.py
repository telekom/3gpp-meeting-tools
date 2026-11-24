"""Write operations: file ingestion, metadata/WI sync, manual upsert, and deletion."""

from __future__ import annotations

from typing import Dict, List, Optional

from modules.specifications.utils.utils import build_dynareport_url


class SpecsIngestionMixin:
    """Mutating database methods for specifications, archive files, and WI mappings."""

    def insert_or_update_file(
        self,
        series_name: str,
        series_url: str,
        spec_number: str,
        spec_url: str,
        filename: str,
        version: str,
        file_url: str,
        upload_date: Optional[str] = None,
    ):
        # Archive directories contain a small number of historical/non-standard ZIPs
        # whose filenames do not encode a valid 3GPP version. They are not selectable
        # specification versions and must not enter the version table.
        clean_version = (version or "").strip()
        if not clean_version:
            self.logger.debug(
                "Skipping specification archive file without a decoded version: %s (%s)",
                filename, spec_number
            )
            return

        version = clean_version

        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute('INSERT OR IGNORE INTO series (name, url) VALUES (?, ?)', (series_name, series_url))
            cursor.execute('SELECT id FROM series WHERE name = ?', (series_name,))
            series_id = cursor.fetchone()[0]

            cursor.execute('''
                INSERT OR IGNORE INTO specifications (series_id, number, url) 
                VALUES (?, ?, ?)
            ''', (series_id, spec_number, spec_url))
            cursor.execute('SELECT id FROM specifications WHERE number = ?', (spec_number,))
            spec_id = cursor.fetchone()[0]

            cursor.execute('''
                INSERT INTO files (spec_id, filename, version, url, upload_date)
                VALUES (?, ?, ?, ?, ?)
                ON CONFLICT(spec_id, version) DO UPDATE SET
                    filename = excluded.filename,
                    url = excluded.url,
                    upload_date = COALESCE(excluded.upload_date, files.upload_date)
            ''', (spec_id, filename, version, file_url, upload_date))

    def update_file_dates(self, spec_number: str, version_date_map: Dict[str, str]):
        """Batch-updates the portal upload dates for all matched versions of a specification."""
        if not version_date_map:
            return
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute('SELECT id FROM specifications WHERE number = ?', (spec_number,))
            spec_row = cursor.fetchone()
            if not spec_row:
                return
            spec_id = spec_row[0]

            for version, date_str in version_date_map.items():
                clean_ver = version.lstrip('v').strip()
                cursor.execute('''
                    UPDATE files 
                    SET upload_date = ?
                    WHERE spec_id = ? AND (version = ? OR version = ?)
                ''', (date_str, spec_id, clean_ver, f"v{clean_ver}"))

    def update_spec_wis(self, spec_number: str, wis_list: List[Dict]):
        """
        Maps related Work Items to a specification and inserts stubs into work_items
        if the WI has not been fully synchronized yet.
        """
        if wis_list is None:
            return

        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute('SELECT id FROM specifications WHERE number = ?', (spec_number,))
            spec_row = cursor.fetchone()
            if not spec_row:
                return
            spec_id = spec_row[0]

            # Clear existing mappings for this spec to prevent stale links
            cursor.execute('DELETE FROM spec_wi_map WHERE spec_id = ?', (spec_id,))

            for wi in wis_list:
                code = str(wi.get('code', '')).strip()
                if not code:
                    continue

                acronym = wi.get('acronym', '')
                title = wi.get('title', '')
                is_primary = 1 if wi.get('is_primary') else 0

                cursor.execute('''
                    INSERT INTO work_items (code, acronym, name, latest_wid, release, start_date, end_date)
                    VALUES (?, ?, ?, '', '', '', '')
                    ON CONFLICT(code) DO UPDATE SET
                        acronym = CASE WHEN (acronym IS NULL OR acronym = '') THEN excluded.acronym ELSE acronym END,
                        name = CASE WHEN (name IS NULL OR name = '') THEN excluded.name ELSE name END
                ''', (code, acronym, title))

                cursor.execute('''
                    INSERT OR REPLACE INTO spec_wi_map (spec_id, wi_code, is_primary)
                    VALUES (?, ?, ?)
                ''', (spec_id, code, is_primary))

    def update_spec_metadata(self, spec_number: str, metadata: dict):
        """Apply successfully parsed metadata without erasing fields that failed to parse."""
        with self._get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute('SELECT id FROM specifications WHERE number = ?', (spec_number,))
            spec_row = cursor.fetchone()
            if not spec_row:
                return
            spec_id = spec_row[0]

            scalar_updates = []
            scalar_values = []
            for column, key in (
                ('title', 'title'), ('type', 'type'),
                ('initial_release', 'initial_release'),
                ('radio_technology', 'radio_technology'),
                ('secondary_groups', 'secondary_groups_raw'),
            ):
                value = metadata.get(key)
                if value is not None:
                    scalar_updates.append(f'{column} = ?')
                    scalar_values.append(value)

            if metadata.get('primary_group') is not None:
                primary_group_id = None
                p_group = metadata.get('primary_group')
                if p_group:
                    cursor.execute('INSERT OR IGNORE INTO working_groups (name) VALUES (?)', (p_group,))
                    cursor.execute('SELECT id FROM working_groups WHERE name = ?', (p_group,))
                    primary_group_id = cursor.fetchone()[0]
                scalar_updates.append('primary_group_id = ?')
                scalar_values.append(primary_group_id)

            if scalar_updates:
                cursor.execute(
                    f"UPDATE specifications SET {', '.join(scalar_updates)} WHERE number = ?",
                    (*scalar_values, spec_number)
                )

            techs = metadata.get('radio_technologies_list')
            if techs is not None:
                cursor.execute('DELETE FROM spec_radio_tech_map WHERE spec_id = ?', (spec_id,))
                for tech in techs:
                    cursor.execute('INSERT OR IGNORE INTO radio_technologies (name) VALUES (?)', (tech,))
                    cursor.execute('SELECT id FROM radio_technologies WHERE name = ?', (tech,))
                    tech_id = cursor.fetchone()[0]
                    cursor.execute(
                        'INSERT OR IGNORE INTO spec_radio_tech_map (spec_id, tech_id) VALUES (?, ?)',
                        (spec_id, tech_id),
                    )

            sec_groups = metadata.get('secondary_groups_list')
            if sec_groups is not None:
                cursor.execute('DELETE FROM spec_secondary_group_map WHERE spec_id = ?', (spec_id,))
                for sg in sec_groups:
                    cursor.execute('INSERT OR IGNORE INTO working_groups (name) VALUES (?)', (sg,))
                    cursor.execute('SELECT id FROM working_groups WHERE name = ?', (sg,))
                    sg_id = cursor.fetchone()[0]
                    cursor.execute(
                        'INSERT OR IGNORE INTO spec_secondary_group_map (spec_id, group_id) VALUES (?, ?)',
                        (spec_id, sg_id),
                    )

    def needs_metadata(self, spec_number: str) -> bool:
        """Return True when core DynaReport metadata is absent or clearly incomplete."""
        query = """
            SELECT title, type, primary_group_id
            FROM specifications
            WHERE number = ?
        """
        with self._get_connection() as conn:
            result = conn.cursor().execute(query, (spec_number,)).fetchone()
            if not result:
                return True
            title, spec_type, primary_group_id = result
            return not title or not spec_type or primary_group_id is None

    def delete_specification(self, spec_number: str) -> bool:
        """Deletes a specification, its file entries, and relational mappings, then removes orphans."""
        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute("SELECT id FROM specifications WHERE number = ?", (spec_number,))
                row = cursor.fetchone()
                if not row:
                    return False
                spec_id = row[0]

                cursor.execute("DELETE FROM files WHERE spec_id = ?", (spec_id,))
                cursor.execute("DELETE FROM spec_radio_tech_map WHERE spec_id = ?", (spec_id,))
                cursor.execute("DELETE FROM spec_secondary_group_map WHERE spec_id = ?", (spec_id,))
                cursor.execute("DELETE FROM spec_wi_map WHERE spec_id = ?", (spec_id,))
                cursor.execute("DELETE FROM specifications WHERE id = ?", (spec_id,))

            self._cleanup_orphans()
            return True
        except Exception as e:
            self.logger.error(f"Failed to delete specification {spec_number}: {e}")
            return False

    def wipe_database(self) -> bool:
        """Completely purges all specifications, files, and relational mapping tables."""
        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute("DELETE FROM files;")
                cursor.execute("DELETE FROM spec_wi_map;")
                cursor.execute("DELETE FROM spec_secondary_group_map;")
                cursor.execute("DELETE FROM spec_radio_tech_map;")
                cursor.execute("DELETE FROM specifications;")
                cursor.execute("DELETE FROM working_groups;")
                cursor.execute("DELETE FROM radio_technologies;")
                cursor.execute("DELETE FROM series;")
            self.vacuum()
            return True
        except Exception as e:
            self.logger.error(f"Failed to wipe specifications database: {e}")
            return False

    def upsert_manual_spec(self, spec_data: dict) -> bool:
        """Inserts or updates a specification record from the manual addition dialog."""
        spec_number = str(spec_data.get("number", "")).strip()
        if not spec_number:
            return False

        series_name = spec_data.get("series")
        if not series_name:
            series_name = spec_number.split(".")[0].strip()

        title = spec_data.get("title", "").strip()
        spec_type = spec_data.get("type", "TS").strip().upper()
        primary_group = spec_data.get("primary_group", "").strip()
        initial_release = spec_data.get("initial_release", "").strip()
        radio_tech = spec_data.get("radio_technology", "").strip()

        try:
            with self._get_connection() as conn:
                cursor = conn.cursor()
                # Ensure series exists
                cursor.execute(
                    "INSERT OR IGNORE INTO series (name, url) VALUES (?, ?)",
                    (series_name, f"https://www.3gpp.org/ftp/Specs/archive/{series_name}_series/")
                )
                cursor.execute("SELECT id FROM series WHERE name = ?", (series_name,))
                series_id = cursor.fetchone()[0]

                # Ensure primary group exists
                primary_group_id = None
                if primary_group:
                    cursor.execute("INSERT OR IGNORE INTO working_groups (name) VALUES (?)", (primary_group,))
                    cursor.execute("SELECT id FROM working_groups WHERE name = ?", (primary_group,))
                    primary_group_id = cursor.fetchone()[0]

                # Insert or update specification
                dyna_url = build_dynareport_url(spec_number)
                cursor.execute("""
                    INSERT INTO specifications (series_id, number, url, title, type, initial_release, radio_technology, primary_group_id)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                    ON CONFLICT(series_id, number) DO UPDATE SET
                        title = CASE WHEN excluded.title != '' THEN excluded.title ELSE specifications.title END,
                        type = excluded.type,
                        initial_release = CASE WHEN excluded.initial_release != '' THEN excluded.initial_release ELSE specifications.initial_release END,
                        radio_technology = CASE WHEN excluded.radio_technology != '' THEN excluded.radio_technology ELSE specifications.radio_technology END,
                        primary_group_id = COALESCE(excluded.primary_group_id, specifications.primary_group_id)
                """, (
                    series_id, spec_number, dyna_url, title, spec_type, initial_release, radio_tech, primary_group_id
                ))

                cursor.execute("SELECT id FROM specifications WHERE number = ?", (spec_number,))
                spec_id = cursor.fetchone()[0]

                # Link radio tech
                if radio_tech:
                    cursor.execute("INSERT OR IGNORE INTO radio_technologies (name) VALUES (?)", (radio_tech,))
                    cursor.execute("SELECT id FROM radio_technologies WHERE name = ?", (radio_tech,))
                    tech_id = cursor.fetchone()[0]
                    cursor.execute(
                        "INSERT OR IGNORE INTO spec_radio_tech_map (spec_id, tech_id) VALUES (?, ?)",
                        (spec_id, tech_id),
                    )

            return True
        except Exception as e:
            self.logger.error(f"Failed to upsert manual specification: {e}")
            return False
