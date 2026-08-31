# --- File: src/modules/specifications/core/scraper.py ---
import logging
import re
from typing import Dict, List, Tuple
from urllib.parse import urljoin
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from PyQt5.QtCore import QThread, pyqtSignal

import requests
from bs4 import BeautifulSoup

from core.network.session import NetworkSession
from modules.specifications.utils.utils import file_version_to_version
from modules.specifications.core.database import SpecsDatabase


class SpecsCrawlerThread(QThread):
    ui_log_msg = pyqtSignal(str, int)
    finished = pyqtSignal()
    finished_path = pyqtSignal(str)

    def __init__(self, db_path: Path, force_metadata_update: bool = False,
                 target_specs: list = None, root_url: str = "https://www.3gpp.org/ftp/Specs/archive/") -> None:
        super().__init__()
        self.db: SpecsDatabase = SpecsDatabase(db_path)
        self.force_metadata_update: bool = force_metadata_update
        self.target_specs: list = target_specs or []
        self.root_url: str = root_url

        self.session: requests.Session = NetworkSession.get_instance()
        self.spec_folder_pattern: re.Pattern = re.compile(r'^(\d{2}\.\d{2,3}(?:-[a-zA-Z0-9]+)?)/?$')
        self.version_pattern: re.Pattern = re.compile(r'-([a-zA-Z0-9]{3})\.zip$')

        # Support YYYY-MM-DD, YYYY/MM/DD, and DD-MM-YYYY / DD.MM.YYYY
        self.date_pattern: re.Pattern = re.compile(
            r'\b(?:(19\d\d|20\d\d)[-/](0[1-9]|1[0-2])[-/](0[1-9]|[12]\d|3[01])|(0[1-9]|[12]\d|3[01])[./-](0[1-9]|1[0-2])[./-](19\d\d|20\d\d))\b'
        )
        self.ver_str_pattern: re.Pattern = re.compile(r'\bv?(\d{1,2}\.\d{1,2}\.\d{1,2})\b', re.IGNORECASE)

    def _normalize_date(self, raw_date: str) -> str:
        """Standardizes extracted date string to YYYY-MM-DD format."""
        clean = raw_date.strip().replace('/', '-').replace('.', '-')
        parts = clean.split('-')
        if len(parts) == 3:
            if len(parts[0]) == 4:  # YYYY-MM-DD
                return f"{parts[0]}-{parts[1].zfill(2)}-{parts[2].zfill(2)}"
            elif len(parts[2]) == 4:  # DD-MM-YYYY -> YYYY-MM-DD
                return f"{parts[2]}-{parts[1].zfill(2)}-{parts[0].zfill(2)}"
        return clean

    def fetch_links(self, url: str) -> List[Tuple[str, str]]:
        try:
            html_text: str = NetworkSession.get_html(url=url, timeout=20)
            soup: BeautifulSoup = BeautifulSoup(html_text, 'html.parser')
            links: List[Tuple[str, str]] = []

            for a_tag in soup.find_all('a', href=True):
                href: str = a_tag['href']
                if ".." in href or "?" in href or href.startswith(("javascript:", "mailto:")):
                    continue

                absolute_url: str = urljoin(url, href)
                if not absolute_url.startswith(url) or absolute_url == url:
                    continue

                links.append((href, absolute_url))

            return list(dict.fromkeys(links))
        except Exception as e:
            self.ui_log_msg.emit(f"⚠️ Error fetching {url}: {e}", logging.WARNING)
            return []

    def fetch_metadata_from_dynareport(self, spec_number: str) -> Dict:
        clean_number: str = spec_number.replace('.', '')
        url: str = f"https://www.3gpp.org/DynaReport/{clean_number}.htm"
        metadata = {
            'title': '', 'type': '', 'initial_release': '',
            'radio_technology': '', 'radio_technologies_list': [],
            'primary_group': '',
            'secondary_groups_raw': '', 'secondary_groups_list': [],
            'version_dates': {},
            'related_wis': []
        }

        try:
            self.ui_log_msg.emit(f"🔍 Fetching DynaReport: {url}", logging.INFO)
            html_text: str = NetworkSession.get_html(url=url, timeout=15)

            if not html_text:
                self.ui_log_msg.emit(f"⚠️ Empty response received from DynaReport for {spec_number}", logging.WARNING)
                return metadata

            soup: BeautifulSoup = BeautifulSoup(html_text, 'html.parser')

            def get_by_id(keyword: str) -> str:
                tag = soup.find(lambda t: t.has_attr('id') and keyword in t['id'].lower())
                return tag.get_text(strip=True) if tag else ''

            def get_field(*label_texts: str) -> str:
                for label_text in label_texts:
                    tags = soup.find_all(
                        lambda tag: tag.name in ['td', 'th', 'span', 'b', 'strong', 'div', 'label']
                                    and tag.get_text(strip=True).strip(':').lower() == label_text.lower()
                    )
                    for tag in tags:
                        sibling = tag.find_next_sibling(
                            lambda t: t.name in ['td', 'span', 'div'] and t.get_text(strip=True)
                        )
                        if sibling:
                            return sibling.get_text(strip=True)

                        parent_cell = tag.find_parent(['td', 'th'])
                        if parent_cell:
                            next_cell = parent_cell.find_next_sibling(['td', 'th'])
                            if next_cell:
                                return next_cell.get_text(strip=True)

                        next_el = tag.find_next(
                            lambda t: t.name in ['td', 'span', 'div', 'a']
                                      and t.get_text(strip=True)
                                      and t not in tag.descendants
                        )
                        if next_el:
                            val = next_el.get_text(strip=True)
                            if len(val) < 200:
                                return val
                return ''

            metadata['title'] = get_by_id('lbltitle') or get_field('Title', 'Specification Title')

            raw_type: str = get_by_id('lblspectype') or get_field('Specification type', 'Spec type', 'Type')
            if raw_type:
                acronym_match = re.search(r'\(([^)]+)\)', raw_type)
                if acronym_match:
                    metadata['type'] = acronym_match.group(1)
                else:
                    if "Technical Specification" in raw_type:
                        metadata['type'] = "TS"
                    elif "Technical Report" in raw_type:
                        metadata['type'] = "TR"
                    else:
                        metadata['type'] = raw_type

            metadata['initial_release'] = get_by_id('lblinitialrel') or get_field('Initial planned Release',
                                                                                  'Initial Release')

            raw_primary = get_by_id('lblprimarywg') or get_field('Primary responsible group', 'Primary WG')
            if raw_primary:
                p_match = re.search(r'([a-zA-Z]+[\s]*\d*)', raw_primary)
                if p_match:
                    metadata['primary_group'] = p_match.group(1).replace(' ', '').upper()
                else:
                    metadata['primary_group'] = raw_primary.strip()

            raw_sec_groups = get_by_id('lblsecondarywg') or get_field('Secondary responsible groups', 'Secondary WG')
            metadata['secondary_groups_raw'] = raw_sec_groups

            if raw_sec_groups:
                matches = re.findall(r'([a-zA-Z]+[\s]*\d*)', raw_sec_groups)
                clean_matches = [m.replace(' ', '').upper() for m in matches if m.strip()]
                metadata['secondary_groups_list'] = list(dict.fromkeys(clean_matches))

            raw_tech = get_by_id('lblradiotech') or get_field('Radio technology')
            metadata['radio_technology'] = raw_tech

            if raw_tech:
                matches = re.findall(r'(2G|3G|4G|LTE|5G|6G|GSM|UMTS|NB-IOT)', raw_tech, re.IGNORECASE)
                metadata['radio_technologies_list'] = list(dict.fromkeys([m.upper() for m in matches]))

            if not metadata['title']:
                title_tag = soup.find('h1') or soup.find('h2')
                if title_tag:
                    metadata['title'] = title_tag.get_text(strip=True)

            # --- Parse Related Work Items Table ---
            related_wis = []
            wi_grid = soup.find(id=re.compile(r'relatedWiGrid', re.IGNORECASE))
            if wi_grid:
                rows = wi_grid.find_all('tr', class_=re.compile(r'rgRow|rgAltRow'))
                for row in rows:
                    cells = row.find_all(['td', 'th'])
                    if len(cells) >= 3:
                        uid_text = cells[0].get_text(strip=True)
                        acronym_text = cells[1].get_text(strip=True)
                        name_text = cells[2].get_text(strip=True)
                        row_classes = row.get("class", [])
                        is_primary = "rgSelectedRow" in row_classes or "selected" in row_classes

                        # Validate UID format (numeric 4 to 8 digits)
                        if uid_text.isdigit() and len(uid_text) >= 4:
                            related_wis.append({
                                'code': uid_text,
                                'acronym': acronym_text,
                                'title': name_text,
                                'is_primary': is_primary
                            })

            metadata['related_wis'] = related_wis
            if related_wis:
                self.ui_log_msg.emit(f"🔗 Extracted {len(related_wis)} related WI(s) for {spec_number}", logging.INFO)

            # --- Parse Portal Upload Dates from Releases / Versions Table ---
            version_dates = {}
            for row in soup.find_all('tr'):
                cells = row.find_all(['td', 'th'])
                if len(cells) < 2:
                    continue

                row_text = " ".join([c.get_text(strip=True) for c in cells])
                ver_match = self.ver_str_pattern.search(row_text)
                date_match = self.date_pattern.search(row_text)

                if ver_match and date_match:
                    version_str = ver_match.group(1)
                    date_str = self._normalize_date(date_match.group(0))
                    version_dates[version_str] = date_str

            metadata['version_dates'] = version_dates

            if version_dates:
                self.ui_log_msg.emit(f"📅 Extracted {len(version_dates)} release dates for {spec_number}", logging.INFO)
            else:
                self.ui_log_msg.emit(f"⚠️ No release dates found in DynaReport for {spec_number}", logging.WARNING)

        except Exception as e:
            self.ui_log_msg.emit(f"❌ Metadata fetch error for {spec_number} at {url}: {e}", logging.ERROR)

        return metadata

    def fetch_spec_files(self, series_name: str, series_url: str, spec_number: str, spec_url: str) -> dict:
        file_links: List[Tuple[str, str]] = self.fetch_links(spec_url)
        files_to_save = []

        for href, file_url in file_links:
            clean_file_name: str = file_url.split('/')[-1]
            if clean_file_name.endswith('.zip'):
                version_str: str = ""
                match = self.version_pattern.search(clean_file_name)
                if match:
                    version_str = file_version_to_version(match.group(1))
                files_to_save.append((clean_file_name, version_str, file_url))

        return {
            'series_name': series_name, 'series_url': series_url,
            'spec_number': spec_number, 'spec_url': spec_url,
            'files': files_to_save
        }

    def run(self) -> None:
        try:
            if not self.root_url.endswith('/'):
                self.root_url += '/'

            spec_tasks: List[Tuple[str, str, str, str, bool]] = []

            # ==========================================
            # GATHER DIRECTORIES
            # ==========================================
            if self.target_specs:
                self.ui_log_msg.emit(f"⏳ Starting Targeted Update for: {', '.join(self.target_specs)}...", logging.INFO)

                for target in self.target_specs:
                    if '.' not in target:
                        series_number = target
                        series_folder = f"{series_number}_series"
                        series_url = urljoin(self.root_url, f"{series_folder}/")

                        self.ui_log_msg.emit(f"⏳ Mapping entire {series_number} series directory...", logging.INFO)
                        spec_links = self.fetch_links(series_url)

                        for href, spec_url in spec_links:
                            folder_name: str = [x for x in spec_url.split('/') if x][-1]
                            match = self.spec_folder_pattern.search(folder_name)
                            if match:
                                clean_spec_number: str = match.group(1)
                                if not spec_url.endswith('/'):
                                    spec_url += '/'
                                needs_meta = self.force_metadata_update or self.db.needs_metadata(clean_spec_number)
                                spec_tasks.append((series_number, series_url, clean_spec_number, spec_url, needs_meta))
                    else:
                        series_number = target.split('.')[0]
                        series_folder = f"{series_number}_series"
                        series_url = urljoin(self.root_url, f"{series_folder}/")
                        spec_url = urljoin(series_url, f"{target}/")

                        # For targeted single specs, always fetch metadata to refresh dates and WIs
                        needs_meta = True if self.force_metadata_update else (self.db.needs_metadata(target) or True)
                        spec_tasks.append((series_number, series_url, target, spec_url, needs_meta))
            else:
                self.ui_log_msg.emit("⏳ Mapping directories in parallel... (This is fast)", logging.INFO)

                raw_links = self.fetch_links(self.root_url)
                series_links = []

                for href, url in raw_links:
                    folder_name = [x for x in url.split('/') if x][-1]
                    match = re.search(r'^(\d{2,3})_series$', folder_name.lower())
                    if match:
                        clean_series_number = match.group(1)
                        series_links.append((clean_series_number, url))

                with ThreadPoolExecutor(max_workers=15) as executor:
                    future_to_series = {
                        executor.submit(self.fetch_links, s_url if s_url.endswith('/') else s_url + '/'): (
                        s_name, s_url)
                        for s_name, s_url in series_links
                    }

                    for future in as_completed(future_to_series):
                        s_name, s_url = future_to_series[future]
                        specs = future.result()

                        for href, spec_url in specs:
                            folder_name: str = [x for x in spec_url.split('/') if x][-1]
                            match = self.spec_folder_pattern.search(folder_name)
                            if match:
                                clean_spec_number: str = match.group(1)
                                if not spec_url.endswith('/'):
                                    spec_url += '/'
                                needs_meta = self.force_metadata_update or self.db.needs_metadata(clean_spec_number)
                                spec_tasks.append((s_name, s_url, clean_spec_number, spec_url, needs_meta))

            total_specs: int = len(spec_tasks)

            # ==========================================
            # PASS 1: FAST FTP SYNC
            # ==========================================
            self.ui_log_msg.emit(f"📥 Pass 1: Fetching available files for {total_specs} specifications...",
                                 logging.INFO)
            completed: int = 0

            with ThreadPoolExecutor(max_workers=15) as executor:
                futures = {
                    executor.submit(self.fetch_spec_files, task[0], task[1], task[2], task[3]): task
                    for task in spec_tasks
                }

                for future in as_completed(futures):
                    completed += 1
                    if completed % 50 == 0 or completed == total_specs:
                        self.ui_log_msg.emit(f"⏳ Files fetched: {completed}/{total_specs}...", logging.INFO)

                    try:
                        result = future.result()
                        files = result['files']
                        spec_num = result['spec_number']

                        if not files:
                            continue

                        for f_name, f_ver, f_url in files:
                            self.db.insert_or_update_file(
                                result['series_name'], result['series_url'],
                                spec_num, result['spec_url'], f_name, f_ver, f_url
                            )
                    except Exception as e:
                        self.ui_log_msg.emit(f"❌ File fetch error: {e}", logging.ERROR)

            self.ui_log_msg.emit("✅ Pass 1 Complete. Unblocking interface...", logging.INFO)
            self.finished_path.emit("SPECS_DB_PASS_ONE")

            # ==========================================
            # PASS 2: PORTAL METADATA, DATES & WIS SYNC
            # ==========================================
            specs_needing_meta = [task for task in spec_tasks if task[4]]

            if specs_needing_meta:
                self.ui_log_msg.emit(
                    f"⏳ Pass 2: Fetching portal metadata, release dates & related WIs for {len(specs_needing_meta)} specifications...",
                    logging.INFO
                )
                completed_meta: int = 0

                with ThreadPoolExecutor(max_workers=10) as executor:
                    meta_futures = {
                        executor.submit(self.fetch_metadata_from_dynareport, task[2]): task
                        for task in specs_needing_meta
                    }

                    for future in as_completed(meta_futures):
                        task = meta_futures[future]
                        spec_num = task[2]
                        completed_meta += 1

                        if completed_meta % 20 == 0 or completed_meta == len(specs_needing_meta):
                            self.ui_log_msg.emit(f"⏳ Metadata fetched: {completed_meta}/{len(specs_needing_meta)}...",
                                                 logging.INFO)

                        try:
                            metadata = future.result()
                            if metadata:
                                if metadata.get('title'):
                                    self.db.update_spec_metadata(spec_num, metadata)
                                if metadata.get('related_wis'):
                                    self.db.update_spec_wis(spec_num, metadata['related_wis'])
                                if metadata.get('version_dates'):
                                    self.db.update_file_dates(spec_num, metadata['version_dates'])
                                    self.ui_log_msg.emit(
                                        f"💾 Updated {len(metadata['version_dates'])} dates in DB for {spec_num}",
                                        logging.INFO)
                        except Exception as e:
                            self.ui_log_msg.emit(f"❌ Metadata DB update error for {spec_num}: {e}", logging.ERROR)
            else:
                self.ui_log_msg.emit(
                    "ℹ️ Pass 2 skipped: All specifications already have cached metadata. Check 'Force Metadata' to refresh.",
                    logging.INFO)

            self.ui_log_msg.emit("✅ 3GPP Database Update Fully Complete!", logging.INFO)
            self.finished_path.emit("SPECS_DB_PASS_TWO")

        except Exception as e:
            self.ui_log_msg.emit(f"❌ Database Update Failed: {str(e)}", logging.ERROR)
        finally:
            self.finished.emit()