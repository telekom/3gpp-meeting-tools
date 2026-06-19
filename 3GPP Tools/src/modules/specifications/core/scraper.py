# --- File: modules/specifications/core/scraper.py ---
import logging
import re
from typing import Dict, List, Tuple

import requests
from bs4 import BeautifulSoup
from urllib.parse import urljoin
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from PyQt5.QtCore import QThread, pyqtSignal

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
        self.spec_folder_pattern: re.Pattern = re.compile(r'^(\d{2}\.\d{2,3})/?$')
        self.version_pattern: re.Pattern = re.compile(r'-([a-zA-Z0-9]{3})\.zip$')

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
            'secondary_groups_raw': '', 'secondary_groups_list': []
        }

        try:
            html_text: str = NetworkSession.get_html(url=url, timeout=15)
            soup: BeautifulSoup = BeautifulSoup(html_text, 'html.parser')

            def get_by_id(keyword: str) -> str:
                tag = soup.find(lambda t: t.has_attr('id') and keyword in t['id'].lower())
                return tag.get_text(strip=True) if tag else ''

            def get_field(*label_texts: str) -> str:
                for label_text in label_texts:
                    tags = soup.find_all(lambda tag: tag.name in ['td', 'th', 'span', 'b', 'strong', 'div', 'label']
                                                     and tag.get_text(strip=True).strip(
                        ':').lower() == label_text.lower())
                    for tag in tags:
                        sibling = tag.find_next_sibling(
                            lambda t: t.name in ['td', 'span', 'div'] and t.get_text(strip=True))
                        if sibling: return sibling.get_text(strip=True)

                        parent_cell = tag.find_parent(['td', 'th'])
                        if parent_cell:
                            next_cell = parent_cell.find_next_sibling(['td', 'th'])
                            if next_cell: return next_cell.get_text(strip=True)

                        next_el = tag.find_next(lambda t: t.name in ['td', 'span', 'div', 'a'] and t.get_text(
                            strip=True) and t not in tag.descendants)
                        if next_el:
                            val = next_el.get_text(strip=True)
                            if len(val) < 200: return val
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

            # ---> UPGRADED: Primary Group Parsing (with strict whitespace sanitization)
            raw_primary = get_by_id('lblprimarywg') or get_field('Primary responsible group', 'Primary WG')
            if raw_primary:
                # Capture the letters and numbers, ignoring surrounding text
                p_match = re.search(r'([a-zA-Z]+[\s]*\d*)', raw_primary)
                if p_match:
                    # Strip out the internal spaces (e.g. "CT 1" -> "CT1") and force uppercase
                    metadata['primary_group'] = p_match.group(1).replace(' ', '').upper()
                else:
                    metadata['primary_group'] = raw_primary.strip()
            else:
                metadata['primary_group'] = ''

            raw_sec_groups = get_by_id('lblsecondarywg') or get_field('Secondary responsible groups', 'Secondary WG')
            metadata['secondary_groups_raw'] = raw_sec_groups

            if raw_sec_groups:
                # ---> UPGRADED: Smart regex that binds letters and trailing numbers together even with spaces
                matches = re.findall(r'([a-zA-Z]+[\s]*\d*)', raw_sec_groups)

                # Strip out extra whitespace inside the acronym (e.g., "SA 2" becomes "SA2") for perfectly clean DB storage
                clean_matches = [m.replace(' ', '').upper() for m in matches if m.strip()]
                metadata['secondary_groups_list'] = list(dict.fromkeys(clean_matches))

            # Radio Technology Parsing
            raw_tech = get_by_id('lblradiotech') or get_field('Radio technology')
            metadata['radio_technology'] = raw_tech

            if raw_tech:
                matches = re.findall(r'(2G|3G|4G|LTE|5G|6G|GSM|UMTS|NB-IOT)', raw_tech, re.IGNORECASE)
                metadata['radio_technologies_list'] = list(dict.fromkeys([m.upper() for m in matches]))

            if not metadata['title']:
                title_tag = soup.find('h1') or soup.find('h2')
                if title_tag: metadata['title'] = title_tag.get_text(strip=True)

        except Exception as e:
            logging.warning(f"Metadata fetch failed for {spec_number} at {url}: {e}")

        return metadata

    def fetch_spec_network_data(self, series_name: str, series_url: str, spec_number: str, spec_url: str,
                                needs_metadata: bool) -> dict:
        file_links: List[Tuple[str, str]] = self.fetch_links(spec_url)
        files_to_save = []

        for href, file_url in file_links:
            clean_file_name: str = file_url.split('/')[-1]
            if clean_file_name.endswith('.zip'):
                version_str: str = ""
                match = self.version_pattern.search(clean_file_name)
                if match: version_str = file_version_to_version(match.group(1))
                files_to_save.append((clean_file_name, version_str, file_url))

        metadata = None
        if len(files_to_save) > 0 and needs_metadata:
            metadata = self.fetch_metadata_from_dynareport(spec_number)

        return {
            'series_name': series_name, 'series_url': series_url,
            'spec_number': spec_number, 'spec_url': spec_url,
            'files': files_to_save, 'metadata': metadata
        }

    def run(self) -> None:
        try:
            if not self.root_url.endswith('/'): self.root_url += '/'

            spec_tasks: List[Tuple[str, str, str, str, bool]] = []

            if self.target_specs:
                self.ui_log_msg.emit(f"⏳ Starting Targeted Update for {len(self.target_specs)} specifications...",
                                     logging.INFO)

                for spec_num in self.target_specs:
                    series_name = f"{spec_num[:2]}_series"
                    series_url = urljoin(self.root_url, f"{series_name}/")
                    spec_url = urljoin(series_url, f"{spec_num}/")
                    needs_meta = self.force_metadata_update or self.db.needs_metadata(spec_num)
                    spec_tasks.append((series_name, series_url, spec_num, spec_url, needs_meta))
            else:
                self.ui_log_msg.emit("⏳ Mapping directories in parallel... (This is fast)", logging.INFO)

                series_links: List[Tuple[str, str]] = [
                    link for link in self.fetch_links(self.root_url) if 'series' in link[0].lower()
                ]

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
                                if not spec_url.endswith('/'): spec_url += '/'
                                needs_meta = self.force_metadata_update or self.db.needs_metadata(clean_spec_number)
                                spec_tasks.append((s_name, s_url, clean_spec_number, spec_url, needs_meta))

            total_specs: int = len(spec_tasks)
            self.ui_log_msg.emit(f"📥 Found {total_specs} specifications. Processing downloads...", logging.INFO)

            completed: int = 0

            with ThreadPoolExecutor(max_workers=15) as executor:
                futures = {executor.submit(self.fetch_spec_network_data, *task): task for task in spec_tasks}

                for future in as_completed(futures):
                    completed += 1
                    if completed % 50 == 0 or completed == total_specs:
                        self.ui_log_msg.emit(f"⏳ Processed {completed}/{total_specs} specifications...", logging.INFO)

                    try:
                        result = future.result()
                        files = result['files']
                        spec_num = result['spec_number']

                        if not files: continue

                        for f_name, f_ver, f_url in files:
                            self.db.insert_or_update_file(
                                result['series_name'], result['series_url'],
                                spec_num, result['spec_url'], f_name, f_ver, f_url
                            )

                        if result['metadata']:
                            if result['metadata']['title']:
                                self.db.update_spec_metadata(spec_num, result['metadata'])
                                self.ui_log_msg.emit(f"✅ {spec_num}: Saved {len(files)} files & updated metadata.",
                                                     logging.INFO)
                            else:
                                self.ui_log_msg.emit(f"⚠️ {spec_num}: Saved {len(files)} files, no metadata found.",
                                                     logging.WARNING)
                        else:
                            self.ui_log_msg.emit(f"⏭️ {spec_num}: Saved {len(files)} files (Metadata skipped).",
                                                 logging.INFO)

                    except Exception as e:
                        self.ui_log_msg.emit(f"❌ Processing error: {e}", logging.ERROR)

            self.ui_log_msg.emit("✅ 3GPP Database Update Complete!", logging.INFO)
            self.finished_path.emit("SPECS_DB_UPDATED")

        except Exception as e:
            self.ui_log_msg.emit(f"❌ Database Update Failed: {str(e)}", logging.ERROR)
        finally:
            self.finished.emit()