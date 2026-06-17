# --- File: modules/specs_db/scraper.py ---
import logging
import urllib.request
import urllib.error
import re
from bs4 import BeautifulSoup
from urllib.parse import urljoin
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from PyQt5.QtCore import QThread, pyqtSignal

from modules.specifications.utils.utils import file_version_to_version
from modules.specifications.core.database import SpecsDatabase


class SpecsCrawlerThread(QThread):
    ui_log_msg = pyqtSignal(str, int)
    finished = pyqtSignal()
    finished_path = pyqtSignal(str)

    def __init__(self, db_path: Path, force_metadata_update: bool = False,
                 root_url="https://www.3gpp.org/ftp/Specs/archive/"):
        super().__init__()
        self.db = SpecsDatabase(db_path)
        self.force_metadata_update = force_metadata_update
        self.root_url = root_url

        self.headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
        }

        # ---> NEW: Strict Regex to ensure we ONLY match valid 3GPP Spec numbers (e.g., 23.501 or 02.30)
        self.spec_folder_pattern = re.compile(r'^(\d{2}\.\d{2,3})/?$')
        self.version_pattern = re.compile(r'-([a-zA-Z0-9]{3})\.zip$')

    def _get_html(self, url: str, timeout: int = 20) -> str:
        req = urllib.request.Request(url, headers=self.headers)
        try:
            with urllib.request.urlopen(req, timeout=timeout) as response:
                return response.read().decode('utf-8', errors='ignore')
        except urllib.error.URLError as e:
            raise Exception(f"Connection failed: {e.reason}")

    def fetch_links(self, url: str) -> list:
        try:
            html_text = self._get_html(url, timeout=20)
            soup = BeautifulSoup(html_text, 'html.parser')
            links = []

            for a_tag in soup.find_all('a', href=True):
                href = a_tag['href']

                # Ignore parent directories and query parameters
                if ".." in href or "?" in href or href.startswith(("javascript:", "mailto:")):
                    continue

                absolute_url = urljoin(url, href)

                # Safety check to prevent crawling backwards
                if not absolute_url.startswith(url) or absolute_url == url:
                    continue

                links.append((href, absolute_url))

            # Remove duplicates
            unique_links = []
            for item in links:
                if item not in unique_links:
                    unique_links.append(item)
            return unique_links

        except Exception as e:
            self.ui_log_msg.emit(f"⚠️ Error fetching {url}: {e}", logging.WARNING)
            return []

    def fetch_metadata_from_dynareport(self, spec_number: str) -> dict:
        clean_number = spec_number.replace('.', '')
        url = f"https://www.3gpp.org/DynaReport/{clean_number}.htm"
        metadata = {'title': '', 'type': '', 'initial_release': '', 'radio_technology': '', 'primary_group': '',
                    'secondary_groups': ''}

        try:
            html_text = self._get_html(url, timeout=15)
            soup = BeautifulSoup(html_text, 'html.parser')

            def get_field(label_text):
                label = soup.find(
                    lambda tag: tag.name in ['td', 'th', 'span', 'b'] and label_text.lower() in tag.text.lower())
                if label:
                    val_cell = label.find_next('td')
                    if val_cell:
                        return val_cell.get_text(strip=True)
                return ''

            metadata['title'] = get_field('Title')

            # --- Extract Type as Acronym ---
            raw_type = get_field('Type')
            if raw_type:
                # Regex looks for content inside parentheses: e.g., "Technical Specification (TS)" -> "TS"
                acronym_match = re.search(r'\(([^)]+)\)', raw_type)
                metadata['type'] = acronym_match.group(1) if acronym_match else raw_type

            metadata['initial_release'] = get_field('Initial planned Release')
            metadata['radio_technology'] = get_field('Radio technology')
            metadata['primary_group'] = get_field('Primary responsible group')
            metadata['secondary_groups'] = get_field('Secondary responsible groups')

            if not metadata['title']:
                title_tag = soup.find('h1') or soup.find('h2')
                if title_tag: metadata['title'] = title_tag.get_text(strip=True)

        except Exception as e:
            # Silently fail for missing metadata to keep logs clean
            pass

        return metadata

    def process_specification(self, series_name, series_url, clean_spec_number, spec_url):
        file_count = 0
        file_links = self.fetch_links(spec_url)

        for href, file_url in file_links:
            clean_file_name = file_url.split('/')[-1]

            if clean_file_name.endswith('.zip'):
                version_str = ""
                match = self.version_pattern.search(clean_file_name)
                if match:
                    version_str = file_version_to_version(match.group(1))

                self.db.insert_or_update_file(
                    series_name, series_url,
                    clean_spec_number, spec_url,
                    clean_file_name, version_str, file_url
                )
                file_count += 1

        # ---> NEW: Optimization - Do not fetch metadata if the folder is completely empty!
        if file_count == 0:
            # We don't even log this to keep your console clean, it's just a ghost folder.
            return

        if self.force_metadata_update or self.db.needs_metadata(clean_spec_number):
            metadata = self.fetch_metadata_from_dynareport(clean_spec_number)

            if metadata['title']:
                self.db.update_spec_metadata(clean_spec_number, metadata)
                self.ui_log_msg.emit(f"✅ {clean_spec_number}: Saved {file_count} files & updated metadata.",
                                     logging.INFO)
            else:
                self.ui_log_msg.emit(
                    f"⚠️ {clean_spec_number}: Saved {file_count} files, but no metadata found on 3GPP.",
                    logging.WARNING)
        else:
            self.ui_log_msg.emit(f"⏭️ {clean_spec_number}: Saved {file_count} files (Metadata skipped).", logging.INFO)

    def run(self):
        try:
            self.ui_log_msg.emit("⏳ Starting 3GPP Database Synchronization...", logging.INFO)

            # Fetch Series (Ensure base URL ends with /)
            if not self.root_url.endswith('/'): self.root_url += '/'
            series_links = [link for link in self.fetch_links(self.root_url) if 'series' in link[0].lower()]

            spec_tasks = []
            for series_name, series_url in series_links:
                # Ensure series URL ends with / so urljoin works perfectly
                if not series_url.endswith('/'): series_url += '/'

                specs = self.fetch_links(series_url)
                for href, spec_url in specs:

                    # ---> FIX: Extract just the folder name (e.g. "23.501") before checking the Regex
                    folder_name = [x for x in spec_url.split('/') if x][-1]

                    match = self.spec_folder_pattern.search(folder_name)
                    if match:
                        clean_spec_number = match.group(1)
                        if not spec_url.endswith('/'): spec_url += '/'

                        spec_tasks.append((series_name, series_url, clean_spec_number, spec_url))

            total_specs = len(spec_tasks)
            self.ui_log_msg.emit(f"📥 Validated {total_specs} specification folders. Processing in parallel...",
                                 logging.INFO)

            completed = 0
            with ThreadPoolExecutor(max_workers=10) as executor:
                futures = {executor.submit(self.process_specification, *task): task for task in spec_tasks}
                for future in as_completed(futures):
                    completed += 1
                    if completed % 50 == 0 or completed == total_specs:
                        self.ui_log_msg.emit(f"⏳ Parsed {completed}/{total_specs} specifications...", logging.INFO)

                    try:
                        future.result()
                    except Exception as e:
                        self.ui_log_msg.emit(f"❌ Thread processing error: {e}", logging.ERROR)

            self.ui_log_msg.emit("✅ 3GPP Database Update Complete!", logging.INFO)
            self.finished_path.emit("SPECS_DB_UPDATED")

        except Exception as e:
            self.ui_log_msg.emit(f"❌ Database Update Failed: {str(e)}", logging.ERROR)
        finally:
            self.finished.emit()