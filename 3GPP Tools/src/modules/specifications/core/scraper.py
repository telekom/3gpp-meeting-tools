# --- File: modules/specs_db/scraper.py ---
import logging
import requests
import re
from bs4 import BeautifulSoup
from urllib.parse import urljoin
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from PyQt5.QtCore import QThread, pyqtSignal

from core.utils.utils import get_proxies
from modules.specifications.utils.utils import file_version_to_version
from modules.specifications.core.database import SpecsDatabase


class SpecsCrawlerThread(QThread):
    ui_log_msg = pyqtSignal(str, int)
    finished = pyqtSignal()

    def __init__(self, db_path: Path, force_metadata_update: bool = False,
                 root_url="https://www.3gpp.org/ftp/Specs/archive/"):
        super().__init__()
        self.db = SpecsDatabase(db_path)
        self.force_metadata_update = force_metadata_update
        self.root_url = root_url
        self.proxies = get_proxies()
        self.session = requests.Session()
        self.session.proxies.update(self.proxies)

        self.link_pattern = re.compile(r'<a href="([^"]+)"')
        self.version_pattern = re.compile(r'-([a-zA-Z0-9]{3})\.zip$')

    def fetch_links(self, url: str) -> list:
        try:
            response = self.session.get(url, timeout=15)
            response.raise_for_status()
            matches = self.link_pattern.findall(response.text)
            links = []
            for match in matches:
                if ".." in match or "?" in match or match.startswith("/"): continue
                absolute_url = urljoin(url, match)
                if not absolute_url.endswith('/') and not '.' in match.split('/')[-1]:
                    absolute_url += '/'
                links.append((match, absolute_url))
            return links
        except Exception as e:
            self.ui_log_msg.emit(f"⚠️ Error fetching {url}: {e}", logging.WARNING)
            return []

    def fetch_metadata_from_dynareport(self, spec_number: str) -> dict:
        clean_number = spec_number.replace('.', '')
        url = f"https://www.3gpp.org/DynaReport/{clean_number}.htm"
        metadata = {'title': '', 'type': '', 'initial_release': '', 'radio_technology': '', 'primary_group': '',
                    'secondary_groups': ''}

        try:
            response = self.session.get(url, timeout=10)
            if response.status_code != 200:
                return metadata

            soup = BeautifulSoup(response.text, 'html.parser')

            def get_field(label_text):
                label = soup.find(
                    lambda tag: tag.name in ['td', 'th', 'span', 'b'] and label_text.lower() in tag.text.lower())
                if label:
                    val_cell = label.find_next('td')
                    if val_cell:
                        return val_cell.get_text(strip=True)
                return ''

            metadata['title'] = get_field('Title')
            metadata['type'] = get_field('Type')
            metadata['initial_release'] = get_field('Initial planned Release')
            metadata['radio_technology'] = get_field('Radio technology')
            metadata['primary_group'] = get_field('Primary responsible group')
            metadata['secondary_groups'] = get_field('Secondary responsible groups')

            if not metadata['title']:
                title_tag = soup.find('h1') or soup.find('h2')
                if title_tag: metadata['title'] = title_tag.get_text(strip=True)

        except Exception as e:
            self.ui_log_msg.emit(f"⚠️ Failed to fetch metadata for {spec_number}: {e}", logging.WARNING)

        return metadata

    def process_specification(self, series_name, series_url, spec_name, spec_url):
        clean_spec_number = spec_name.strip('/')

        # 1. Store Files
        file_links = self.fetch_links(spec_url)
        for file_name, file_url in file_links:
            if file_name.endswith('.zip'):
                version_str = ""
                match = self.version_pattern.search(file_name)
                if match:
                    version_str = file_version_to_version(match.group(1))

                self.db.insert_or_update_file(
                    series_name, series_url,
                    clean_spec_number, spec_url,
                    file_name, version_str, file_url
                )

        # 2. Store Metadata
        if self.force_metadata_update or self.db.needs_metadata(clean_spec_number):
            metadata = self.fetch_metadata_from_dynareport(clean_spec_number)
            if metadata['title']:
                self.db.update_spec_metadata(clean_spec_number, metadata)

    def run(self):
        try:
            self.ui_log_msg.emit("⏳ Starting 3GPP Database Synchronization...", logging.INFO)
            series_links = [link for link in self.fetch_links(self.root_url) if 'series' in link[0].lower()]

            spec_tasks = []
            for series_name, series_url in series_links:
                specs = self.fetch_links(series_url)
                for spec_name, spec_url in specs:
                    spec_tasks.append((series_name, series_url, spec_name, spec_url))

            total_specs = len(spec_tasks)
            self.ui_log_msg.emit(f"📥 Found {total_specs} specifications. Processing in parallel...", logging.INFO)

            completed = 0
            with ThreadPoolExecutor(max_workers=10) as executor:
                futures = {executor.submit(self.process_specification, *task): task for task in spec_tasks}
                for future in as_completed(futures):
                    completed += 1
                    if completed % 50 == 0 or completed == total_specs:
                        self.ui_log_msg.emit(f"⏳ Parsed {completed}/{total_specs} specifications...", logging.INFO)
                    future.result()

            self.ui_log_msg.emit("✅ 3GPP Database Update Complete!", logging.INFO)
        except Exception as e:
            self.ui_log_msg.emit(f"❌ Database Update Failed: {str(e)}", logging.ERROR)
        finally:
            self.finished.emit()