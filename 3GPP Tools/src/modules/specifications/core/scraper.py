# --- File: modules/specs_db/scraper.py ---
import logging
import re
from typing import Dict, List, Tuple
from bs4 import BeautifulSoup
from urllib.parse import urljoin
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from PyQt5.QtCore import QThread, pyqtSignal

from core.network.session import NetworkSession
from modules.specifications.utils.utils import file_version_to_version
from modules.specifications.core.database import SpecsDatabase

class SpecsCrawlerThread(QThread):
    ui_log_msg: pyqtSignal = pyqtSignal(str, int)
    finished: pyqtSignal = pyqtSignal()
    finished_path: pyqtSignal = pyqtSignal(str)

    def __init__(self, db_path: Path, force_metadata_update: bool = False,
                 root_url: str = "https://www.3gpp.org/ftp/Specs/archive/") -> None:
        super().__init__()
        self.db: SpecsDatabase = SpecsDatabase(db_path)
        self.force_metadata_update: bool = force_metadata_update
        self.root_url: str = root_url

        self.spec_folder_pattern: re.Pattern = re.compile(r'^(\d{2}\.\d{2,3})/?$')
        self.version_pattern: re.Pattern = re.compile(r'-([a-zA-Z0-9]{3})\.zip$')

    def fetch_links(self, url: str) -> List[Tuple[str, str]]:
        try:
            # ---> Uses the shared session automatically!
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

    def fetch_metadata_from_dynareport(self, spec_number: str) -> Dict[str, str]:
        clean_number: str = spec_number.replace('.', '')
        url: str = f"https://www.3gpp.org/DynaReport/{clean_number}.htm"
        metadata: Dict[str, str] = {
            'title': '', 'type': '', 'initial_release': '',
            'radio_technology': '', 'primary_group': '', 'secondary_groups': ''
        }

        try:
            # ---> Uses the shared session automatically!
            html_text: str = NetworkSession.get_html(url=url, timeout=15)
            soup: BeautifulSoup = BeautifulSoup(html_text, 'html.parser')

            def get_field(label_text: str) -> str:
                label = soup.find(
                    lambda tag: tag.name in ['td', 'th', 'span', 'b'] and label_text.lower() in tag.text.lower())
                if label:
                    val_cell = label.find_next('td')
                    if val_cell:
                        return val_cell.get_text(strip=True)
                return ''

            metadata['title'] = get_field('Title')

            raw_type: str = get_field('Type')
            if raw_type:
                acronym_match = re.search(r'\(([^)]+)\)', raw_type)
                metadata['type'] = acronym_match.group(1) if acronym_match else raw_type

            metadata['initial_release'] = get_field('Initial planned Release')
            metadata['radio_technology'] = get_field('Radio technology')
            metadata['primary_group'] = get_field('Primary responsible group')
            metadata['secondary_groups'] = get_field('Secondary responsible groups')

            if not metadata['title']:
                title_tag = soup.find('h1') or soup.find('h2')
                if title_tag: metadata['title'] = title_tag.get_text(strip=True)

        except Exception:
            pass

        return metadata

    def process_specification(self, series_name: str, series_url: str, clean_spec_number: str, spec_url: str) -> None:
        file_count: int = 0
        file_links: List[Tuple[str, str]] = self.fetch_links(spec_url)

        for href, file_url in file_links:
            clean_file_name: str = file_url.split('/')[-1]

            if clean_file_name.endswith('.zip'):
                version_str: str = ""
                match = self.version_pattern.search(clean_file_name)
                if match:
                    version_str = file_version_to_version(match.group(1))

                self.db.insert_or_update_file(
                    series_name, series_url,
                    clean_spec_number, spec_url,
                    clean_file_name, version_str, file_url
                )
                file_count += 1

        if file_count == 0:
            return

        if self.force_metadata_update or self.db.needs_metadata(clean_spec_number):
            metadata: Dict[str, str] = self.fetch_metadata_from_dynareport(clean_spec_number)

            if metadata['title']:
                self.db.update_spec_metadata(clean_spec_number, metadata)
                self.ui_log_msg.emit(f"✅ {clean_spec_number}: Saved {file_count} files & updated metadata.", logging.INFO)
            else:
                self.ui_log_msg.emit(f"⚠️ {clean_spec_number}: Saved {file_count} files, but no metadata found.", logging.WARNING)
        else:
            self.ui_log_msg.emit(f"⏭️ {clean_spec_number}: Saved {file_count} files (Metadata skipped).", logging.INFO)

    def run(self) -> None:
        try:
            self.ui_log_msg.emit("⏳ Starting 3GPP Database Synchronization...", logging.INFO)

            if not self.root_url.endswith('/'): self.root_url += '/'
            series_links: List[Tuple[str, str]] = [link for link in self.fetch_links(self.root_url) if 'series' in link[0].lower()]

            spec_tasks: List[Tuple[str, str, str, str]] = []
            for series_name, series_url in series_links:
                if not series_url.endswith('/'): series_url += '/'

                specs = self.fetch_links(series_url)
                for href, spec_url in specs:

                    folder_name: str = [x for x in spec_url.split('/') if x][-1]

                    match = self.spec_folder_pattern.search(folder_name)
                    if match:
                        clean_spec_number: str = match.group(1)
                        if not spec_url.endswith('/'): spec_url += '/'

                        spec_tasks.append((series_name, series_url, clean_spec_number, spec_url))

            total_specs: int = len(spec_tasks)
            self.ui_log_msg.emit(f"📥 Validated {total_specs} specification folders. Processing in parallel...", logging.INFO)

            completed: int = 0
            with ThreadPoolExecutor(max_workers=3) as executor:
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