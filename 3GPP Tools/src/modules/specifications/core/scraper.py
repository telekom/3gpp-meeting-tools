# --- File: src/modules/specifications/core/scraper.py ---
import logging
import re
from typing import Callable, Dict, List, Optional, Tuple
from urllib.parse import urljoin
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from PyQt5.QtCore import QThread, pyqtSignal
from bs4 import BeautifulSoup

from core.network.session import HttpError, NetworkError, NetworkSession
from modules.specifications.utils.utils import file_version_to_version
from modules.specifications.core.database import SpecsDatabase

logger = logging.getLogger(__name__)

RE_VERSION_STR = re.compile(r'\bv?(\d{1,2}\.\d{1,2}\.\d{1,2})\b', re.IGNORECASE)
RE_DATE = re.compile(
    r'\b(?:(19\d\d|20\d\d)[-/](0[1-9]|1[0-2])[-/](0[1-9]|[12]\d|3[01])|(0[1-9]|[12]\d|3[01])[./-](0[1-9]|1[0-2])[./-](19\d\d|20\d\d))\b'
)


def normalize_date(raw_date: str) -> str:
    """Standardizes extracted date string to YYYY-MM-DD format."""
    clean = raw_date.strip().replace('/', '-').replace('.', '-')
    parts = clean.split('-')
    if len(parts) == 3:
        if len(parts[0]) == 4:
            return f"{parts[0]}-{parts[1].zfill(2)}-{parts[2].zfill(2)}"
        elif len(parts[2]) == 4:
            return f"{parts[2]}-{parts[1].zfill(2)}-{parts[0].zfill(2)}"
    return clean


def normalize_working_group(raw_group: str) -> str:
    """Converts 3GPP group descriptions (e.g., 'S2 (SA WG2)', 'R1') to standard codes (SA2, RAN1)."""
    if not raw_group:
        return ""
    g = raw_group.strip()
    if g in ("-", "N/A", "none"):
        return ""

    match_full = re.search(r'\b(SA|RAN|CT|GERAN)\s*(?:WG)?\s*([1-6])\b', g, re.IGNORECASE)
    if match_full:
        return f"{match_full.group(1).upper()}{match_full.group(2)}"

    legacy_map = {
        "S1": "SA1", "S2": "SA2", "S3": "SA3", "S4": "SA4", "S5": "SA5", "S6": "SA6", "SP": "SA",
        "R1": "RAN1", "R2": "RAN2", "R3": "RAN3", "R4": "RAN4", "R5": "RAN5", "RP": "RAN",
        "C1": "CT1", "C3": "CT3", "C4": "CT4", "C6": "CT6", "CP": "CT"
    }
    first_token = re.split(r'[\s(]', g)[0].upper()
    return legacy_map.get(first_token, g)


def normalize_release(raw_rel: str) -> str:
    """Normalizes release designations like 'Release 15' or '15' to 'Rel-15'."""
    if not raw_rel or raw_rel.strip() in ("-", "N/A"):
        return ""
    match = re.search(r'\b(?:Rel(?:ease)?[- ]?)?(\d+)\b', raw_rel.strip(), re.IGNORECASE)
    return f"Rel-{match.group(1)}" if match else raw_rel.strip()


def deduce_spec_type(spec_number: str, raw_type: str = "") -> str:
    """Deduces TS vs TR using official type description and 3GPP numbering rules."""
    if raw_type:
        if "Technical Report" in raw_type or "(TR)" in raw_type.upper():
            return "TR"
        if "Technical Specification" in raw_type or "(TS)" in raw_type.upper():
            return "TS"
    if re.search(r'\b\d{2}\.[89]\d{2}\b', spec_number):
        return "TR"
    return "TS"


def _parse_dynareport_content(html_text: str, clean_spec: str) -> Dict:
    """Extracts specification attributes from DynaReport HTML, ignoring data grids."""
    metadata = {
        'title': '', 'type': '', 'initial_release': '',
        'radio_technology': '', 'radio_technologies_list': [],
        'primary_group': '',
        'secondary_groups_raw': '', 'secondary_groups_list': [],
        'version_dates': {},
        'related_wis': []
    }
    soup = BeautifulSoup(html_text, 'html.parser')

    # Decompose data grids so table headers (e.g. UID) don't collide with spec metadata
    for grid in soup.find_all(lambda tag: tag.name in ('table', 'div') and tag.has_attr('id') and (
        'grid' in tag['id'].lower() or 'releases' in tag['id'].lower()
    )):
        grid.decompose()

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
        return ''

    # 1. Title
    title = get_by_id('lbltitle') or get_field('Specification Title', 'Title')
    if title.lower() in ("uid", "3gpp specification detail", "specification detail", "title"):
        title = ""
    metadata['title'] = title

    # 2. Type (TS vs TR)
    raw_type = get_by_id('lblspectype') or get_field('Specification type', 'Spec type', 'Type')
    metadata['type'] = deduce_spec_type(clean_spec, raw_type)

    # 3. Initial Planned Release
    raw_rel = get_by_id('lblinitialrel') or get_field('Initial planned Release', 'Initial Release')
    metadata['initial_release'] = normalize_release(raw_rel)

    # 4. Primary Responsible Group
    raw_primary = get_by_id('lblprimarywg') or get_field('Primary responsible group', 'Primary WG')
    metadata['primary_group'] = normalize_working_group(raw_primary)

    # 5. Secondary Responsible Groups
    raw_sec = get_by_id('lblsecondarywg') or get_field('Secondary responsible groups', 'Secondary WG')
    metadata['secondary_groups_raw'] = raw_sec
    if raw_sec:
        matches = re.findall(r'([a-zA-Z]+[\s]*\d*)', raw_sec)
        clean_matches = [normalize_working_group(m) for m in matches if m.strip()]
        metadata['secondary_groups_list'] = list(dict.fromkeys([c for c in clean_matches if c]))

    # 6. Radio Technology
    raw_tech = get_by_id('lblradiotech') or get_field('Radio technology')
    if raw_tech:
        matches = re.findall(r'(2G|3G|4G|LTE|5G|6G|GSM|UMTS|NB-IOT)', raw_tech, re.IGNORECASE)
        metadata['radio_technologies_list'] = list(dict.fromkeys([m.upper() for m in matches]))
        metadata['radio_technology'] = ", ".join(metadata['radio_technologies_list'])

    return metadata


def fetch_metadata_from_dynareport(
    spec_number: str,
    log_cb: Optional[Callable[[str, int], None]] = None
) -> Dict:
    """
    Fetches and parses specification metadata from 3GPP DynaReport HTML with logging.
    Tests user-entered candidate format first and only accepts pages with valid specification titles.
    """
    def log(msg: str, level: int = logging.INFO):
        logger.log(level, msg)
        if log_cb:
            log_cb(msg, level)

    clean_spec = spec_number.strip()
    cleaned_num = re.sub(r'^(?:3GPP\s+)?(?:TS|TR)\s*', '', clean_spec, flags=re.IGNORECASE).strip()
    match_parts = re.search(r'^(\d{2})\.?(\d{3})(?:[-_.](\d{1,2}))?', cleaned_num)

    candidates: List[str] = []
    if match_parts:
        series, core, part = match_parts.group(1), match_parts.group(2), match_parts.group(3)
        base = f"{series}{core}"
        if part:
            # 1. Exact user-entered format first (e.g., 23801-01.htm)
            candidates.append(f"https://www.3gpp.org/DynaReport/{base}-{part}.htm")

            # 2. Integer format without leading zeros (e.g., 23801-1.htm)
            int_part = str(int(part))
            int_url = f"https://www.3gpp.org/DynaReport/{base}-{int_part}.htm"
            if int_url not in candidates:
                candidates.append(int_url)

            # 3. Two-digit zero-padded format (e.g., 23801-01.htm)
            padded = int_part.zfill(2)
            padded_url = f"https://www.3gpp.org/DynaReport/{base}-{padded}.htm"
            if padded_url not in candidates:
                candidates.append(padded_url)

        # 4. Base specification fallback (e.g., 23801.htm)
        candidates.append(f"https://www.3gpp.org/DynaReport/{base}.htm")
    else:
        candidates.append(f"https://www.3gpp.org/DynaReport/{clean_spec.replace('.', '')}.htm")

    metadata = {
        'number': clean_spec,
        'title': '', 'type': '', 'initial_release': '',
        'radio_technology': '', 'radio_technologies_list': [],
        'primary_group': '',
        'secondary_groups_raw': '', 'secondary_groups_list': [],
        'version_dates': {},
        'related_wis': [],
        'error': ''
    }

    attempt_errors: List[str] = []

    for url in candidates:
        url_name = url.split('/')[-1]
        log(f"🔍 Trying DynaReport: {url}...", logging.INFO)
        try:
            resp_text = NetworkSession.get_html(url=url, timeout=12)
            if not resp_text:
                attempt_errors.append(f"{url_name} (Empty response)")
                continue

            parsed = _parse_dynareport_content(resp_text, clean_spec)
            if parsed.get('title'):
                metadata.update(parsed)
                log(f"✅ Found valid DynaReport at: {url} (Title: '{metadata['title'][:30]}...')", logging.INFO)
                return metadata
            else:
                attempt_errors.append(f"{url_name} (Blank report template - no title)")
                log(f"ℹ️ {url_name} returned blank template without title; checking next candidate...", logging.INFO)

        except HttpError as http_err:
            attempt_errors.append(f"{url_name} (HTTP {http_err.status_code})")
            log(f"⚠️ {url_name} returned HTTP {http_err.status_code}", logging.WARNING)
        except NetworkError as net_err:
            attempt_errors.append(f"{url_name} (Network error)")
            log(f"⚠️ {url_name} network error: {net_err}", logging.WARNING)
        except Exception as e:
            attempt_errors.append(f"{url_name} ({e})")
            log(f"⚠️ {url_name} error: {e}", logging.WARNING)

    err_detail = ", ".join(attempt_errors) if attempt_errors else "All candidate URLs failed"
    metadata['error'] = f"Report not found on 3GPP server ({err_detail})"
    log(f"❌ {metadata['error']} for '{spec_number}'", logging.ERROR)
    return metadata


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

        self.spec_folder_pattern: re.Pattern = re.compile(r'^(\d{2}\.\d{2,3}(?:-[a-zA-Z0-9]+)?)/?$')
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
        return fetch_metadata_from_dynareport(spec_number, log_cb=self.ui_log_msg.emit)

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
                        executor.submit(self.fetch_links, s_url if s_url.endswith('/') else s_url + '/'): (s_name, s_url)
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

            # PASS 1: FAST FTP SYNC
            self.ui_log_msg.emit(f"📥 Pass 1: Fetching available files for {total_specs} specifications...", logging.INFO)
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

            # PASS 2: PORTAL METADATA, DATES & WIS SYNC
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
                            self.ui_log_msg.emit(f"⏳ Metadata fetched: {completed_meta}/{len(specs_needing_meta)}...", logging.INFO)

                        try:
                            metadata = future.result()
                            if metadata:
                                if metadata.get('title'):
                                    self.db.update_spec_metadata(spec_num, metadata)
                                if metadata.get('related_wis'):
                                    self.db.update_spec_wis(spec_num, metadata['related_wis'])
                                if metadata.get('version_dates'):
                                    self.db.update_file_dates(spec_num, metadata['version_dates'])
                                    self.ui_log_msg.emit(f"💾 Updated {len(metadata['version_dates'])} dates in DB for {spec_num}", logging.INFO)
                        except Exception as e:
                            self.ui_log_msg.emit(f"❌ Metadata DB update error for {spec_num}: {e}", logging.ERROR)
            else:
                self.ui_log_msg.emit("ℹ️ Pass 2 skipped: All specifications already have cached metadata.", logging.INFO)

            self.ui_log_msg.emit("✅ 3GPP Database Update Fully Complete!", logging.INFO)
            self.finished_path.emit("SPECS_DB_PASS_TWO")

        except Exception as e:
            self.ui_log_msg.emit(f"❌ Database Update Failed: {str(e)}", logging.ERROR)
        finally:
            self.finished.emit()