import concurrent.futures
import logging
import re

from PyQt5.QtCore import QThread, pyqtSignal
from bs4 import BeautifulSoup

from core.network.session import NetworkSession
from modules.work_items.core.wi_database import WorkItemsDatabase

logger = logging.getLogger(__name__)


# Mapping standard Working Group names to the 3GPP dynareport URL codes
PORTAL_TO_WG_MAP = {
    "SP": "SA", "S1": "SA1", "S2": "SA2", "S3": "SA3", "S4": "SA4", "S5": "SA5", "S6": "SA6",
    "RP": "RAN", "R1": "RAN1", "R2": "RAN2", "R3": "RAN3", "R4": "RAN4", "R5": "RAN5", "R6": "RAN6",
    "CP": "CT", "C1": "CT1", "C3": "CT3", "C4": "CT4", "C6": "CT6"
}


def extract_wi_code(query: str) -> str:
    """
    Extracts a numeric Work Item code from diverse user input formats:
    - Pure numbers: '1121123'
    - Hash / prefix: '#1121123', 'WI #1121123', 'WI 1121123'
    - Portal URLs: 'https://portal.3gpp.org/desktopmodules/WorkItem/WorkItemDetails.aspx?workitemId=1121123'
    """
    query = (query or "").strip()
    if not query:
        return ""
    m = re.search(r'workitemId=(\d+)', query, re.IGNORECASE)
    if m:
        return m.group(1)
    m = re.search(r'(?:WI\s*#?|#)?\s*(\d{4,8})\b', query, re.IGNORECASE)
    if m:
        return m.group(1)
    m = re.search(r'(\d+)', query)
    return m.group(1) if m else query


def fetch_wi_details_from_portal(wi_code_or_url: str, log_cb=None) -> dict:
    """
    Scrapes detailed metadata, responsible working groups, and remarks history
    directly from 3GPP Portal WorkItemDetails.aspx.
    """
    code = extract_wi_code(wi_code_or_url)
    if not code:
        if log_cb:
            log_cb("Invalid Work Item ID or URL provided.", logging.ERROR)
        return {"error": "Invalid Work Item ID"}

    url = f"https://portal.3gpp.org/desktopmodules/WorkItem/WorkItemDetails.aspx?workitemId={code}"
    if log_cb:
        log_cb(f"Fetching details from {url}...", logging.INFO)
    logger.info(f"Fetching WI details from: {url}")

    session = NetworkSession.get_instance()
    NetworkSession.apply_humanness(session)

    try:
        response = session.get(url, timeout=30)
        if response.status_code in (404, 500):
            msg = f"Work Item #{code} was not found on the 3GPP Portal (HTTP {response.status_code})."
            if log_cb:
                log_cb(msg, logging.WARNING)
            return {"error": msg, "code": code}
        response.raise_for_status()
    except Exception as e:
        msg = f"Network or HTTP error fetching WI #{code}: {e}"
        if log_cb:
            log_cb(msg, logging.ERROR)
        return {"error": str(e), "code": code}

    soup = BeautifulSoup(response.text, "html.parser")

    metadata = {
        "code": code,
        "name": "",
        "acronym": "",
        "release": "",
        "status": "",
        "start_date": "",
        "end_date": "",
        "latest_wid": "",
        "responsible_groups": [],
        "remarks": [],
        "rapporteur": "",
        "parent_wi": "",
    }

    # Name / Title
    name_tag = soup.find("span", id="lblName")
    if name_tag:
        metadata["name"] = name_tag.get_text(strip=True)

    # Acronym & Effective Acronym
    acr_tag = soup.find("span", id="lblAcronym") or soup.find("span", id="lblEffectiveAcronym")
    if acr_tag:
        metadata["acronym"] = acr_tag.get_text(strip=True)

    # Release
    rel_tag = soup.find("span", id="lblRelease")
    if rel_tag:
        metadata["release"] = rel_tag.get_text(strip=True)

    # Status
    st_tag = soup.find("span", id="lblStatus")
    if st_tag:
        metadata["status"] = st_tag.get_text(strip=True)

    # Dates
    start_tag = soup.find("span", id="lblStartDate")
    if start_tag:
        metadata["start_date"] = start_tag.get_text(strip=True)

    end_tag = soup.find("span", id="lblEndDate")
    if end_tag:
        metadata["end_date"] = end_tag.get_text(strip=True)

    # Latest WID Version
    wid_tag = soup.find("a", id="lnkWiVersion")
    if wid_tag:
        metadata["latest_wid"] = wid_tag.get_text(strip=True)

    # Parent WI
    parent_tag = soup.find("a", id="lnkParentWi")
    if parent_tag:
        metadata["parent_wi"] = parent_tag.get_text(strip=True)

    # Rapporteur
    rap_a = soup.find("a", id="lnkRapporteur")
    rap_lbl = soup.find("span", id="lblRapporteur")
    rap_parts = []
    if rap_a and rap_a.get_text(strip=True):
        rap_parts.append(rap_a.get_text(strip=True))
    if rap_lbl and rap_lbl.get_text(strip=True):
        rap_parts.append(rap_lbl.get_text(strip=True))
    metadata["rapporteur"] = " ".join(rap_parts)

    # Responsible Group(s)
    for td in soup.find_all("td"):
        if "Responsible group" in td.get_text():
            next_td = td.find_next_sibling("td")
            if next_td:
                raw_groups = [g.strip() for g in re.split(r"[,;/|\s]+", next_td.get_text()) if g.strip()]
                mapped_groups = []
                for g in raw_groups:
                    std = PORTAL_TO_WG_MAP.get(g.upper(), g.upper())
                    if std and std not in mapped_groups:
                        mapped_groups.append(std)
                metadata["responsible_groups"] = mapped_groups
            break

    # Remarks Table
    remarks_table = soup.find("table", id="wiRemarks_remarksGrid_ctl00") or soup.find("table", id=re.compile(r"wiRemarks.*Grid"))
    if remarks_table:
        tbody = remarks_table.find("tbody") or remarks_table
        for row in tbody.find_all("tr"):
            cols = row.find_all("td")
            if len(cols) >= 3:
                date_span = cols[0].find("span") or cols[0]
                remark_div = cols[2].find("div") or cols[2]
                d_text = date_span.get_text(strip=True)
                r_text = remark_div.get_text(strip=True)
                if d_text and r_text and "Creation date" not in d_text:
                    metadata["remarks"].append({"date": d_text, "text": r_text})

    # Ensure remarks are sorted from most recent to oldest
    metadata["remarks"].sort(key=lambda x: x["date"], reverse=True)

    if not metadata["name"] and not metadata["acronym"]:
        err_msg = f"Work Item #{code} returned no title or acronym. It may not exist on 3GPP Portal."
        if log_cb:
            log_cb(err_msg, logging.WARNING)
        return {"error": err_msg, "code": code}

    return metadata


class WorkItemsScraperThread(QThread):
    # Emits (current_completed, total_wgs, message_text)
    progress = pyqtSignal(int, int, str)
    # Emits (success_boolean, final_message_text)
    finished_sync = pyqtSignal(bool, str)

    def __init__(self, db_path, parent=None):
        super().__init__(parent)
        self.db_path = db_path

        # Mapping standard Working Group names to the 3GPP dynareport URL codes
        self.wgs = {
            "SA": "SP", "SA1": "S1", "SA2": "S2", "SA3": "S3", "SA4": "S4", "SA5": "S5", "SA6": "S6",
            "RAN": "RP", "RAN1": "R1", "RAN2": "R2", "RAN3": "R3", "RAN4": "R4", "RAN5": "R5", "RAN6": "R6",
            "CT": "CP", "CT1": "C1", "CT3": "C3", "CT4": "C4", "CT6": "C6"
        }

    def _report(self, current: int, total: int, message: str, level: int = logging.INFO):
        logger.log(level, message)
        self.progress.emit(current, total, message)

    def run(self):
        # Instantiate a local DB connection inside the thread
        db = WorkItemsDatabase(self.db_path)
        total_wgs = len(self.wgs)
        self._report(0, total_wgs, "🔄 Initializing Work Items parallel sync...")

        completed = 0

        # Use a ThreadPool to download up to 5 WG pages concurrently
        with concurrent.futures.ThreadPoolExecutor(max_workers=5) as executor:
            future_to_wg = {
                executor.submit(self._fetch_and_parse, wg_name, wg_code): wg_name
                for wg_name, wg_code in self.wgs.items()
            }

            for future in concurrent.futures.as_completed(future_to_wg):
                if self.isInterruptionRequested():
                    break
                wg_name = future_to_wg[future]
                completed += 1
                try:
                    items = future.result()
                    if items:
                        # Push the parsed items into the database atomically
                        db.upsert_work_items(wg_name, items)
                        msg = f"Synced {len(items)} WIs for {wg_name}."
                    else:
                        msg = f"No active WIs found for {wg_name}."

                    self._report(completed, total_wgs, msg)
                except Exception as e:
                    self._report(completed, total_wgs, f"❌ Error syncing {wg_name}: {e}", logging.ERROR)

        if self.isInterruptionRequested():
            self.finished_sync.emit(False, "Work Items sync cancelled.")
            return

        logger.info("✅ Successfully synced Work Items for all Working Groups.")
        self.finished_sync.emit(True, "Successfully synced Work Items for all Working Groups.")

    def _fetch_and_parse(self, wg_name: str, wg_code: str) -> list:
        if self.isInterruptionRequested():
            return []
        url = f"https://www.3gpp.org/dynareport?code=TSG-WG--{wg_code}--wis.htm"

        # Utilize the global session to inherit proxies and humanness settings
        session = NetworkSession.get_instance()
        NetworkSession.apply_humanness(session)

        response = session.get(url, timeout=30)
        response.raise_for_status()

        soup = BeautifulSoup(response.text, "html.parser")

        # Locate the specific table holding the WIs
        table = soup.find("table", class_="dsp-tsgwgxwis")
        parsed_items = []

        if not table:
            return parsed_items

        for row in table.find_all("tr"):
            cols = row.find_all("td")
            if len(cols) >= 3:
                a_tag = cols[0].find("a")
                if not a_tag:
                    continue

                # Column 1: WI Name
                name = a_tag.get_text(strip=True)

                # Extract WI Code from the href attribute
                href = a_tag.get("href", "")
                code_match = re.search(r'workitemId=(\d+)', href)
                if not code_match:
                    continue
                wi_code = code_match.group(1)

                # Column 2: Acronym | Column 3: Release
                acronym = cols[1].get_text(strip=True)
                release = cols[2].get_text(strip=True)

                parsed_items.append({
                    "code": wi_code,
                    "name": name,
                    "acronym": acronym,
                    "release": release
                })

        return parsed_items


class TargetedWIScraperThread(QThread):
    progress = pyqtSignal(int, int, str)
    finished_sync = pyqtSignal(bool, str)

    def __init__(self, db_path, target_wi_codes: list, parent=None):
        super().__init__(parent)
        self.db_path = db_path
        self.target_wi_codes = target_wi_codes

    def run(self):
        if not self.target_wi_codes:
            self.finished_sync.emit(False, "No Work Items selected for update.")
            return

        db = WorkItemsDatabase(self.db_path)
        total_targets = len(self.target_wi_codes)

        logger.info(f"TargetedWIScraperThread starting for {total_targets} WIs: {self.target_wi_codes}")
        self.progress.emit(0, total_targets, "Initializing targeted Work Item update...")

        completed = 0
        batch_metadata = []

        # Pool HTTP requests to prevent bottlenecks, capping at 10 simultaneous workers
        with concurrent.futures.ThreadPoolExecutor(max_workers=10) as executor:
            future_to_wi = {
                executor.submit(self._fetch_and_parse_details, wi_code): wi_code
                for wi_code in self.target_wi_codes
            }

            for future in concurrent.futures.as_completed(future_to_wi):
                if self.isInterruptionRequested():
                    break
                wi_code = future_to_wi[future]
                completed += 1
                try:
                    metadata = future.result()
                    if metadata:
                        # Append the WI Code so the database knows which row to update
                        metadata['code'] = wi_code
                        batch_metadata.append(metadata)
                        msg = f"Parsed metadata for WI {wi_code}."
                        logger.debug(f"Successfully scraped WI {wi_code}: {metadata}")
                    else:
                        msg = f"No metadata found for WI {wi_code}."
                        logger.warning(msg)

                    self.progress.emit(completed, total_targets, msg)
                except Exception as e:
                    error_msg = f"Error parsing WI {wi_code}: {str(e)}"
                    logger.error(error_msg, exc_info=True)
                    self.progress.emit(completed, total_targets, error_msg)

        if self.isInterruptionRequested():
            self.finished_sync.emit(False, "Targeted Work Item update cancelled.")
            return

        # Perform a single, high-speed atomic transaction for all updated records
        if batch_metadata:
            logger.info(f"Sending {len(batch_metadata)} parsed WIs to the database for update.")
            self.progress.emit(completed, total_targets, "Saving batch to database...")
            try:
                db.update_work_items_metadata(batch_metadata)
            except Exception as e:
                logger.error(f"Database transaction failed: {e}", exc_info=True)
                self.finished_sync.emit(False, f"Database transaction failed: {str(e)}")
                return
        else:
            logger.warning("batch_metadata is empty! Skipping database update.")

        logger.info(f"✅ Successfully processed {len(batch_metadata)} Work Items.")
        self.finished_sync.emit(True, f"Successfully processed {len(batch_metadata)} Work Items.")

    def _fetch_and_parse_details(self, wi_code: str) -> dict:
        if self.isInterruptionRequested():
            return {}
        data = fetch_wi_details_from_portal(wi_code)
        if "error" in data:
            logger.warning(f"Failed to fetch details for WI {wi_code}: {data['error']}")
            return {}
        return data