import re
import concurrent.futures
from bs4 import BeautifulSoup
from PyQt5.QtCore import QThread, pyqtSignal

from core.network.session import NetworkSession
from modules.work_items.core.wi_database import WorkItemsDatabase


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

    def run(self):
        # Instantiate a local DB connection inside the thread
        db = WorkItemsDatabase(self.db_path)
        total_wgs = len(self.wgs)
        self.progress.emit(0, total_wgs, "Initializing Work Items parallel sync...")

        completed = 0

        # Use a ThreadPool to download up to 5 WG pages concurrently
        with concurrent.futures.ThreadPoolExecutor(max_workers=5) as executor:
            future_to_wg = {
                executor.submit(self._fetch_and_parse, wg_name, wg_code): wg_name
                for wg_name, wg_code in self.wgs.items()
            }

            for future in concurrent.futures.as_completed(future_to_wg):
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

                    self.progress.emit(completed, total_wgs, msg)
                except Exception as e:
                    self.progress.emit(completed, total_wgs, f"Error syncing {wg_name}: {str(e)}")

        self.finished_sync.emit(True, "Successfully synced Work Items for all Working Groups.")

    def _fetch_and_parse(self, wg_name: str, wg_code: str) -> list:
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