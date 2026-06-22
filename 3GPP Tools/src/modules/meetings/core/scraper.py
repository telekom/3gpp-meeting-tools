# --- File: modules/meetings/core/scraper.py ---
import logging
import re
from pathlib import Path
from urllib.parse import urljoin
from concurrent.futures import ThreadPoolExecutor, as_completed
from bs4 import BeautifulSoup
from PyQt5.QtCore import QThread, pyqtSignal

from core.network.session import NetworkSession
from modules.meetings.core.meetings_db import MeetingsDatabase

MEETING_SOURCES = {
    "RAN": {"ftp": "https://www.3gpp.org/ftp/tsg_ran/TSG_RAN/",
            "dyna": "https://www.3gpp.org/dynareport?code=Meetings-RP.htm"},
    "RAN1": {"ftp": "https://www.3gpp.org/ftp/tsg_ran/WG1_RL1/",
             "dyna": "https://www.3gpp.org/dynareport?code=Meetings-R1.htm"},
    "RAN2": {"ftp": "https://www.3gpp.org/ftp/tsg_ran/WG2_RL2/",
             "dyna": "https://www.3gpp.org/dynareport?code=Meetings-R2.htm"},
    "RAN3": {"ftp": "https://www.3gpp.org/ftp/tsg_ran/WG3_Iu/",
             "dyna": "https://www.3gpp.org/dynareport?code=Meetings-R3.htm"},
    "RAN4": {"ftp": "https://www.3gpp.org/ftp/tsg_ran/WG4_Radio/",
             "dyna": "https://www.3gpp.org/dynareport?code=Meetings-R4.htm"},
    "SA": {"ftp": "https://www.3gpp.org/ftp/tsg_sa/TSG_SA/",
           "dyna": "https://www.3gpp.org/dynareport?code=Meetings-SP.htm"},
    "SA1": {"ftp": "https://www.3gpp.org/ftp/tsg_sa/WG1_Serv/",
            "dyna": "https://www.3gpp.org/dynareport?code=Meetings-S1.htm"},
    "SA2": {"ftp": "https://www.3gpp.org/ftp/tsg_sa/WG3_Security/",
            "dyna": "https://www.3gpp.org/dynareport?code=Meetings-S2.htm"},
    "SA3": {"ftp": "https://www.3gpp.org/ftp/tsg_sa/WG2_Arch/",
            "dyna": "https://www.3gpp.org/dynareport?code=Meetings-S3.htm"},
    "SA4": {"ftp": "https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/",
            "dyna": "https://www.3gpp.org/dynareport?code=Meetings-S4.htm"},
    "SA5": {"ftp": "https://www.3gpp.org/ftp/tsg_sa/WG5_TM/",
            "dyna": "https://www.3gpp.org/dynareport?code=Meetings-S5.htm"},
    "SA6": {"ftp": "https://www.3gpp.org/ftp/tsg_sa/WG6_MissionCritical/",
            "dyna": "https://www.3gpp.org/dynareport?code=Meetings-S6.htm"},
    "CT": {"ftp": "https://www.3gpp.org/ftp/tsg_ct/TSG_CT/",
           "dyna": "https://www.3gpp.org/dynareport?code=Meetings-CP.htm"},
    "CT1": {"ftp": "https://www.3gpp.org/ftp/tsg_ct/WG1_mm-cc-sm_ex-CN1/",
            "dyna": "https://www.3gpp.org/dynareport?code=Meetings-C1.htm"},
    "CT2": {"ftp": "https://www.3gpp.org/ftp/tsg_ct/WG2_capability_ex-T2/",
            "dyna": "https://www.3gpp.org/dynareport?code=Meetings-C2.htm"},
    "CT3": {"ftp": "https://www.3gpp.org/ftp/tsg_ct/WG3_interworking_ex-CN3/",
            "dyna": "https://www.3gpp.org/dynareport?code=Meetings-C3.htm"},
    "CT4": {"ftp": "https://www.3gpp.org/ftp/tsg_ct/WG4_protocollars_ex-CN4/",
            "dyna": "https://www.3gpp.org/dynareport?code=Meetings-C4.htm"},
    "CT5": {"ftp": "https://www.3gpp.org/ftp/tsg_ct/WG5_osa_ex-CN5/",
            "dyna": "https://www.3gpp.org/dynareport?code=Meetings-C5.htm"},
    "CT6": {"ftp": "https://www.3gpp.org/ftp/tsg_ct/WG6_Smartcard_Ex-T3/",
            "dyna": "https://www.3gpp.org/dynareport?code=Meetings-C6.htm"},
}

# 1. Matches patterns like: TSGR2_05, TSGS2_109bis-e, TSGS2_154AHE_Electronic_2023-01
meeting_pattern = re.compile(r'^TSG[A-Z0-9]+_\d+', re.IGNORECASE)


class MeetingsCrawlerThread(QThread):
    ui_log_msg = pyqtSignal(str, int)
    finished = pyqtSignal()
    finished_path = pyqtSignal(str)

    def __init__(self, db_path: Path, target_meetings: list = None):
        super().__init__()
        self.db = MeetingsDatabase(db_path)
        self.target_meetings = target_meetings or []

    def is_meeting(self, folder_name: str) -> bool:
        # Check if it matches the TSG pattern
        if meeting_pattern.match(folder_name):
            return True

        return False

    def extract_meeting_number(self, folder_name: str) -> str:
        """Extracts the meeting number, ignoring 3G prefixes and normalizing hyphens."""
        match = re.search(r'(?:^|_|-)(AH\d*|\d+(?:-?bis|-?e|-?a|-?b)?)(?:_|-|$)', folder_name, re.IGNORECASE)
        if match:
            return match.group(1).replace('-', '').upper()
        return folder_name

    def fetch_wg_directories(self, wg_name: str, ftp_base_url: str) -> list:
        """Grabs the URLs of all meeting folders for a specific Working Group."""
        meeting_tasks = []
        try:
            html = NetworkSession.get_html(ftp_base_url)
            self.ui_log_msg.emit(f"Parsing {ftp_base_url}", logging.INFO)
            soup = BeautifulSoup(html, 'html.parser')

            for a_tag in soup.find_all('a', href=True):
                href = a_tag['href']
                if ".." in href or "?" in href:
                    continue

                folder_name = href.strip('/')
                folder_name = folder_name.split('/')[-1]
                absolute_url = urljoin(ftp_base_url, href)

                folder_lowercase = folder_name.lower()
                if folder_lowercase in ["docs", "inbox", "info", "specs", "drafts", "outgoing"]:
                    continue

                if not self.is_meeting(folder_lowercase):
                    self.ui_log_msg.emit(f"Folder {folder_lowercase} not a meeting. Skipping", logging.INFO)
                    continue

                meeting_num = self.extract_meeting_number(folder_name)

                if self.target_meetings:
                    is_target = any(t["wg"] == wg_name and t["meeting"] == meeting_num for t in self.target_meetings)
                    if not is_target:
                        continue

                url_key = absolute_url.split('ftp/', 1)[-1] if 'ftp/' in absolute_url else absolute_url
                meeting_tasks.append({
                    "wg_name": wg_name,
                    "folder_name": folder_name,
                    "meeting_num": meeting_num,
                    "url_key": url_key,
                    "absolute_url": absolute_url
                })
        except Exception as e:
            self.ui_log_msg.emit(f"⚠️ Directory Fetch Error for {wg_name}: {e}", logging.WARNING)

        return meeting_tasks

    def process_individual_meeting(self, task: dict) -> int:
        """Checks the Docs/ folder for a single meeting. Returns the number of TDocs found."""
        absolute_url = task["absolute_url"]
        docs_url = ""
        first_tdoc, last_tdoc = "", ""
        tdoc_count = 0

        for doc_folder in ["Docs/", "docs/"]:
            test_docs_url = urljoin(absolute_url, doc_folder)
            try:
                docs_html = NetworkSession.get_html(test_docs_url, timeout=5)
                docs_url = test_docs_url

                d_soup = BeautifulSoup(docs_html, 'html.parser')
                tdoc_files = [a.text for a in d_soup.find_all('a', href=True) if
                              a.text.endswith(('.zip', '.doc', '.docx', '.pdf'))]

                if tdoc_files:
                    tdoc_files.sort()
                    first_tdoc = tdoc_files[0]
                    last_tdoc = tdoc_files[-1]
                    tdoc_count = len(tdoc_files)  # <--- COUNT EXTRACTED HERE
                break
            except Exception:
                pass

        self.db.insert_or_update_meeting_pass1(
            task["wg_name"], task["folder_name"], task["meeting_num"],
            task["url_key"], docs_url, first_tdoc, last_tdoc
        )
        return tdoc_count

    def process_dynareport(self, wg_name: str, dyna_url: str):
        try:
            html = NetworkSession.get_html(dyna_url)
            soup = BeautifulSoup(html, 'html.parser')

            rows = soup.find_all('tr')
            for row in rows:
                cols = row.find_all(['td', 'th'])
                if len(cols) >= 5:
                    meeting_name = cols[0].get_text(strip=True)
                    meeting_num = cols[1].get_text(strip=True)
                    sub_num = cols[2].get_text(strip=True)
                    dates_raw = cols[3].get_text(strip=True)
                    location = cols[4].get_text(strip=True)

                    if not meeting_num or meeting_name.lower() == "meeting":
                        continue

                    full_meeting_num = f"{meeting_num}{sub_num}".replace('-', '').strip().upper()

                    if self.target_meetings:
                        is_target = any(
                            t["wg"] == wg_name and t["meeting"] == full_meeting_num for t in self.target_meetings)
                        if not is_target:
                            continue

                    start_date, end_date = "", ""
                    if "..." in dates_raw:
                        parts = dates_raw.split("...")
                        start_date = parts[0].strip()
                        end_date = parts[1].strip() if len(parts) > 1 else ""

                    self.db.update_meeting_metadata_pass2(
                        wg_name, full_meeting_num, meeting_name, location, start_date, end_date
                    )
        except Exception as e:
            self.ui_log_msg.emit(f"⚠️ DynaReport Error for {wg_name}: {e}", logging.WARNING)

    def run(self):
        import time
        start_time = time.time()

        try:
            # --- 1. Mapping Phase ---
            self.ui_log_msg.emit("⏳ Mapping FTP directories...", logging.INFO)
            all_meeting_tasks = []

            with ThreadPoolExecutor(max_workers=10) as executor:
                futures = {executor.submit(self.fetch_wg_directories, wg, data["ftp"]): wg for wg, data in
                           MEETING_SOURCES.items()}
                for future in as_completed(futures):
                    if future.result(): all_meeting_tasks.extend(future.result())

            total_meetings = len(all_meeting_tasks)
            if total_meetings == 0:
                self.ui_log_msg.emit("⚠️ No meeting folders found. Check network connection.", logging.WARNING)
                return

            self.ui_log_msg.emit(f"📥 Found {total_meetings} meetings. Initiating deep scrape...", logging.INFO)

            # --- 2. Deep Scrape Phase (with enhanced logging) ---
            completed = 0
            total_tdocs_found = 0

            # Using a smaller chunk size for logging progress
            with ThreadPoolExecutor(max_workers=20) as executor:
                # Map futures to tasks to track which one might be hanging
                future_to_task = {executor.submit(self.process_individual_meeting, task): task for task in
                                  all_meeting_tasks}

                for future in as_completed(future_to_task):
                    task = future_to_task[future]
                    completed += 1

                    try:
                        files_in_this_meeting = future.result()
                        total_tdocs_found += files_in_this_meeting
                    except Exception as e:
                        # LOG THE FOLDER THAT FAILED
                        self.ui_log_msg.emit(f"❌ Failed to scrape {task['folder_name']}: {e}", logging.ERROR)

                    # LOGGING HEARTBEAT
                    if completed % 20 == 0 or completed == total_meetings:
                        elapsed = time.time() - start_time
                        rate = completed / elapsed if elapsed > 0 else 0
                        self.ui_log_msg.emit(
                            f"⏳ Scanned {completed}/{total_meetings} meetings "
                            f"| TDocs found: {total_tdocs_found} | "
                            f"Speed: {rate:.1f} mtg/sec",
                            logging.INFO
                        )

            self.ui_log_msg.emit(f"✅ Pass 1 Complete. Indexed {total_tdocs_found} TDocs in {elapsed:.1f}s",
                                 logging.INFO)
            self.finished_path.emit("MEETINGS_DB_PASS_ONE")

            # --- 3. Pass 2 (DynaReports) ---
            self.ui_log_msg.emit("⏳ Pass 2: Updating metadata from DynaReports...", logging.INFO)
            wgs_to_fetch = {t["wg"] for t in self.target_meetings} if self.target_meetings else MEETING_SOURCES.keys()

            with ThreadPoolExecutor(max_workers=5) as executor:
                dyna_futures = [executor.submit(self.process_dynareport, wg, MEETING_SOURCES[wg]["dyna"]) for wg in
                                wgs_to_fetch]
                for future in as_completed(dyna_futures):
                    future.result()  # Trigger exceptions if any

            self.ui_log_msg.emit("✅ 3GPP Database Update Fully Complete!", logging.INFO)
            self.finished_path.emit("MEETINGS_DB_PASS_TWO")

        except Exception as e:
            self.ui_log_msg.emit(f"❌ Critical Failure: {str(e)}", logging.ERROR)
        finally:
            self.finished.emit()