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

# Mapping Working Groups to their FTP and DynaReport endpoints
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


class MeetingsCrawlerThread(QThread):
    ui_log_msg = pyqtSignal(str, int)
    finished = pyqtSignal()
    finished_path = pyqtSignal(str)

    def __init__(self, db_path: Path):
        super().__init__()
        self.db = MeetingsDatabase(db_path)

    def extract_meeting_number(self, folder_name: str) -> str:
        """Extracts the meeting number, maintaining 'AH' or 'E' flags."""
        match = re.search(r'(AH|\d+[A-Z]?)', folder_name.upper())
        return match.group(1) if match else folder_name

    def process_ftp_folder(self, wg_name: str, ftp_base_url: str):
        try:
            html = NetworkSession.get_html(ftp_base_url)
            soup = BeautifulSoup(html, 'html.parser')

            for a_tag in soup.find_all('a', href=True):
                href = a_tag['href']
                if ".." in href or "?" in href:
                    continue

                folder_name = href.strip('/')
                absolute_url = urljoin(ftp_base_url, href)

                # Exclude obvious non-meeting folders
                if folder_name.lower() in ["docs", "inbox", "info", "specs", "drafts", "outgoing"]:
                    continue

                meeting_num = self.extract_meeting_number(folder_name)
                url_key = absolute_url.split('ftp/', 1)[-1] if 'ftp/' in absolute_url else absolute_url

                # Check for Docs directory
                docs_url = ""
                first_tdoc, last_tdoc = "", ""

                for doc_folder in ["Docs/", "docs/"]:
                    test_docs_url = urljoin(absolute_url, doc_folder)
                    try:
                        docs_html = NetworkSession.get_html(test_docs_url, timeout=5)
                        docs_url = test_docs_url

                        # Find all zip/doc files to identify first and last Tdocs
                        d_soup = BeautifulSoup(docs_html, 'html.parser')
                        tdoc_files = [a.text for a in d_soup.find_all('a', href=True) if
                                      a.text.endswith(('.zip', '.doc', '.docx', '.pdf'))]

                        if tdoc_files:
                            tdoc_files.sort()
                            first_tdoc = tdoc_files[0]
                            last_tdoc = tdoc_files[-1]
                        break  # Found the docs folder, stop checking variations
                    except Exception:
                        pass  # Folder doesn't exist or timed out, try the next casing

                self.db.insert_or_update_meeting_pass1(
                    wg_name, folder_name, meeting_num, url_key, docs_url, first_tdoc, last_tdoc
                )
        except Exception as e:
            self.ui_log_msg.emit(f"⚠️ FTP Fetch Error for {wg_name}: {e}", logging.WARNING)

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
                    sub_num = cols[2].get_text(strip=True)  # e.g., 'e' or 'AH'
                    dates_raw = cols[3].get_text(strip=True)
                    location = cols[4].get_text(strip=True)

                    if not meeting_num or meeting_name.lower() == "meeting":
                        continue

                    # Combine meeting num and sub letter (e.g., 149 + E = 149E)
                    full_meeting_num = f"{meeting_num}{sub_num}".strip().upper()

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
        try:
            self.ui_log_msg.emit("⏳ Pass 1: Fetching FTP Meeting directories...", logging.INFO)

            with ThreadPoolExecutor(max_workers=10) as executor:
                ftp_futures = {executor.submit(self.process_ftp_folder, wg, data["ftp"]): wg for wg, data in
                               MEETING_SOURCES.items()}
                for count, future in enumerate(as_completed(ftp_futures), 1):
                    wg = ftp_futures[future]
                    self.ui_log_msg.emit(f"⏳ FTP Scanned {wg} ({count}/{len(MEETING_SOURCES)})...", logging.INFO)

            self.ui_log_msg.emit("✅ Pass 1 Complete. Unblocking interface...", logging.INFO)
            self.finished_path.emit("MEETINGS_DB_PASS_ONE")

            self.ui_log_msg.emit("⏳ Pass 2: Updating meeting metadata from DynaReports...", logging.INFO)

            with ThreadPoolExecutor(max_workers=5) as executor:
                dyna_futures = {executor.submit(self.process_dynareport, wg, data["dyna"]): wg for wg, data in
                                MEETING_SOURCES.items()}
                for future in as_completed(dyna_futures):
                    pass  # Errors logged internally

            self.ui_log_msg.emit("✅ 3GPP Meetings Database Fully Updated!", logging.INFO)
            self.finished_path.emit("MEETINGS_DB_PASS_TWO")

        except Exception as e:
            self.ui_log_msg.emit(f"❌ Meetings Sync Failed: {str(e)}", logging.ERROR)
        finally:
            self.finished.emit()