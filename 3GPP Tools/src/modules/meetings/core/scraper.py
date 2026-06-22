# --- File: modules/meetings/core/scraper.py ---
import logging
import re
import time
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


class MeetingsCrawlerThread(QThread):
    ui_log_msg = pyqtSignal(str, int)
    finished = pyqtSignal()
    finished_path = pyqtSignal(str)

    def __init__(self, db_path: Path, target_meetings: list = None, sync_wg=True, sync_docs=True, sync_dyna=True):
        super().__init__()
        self.db = MeetingsDatabase(db_path)
        self.target_meetings = target_meetings or []
        self.sync_wg = sync_wg
        self.sync_docs = sync_docs
        self.sync_dyna = sync_dyna

        self.meeting_pattern = re.compile(r'^TSG[A-Z0-9]+_\d+', re.IGNORECASE)
        self.deny_list = {
            "cr_implementation", "tor", "tool_automation_6g", "specifications",
            "r2_tss_trs_early_versions", "outgoing_liaisons", "_doc_list_archive",
            "approved_reports", "docs", "latest_sa2_specs"
        }

    def is_meeting(self, folder_name: str) -> bool:
        if folder_name.lower() in self.deny_list: return False
        return bool(self.meeting_pattern.match(folder_name))

    def extract_meeting_number(self, folder_name: str) -> str:
        match = re.search(r'(?:^|_|-)(AH\d*|\d+(?:-?bis|-?e|-?a|-?b)?)(?:_|-|$)', folder_name, re.IGNORECASE)
        return match.group(1).replace('-', '').upper() if match else folder_name

    def fetch_wg_directories(self, wg_name: str, ftp_base_url: str) -> list:
        meeting_tasks = []
        try:
            self.ui_log_msg.emit(f"Parsing {ftp_base_url}", logging.INFO)
            soup = BeautifulSoup(NetworkSession.get_html(ftp_base_url), 'html.parser')
            for a_tag in soup.find_all('a', href=True):
                href = a_tag['href']
                if ".." in href or "?" in href: continue
                folder_name = href.strip('/').split('/')[-1]
                if folder_name in ["..", ".", ""] or not self.is_meeting(folder_name): continue

                absolute_url = urljoin(ftp_base_url, href)
                meeting_num = self.extract_meeting_number(folder_name)

                if self.target_meetings and not any(
                        t["wg"] == wg_name and t["meeting"] == meeting_num for t in self.target_meetings):
                    continue

                url_key = absolute_url.split('ftp/', 1)[-1] if 'ftp/' in absolute_url else absolute_url
                meeting_tasks.append({
                    "wg_name": wg_name, "folder_name": folder_name, "meeting_num": meeting_num,
                    "url_key": url_key, "absolute_url": absolute_url
                })
        except Exception as e:
            self.ui_log_msg.emit(f"⚠️ Directory Fetch Error for {wg_name}: {e}", logging.WARNING)
        return meeting_tasks

    def process_individual_meeting(self, task: dict) -> int:
        docs_url, first_tdoc, last_tdoc, tdoc_count = "", "", "", 0
        for doc_folder in ["Docs/", "docs/"]:
            test_docs_url = urljoin(task["absolute_url"], doc_folder)
            try:
                soup = BeautifulSoup(NetworkSession.get_html(test_docs_url, timeout=5), 'html.parser')
                docs_url = test_docs_url
                tdocs = sorted([a.text for a in soup.find_all('a', href=True) if
                                a.text.endswith(('.zip', '.doc', '.docx', '.pdf'))])
                if tdocs:
                    first_tdoc, last_tdoc, tdoc_count = tdocs[0], tdocs[-1], len(tdocs)
                break
            except Exception:
                pass

        self.db.update_meeting_docs(task["url_key"], docs_url, first_tdoc, last_tdoc)
        return tdoc_count

    def process_dynareport(self, wg_name: str, dyna_url: str):
        try:
            soup = BeautifulSoup(NetworkSession.get_html(dyna_url), 'html.parser')
            for row in soup.find_all('tr'):
                cols = row.find_all(['td', 'th'])
                if len(cols) >= 5:
                    m_name = cols[0].get_text(strip=True)
                    m_num = cols[1].get_text(strip=True)
                    if not m_num or m_name.lower() == "meeting": continue

                    full_num = f"{m_num}{cols[2].get_text(strip=True)}".replace('-', '').strip().upper()
                    if self.target_meetings and not any(
                            t["wg"] == wg_name and t["meeting"] == full_num for t in self.target_meetings):
                        continue

                    dates = cols[3].get_text(strip=True).split("...")
                    self.db.update_meeting_metadata_pass2(
                        wg_name, full_num, m_name, cols[4].get_text(strip=True),
                        dates[0].strip(), dates[1].strip() if len(dates) > 1 else ""
                    )
        except Exception as e:
            self.ui_log_msg.emit(f"⚠️ DynaReport Error for {wg_name}: {e}", logging.WARNING)

    def run(self):
        start_time = time.time()
        try:
            all_tasks = []

            # --- PHASE 1: DIRECTORIES ---
            if self.sync_wg:
                self.ui_log_msg.emit("⏳ [Phase 1/3] Mapping WG directories...", logging.INFO)
                mapped = set()
                with ThreadPoolExecutor(max_workers=10) as executor:
                    futures = {executor.submit(self.fetch_wg_directories, wg, data["ftp"]): wg for wg, data in
                               MEETING_SOURCES.items()}
                    for future in as_completed(futures):
                        if res := future.result():
                            all_tasks.extend(res)
                            for t in res:
                                mapped.add(f"{t['wg_name']}:{t['meeting_num']}")
                                self.db.insert_meeting_basic(t["wg_name"], t["folder_name"], t["meeting_num"],
                                                             t["url_key"])

                if self.target_meetings:
                    for t in self.target_meetings:
                        if f"{t['wg']}:{t['meeting']}" not in mapped:
                            self.ui_log_msg.emit(f"⚠️ Target {t['wg']}:{t['meeting']} not found on FTP!",
                                                 logging.WARNING)
            else:
                self.ui_log_msg.emit("⏭️ [Phase 1/3] Skipping Directory Mapping (loading DB)...", logging.INFO)
                for m in self.db.search_meetings():
                    if self.target_meetings and not any(
                            t["wg"] == m["wg_name"] and t["meeting"] == m["meeting_number"] for t in
                            self.target_meetings):
                        continue
                    if m.get('url_key'):
                        all_tasks.append({
                            "wg_name": m["wg_name"], "folder_name": m.get("folder_name", m["meeting_number"]),
                            "meeting_num": m["meeting_number"], "url_key": m["url_key"],
                            "absolute_url": f"https://www.3gpp.org/ftp/{m['url_key']}"
                        })

            # --- PHASE 2: DOCS ---
            if self.sync_docs:
                if not all_tasks:
                    self.ui_log_msg.emit("⚠️ No meetings available to scan for Docs.", logging.WARNING)
                else:
                    self.ui_log_msg.emit(f"⏳ [Phase 2/3] Deep scraping Docs for {len(all_tasks)} meetings...",
                                         logging.INFO)
                    completed, tdocs_found = 0, 0

                    with ThreadPoolExecutor(max_workers=5) as executor:
                        future_to_task = {executor.submit(self.process_individual_meeting, task): task for task in
                                          all_tasks}
                        for future in as_completed(future_to_task):
                            task = future_to_task[future]
                            completed += 1
                            try:
                                tdocs_found += future.result()
                            except Exception as e:
                                self.ui_log_msg.emit(f"❌ Error scraping {task['folder_name']}: {e}", logging.ERROR)

                            if completed % 10 == 0 or completed == len(all_tasks):
                                self.ui_log_msg.emit(
                                    f"⏳ Scanned {completed}/{len(all_tasks)} Docs folders | TDocs: {tdocs_found}",
                                    logging.INFO)
                    self.ui_log_msg.emit(f"✅ Pass 2 Complete. Indexed {tdocs_found} total TDocs.", logging.INFO)
                    self.finished_path.emit("MEETINGS_DB_PASS_ONE")
            else:
                self.ui_log_msg.emit("⏭️ [Phase 2/3] Skipping Docs folder deep scrape...", logging.INFO)

            # --- PHASE 3: METADATA ---
            if self.sync_dyna:
                self.ui_log_msg.emit("⏳ [Phase 3/3] Updating metadata from DynaReports...", logging.INFO)
                wgs_to_fetch = {t["wg"] for t in
                                self.target_meetings} if self.target_meetings else MEETING_SOURCES.keys()

                with ThreadPoolExecutor(max_workers=5) as executor:
                    dyna_futures = {executor.submit(self.process_dynareport, wg, MEETING_SOURCES[wg]["dyna"]): wg for wg
                                    in wgs_to_fetch}
                    for future in as_completed(dyna_futures): future.result()
                self.ui_log_msg.emit("✅ Pass 3 Complete.", logging.INFO)
            else:
                self.ui_log_msg.emit("⏭️ [Phase 3/3] Skipping DynaReports metadata update...", logging.INFO)

            self.ui_log_msg.emit(f"✅ 3GPP Sync Fully Complete in {time.time() - start_time:.1f}s!", logging.INFO)
            self.finished_path.emit("MEETINGS_DB_PASS_TWO")

        except Exception as e:
            self.ui_log_msg.emit(f"❌ Critical Failure: {str(e)}", logging.ERROR)
        finally:
            self.finished.emit()