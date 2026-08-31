# --- File: modules/meetings/core/scraper.py ---
import logging
import re
import time
from pathlib import Path
from urllib.parse import urljoin, unquote
from concurrent.futures import ThreadPoolExecutor, as_completed
from bs4 import BeautifulSoup
from PyQt5.QtCore import QThread, pyqtSignal

from core.network.session import NetworkSession
from modules.meetings.core.meetings_db import MeetingsDatabase

MEETING_SOURCES = {
    "RAN": {"ftp": ["https://www.3gpp.org/ftp/tsg_ran/TSG_RAN/", "https://www.3gpp.org/ftp/tsg_ran/TSG_RAN/TSGR_AHs/"],
            "dyna": "https://www.3gpp.org/dynareport?code=Meetings-RP.htm"},
    "RAN1": {"ftp": ["https://www.3gpp.org/ftp/tsg_ran/WG1_RL1/", "https://www.3gpp.org/ftp/tsg_ran/WG1_RL1/TSGR1_AH/"],
             "dyna": "https://www.3gpp.org/dynareport?code=Meetings-R1.htm"},
    "RAN2": {
        "ftp": ["https://www.3gpp.org/ftp/tsg_ran/WG2_RL2/", "https://www.3gpp.org/ftp/tsg_ran/WG2_RL2/TSGR2_AHs/"],
        "dyna": "https://www.3gpp.org/dynareport?code=Meetings-R2.htm"},
    "RAN3": {"ftp": ["https://www.3gpp.org/ftp/tsg_ran/WG3_Iu/", "https://www.3gpp.org/ftp/tsg_ran/WG3_Iu/TSGR3_AHGs/"],
             "dyna": "https://www.3gpp.org/dynareport?code=Meetings-R3.htm"},
    "RAN4": {
        "ftp": ["https://www.3gpp.org/ftp/tsg_ran/WG4_Radio/", "https://www.3gpp.org/ftp/tsg_ran/WG4_Radio/TSGR4_AHs/"],
        "dyna": "https://www.3gpp.org/dynareport?code=Meetings-R4.htm"},
    "SA": {"ftp": ["https://www.3gpp.org/ftp/tsg_sa/TSG_SA/"],
           "dyna": "https://www.3gpp.org/dynareport?code=Meetings-SP.htm"},
    "SA1": {"ftp": ["https://www.3gpp.org/ftp/tsg_sa/WG1_Serv/"],
            "dyna": "https://www.3gpp.org/dynareport?code=Meetings-S1.htm"},
    "SA2": {"ftp": ["https://www.3gpp.org/ftp/tsg_sa/WG2_Arch/"],
            "dyna": "https://www.3gpp.org/dynareport?code=Meetings-S2.htm"},
    "SA3": {"ftp": ["https://www.3gpp.org/ftp/tsg_sa/WG3_Security/"],
            "dyna": "https://www.3gpp.org/dynareport?code=Meetings-S3.htm"},
    "SA4": {"ftp": ["https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/"],
            "dyna": "https://www.3gpp.org/dynareport?code=Meetings-S4.htm"},
    "SA5": {"ftp": ["https://www.3gpp.org/ftp/tsg_sa/WG5_TM/"],
            "dyna": "https://www.3gpp.org/dynareport?code=Meetings-S5.htm"},
    "SA6": {"ftp": ["https://www.3gpp.org/ftp/tsg_sa/WG6_MissionCritical/"],
            "dyna": "https://www.3gpp.org/dynareport?code=Meetings-S6.htm"},
    "CT": {"ftp": ["https://www.3gpp.org/ftp/tsg_ct/TSG_CT/"],
           "dyna": "https://www.3gpp.org/dynareport?code=Meetings-CP.htm"},
    "CT1": {"ftp": ["https://www.3gpp.org/ftp/tsg_ct/WG1_mm-cc-sm_ex-CN1/"],
            "dyna": "https://www.3gpp.org/dynareport?code=Meetings-C1.htm"},
    "CT2": {"ftp": ["https://www.3gpp.org/ftp/tsg_ct/WG2_capability_ex-T2/"],
            "dyna": "https://www.3gpp.org/dynareport?code=Meetings-C2.htm"},
    "CT3": {"ftp": ["https://www.3gpp.org/ftp/tsg_ct/WG3_interworking_ex-CN3/"],
            "dyna": "https://www.3gpp.org/dynareport?code=Meetings-C3.htm"},
    "CT4": {"ftp": ["https://www.3gpp.org/ftp/tsg_ct/WG4_protocollars_ex-CN4/"],
            "dyna": "https://www.3gpp.org/dynareport?code=Meetings-C4.htm"},
    "CT5": {"ftp": ["https://www.3gpp.org/ftp/tsg_ct/WG5_osa_ex-CN5/"],
            "dyna": "https://www.3gpp.org/dynareport?code=Meetings-C5.htm"},
    "CT6": {"ftp": ["https://www.3gpp.org/ftp/tsg_ct/WG6_Smartcard_Ex-T3/"],
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

        self.meeting_pattern = re.compile(r'^[A-Z0-9]+_(?:\d+|AH)', re.IGNORECASE)
        self.num_pattern = re.compile(r'(?:^|_)(AH\d*|\d+(?:[a-z]+)?(?:-?AH)?(?:-?e)?)(?:_|-|$)', re.IGNORECASE)
        self.href_pattern = re.compile(r'href=["\']([^"\'>]+)["\']', re.IGNORECASE)
        self.tdoc_pattern = re.compile(r'>\s*([^<]+\.(?:zip|doc|docx|pdf))\s*</a>', re.IGNORECASE)

        self.deny_list = {
            "cr_implementation", "tor", "tool_automation_6g", "specifications",
            "r2_tss_trs_early_versions", "outgoing_liaisons", "_doc_list_archive",
            "approved_reports", "docs", "latest_sa2_specs"
        }

    def is_meeting(self, folder_name: str) -> bool:
        if folder_name.lower() in self.deny_list: return False
        if re.search(r'\.[a-z0-9]{2,4}$', folder_name, re.IGNORECASE): return False
        return bool(self.meeting_pattern.match(folder_name))

    def extract_meeting_number(self, folder_name: str) -> str:
        match = self.num_pattern.search(folder_name)
        if match:
            raw = match.group(1).replace('-', '').upper()
            return re.sub(r'(?<!\d)0+(\d)', r'\1', raw)
        return folder_name

    def fetch_wg_directories(self, wg_name: str, ftp_base_url: str, is_ah_folder: bool = False) -> list:
        meeting_tasks = []
        try:
            self.ui_log_msg.emit(f"🌐 Parsing {ftp_base_url}", logging.INFO)
            html = NetworkSession.get_html(ftp_base_url)

            hrefs = self.href_pattern.findall(html)
            if not hrefs:
                self.ui_log_msg.emit(f"⚠️ [Debug] NO links found for {ftp_base_url.split('/')[-2]}.", logging.WARNING)
                return meeting_tasks

            for href in hrefs:
                if ".." in href or "?" in href: continue
                folder_name = href.strip('/').split('/')[-1]
                if folder_name in ["..", ".", ""]: continue

                if folder_name.lower() in self.deny_list: continue
                if re.search(r'\.[a-z0-9]{2,4}$', folder_name, re.IGNORECASE): continue

                if not is_ah_folder and not self.is_meeting(folder_name): continue

                absolute_url = urljoin(ftp_base_url, href)
                meeting_num = self.extract_meeting_number(folder_name)

                if self.target_meetings and not any(
                        t["wg"] == wg_name and t["meeting"] == meeting_num for t in self.target_meetings):
                    continue

                url_key = absolute_url.split('ftp/', 1)[-1] if 'ftp/' in absolute_url else absolute_url
                docs_url = urljoin(absolute_url, "Docs/") if absolute_url.endswith('/') else f"{absolute_url}/Docs/"

                meeting_tasks.append({
                    "wg_name": wg_name, "folder_name": folder_name, "meeting_num": meeting_num,
                    "url_key": url_key, "absolute_url": absolute_url,
                    "is_ad_hoc": is_ah_folder,
                    "docs_url": docs_url
                })

            self.ui_log_msg.emit(
                f"✅ {wg_name}: Found {len(meeting_tasks)} meeting folders in {ftp_base_url.split('/')[-2]}.",
                logging.INFO)

        except Exception as e:
            self.ui_log_msg.emit(f"⚠️ Directory Fetch Error for {ftp_base_url}: {e}", logging.WARNING)
        return meeting_tasks

    def process_individual_meeting(self, task: dict) -> tuple:
        base_docs_url = task.get("docs_url", urljoin(task["absolute_url"], "Docs/"))
        first_tdoc, first_pfx, first_num = "", "", 0
        last_tdoc, last_pfx, last_num = "", "", 0
        tdoc_count = 0
        final_docs_url = base_docs_url

        def fetch_and_parse(url):
            html = NetworkSession.get_html(url, timeout=5)
            tdocs_raw = self.tdoc_pattern.findall(html)
            parsed = []
            for f in tdocs_raw:
                clean_name = re.sub(r'\.[a-zA-Z0-9]{2,4}$', '', f.strip())
                match = re.match(r'^([A-Za-z0-9]+)-?(\d+)', clean_name)
                if match:
                    parsed.append({"clean": clean_name, "prefix": match.group(1).upper(), "num": int(match.group(2))})
            if parsed:
                parsed.sort(key=lambda x: (x["num"], x["clean"]))
            return parsed

        try:
            parsed_list = fetch_and_parse(base_docs_url)
            if not parsed_list and "Docs/" in base_docs_url:
                fallback_url = base_docs_url.replace("Docs/", "docs/")
                parsed_list = fetch_and_parse(fallback_url)
                if parsed_list: final_docs_url = fallback_url
        except Exception:
            parsed_list = []

        if parsed_list:
            tdoc_count = len(parsed_list)
            first_tdoc, first_pfx, first_num = parsed_list[0]["clean"], parsed_list[0]["prefix"], parsed_list[0]["num"]
            last_tdoc, last_pfx, last_num = parsed_list[-1]["clean"], parsed_list[-1]["prefix"], parsed_list[-1]["num"]

        docs_data = (final_docs_url, first_tdoc, first_pfx, first_num, last_tdoc, last_pfx, last_num, task["url_key"])
        return docs_data, tdoc_count

    def process_dynareport(self, wg_name: str, dyna_url: str) -> list:
        results = []
        try:
            html = NetworkSession.get_html(dyna_url)
            soup = BeautifulSoup(html, 'html.parser')

            for row in soup.find_all('tr'):
                cols = row.find_all(['td', 'th'])
                if len(cols) < 2:
                    continue

                col_texts = []
                for td in cols:
                    text = td.get_text(separator=" ", strip=True)
                    text = re.sub(r'[\u2010-\u2015\u2212]', '-', text)
                    col_texts.append(text)

                m_name = col_texts[0].strip()
                if not m_name or m_name.lower() == 'meeting':
                    continue

                mtg_id = ""
                url_key = ""
                for a in row.find_all('a', href=True):
                    href = a['href'].replace('\\', '/')
                    match_id = re.search(r'MtgId=(\d+)', href, re.IGNORECASE)
                    if match_id: mtg_id = match_id.group(1)

                    if '/ftp/' in href.lower() and 'tsg_' in href.lower():
                        parts = re.split(r'/ftp/', href, flags=re.IGNORECASE)
                        if len(parts) > 1:
                            path_after_ftp = unquote(parts[1].split('?')[0].strip('/'))
                            if not url_key or len(path_after_ftp) < len(url_key):
                                url_key = path_after_ftp

                start_d, end_d, town = "", "", ""
                all_dates = []
                date_idx = -1

                for i in range(1, len(col_texts)):
                    found_dates = re.findall(r'(?:19|20)\d{2}[-/]\d{2}[-/]\d{2}', col_texts[i])
                    if found_dates:
                        found_dates = [d.replace('/', '-') for d in found_dates]
                        all_dates.extend(found_dates)
                        if date_idx == -1:
                            date_idx = i

                if all_dates:
                    start_d = min(all_dates)
                    end_d = max(all_dates)

                if date_idx > 0:
                    candidate = col_texts[date_idx - 1].strip()
                    if candidate and not re.search(r'3GPP|\bRAN[1-6]?\b|\bSA[1-6]?\b|\bCT[1-6]?\b', candidate,
                                                   re.IGNORECASE):
                        town = candidate

                full_num = ""
                search_text = m_name + " " + (col_texts[1] if len(col_texts) > 1 else "")
                match_explicit = re.search(r'#(\d+[a-z0-9\-]*)', search_text, re.IGNORECASE)

                if match_explicit:
                    full_num = match_explicit.group(1).replace('-', '').strip().upper()
                else:
                    tokens = m_name.split()
                    found_token = ""
                    for token in reversed(tokens):
                        if re.search(r'\d|AH', token, re.IGNORECASE):
                            found_token = token
                            break

                    if not found_token and len(col_texts) > 1:
                        if not re.search(r'(?:19|20)\d{2}', col_texts[1]):
                            found_token = col_texts[1]

                    if found_token:
                        clean_token = re.sub(r'^(?:3GPP)?(?:R|S|C|RAN|SA|CT|SP|RP|CP)[P0-6]?\s*-?', '', found_token,
                                             flags=re.IGNORECASE)

                        suffix = ""
                        if date_idx >= 4 and len(col_texts) > 2:
                            potential_suffix = col_texts[2].strip()
                            if len(potential_suffix) <= 6 and not re.search(r'(?:19|20)\d{2}', potential_suffix):
                                suffix = potential_suffix

                        full_num = f"{clean_token}{suffix}".replace('-', '').strip().upper()

                full_num = re.sub(r'(?<!\d)0+(\d)', r'\1', full_num)

                if self.target_meetings and not any(
                        t["wg"] == wg_name and t["meeting"] == full_num for t in self.target_meetings):
                    continue

                new_m_num = ""
                if wg_name.startswith("RAN"):
                    is_adhoc = False
                    if re.search(r'(?:AH|Ad\s*Hoc|Workshop|Release|Evolution)', m_name, re.IGNORECASE):
                        is_adhoc = True
                    elif re.search(r'^R[P1-4]-\d', m_name, re.IGNORECASE):
                        is_adhoc = True
                    elif "AH" in url_key.upper() or "AH" in full_num:
                        is_adhoc = True

                    if is_adhoc:
                        new_m_num = m_name

                results.append((wg_name, full_num, url_key, mtg_id, m_name, town, start_d, end_d, new_m_num))

        except Exception as e:
            self.ui_log_msg.emit(f"⚠️ DynaReport Error for {wg_name}: {e}", logging.WARNING)

        return results

    def run(self):
        start_time = time.time()
        try:
            all_tasks = []

            if self.sync_wg:
                self.ui_log_msg.emit("⏳ [Phase 1/3] Mapping WG directories...", logging.INFO)
                mapped = set()

                with ThreadPoolExecutor(max_workers=15) as executor:
                    futures = {}
                    for wg_name, source_info in MEETING_SOURCES.items():
                        urls = source_info["ftp"] if isinstance(source_info["ftp"], list) else [source_info["ftp"]]
                        for url in urls:
                            is_ah = "AH" in url
                            future = executor.submit(self.fetch_wg_directories, wg_name, url, is_ah)
                            futures[future] = wg_name

                    for future in as_completed(futures):
                        if res := future.result():
                            all_tasks.extend(res)
                            for t in res:
                                mapped.add(f"{t['wg_name']}:{t['meeting_num']}")

                if all_tasks:
                    self.ui_log_msg.emit(f"⏳ Bulk-saving {len(all_tasks)} folders to local Database... (This is fast)",
                                         logging.INFO)
                    self.db.insert_meetings_bulk(all_tasks)
                    self.ui_log_msg.emit(f"✅ [Phase 1/3] Successfully mapped & saved {len(all_tasks)} meeting folders.",
                                         logging.INFO)

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

            self.finished_path.emit("MEETINGS_DB_PHASE_1")

            if self.sync_docs:
                if not all_tasks:
                    self.ui_log_msg.emit("⚠️ No meetings available to scan for Docs.", logging.WARNING)
                else:
                    self.ui_log_msg.emit(f"⏳ [Phase 2/3] Deep scraping Docs for {len(all_tasks)} meetings...",
                                         logging.INFO)
                    completed, tdocs_found = 0, 0
                    p2_start = time.time()

                    all_docs_data = []

                    with ThreadPoolExecutor(max_workers=10) as executor:
                        future_to_task = {}
                        for task in all_tasks:
                            future = executor.submit(self.process_individual_meeting, task)
                            future_to_task[future] = task

                        for future in as_completed(future_to_task):
                            task = future_to_task[future]
                            completed += 1
                            try:
                                docs_tuple, count = future.result()
                                tdocs_found += count
                                all_docs_data.append(docs_tuple)
                            except Exception as e:
                                self.ui_log_msg.emit(f"❌ Error scraping {task['folder_name']}: {e}", logging.ERROR)

                            if completed % 10 == 0 or completed == len(all_tasks):
                                elapsed = time.time() - p2_start
                                rate = completed / elapsed if elapsed > 0 else 0
                                self.ui_log_msg.emit(
                                    f"⏳ Scanned {completed}/{len(all_tasks)} Docs folders "
                                    f"| TDocs: {tdocs_found} | Speed: {rate:.1f} mtg/sec",
                                    logging.INFO
                                )

                    if all_docs_data:
                        self.db.update_meeting_docs_bulk(all_docs_data)

                    self.ui_log_msg.emit(f"✅ Pass 2 Complete. Indexed {tdocs_found} total TDocs.", logging.INFO)
            else:
                self.ui_log_msg.emit("⏭️ [Phase 2/3] Skipping Docs folder deep scrape...", logging.INFO)

            self.finished_path.emit("MEETINGS_DB_PHASE_2")

            if self.sync_dyna:
                self.ui_log_msg.emit("⏳ [Phase 3/3] Updating metadata from DynaReports...", logging.INFO)
                wgs_to_fetch = {t["wg"] for t in
                                self.target_meetings} if self.target_meetings else MEETING_SOURCES.keys()

                completed_dyna = 0
                total_dyna = len(wgs_to_fetch)

                all_metadata = []

                with ThreadPoolExecutor(max_workers=5) as executor:
                    dyna_futures = []
                    for wg_name in wgs_to_fetch:
                        urls = MEETING_SOURCES[wg_name]["ftp"]
                        dyna_url = MEETING_SOURCES[wg_name]["dyna"]
                        future = executor.submit(self.process_dynareport, wg_name, dyna_url)
                        dyna_futures.append(future)

                    for future in as_completed(dyna_futures):
                        if res := future.result():
                            all_metadata.extend(res)
                        completed_dyna += 1
                        self.ui_log_msg.emit(f"⏳ DynaReports: {completed_dyna}/{total_dyna} pages processed...",
                                             logging.INFO)

                if all_metadata:
                    self.ui_log_msg.emit(f"⏳ Bulk-saving metadata for {len(all_metadata)} meetings...", logging.INFO)
                    self.db.update_meeting_metadata_bulk(all_metadata)

                self.ui_log_msg.emit("✅ Pass 3 Complete.", logging.INFO)
            else:
                self.ui_log_msg.emit("⏭️ [Phase 3/3] Skipping DynaReports metadata update...", logging.INFO)

            self.ui_log_msg.emit(f"✅ 3GPP Sync Fully Complete in {time.time() - start_time:.1f}s!", logging.INFO)
            self.finished_path.emit("MEETINGS_DB_PHASE_3")

        except Exception as e:
            self.ui_log_msg.emit(f"❌ Critical Failure: {str(e)}", logging.ERROR)
        finally:
            self.finished.emit()


# ==========================================
# --- MANUAL SINGLE MEETING FETCHER ---
# ==========================================
class ManualMeetingFetcherThread(QThread):
    progress_msg = pyqtSignal(str)
    fetch_finished = pyqtSignal(bool, dict, str)

    def __init__(self, db_path: Path, identifier: str, wg_hint: str = None):
        super().__init__()
        self.db_path = db_path
        self.identifier = identifier.strip()
        self.wg_hint = wg_hint.strip() if wg_hint and wg_hint not in ["All", "All / Auto-Detect"] else None

    def run(self):
        try:
            crawler = MeetingsCrawlerThread(self.db_path)
            target_id = ""
            target_num = ""
            target_wg = self.wg_hint

            # 1. Parse the input query
            # Case A: Direct FTP URL or path
            if "ftp/" in self.identifier.lower() or self.identifier.lower().startswith("tsg_"):
                cleaned = re.sub(r'^https?://[^/]+/ftp/', '', self.identifier, flags=re.IGNORECASE).strip('/')
                parts = cleaned.split('/')
                folder = parts[-1]
                # Try to resolve WG from path
                for wg, s in MEETING_SOURCES.items():
                    for u in (s["ftp"] if isinstance(s["ftp"], list) else [s["ftp"]]):
                        if parts[0] in u:
                            target_wg = wg
                            break
                meeting_num = crawler.extract_meeting_number(folder)
                result = {
                    "wg_name": target_wg or "SA3",
                    "meeting_number": meeting_num,
                    "url_key": cleaned,
                    "folder_name": folder,
                    "docs_folder_url": f"https://www.3gpp.org/ftp/{cleaned}/Docs/",
                    "mtg_id": "",
                    "name": f"{target_wg or ''} #{meeting_num}",
                    "location": "",
                    "start_date": "",
                    "end_date": "",
                    "first_tdoc": "",
                    "last_tdoc": "",
                    "is_ad_hoc": 1 if ("AH" in meeting_num.upper() or "BIS" in meeting_num.upper()) else 0,
                    "is_electronic": 1 if ("E" in meeting_num.upper()) else 0
                }
                self._inspect_docs_and_finish(crawler, result)
                return

            # Case B: Standard Compound Tag (e.g. "SA3-130", "SA3#130", "SA2_160e")
            tag_match = re.match(r'^(?:3GPP\s*)?([A-Za-z0-9]+)[_\-#\s]+([0-9a-zA-Z\-]+)$', self.identifier)
            if tag_match:
                raw_wg = tag_match.group(1).upper()
                target_num = tag_match.group(2).replace('-', '').upper()

                # Normalize WG names
                wg_map = {
                    "SP": "SA", "S1": "SA1", "S2": "SA2", "S3": "SA3", "S4": "SA4", "S5": "SA5", "S6": "SA6",
                    "RP": "RAN", "R1": "RAN1", "R2": "RAN2", "R3": "RAN3", "R4": "RAN4",
                    "CP": "CT", "C1": "CT1", "C2": "CT2", "C3": "CT3", "C4": "CT4", "C5": "CT5", "C6": "CT6"
                }
                target_wg = wg_map.get(raw_wg, raw_wg if raw_wg in MEETING_SOURCES else target_wg)

            # Case C: Pure Numeric (MtgId or Meeting Number)
            elif self.identifier.isdigit():
                if len(self.identifier) >= 4:  # Likely a 3GPP Portal MtgId (e.g. 33120)
                    target_id = self.identifier
                else:  # Likely a Meeting number (e.g. 130)
                    target_num = self.identifier
            else:
                target_num = self.identifier

            # 2. Query DynaReports for Metadata
            self.progress_msg.emit("⏳ Querying 3GPP DynaReports...")
            wgs_to_check = [target_wg] if (target_wg and target_wg in MEETING_SOURCES) else list(MEETING_SOURCES.keys())

            found_meta = None
            for wg in wgs_to_check:
                dyna_url = MEETING_SOURCES[wg]["dyna"]
                reports = crawler.process_dynareport(wg, dyna_url)
                for item in reports:
                    wg_name, full_num, url_key, mtg_id, m_name, town, start_d, end_d, new_m_num = item

                    # Match conditions
                    if target_id and mtg_id == target_id:
                        found_meta = item
                        break
                    if target_num and (full_num.upper() == target_num.upper() or new_m_num.upper() == target_num.upper()):
                        found_meta = item
                        break
                    if target_num and target_num.upper() in m_name.upper():
                        found_meta = item
                        break

                if found_meta:
                    break

            result = {}
            if found_meta:
                wg_name, full_num, url_key, mtg_id, m_name, town, start_d, end_d, new_m_num = found_meta
                meeting_number = new_m_num if new_m_num else full_num
                result = {
                    "wg_name": wg_name,
                    "meeting_number": meeting_number,
                    "url_key": url_key,
                    "mtg_id": mtg_id,
                    "name": m_name,
                    "location": town,
                    "start_date": start_d,
                    "end_date": end_d,
                    "folder_name": url_key.split('/')[-1] if url_key else meeting_number,
                    "docs_folder_url": f"https://www.3gpp.org/ftp/{url_key}/Docs/" if url_key else "",
                    "first_tdoc": "",
                    "last_tdoc": "",
                    "is_ad_hoc": 1 if ("AH" in meeting_number.upper() or "BIS" in meeting_number.upper()) else 0,
                    "is_electronic": 1 if ("E" in meeting_number.upper() or "electronic" in m_name.lower()) else 0,
                }
            else:
                # Initialize placeholder result if DynaReport didn't index it yet
                result = {
                    "wg_name": target_wg or "SA3",
                    "meeting_number": target_num or self.identifier,
                    "url_key": "",
                    "mtg_id": target_id or "",
                    "name": f"{target_wg or ''} #{target_num or self.identifier}",
                    "location": "",
                    "start_date": "",
                    "end_date": "",
                    "folder_name": "",
                    "docs_folder_url": "",
                    "first_tdoc": "",
                    "last_tdoc": "",
                    "is_ad_hoc": 1 if ("AH" in (target_num or self.identifier).upper()) else 0,
                    "is_electronic": 1 if ("E" in (target_num or self.identifier).upper()) else 0
                }

            # 3. Locate FTP folder if url_key is missing or needs verification
            active_wg = result.get("wg_name")
            if active_wg and active_wg in MEETING_SOURCES and not result.get("url_key"):
                self.progress_msg.emit(f"🌐 Searching FTP directory for {active_wg}...")
                urls = MEETING_SOURCES[active_wg]["ftp"]
                urls = urls if isinstance(urls, list) else [urls]
                for ftp_url in urls:
                    is_ah = "AH" in ftp_url
                    folders = crawler.fetch_wg_directories(active_wg, ftp_url, is_ah)
                    for f_task in folders:
                        m_num = f_task["meeting_num"].upper()
                        f_name = f_task["folder_name"].upper()
                        req_num = (result.get("meeting_number") or target_num or self.identifier).upper()

                        if m_num == req_num or req_num in f_name or f_name.endswith(f"_{req_num}"):
                            result["url_key"] = f_task["url_key"]
                            result["folder_name"] = f_task["folder_name"]
                            result["docs_folder_url"] = f_task["docs_url"]
                            if not result.get("meeting_number"):
                                result["meeting_number"] = f_task["meeting_num"]
                            if f_task.get("is_ad_hoc"):
                                result["is_ad_hoc"] = 1
                            break
                    if result.get("url_key"):
                        break

            # 4. Deep scrape Docs folder for TDoc range
            self._inspect_docs_and_finish(crawler, result)

        except Exception as e:
            self.fetch_finished.emit(False, {}, f"Error querying meeting: {str(e)}")

    def _inspect_docs_and_finish(self, crawler: MeetingsCrawlerThread, result: dict):
        if result.get("url_key"):
            self.progress_msg.emit("📄 Inspecting Docs/ folder for TDoc range...")
            docs_url = result.get("docs_folder_url") or f"https://www.3gpp.org/ftp/{result['url_key']}/Docs/"
            task = {
                "wg_name": result["wg_name"],
                "folder_name": result.get("folder_name", result["meeting_number"]),
                "meeting_num": result["meeting_number"],
                "url_key": result["url_key"],
                "absolute_url": f"https://www.3gpp.org/ftp/{result['url_key']}",
                "docs_url": docs_url
            }
            docs_tuple, tdoc_count = crawler.process_individual_meeting(task)
            final_docs_url, first_tdoc, first_pfx, first_num, last_tdoc, last_pfx, last_num, _ = docs_tuple

            result["docs_folder_url"] = final_docs_url
            result["first_tdoc"] = first_tdoc
            result["first_tdoc_prefix"] = first_pfx
            result["first_tdoc_num"] = first_num
            result["last_tdoc"] = last_tdoc
            result["last_tdoc_prefix"] = last_pfx
            result["last_tdoc_num"] = last_num
            result["tdoc_count"] = tdoc_count

        if not result.get("wg_name") and not result.get("meeting_number") and not result.get("url_key"):
            self.fetch_finished.emit(False, {}, f"Could not find any meeting matching '{self.identifier}'.")
        else:
            self.fetch_finished.emit(True, result, f"Ready: {result.get('wg_name', '')} #{result.get('meeting_number', '')}")