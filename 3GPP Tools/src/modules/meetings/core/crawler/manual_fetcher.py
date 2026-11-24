"""Background single-meeting resolver querying DynaReports and FTP Docs/."""

from __future__ import annotations

import re
from pathlib import Path

from PyQt5.QtCore import QThread, pyqtSignal

from .crawler_thread import MeetingsCrawlerThread
from .sources import MEETING_SOURCES

_WG_ALIAS_MAP = {
    "SP": "SA",
    "S1": "SA1",
    "S2": "SA2",
    "S3": "SA3",
    "S4": "SA4",
    "S5": "SA5",
    "S6": "SA6",
    "RP": "RAN",
    "R1": "RAN1",
    "R2": "RAN2",
    "R3": "RAN3",
    "R4": "RAN4",
    "CP": "CT",
    "C1": "CT1",
    "C2": "CT2",
    "C3": "CT3",
    "C4": "CT4",
    "C5": "CT5",
    "C6": "CT6",
}


class ManualMeetingFetcherThread(QThread):
    """Fetch metadata and TDoc bounds for a user-specified meeting identifier or URL."""

    progress_msg = pyqtSignal(str)
    fetch_finished = pyqtSignal(bool, dict, str)

    def __init__(self, db_path: Path, identifier: str, wg_hint: str = None):
        super().__init__()
        self.db_path = db_path
        self.identifier = identifier.strip()
        self.wg_hint = (
            wg_hint.strip() if wg_hint and wg_hint not in ["All", "All / Auto-Detect"] else None
        )

    def run(self):
        try:
            crawler = MeetingsCrawlerThread(self.db_path)
            target_id = ""
            target_num = ""
            target_wg = self.wg_hint

            # 1. Parse the input query
            # Case A: Direct FTP URL or path
            if "ftp/" in self.identifier.lower() or self.identifier.lower().startswith("tsg_"):
                cleaned = re.sub(
                    r"^https?://[^/]+/ftp/", "", self.identifier, flags=re.IGNORECASE
                ).strip("/")
                parts = cleaned.split("/")
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
                    "is_ad_hoc": 1
                    if ("AH" in meeting_num.upper() or "BIS" in meeting_num.upper())
                    else 0,
                    "is_electronic": 1 if ("E" in meeting_num.upper()) else 0,
                }
                self._inspect_docs_and_finish(crawler, result)
                return

            # Case B: Standard Compound Tag (e.g. "SA3-130", "SA3#130", "SA2_160e")
            tag_match = re.match(
                r"^(?:3GPP\s*)?([A-Za-z0-9]+)[_\-#\s]+([0-9a-zA-Z\-]+)$",
                self.identifier,
            )
            if tag_match:
                raw_wg = tag_match.group(1).upper()
                target_num = tag_match.group(2).replace("-", "").upper()
                target_wg = _WG_ALIAS_MAP.get(
                    raw_wg, raw_wg if raw_wg in MEETING_SOURCES else target_wg
                )

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
            wgs_to_check = (
                [target_wg]
                if (target_wg and target_wg in MEETING_SOURCES)
                else list(MEETING_SOURCES.keys())
            )

            found_meta = None
            for wg in wgs_to_check:
                dyna_url = MEETING_SOURCES[wg]["dyna"]
                reports = crawler.process_dynareport(wg, dyna_url)
                for item in reports:
                    _wg_name, full_num, _url_key, mtg_id, m_name, _town, _sd, _ed, new_m_num = item

                    if target_id and mtg_id == target_id:
                        found_meta = item
                        break
                    if target_num and (
                        full_num.upper() == target_num.upper()
                        or new_m_num.upper() == target_num.upper()
                    ):
                        found_meta = item
                        break
                    if target_num and target_num.upper() in m_name.upper():
                        found_meta = item
                        break

                if found_meta:
                    break

            if found_meta:
                wg_name, full_num, url_key, mtg_id, m_name, town, start_d, end_d, new_m_num = (
                    found_meta
                )
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
                    "folder_name": url_key.split("/")[-1] if url_key else meeting_number,
                    "docs_folder_url": f"https://www.3gpp.org/ftp/{url_key}/Docs/"
                    if url_key
                    else "",
                    "first_tdoc": "",
                    "last_tdoc": "",
                    "is_ad_hoc": 1
                    if ("AH" in meeting_number.upper() or "BIS" in meeting_number.upper())
                    else 0,
                    "is_electronic": 1
                    if ("E" in meeting_number.upper() or "electronic" in m_name.lower())
                    else 0,
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
                    "is_electronic": 1 if ("E" in (target_num or self.identifier).upper()) else 0,
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
                        req_num = (
                            result.get("meeting_number") or target_num or self.identifier
                        ).upper()

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
            docs_url = (
                result.get("docs_folder_url")
                or f"https://www.3gpp.org/ftp/{result['url_key']}/Docs/"
            )
            task = {
                "wg_name": result["wg_name"],
                "folder_name": result.get("folder_name", result["meeting_number"]),
                "meeting_num": result["meeting_number"],
                "url_key": result["url_key"],
                "absolute_url": f"https://www.3gpp.org/ftp/{result['url_key']}",
                "docs_url": docs_url,
            }
            docs_tuple, tdoc_count = crawler.process_individual_meeting(task)
            (
                final_docs_url,
                first_tdoc,
                first_pfx,
                first_num,
                last_tdoc,
                last_pfx,
                last_num,
                _,
            ) = docs_tuple

            result["docs_folder_url"] = final_docs_url
            result["first_tdoc"] = first_tdoc
            result["first_tdoc_prefix"] = first_pfx
            result["first_tdoc_num"] = first_num
            result["last_tdoc"] = last_tdoc
            result["last_tdoc_prefix"] = last_pfx
            result["last_tdoc_num"] = last_num
            result["tdoc_count"] = tdoc_count

        if (
            not result.get("wg_name")
            and not result.get("meeting_number")
            and not result.get("url_key")
        ):
            self.fetch_finished.emit(
                False, {}, f"Could not find any meeting matching '{self.identifier}'."
            )
        else:
            self.fetch_finished.emit(
                True,
                result,
                f"Ready: {result.get('wg_name', '')} #{result.get('meeting_number', '')}",
            )
