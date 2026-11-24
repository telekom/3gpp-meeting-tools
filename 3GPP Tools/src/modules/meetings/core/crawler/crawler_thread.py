"""Multi-phase QThread orchestrator for syncing 3GPP meetings into SQLite."""

from __future__ import annotations

import logging
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

from PyQt5.QtCore import QThread, pyqtSignal

from modules.meetings.core.meetings_db import MeetingsDatabase
from .dynareport import process_dynareport as _process_dynareport_impl
from .ftp_and_docs import (
    fetch_wg_directories as _fetch_wg_directories_impl,
    process_individual_meeting as _process_individual_meeting_impl,
)
from .sources import (
    HREF_PATTERN,
    MEETING_DENY_LIST,
    MEETING_NAME_PATTERN,
    MEETING_NUM_PATTERN,
    MEETING_SOURCES,
    TDOC_LINK_PATTERN,
    extract_meeting_number as _extract_meeting_number_impl,
    is_meeting_name as _is_meeting_name_impl,
)

logger = logging.getLogger(__name__)


class MeetingsCrawlerThread(QThread):
    """Background 3-phase crawler syncing WG folders, Docs/ ranges, and DynaReports."""

    ui_log_msg = pyqtSignal(str, int)
    finished = pyqtSignal()
    finished_path = pyqtSignal(str)

    def __init__(
        self,
        db_path: Path,
        target_meetings: list = None,
        sync_wg=True,
        sync_docs=True,
        sync_dyna=True,
    ):
        super().__init__()
        self.db = MeetingsDatabase(db_path)
        self.target_meetings = target_meetings or []
        self.sync_wg = sync_wg
        self.sync_docs = sync_docs
        self.sync_dyna = sync_dyna

        self.meeting_pattern = MEETING_NAME_PATTERN
        self.num_pattern = MEETING_NUM_PATTERN
        self.href_pattern = HREF_PATTERN
        self.tdoc_pattern = TDOC_LINK_PATTERN
        self.deny_list = set(MEETING_DENY_LIST)

    def is_meeting(self, folder_name: str) -> bool:
        """Return True if an FTP folder name matches a 3GPP meeting directory."""
        return _is_meeting_name_impl(
            folder_name, deny_list=self.deny_list, pattern=self.meeting_pattern
        )

    def extract_meeting_number(self, folder_name: str) -> str:
        """Extract a normalized meeting number from a directory name."""
        return _extract_meeting_number_impl(folder_name, pattern=self.num_pattern)

    def fetch_wg_directories(
        self, wg_name: str, ftp_base_url: str, is_ah_folder: bool = False
    ) -> list:
        """Fetch WG meeting folder descriptors from an FTP index URL."""
        return _fetch_wg_directories_impl(
            wg_name,
            ftp_base_url,
            is_ah_folder=is_ah_folder,
            target_meetings=self.target_meetings,
            href_pattern=self.href_pattern,
            deny_list=self.deny_list,
            is_meeting_fn=self.is_meeting,
            extract_num_fn=self.extract_meeting_number,
        )

    def process_individual_meeting(self, task: dict) -> tuple:
        """Scan one meeting's Docs/ listing and return `(docs_data, tdoc_count)`."""
        return _process_individual_meeting_impl(task, tdoc_pattern=self.tdoc_pattern)

    def process_dynareport(self, wg_name: str, dyna_url: str) -> list:
        """Parse metadata rows from a WG's DynaReport page."""
        return _process_dynareport_impl(
            wg_name, dyna_url, target_meetings=self.target_meetings
        )

    def run(self):
        start_time = time.time()
        try:
            all_tasks = []

            if self.sync_wg:
                logger.info("⏳ [Phase 1/3] Mapping WG directories...")
                mapped = set()

                with ThreadPoolExecutor(max_workers=15) as executor:
                    futures = {}
                    for wg_name, source_info in MEETING_SOURCES.items():
                        urls = (
                            source_info["ftp"]
                            if isinstance(source_info["ftp"], list)
                            else [source_info["ftp"]]
                        )
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
                    logger.info(
                        f"⏳ Bulk-saving {len(all_tasks)} folders to local Database... (This is fast)"
                    )
                    self.db.insert_meetings_bulk(all_tasks)
                    logger.info(
                        f"✅ [Phase 1/3] Successfully mapped & saved {len(all_tasks)} meeting folders."
                    )

                if self.target_meetings:
                    for t in self.target_meetings:
                        if f"{t['wg']}:{t['meeting']}" not in mapped:
                            logger.warning(
                                f"⚠️ Target {t['wg']}:{t['meeting']} not found on FTP!"
                            )
            else:
                logger.info("⏭️ [Phase 1/3] Skipping Directory Mapping (loading DB)...")
                for m in self.db.search_meetings():
                    if self.target_meetings and not any(
                        t["wg"] == m["wg_name"] and t["meeting"] == m["meeting_number"]
                        for t in self.target_meetings
                    ):
                        continue
                    if m.get("url_key"):
                        url_key = str(m["url_key"]).strip("/")
                        absolute_url = f"https://www.3gpp.org/ftp/{url_key}"
                        all_tasks.append(
                            {
                                "wg_name": m["wg_name"],
                                "folder_name": m.get("folder_name") or m["meeting_number"],
                                "meeting_num": m["meeting_number"],
                                "url_key": url_key,
                                "absolute_url": absolute_url,
                                "docs_url": absolute_url.rstrip("/") + "/Docs/",
                            }
                        )

            self.finished_path.emit("MEETINGS_DB_PHASE_1")

            if self.sync_docs:
                if not all_tasks:
                    logger.warning("⚠️ No meetings available to scan for Docs.")
                else:
                    logger.info(
                        f"⏳ [Phase 2/3] Deep scraping Docs for {len(all_tasks)} meetings..."
                    )
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
                                logger.error(f"❌ Error scraping {task['folder_name']}: {e}")

                            if completed % 10 == 0 or completed == len(all_tasks):
                                elapsed = time.time() - p2_start
                                rate = completed / elapsed if elapsed > 0 else 0
                                logger.info(
                                    f"⏳ Scanned {completed}/{len(all_tasks)} Docs folders "
                                    f"| TDocs: {tdocs_found} | Speed: {rate:.1f} mtg/sec"
                                )

                    if all_docs_data:
                        self.db.update_meeting_docs_bulk(all_docs_data)

                    logger.info(f"✅ Pass 2 Complete. Indexed {tdocs_found} total TDocs.")
            else:
                logger.info("⏭️ [Phase 2/3] Skipping Docs folder deep scrape...")

            self.finished_path.emit("MEETINGS_DB_PHASE_2")

            if self.sync_dyna:
                logger.info("⏳ [Phase 3/3] Updating metadata from DynaReports...")
                wgs_to_fetch = (
                    {t["wg"] for t in self.target_meetings}
                    if self.target_meetings
                    else MEETING_SOURCES.keys()
                )

                completed_dyna = 0
                total_dyna = len(wgs_to_fetch)

                all_metadata = []

                with ThreadPoolExecutor(max_workers=5) as executor:
                    dyna_futures = []
                    for wg_name in wgs_to_fetch:
                        dyna_url = MEETING_SOURCES[wg_name]["dyna"]
                        future = executor.submit(self.process_dynareport, wg_name, dyna_url)
                        dyna_futures.append(future)

                    for future in as_completed(dyna_futures):
                        if res := future.result():
                            all_metadata.extend(res)
                        completed_dyna += 1
                        logger.info(
                            f"⏳ DynaReports: {completed_dyna}/{total_dyna} pages processed..."
                        )

                if all_metadata:
                    logger.info(
                        f"⏳ Bulk-saving metadata for {len(all_metadata)} meetings..."
                    )
                    self.db.update_meeting_metadata_bulk(all_metadata)

                logger.info("✅ Pass 3 Complete.")
            else:
                logger.info("⏭️ [Phase 3/3] Skipping DynaReports metadata update...")

            logger.info(f"✅ 3GPP Sync Fully Complete in {time.time() - start_time:.1f}s!")
            self.finished_path.emit("MEETINGS_DB_PHASE_3")

        except Exception as e:
            logger.error(f"❌ Critical Failure: {str(e)}")
        finally:
            self.finished.emit()
