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
    """Background 3-phase crawler syncing WG folders, DynaReports, and Docs/ ranges."""

    ui_log_msg = pyqtSignal(str, int)
    progress_updated = pyqtSignal(str, int, int)
    task_summary = pyqtSignal(dict)
    finished = pyqtSignal()
    finished_path = pyqtSignal(str)

    def __init__(
        self,
        db_path: Path,
        target_meetings: list = None,
        sync_wg: bool = True,
        sync_docs: bool = True,
        sync_dyna: bool = True,
        target_wgs: list = None,
        docs_mode: str = "smart",
    ):
        super().__init__()
        self.db = MeetingsDatabase(db_path)
        self.target_meetings = target_meetings or []
        self.sync_wg = sync_wg
        self.sync_docs = sync_docs
        self.sync_dyna = sync_dyna
        self.target_wgs = [w for w in (target_wgs or []) if w in MEETING_SOURCES]
        if not sync_docs:
            self.docs_mode = "none"
        elif self.target_meetings:
            self.docs_mode = "all"
        elif docs_mode in ("smart", "all", "none"):
            self.docs_mode = docs_mode
        else:
            self.docs_mode = "smart"

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

    def _load_db_docs_tasks(self, wgs_to_sync: list) -> list:
        """Build Docs/ scan task descriptors from meetings stored in SQLite."""
        tasks = []
        for m in self.db.search_meetings(wg_name=wgs_to_sync):
            if self.target_meetings and not any(
                t["wg"] == m["wg_name"]
                and str(t["meeting"]).upper() == str(m["meeting_number"]).upper()
                for t in self.target_meetings
            ):
                continue
            if m.get("url_key"):
                url_key = str(m["url_key"]).strip("/")
                absolute_url = f"https://www.3gpp.org/ftp/{url_key}"
                tasks.append(
                    {
                        "wg_name": m["wg_name"],
                        "folder_name": m.get("folder_name") or m["meeting_number"],
                        "meeting_num": m["meeting_number"],
                        "url_key": url_key,
                        "absolute_url": absolute_url,
                        "docs_url": absolute_url.rstrip("/") + "/Docs/",
                    }
                )
        return tasks

    def run(self):
        start_time = time.time()
        new_meetings = 0
        updated_meetings = 0
        docs_scanned = 0
        tdocs_found = 0
        cancelled = False

        if self.target_meetings:
            wgs_to_sync = list(
                dict.fromkeys(
                    t["wg"] for t in self.target_meetings if t.get("wg") in MEETING_SOURCES
                )
            )
        elif self.target_wgs:
            wgs_to_sync = list(dict.fromkeys(self.target_wgs))
        else:
            wgs_to_sync = list(MEETING_SOURCES.keys())

        try:
            all_tasks = []
            p1_inserted_keys = set()

            # --- STEP 1/3: FTP Directory Discovery ---
            if self.sync_wg and wgs_to_sync:
                logger.info(f"⏳ [Step 1/3] Mapping FTP directories for {len(wgs_to_sync)} WG(s)...")
                mapped = set()
                url_jobs = []
                for wg_name in wgs_to_sync:
                    source_info = MEETING_SOURCES[wg_name]
                    urls = (
                        source_info["ftp"]
                        if isinstance(source_info["ftp"], list)
                        else [source_info["ftp"]]
                    )
                    for url in urls:
                        url_jobs.append((wg_name, url, "AH" in url))

                self.progress_updated.emit("Step 1/3: Scanning FTP folders...", 0, len(url_jobs))
                done_urls = 0

                with ThreadPoolExecutor(max_workers=15) as executor:
                    futures = {
                        executor.submit(self.fetch_wg_directories, wg, url, is_ah): wg
                        for wg, url, is_ah in url_jobs
                    }
                    for future in as_completed(futures):
                        if self.isInterruptionRequested():
                            cancelled = True
                            for pending in futures:
                                pending.cancel()
                            break
                        done_urls += 1
                        if res := future.result():
                            all_tasks.extend(res)
                            for t in res:
                                mapped.add(f"{t['wg_name']}:{t['meeting_num']}")
                        self.progress_updated.emit(
                            f"Step 1/3: Scanning FTP folders ({done_urls}/{len(url_jobs)})...",
                            done_urls,
                            len(url_jobs),
                        )

                if hasattr(self.db, "cleanup_container_rows"):
                    self.db.cleanup_container_rows()

                if all_tasks:
                    logger.info(
                        f"⏳ Bulk-saving {len(all_tasks)} folders to local Database..."
                    )
                    p1_stats = self.db.insert_meetings_bulk(all_tasks) or {}
                    new_meetings += int(p1_stats.get("inserted", 0))
                    updated_meetings += int(p1_stats.get("updated", 0)) + int(
                        p1_stats.get("linked", 0)
                    )
                    p1_inserted_keys = set(p1_stats.get("inserted_url_keys", []))
                    logger.info(
                        f"✅ [Step 1/3] Mapped {len(all_tasks)} folders "
                        f"({p1_stats.get('inserted', 0)} new, {p1_stats.get('linked', 0)} linked)."
                    )

                if self.target_meetings and not cancelled:
                    for t in self.target_meetings:
                        if f"{t['wg']}:{t['meeting']}" not in mapped:
                            logger.warning(
                                f"⚠️ Target {t['wg']}:{t['meeting']} not found on FTP!"
                            )

                self.finished_path.emit("MEETINGS_DB_PHASE_1")
            else:
                logger.info("⏭️ [Step 1/3] Skipping FTP Directory Mapping...")
                self.finished_path.emit("MEETINGS_DB_PHASE_1")

            # --- STEP 2/3: DynaReport Metadata (Dates, Locations, Portal MtgIds) ---
            if self.sync_dyna and wgs_to_sync and not cancelled and not self.isInterruptionRequested():
                logger.info(f"⏳ [Step 2/3] Updating metadata from DynaReports for {len(wgs_to_sync)} WG(s)...")
                completed_dyna = 0
                total_dyna = len(wgs_to_sync)
                all_metadata = []
                self.progress_updated.emit(
                    f"Step 2/3: Updating dates & Portal IDs (0/{total_dyna})...",
                    0,
                    total_dyna,
                )

                with ThreadPoolExecutor(max_workers=5) as executor:
                    dyna_futures = [
                        executor.submit(
                            self.process_dynareport, wg_name, MEETING_SOURCES[wg_name]["dyna"]
                        )
                        for wg_name in wgs_to_sync
                    ]
                    for future in as_completed(dyna_futures):
                        if self.isInterruptionRequested():
                            cancelled = True
                            for pending in dyna_futures:
                                pending.cancel()
                            break
                        if res := future.result():
                            all_metadata.extend(res)
                        completed_dyna += 1
                        self.progress_updated.emit(
                            f"Step 2/3: Updating dates & Portal IDs ({completed_dyna}/{total_dyna})...",
                            completed_dyna,
                            total_dyna,
                        )
                        logger.info(
                            f"⏳ DynaReports: {completed_dyna}/{total_dyna} pages processed..."
                        )

                if all_metadata:
                    logger.info(
                        f"⏳ Bulk-saving metadata for {len(all_metadata)} meetings..."
                    )
                    p2_stats = self.db.update_meeting_metadata_bulk(all_metadata) or {}
                    new_meetings += int(p2_stats.get("inserted", 0))
                    updated_meetings += int(p2_stats.get("updated", 0))

                logger.info("✅ Step 2/3 (DynaReport) Complete.")
                self.finished_path.emit("MEETINGS_DB_PHASE_2")
            else:
                if self.isInterruptionRequested():
                    cancelled = True
                else:
                    logger.info("⏭️ [Step 2/3] Skipping DynaReports metadata update...")
                self.finished_path.emit("MEETINGS_DB_PHASE_2")

            # --- STEP 3/3: Docs/ TDoc Range Scan ---
            if self.docs_mode != "none" and wgs_to_sync and not cancelled and not self.isInterruptionRequested():
                if self.target_meetings:
                    docs_tasks = all_tasks or self._load_db_docs_tasks(wgs_to_sync)
                elif self.docs_mode == "smart":
                    docs_tasks = self.db.get_smart_docs_candidates(
                        wg_names=wgs_to_sync,
                        force_url_keys=p1_inserted_keys,
                    )
                else:
                    docs_tasks = self._load_db_docs_tasks(wgs_to_sync)

                if not docs_tasks:
                    logger.info("ℹ️ [Step 3/3] No meetings required a Docs/ TDoc-range scan.")
                else:
                    logger.info(
                        f"⏳ [Step 3/3] Scraping Docs/ ({self.docs_mode}) for {len(docs_tasks)} meeting(s)..."
                    )
                    total_docs = len(docs_tasks)
                    p3_start = time.time()
                    all_docs_data = []
                    self.progress_updated.emit(
                        f"Step 3/3: Scanning Docs/ (0/{total_docs})...",
                        0,
                        total_docs,
                    )

                    with ThreadPoolExecutor(max_workers=10) as executor:
                        future_to_task = {
                            executor.submit(self.process_individual_meeting, task): task
                            for task in docs_tasks
                        }
                        for future in as_completed(future_to_task):
                            if self.isInterruptionRequested():
                                cancelled = True
                                for pending in future_to_task:
                                    pending.cancel()
                                break
                            task = future_to_task[future]
                            docs_scanned += 1
                            try:
                                docs_tuple, count = future.result()
                                tdocs_found += count
                                all_docs_data.append(docs_tuple)
                            except Exception as e:
                                logger.error(f"❌ Error scraping {task['folder_name']}: {e}")

                            self.progress_updated.emit(
                                f"Step 3/3: Scanning Docs/ ({docs_scanned}/{total_docs})...",
                                docs_scanned,
                                total_docs,
                            )

                            if docs_scanned % 10 == 0 or docs_scanned == total_docs:
                                elapsed = time.time() - p3_start
                                rate = docs_scanned / elapsed if elapsed > 0 else 0
                                logger.info(
                                    f"⏳ Scanned {docs_scanned}/{total_docs} Docs folders "
                                    f"| TDocs: {tdocs_found} | Speed: {rate:.1f} mtg/sec"
                                )

                    if all_docs_data:
                        self.db.update_meeting_docs_bulk(all_docs_data)

                    logger.info(f"✅ Step 3/3 Complete. Indexed {tdocs_found} total TDocs.")
            else:
                if self.isInterruptionRequested():
                    cancelled = True
                else:
                    logger.info("⏭️ [Step 3/3] Skipping Docs/ deep scrape...")

            self.finished_path.emit("MEETINGS_DB_PHASE_3")
            elapsed_total = time.time() - start_time
            if cancelled:
                logger.warning(f"🛑 3GPP Meetings Sync stopped by user after {elapsed_total:.1f}s.")
            else:
                logger.info(f"✅ 3GPP Sync Fully Complete in {elapsed_total:.1f}s!")

        except Exception as e:
            logger.error(f"❌ Critical Failure: {str(e)}")
        finally:
            self.task_summary.emit(
                {
                    "new_meetings": new_meetings,
                    "updated_meetings": updated_meetings,
                    "docs_scanned": docs_scanned,
                    "tdocs_found": tdocs_found,
                    "wgs_count": len(wgs_to_sync),
                    "elapsed_s": round(time.time() - start_time, 1),
                    "cancelled": cancelled or self.isInterruptionRequested(),
                    "docs_mode": self.docs_mode,
                }
            )
            self.finished.emit()
