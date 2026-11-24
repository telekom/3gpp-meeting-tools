"""QThread orchestrator for Cross-Meeting Study Item & Topic Agreement tracking."""

from __future__ import annotations

import logging
import os
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from typing import List, Optional, Set

from PyQt5.QtCore import QThread, pyqtSignal

from modules.meetings.core.agreement_tracker.fetchers import AgreementFetchersMixin
from modules.meetings.core.agreement_tracker.filters import AgreementFilterMixin
from modules.meetings.core.agreement_tracker.processor import AgreementProcessorMixin
from modules.meetings.core.meetings_db import MeetingsDatabase
from modules.meetings.ui.tdocs_models import TDocsTableModel

logger = logging.getLogger(__name__)


class AgreementTrackerWorker(
    QThread,
    AgreementFilterMixin,
    AgreementFetchersMixin,
    AgreementProcessorMixin,
):
    """Background worker scanning 3GPP meetings for Chairman's Notes & Agreed TDoc agreements."""

    stage_progress = pyqtSignal(str, int, int)
    log_msg = pyqtSignal(str)
    results_ready = pyqtSignal(list)
    finished = pyqtSignal(bool, str)

    def __init__(
        self,
        meetings_db: MeetingsDatabase,
        wgs: List[str],
        date_from: str,
        date_to: str,
        target_wis: List[str],
        target_ais: List[str],
        keyword_query: str = "",
        wi_ai_or: bool = False,
        include_agreements: bool = True,
        include_ffs: bool = True,
        include_agreed_tdocs: bool = True,
        include_prev_recaps: bool = False,
        auto_download_missing: bool = True,
        bypass_cache: bool = False,
        cache_dir: str = "",
        parent=None,
    ):
        super().__init__(parent)
        self.meetings_db = meetings_db
        self.wgs = [w.strip() for w in (wgs or []) if w.strip()]
        self.date_from = date_from
        self.date_to = date_to
        self.target_wis = [w.strip() for w in (target_wis or []) if w.strip()]
        self.target_wi_base_tokens: Set[str] = set()
        for w in self.target_wis:
            self.target_wi_base_tokens.update(TDocsTableModel._extract_base_tokens(w))
            self.target_wi_base_tokens.add(w.strip().lower())

        self.target_ais = [self._normalize_ai(ai) for ai in (target_ais or []) if self._normalize_ai(ai)]
        self.keyword_query = (keyword_query or "").strip()
        self.wi_ai_or = bool(wi_ai_or)
        self.include_agreements = bool(include_agreements)
        self.include_ffs = bool(include_ffs)
        self.include_agreed_tdocs = bool(include_agreed_tdocs)
        self.include_prev_recaps = bool(include_prev_recaps)
        self.auto_download_missing = bool(auto_download_missing)
        self.bypass_cache = bool(bypass_cache)
        self.cache_dir = Path(cache_dir)
        self._is_cancelled = False
        self._executor: Optional[ThreadPoolExecutor] = None

    def _log(self, message: str, level: int = logging.INFO):
        logger.log(level, message)
        self.log_msg.emit(message)

    @property
    def is_cancelled(self) -> bool:
        return self._is_cancelled or self.isInterruptionRequested()

    def cancel(self):
        self._is_cancelled = True
        self.requestInterruption()
        if self._executor:
            try:
                self._executor.shutdown(wait=False, cancel_futures=True)
            except Exception:
                pass

    def run(self):
        try:
            self.stage_progress.emit("Stage 1/3: Resolving matching meetings...", 0, 100)
            self._log("🔍 Querying meetings database for target Working Groups and dates...")

            meetings = self.meetings_db.search_meetings(
                wg_name=self.wgs if self.wgs else None,
                date_from=self.date_from,
                date_to=self.date_to,
            )

            total_meetings = len(meetings)
            if total_meetings == 0:
                self._log("⚠️ No meetings found matching the selected criteria.", logging.WARNING)
                self.results_ready.emit([])
                self.finished.emit(True, "No meetings found.")
                return

            self._log(f"✅ Found {total_meetings} meeting(s) to scan for agreements.")

            all_matched_agreements: List[dict] = []
            processed_count = 0
            max_workers = min(3, os.cpu_count() or 2)
            self._log(f"⚡ Starting agreement extraction pool ({max_workers} threads)...")

            with ThreadPoolExecutor(max_workers=max_workers) as executor:
                self._executor = executor
                future_to_meeting = {
                    executor.submit(self._process_single_meeting, mtg): mtg
                    for mtg in meetings
                }

                for future in as_completed(future_to_meeting):
                    if self.is_cancelled:
                        self._log("🛑 Operation cancelled by user.", logging.INFO)
                        break

                    processed_count += 1
                    status_desc = f"Stage 2/3: Scanning meetings ({processed_count}/{total_meetings})..."
                    self.stage_progress.emit(status_desc, processed_count, total_meetings)

                    try:
                        matches = future.result()
                        all_matched_agreements.extend(matches)
                    except Exception as e:
                        mtg_ref = future_to_meeting[future]
                        self._log(
                            f"❌ Error in meeting {mtg_ref.get('wg_name')} #{mtg_ref.get('meeting_number')}: {e}",
                            logging.ERROR,
                        )

            self._executor = None
            if self.is_cancelled:
                self.finished.emit(False, "Search cancelled.")
                return

            self.stage_progress.emit("Stage 3/3: Sorting & organizing agreements...", 100, 100)
            all_matched_agreements.sort(
                key=lambda r: (
                    r.get("start_date", ""),
                    r.get("WG", ""),
                    r.get("meeting_number", ""),
                    self._ai_sort_key(r.get("AI", "")),
                    r.get("TDoc", ""),
                )
            )

            summary = (
                f"Scan complete! Found {len(all_matched_agreements)} matching agreement(s)/item(s) "
                f"across {total_meetings} meeting(s)."
            )
            self._log(f"🏁 {summary}")
            self.results_ready.emit(all_matched_agreements)
            self.finished.emit(True, summary)

        except Exception as e:
            err = f"Error during agreement tracking: {e}"
            logger.error(err, exc_info=True)
            self._log(f"❌ {err}", logging.ERROR)
            self.finished.emit(False, str(e))
