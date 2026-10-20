# --- File: src/modules/meetings/core/contribution_worker.py ---
import os
import re
import logging
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

from PyQt5.QtCore import QThread, pyqtSignal

from core.utils.company_sanitizer import CompanySanitizer
from modules.meetings.core.meetings_db import MeetingsDatabase
from modules.meetings.core.tdocs_downloader import TDocsDownloaderThread
from modules.meetings.core.tdocs_parser import TDocsParser

logger = logging.getLogger(__name__)

class ContributionSearchWorker(QThread):
    stage_progress = pyqtSignal(str, int, int)
    log_msg = pyqtSignal(str)
    results_ready = pyqtSignal(list)
    finished = pyqtSignal(bool, str)

    def __init__(
        self,
        meetings_db: MeetingsDatabase,
        wgs: list,
        date_from: str,
        date_to: str,
        target_companies: set,
        primary_only: bool,
        target_wis: list,
        target_ais: list,
        wi_ai_or: bool,
        bypass_cache: bool,
        cache_dir: str,
        parent=None
    ):
        super().__init__(parent)
        self.meetings_db = meetings_db
        self.wgs = wgs
        self.date_from = date_from
        self.date_to = date_to
        self.target_companies = target_companies
        self.primary_only = primary_only
        self.target_wis = [w.strip().lower() for w in target_wis if w.strip()]
        self.target_ais = {self._normalize_ai(ai) for ai in target_ais if self._normalize_ai(ai)}
        self.wi_ai_or = bool(wi_ai_or)
        self.bypass_cache = bypass_cache
        self.cache_dir = Path(cache_dir)
        self._is_cancelled = False
        self._executor = None

    def _log(self, message: str, level: int = logging.INFO):
        logger.log(level, message)
        self.log_msg.emit(message)

    def cancel(self):
        self._is_cancelled = True
        if self._executor:
            try:
                self._executor.shutdown(wait=False, cancel_futures=True)
            except Exception:
                pass

    def run(self):
        try:
            self.stage_progress.emit("Stage 1/3: Resolving matching meetings...", 0, 100)
            self._log("🔍 Querying meetings database for target criteria...")

            meetings = self.meetings_db.search_meetings(
                wg_name=self.wgs if self.wgs else None,
                date_from=self.date_from,
                date_to=self.date_to
            )

            total_meetings = len(meetings)
            if total_meetings == 0:
                self._log("⚠️ No meetings found matching the selected criteria.", logging.WARNING)
                self.results_ready.emit([])
                self.finished.emit(True, "No meetings found.")
                return

            self._log(f"✅ Found {total_meetings} meeting(s) to process.")

            all_matched_tdocs = []
            processed_count = 0

            max_workers = min(3, os.cpu_count() or 2)
            self._log(f"⚡ Starting worker pool ({max_workers} threads)...")

            with ThreadPoolExecutor(max_workers=max_workers) as executor:
                self._executor = executor
                future_to_meeting = {
                    executor.submit(self._process_single_meeting, mtg): mtg
                    for mtg in meetings
                }

                for future in as_completed(future_to_meeting):
                    if self._is_cancelled:
                        self._log("🛑 Operation cancelled by user.", logging.INFO)
                        break

                    processed_count += 1
                    status_desc = f"Stage 2/3: Processing ({processed_count}/{total_meetings})..."
                    self.stage_progress.emit(status_desc, processed_count, total_meetings)

                    try:
                        matches = future.result()
                        all_matched_tdocs.extend(matches)
                    except Exception as e:
                        mtg_ref = future_to_meeting[future]
                        self._log(f"❌ Error in meeting {mtg_ref.get('meeting_number')}: {e}", logging.ERROR)

            self._executor = None
            if self._is_cancelled:
                self.finished.emit(False, "Search cancelled.")
                return

            self.stage_progress.emit("Stage 3/3: Aggregating results...", 100, 100)
            all_matched_tdocs.sort(key=lambda r: (r.get("WG", ""), r.get("Meeting", ""), r.get("TDoc", "")))

            summary = f"Audit complete! Found {len(all_matched_tdocs)} contributions across {total_meetings} meeting(s)."
            self._log(f"🏁 {summary}")
            self.results_ready.emit(all_matched_tdocs)
            self.finished.emit(True, summary)

        except Exception as e:
            err = f"Error during contribution search: {e}"
            logging.error(err, exc_info=True)
            self._log(f"❌ {err}", logging.ERROR)
            self.finished.emit(False, str(e))

    def _process_single_meeting(self, mtg: dict) -> list:
        if self._is_cancelled:
            return []

        mtg_num = mtg.get("meeting_number", "Unknown")
        wg_name = mtg.get("wg_name", "")
        mtg_id = mtg.get("mtg_id")
        tag = f"{wg_name} #{mtg_num}"
        folder_name = mtg.get("folder_name") or mtg_num

        local_mtg_dir = self.cache_dir / folder_name
        agenda_dir = local_mtg_dir / "Agenda"
        excel_file = None

        if agenda_dir.exists():
            excel_file = next(
                (f for f in agenda_dir.iterdir()
                 if ("tdoc_list_meeting_" in f.name.lower() or "tdocs_list_" in f.name.lower())
                 and f.name.endswith(".xlsx")),
                None
            )
            if not excel_file and mtg_id:
                fallback = agenda_dir / f"TDoc_List_Meeting_{mtg_id}.xlsx"
                if fallback.exists():
                    excel_file = fallback

        if self.bypass_cache or not excel_file:
            if not mtg_id:
                self._log(f"⚠️ [{tag}] Missing 3GPP portal ID; skipping download.", logging.WARNING)
                return []
            if self._is_cancelled:
                return []

            self._log(f"📥 [{tag}] Downloading TDocs workbook from 3GPP...")
            dl = TDocsDownloaderThread(mtg_id, local_mtg_dir, cancel_check=lambda: self._is_cancelled)
            dl.run()
            if self._is_cancelled:
                return []

            if agenda_dir.exists():
                excel_file = next(
                    (f for f in agenda_dir.iterdir()
                     if ("tdoc_list_meeting_" in f.name.lower() or "tdocs_list_" in f.name.lower())
                     and f.name.endswith(".xlsx")),
                    None
                )
                if not excel_file:
                    fallback = agenda_dir / f"TDoc_List_Meeting_{mtg_id}.xlsx"
                    if fallback.exists():
                        excel_file = fallback

        if not excel_file or not excel_file.exists():
            self._log(f"⚠️ [{tag}] No TDocs list available.", logging.WARNING)
            return []

        if self._is_cancelled:
            return []

        json_cache = str(excel_file) + ".json"
        if not self.bypass_cache and os.path.exists(json_cache):
            self._log(f"⚡ [{tag}] Loaded via JSON cache ({excel_file.name}).")
        else:
            self._log(f"⚙️ [{tag}] Ingesting spreadsheet ({excel_file.name})...")

        tdocs_data = TDocsParser.parse_tdocs_excel(str(excel_file))

        matched = []
        for row in tdocs_data:
            resolved_wi = self._extract_related_wis(row)
            row["Related WIs"] = resolved_wi
            row["WG"] = wg_name

            if self._matches_filter(row):
                row_copy = dict(row)
                row_copy["WG"] = wg_name
                row_copy["Meeting"] = tag
                row_copy["meeting_number"] = mtg_num
                row_copy["folder_name"] = folder_name
                row_copy["url_key"] = mtg.get("url_key", "")
                row_copy["docs_folder_url"] = mtg.get("docs_folder_url", "")
                row_copy["end_date"] = mtg.get("end_date", "")
                row_copy["start_date"] = mtg.get("start_date", "")
                row_copy["is_electronic"] = mtg.get("is_electronic", 0)
                row_copy["wg_name"] = wg_name
                row_copy["Related WIs"] = resolved_wi
                matched.append(row_copy)

        self._log(f"   ↳ [{tag}] {len(matched)} matching contribution(s).")
        return matched

    def _extract_related_wis(self, row: dict) -> str:
        # 1. Primary Check: Official 3GPP 'Related WIs' field[cite: 32]
        for key in ["Related WIs", "Related WI", "Related WI(s)", "Work Item", "Work Items", "WI", "WID"]:
            val = str(row.get(key, "")).strip()
            if val and val.lower() not in ["", "none", "unknown", "-"]:
                items = [w.strip() for w in re.split(r'[,;]+', val) if w.strip()]
                if items:
                    return ", ".join(items)

        # 2. Instant Fallback: Check 'Agenda item description' for parenthesized acronym[cite: 32]
        desc = str(row.get("Agenda item description", "")).strip()
        if desc:
            match = re.search(r'\(([A-Za-z0-9_ -]{3,35})\)', desc)
            if match:
                candidate = match.group(1).strip()
                if '_' in candidate or any(c.isdigit() for c in candidate):
                    return candidate

        # 3. Instant Fallback: Check Title for bracketed tags [WI_TAG][cite: 6]
        title = str(row.get("Title", "")).strip()
        if title:
            match = re.search(r'\[([A-Za-z0-9_ -]{3,35})\]', title)
            if match:
                candidate = match.group(1).strip()
                if candidate.lower() not in ["draft", "reply", "update", "revision", "discussion", "pcr", "cr", "ls", "none"]:
                    if '_' in candidate or any(c.isdigit() for c in candidate):
                        return candidate

        return ""

    @staticmethod
    def _normalize_ai(value: str) -> str:
        text = str(value or "").strip()
        text = re.sub(r"^AI\s*[:#-]?\s*", "", text, flags=re.IGNORECASE)
        return re.sub(r"\s+", "", text).lower()

    def _matches_filter(self, row: dict) -> bool:
        source_raw = str(row.get("Source", "")).strip()

        if self.target_companies:
            if not source_raw:
                return False

            if self.primary_only:
                parts = re.split(r'[,;]|\band\b', source_raw, maxsplit=1)
                first_author = parts[0] if parts else source_raw
                matched_companies = set(CompanySanitizer.get_matching_contributors(first_author))
            else:
                matched_companies = set(CompanySanitizer.get_matching_contributors(source_raw))

            intersect = matched_companies & self.target_companies
            if not intersect:
                return False

            row["Matched Company"] = ", ".join(sorted(list(intersect)))
        else:
            all_contributors = CompanySanitizer.get_matching_contributors(source_raw)
            row["Matched Company"] = ", ".join(all_contributors) if all_contributors else source_raw

        # WI and AI are separate facets. By default both active facets must match.
        # Optional OR mode is useful for queries spanning a WI plus specific AIs
        # that are not tagged with that WI (e.g. AIML WI OR AI 20.6.21).
        wi_hit = None
        if self.target_wis:
            wi_val = str(row.get("Related WIs", "")).strip().lower()
            title_val = str(row.get("Title", "")).strip().lower()
            desc_val = str(row.get("Agenda item description", "")).strip().lower()
            wi_hit = any((token in wi_val or token in title_val or token in desc_val) for token in self.target_wis)

        ai_hit = None
        if self.target_ais:
            row_ai = self._normalize_ai(row.get("Agenda Item", ""))
            ai_hit = row_ai in self.target_ais

        active_hits = [hit for hit in (wi_hit, ai_hit) if hit is not None]
        if active_hits:
            if self.wi_ai_or:
                if not any(active_hits):
                    return False
            elif not all(active_hits):
                return False

        # Preserve the actual reason this row passed the WI/AI facet. The table
        # uses this annotation for restrained match-origin highlighting.
        match_parts = []
        if wi_hit is True:
            match_parts.append("WI")
        if ai_hit is True:
            match_parts.append("AI")
        row["_Filter Match"] = " + ".join(match_parts)

        return True
