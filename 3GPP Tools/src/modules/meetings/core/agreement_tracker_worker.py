# --- File: src/modules/meetings/core/agreement_tracker_worker.py ---
"""
Cross-Meeting Study Item & Topic Agreement Tracker Worker (`AgreementTrackerWorker`).

Asynchronously scans selected 3GPP Working Groups (e.g., RAN2, RAN3, RAN1, SA2) across a
date range, automatically locating and downloading Final/EOM Chairman's Notes from
`/Inbox/Chairs_Notes` or `/Inbox/Chair_Notes` when needed, parsing structured
Agreements, Working Assumptions, Conclusions, FFS/Open Issues, and Agreed TPs/pCRs/CRs,
enriching them with official Excel TDoc metadata & Agenda Item mappings, and filtering
by Study Item / Work Item, Agenda Item prefix, and topic keywords.
"""
import csv
import logging
import os
import re
import subprocess
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from typing import Dict, List, Optional, Set, Tuple

from PyQt5.QtCore import QThread, pyqtSignal

from core.network.session import NetworkSession
from core.utils.paths import to_extended_path
from modules.word_tools.core.libreoffice_converter import (
    convert_doc_to_docx_libreoffice,
    is_libreoffice_available,
)
from modules.meetings.core.chair_notes_downloader import ChairNotesDownloaderThread
from modules.meetings.core.meetings_db import MeetingsDatabase
from modules.meetings.core.tdocs.chair_notes_agreements_parser import (
    ChairNotesAgreementsParser,
    _resolve_inherited_field,
)
from modules.meetings.core.tdocs.tdocs_downloader import TDocsDownloaderThread
from modules.meetings.core.tdocs.tdocs_parser import TDocsParser
from modules.meetings.ui.tdocs_models import TDocsTableModel

logger = logging.getLogger(__name__)


class AgreementTrackerWorker(QThread):
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

    @staticmethod
    def _normalize_ai(value: str) -> str:
        text = str(value or "").strip()
        text = re.sub(r"^AI\s*[:#-]?\s*", "", text, flags=re.IGNORECASE)
        return re.sub(r"\s+", "", text).lower().rstrip(".")

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
            # Sort chronologically by start_date, WG, meeting_number, AI
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

    @staticmethod
    def _ai_sort_key(ai_str: str) -> Tuple:
        if not ai_str:
            return (9999,)
        parts = re.findall(r'\d+', str(ai_str))
        if parts:
            return tuple(int(p) for p in parts)
        return (9998, str(ai_str))

    def _process_single_meeting(self, mtg: dict) -> List[dict]:
        if self.is_cancelled:
            return []

        mtg_num = str(mtg.get("meeting_number", "Unknown")).strip()
        wg_name = str(mtg.get("wg_name", "")).strip()
        mtg_id = mtg.get("mtg_id")
        tag = f"{wg_name} #{mtg_num}"
        folder_name = mtg.get("folder_name") or mtg_num

        local_mtg_dir = self.cache_dir / folder_name
        agenda_dir = local_mtg_dir / "Agenda"
        os.makedirs(to_extended_path(agenda_dir), exist_ok=True)

        # 1. Load or download TDocs Excel list (to enrich TDoc -> WI/Title/Source and extract agreed pCRs/CRs)
        tdocs_by_id, ai_wi_from_excel, ai_title_from_excel = self._load_or_fetch_tdocs_excel(
            mtg, local_mtg_dir, agenda_dir, tag
        )
        if self.is_cancelled:
            return []

        # 2. Also load agenda.csv if present in Agenda/
        ai_wi_from_csv, ai_title_from_csv = self._load_agenda_csv_map(agenda_dir)
        combined_ai_wis = {**ai_wi_from_excel, **ai_wi_from_csv}
        combined_ai_titles = {**ai_title_from_excel, **ai_title_from_csv}

        # 3. Load or download/parse Chairman's Notes agreements
        chair_agenda_data, chair_agreements = self._load_or_fetch_chair_notes(
            mtg, local_mtg_dir, agenda_dir, tag
        )
        if self.is_cancelled:
            return []

        # Merge any AI -> WI / Title discovered from chair_agenda_data
        for tid, tinfo in (chair_agenda_data or {}).items():
            ai_val = str(tinfo.get("Agenda Item", "")).strip()
            wi_val = str(tinfo.get("Related WIs", "")).strip()
            sub_val = str(tinfo.get("Sub-Item", "")).strip()
            if ai_val and ai_val != "N/A":
                if wi_val and ai_val not in combined_ai_wis:
                    combined_ai_wis[ai_val] = wi_val
                if sub_val and ai_val not in combined_ai_titles:
                    combined_ai_titles[ai_val] = sub_val

        # 4. Enrich Chair Notes agreements with Excel TDoc metadata & inherited AI metadata
        candidates: List[dict] = []
        seen_tdoc_agreements: Set[str] = set()

        for ag in (chair_agreements or []):
            is_recap = bool(ag.get("is_previous_meeting_recap", False))
            if is_recap and not self.include_prev_recaps:
                continue

            kind = str(ag.get("kind", "Agreement")).strip()
            if not self._kind_enabled(kind):
                continue

            ai_num = str(ag.get("ai", "")).strip()
            tdoc_id = str(ag.get("tdoc_id", "")).strip().upper()
            wi = str(ag.get("wi", "")).strip()
            ai_title = str(ag.get("ai_title", "")).strip()
            release = str(ag.get("release", "")).strip()
            text = str(ag.get("text", "")).strip()
            source_file = str(ag.get("source_file", "")).strip()

            tdoc_row = tdocs_by_id.get(tdoc_id, {}) if tdoc_id else {}
            chair_tdoc_row = chair_agenda_data.get(tdoc_id, {}) if tdoc_id else {}

            tdoc_title = str(tdoc_row.get("Title") or chair_tdoc_row.get("Title") or "").strip()
            tdoc_source = str(tdoc_row.get("Source") or chair_tdoc_row.get("Source") or "").strip()
            tdoc_type = str(tdoc_row.get("Type") or chair_tdoc_row.get("Type") or "").strip()
            tdoc_status = str(tdoc_row.get("TDoc Status") or chair_tdoc_row.get("Result") or "").strip()

            if not ai_num and tdoc_row.get("Agenda item"):
                ai_num = str(tdoc_row.get("Agenda item", "")).strip()
            if not wi:
                wi = (
                    str(tdoc_row.get("Related WIs", "")).strip()
                    or _resolve_inherited_field(ai_num, combined_ai_wis)
                )
            if not ai_title:
                ai_title = (
                    _resolve_inherited_field(ai_num, combined_ai_titles)
                    or str(tdoc_row.get("Agenda item description", "")).strip()
                )
            if not release:
                release = str(tdoc_row.get("Release", "")).strip()

            if tdoc_id:
                seen_tdoc_agreements.add(tdoc_id)

            record = {
                "WG": wg_name,
                "Meeting": tag,
                "meeting_number": mtg_num,
                "start_date": str(mtg.get("start_date", "")).strip(),
                "end_date": str(mtg.get("end_date", "")).strip(),
                "folder_name": folder_name,
                "url_key": mtg.get("url_key", ""),
                "docs_folder_url": mtg.get("docs_folder_url", ""),
                "is_electronic": mtg.get("is_electronic", 0),
                "wg_name": wg_name,
                "AI": ai_num,
                "AI Title": ai_title,
                "Work Item": wi,
                "Release": release,
                "Kind": f"{kind} (Prev Mtg Recap)" if is_recap else kind,
                "Agreement Text": text,
                "TDoc": tdoc_id,
                "TDoc Title": tdoc_title,
                "Source": tdoc_source,
                "TDoc Type": tdoc_type,
                "TDoc Status": tdoc_status,
                "Chair Notes File": source_file,
                "is_previous_meeting_recap": is_recap,
            }
            if self._matches_filter(record):
                candidates.append(record)

        # 5. Also include Agreed pCRs / CRs / TPs from official Excel TDoc list (so SA2 and agreed CRs in RAN2/RAN3 are included!)
        if self.include_agreed_tdocs:
            for tid, row in tdocs_by_id.items():
                if tid in seen_tdoc_agreements:
                    continue
                status_raw = str(row.get("TDoc Status", "") or row.get("Result", "")).strip()
                type_raw = str(row.get("Type", "")).strip()
                for_raw = str(row.get("For", "")).strip()
                low_status = status_raw.lower()
                low_type = type_raw.lower()
                low_for = for_raw.lower()

                is_agreed_status = any(
                    s in low_status for s in ("agreed", "approved", "endorsed")
                ) and not any(
                    neg in low_status for neg in ("not agreed", "not approved", "not endorsed")
                )
                if not is_agreed_status:
                    continue

                # Focus on technical outputs (pCR, CR, TP, Draft TR/TS, LS out, or items for Agreement/Approval)
                is_substantive_type = (
                    any(t in low_type for t in ("pcr", "cr", "tp", "draft", "tr", "ts", "ls out", "lsout", "sid", "wid"))
                    or any(f in low_for for f in ("agreement", "approval", "decision"))
                    or wg_name.upper().startswith("SA")
                )
                if not is_substantive_type:
                    continue

                ai_num = str(row.get("Agenda item", "") or row.get("Agenda Item", "")).strip()
                wi = str(row.get("Related WIs", "")).strip() or _resolve_inherited_field(ai_num, combined_ai_wis)
                ai_title = (
                    str(row.get("Agenda item description", "")).strip()
                    or _resolve_inherited_field(ai_num, combined_ai_titles)
                )
                tdoc_title = str(row.get("Title", "")).strip()
                tdoc_source = str(row.get("Source", "")).strip()
                release = str(row.get("Release", "")).strip()
                chair_comm = ""
                if tid in chair_agenda_data:
                    chair_comm = str(chair_agenda_data[tid].get("Comments", "")).strip()

                text_summary = f"[{status_raw}] {tid}: {tdoc_title}"
                if chair_comm:
                    text_summary += f"\n\nChair Notes Remarks:\n{chair_comm}"

                record = {
                    "WG": wg_name,
                    "Meeting": tag,
                    "meeting_number": mtg_num,
                    "start_date": str(mtg.get("start_date", "")).strip(),
                    "end_date": str(mtg.get("end_date", "")).strip(),
                    "folder_name": folder_name,
                    "url_key": mtg.get("url_key", ""),
                    "docs_folder_url": mtg.get("docs_folder_url", ""),
                    "is_electronic": mtg.get("is_electronic", 0),
                    "wg_name": wg_name,
                    "AI": ai_num,
                    "AI Title": ai_title,
                    "Work Item": wi,
                    "Release": release,
                    "Kind": "Agreed TDoc / TP / pCR",
                    "Agreement Text": text_summary,
                    "TDoc": tid,
                    "TDoc Title": tdoc_title,
                    "Source": tdoc_source,
                    "TDoc Type": type_raw,
                    "TDoc Status": status_raw,
                    "Chair Notes File": "TDoc_List_Meeting.xlsx",
                    "is_previous_meeting_recap": False,
                }
                if self._matches_filter(record):
                    candidates.append(record)

        self._log(f"   ↳ [{tag}] {len(candidates)} matching agreement(s)/item(s).")
        return candidates

    def _kind_enabled(self, kind: str) -> bool:
        low = kind.lower()
        if "ffs" in low or "open issue" in low:
            return self.include_ffs
        if "agreed tdoc" in low or "pcr" in low or "tp" in low:
            return self.include_agreed_tdocs
        return self.include_agreements

    def _matches_filter(self, record: dict) -> bool:
        # 1. WI Filter (supports multi-token base extraction just like TDocsTableModel!)
        wi_hit = None
        if self.target_wis:
            rec_wi = str(record.get("Work Item", "")).strip()
            rec_ai_title = str(record.get("AI Title", "")).strip()
            rec_tdoc_title = str(record.get("TDoc Title", "")).strip()
            combined_wi_haystack = f"{rec_wi} {rec_ai_title} {rec_tdoc_title}".lower()
            rec_base_tokens = TDocsTableModel._extract_base_tokens(rec_wi)

            wi_hit = bool(rec_base_tokens & self.target_wi_base_tokens) or any(
                w.lower() in combined_wi_haystack for w in self.target_wis
            )

        # 2. AI Filter (supports exact match AND prefix match, e.g., `9.3` matches `9.3`, `9.3.1`, `9.3.1.2`)
        ai_hit = None
        if self.target_ais:
            rec_ai = self._normalize_ai(record.get("AI", ""))
            if not rec_ai:
                ai_hit = False
            else:
                ai_hit = any(
                    rec_ai == target_ai or rec_ai.startswith(f"{target_ai}.")
                    for target_ai in self.target_ais
                )

        active_facets = [h for h in (wi_hit, ai_hit) if h is not None]
        if active_facets:
            if self.wi_ai_or:
                if not any(active_facets):
                    return False
            elif not all(active_facets):
                return False

        # 3. Keyword / Topic Query Filter
        if self.keyword_query:
            haystack = " ".join([
                str(record.get("Agreement Text", "")),
                str(record.get("AI Title", "")),
                str(record.get("Work Item", "")),
                str(record.get("TDoc", "")),
                str(record.get("TDoc Title", "")),
                str(record.get("Source", "")),
                str(record.get("AI", "")),
            ]).lower()
            # Support comma/OR-separated terms or space-separated AND tokens
            if "," in self.keyword_query or " OR " in self.keyword_query:
                or_terms = [
                    t.strip().lower()
                    for t in re.split(r',|\bOR\b', self.keyword_query)
                    if t.strip()
                ]
                if or_terms and not any(term in haystack for term in or_terms):
                    return False
            else:
                and_terms = [t.strip().lower() for t in self.keyword_query.split() if t.strip()]
                if and_terms and not all(term in haystack for term in and_terms):
                    return False

        return True

    def _load_or_fetch_tdocs_excel(
        self, mtg: dict, local_mtg_dir: Path, agenda_dir: Path, tag: str
    ) -> Tuple[Dict[str, dict], Dict[str, str], Dict[str, str]]:
        """Loads or downloads the official TDoc Excel list and builds lookup dictionaries."""
        mtg_id = mtg.get("mtg_id")
        excel_file = None
        if agenda_dir.exists():
            excel_file = next(
                (
                    f for f in agenda_dir.iterdir()
                    if ("tdoc_list_meeting_" in f.name.lower() or "tdocs_list_" in f.name.lower())
                    and f.name.endswith(".xlsx")
                ),
                None,
            )
            if not excel_file and mtg_id:
                fallback = agenda_dir / f"TDoc_List_Meeting_{mtg_id}.xlsx"
                if fallback.exists():
                    excel_file = fallback

        if (self.bypass_cache or not excel_file) and self.auto_download_missing and mtg_id:
            if not self.is_cancelled:
                self._log(f"📥 [{tag}] Downloading TDocs list workbook...")
                dl = TDocsDownloaderThread(mtg_id, local_mtg_dir, cancel_check=lambda: self.is_cancelled)
                dl.run()
                if agenda_dir.exists():
                    excel_file = next(
                        (
                            f for f in agenda_dir.iterdir()
                            if ("tdoc_list_meeting_" in f.name.lower() or "tdocs_list_" in f.name.lower())
                            and f.name.endswith(".xlsx")
                        ),
                        None,
                    )

        tdocs_by_id: Dict[str, dict] = {}
        ai_wis: Dict[str, str] = {}
        ai_titles: Dict[str, str] = {}

        if not excel_file or not excel_file.exists():
            return tdocs_by_id, ai_wis, ai_titles

        try:
            rows = TDocsParser.parse_tdocs_excel(str(excel_file))
            for r in rows:
                tid = str(r.get("TDoc", "")).strip().upper()
                if not tid:
                    continue
                wi = self._extract_related_wis_from_row(r)
                r["Related WIs"] = wi
                tdocs_by_id[tid] = r
                ai_val = str(r.get("Agenda item", "") or r.get("Agenda Item", "")).strip()
                desc_val = str(r.get("Agenda item description", "")).strip()
                if ai_val and ai_val != "N/A":
                    if wi and ai_val not in ai_wis:
                        ai_wis[ai_val] = wi
                    if desc_val and ai_val not in ai_titles:
                        ai_titles[ai_val] = desc_val
        except Exception as e:
            logger.warning(f"[AgreementTrackerWorker] Failed parsing Excel for {tag}: {e}")

        return tdocs_by_id, ai_wis, ai_titles

    @staticmethod
    def _extract_related_wis_from_row(row: dict) -> str:
        for key in ["Related WIs", "Related WI", "Related WI(s)", "Work Item", "Work Items", "WI", "WID"]:
            val = str(row.get(key, "")).strip()
            if val and val.lower() not in ["", "none", "unknown", "-"]:
                items = [w.strip() for w in re.split(r'[,;]+', val) if w.strip()]
                if items:
                    return ", ".join(items)
        desc = str(row.get("Agenda item description", "")).strip()
        if desc:
            m = re.search(r'\(([A-Za-z0-9_ -]{3,35})\)', desc)
            if m:
                cand = m.group(1).strip()
                if "_" in cand or any(c.isdigit() for c in cand):
                    return cand
        return ""

    @staticmethod
    def _load_agenda_csv_map(agenda_dir: Path) -> Tuple[Dict[str, str], Dict[str, str]]:
        ai_wis: Dict[str, str] = {}
        ai_titles: Dict[str, str] = {}
        csv_path = agenda_dir / "agenda.csv"
        if not os.path.exists(to_extended_path(csv_path)):
            return ai_wis, ai_titles
        try:
            with open(to_extended_path(csv_path), "r", encoding="utf-8", errors="ignore") as f:
                reader = csv.DictReader(f)
                for row in reader:
                    ai = str(row.get("ai_number", "")).strip()
                    desc = str(row.get("description", "")).strip()
                    acr = str(row.get("acronym", "")).strip()
                    if ai:
                        if acr:
                            ai_wis[ai] = acr
                        if desc:
                            ai_titles[ai] = desc
        except Exception:
            pass
        return ai_wis, ai_titles

    def _load_or_fetch_chair_notes(
        self, mtg: dict, local_mtg_dir: Path, agenda_dir: Path, tag: str
    ) -> Tuple[Dict[str, dict], List[dict]]:
        """
        Loads cached Chair Notes agreements from `Agenda/chair_notes_cache.json`,
        or parses local files in `Agenda/ChairNotes/`, or automatically locates and downloads
        the Final/EOM Chairman's Notes from the 3GPP FTP server (`/Inbox/Chairs_Notes`).
        """
        if not self.bypass_cache:
            cached_agenda, cached_agreements = ChairNotesAgreementsParser.load_cached_chair_notes_data(
                local_mtg_dir
            )
            if cached_agreements or cached_agenda:
                self._log(
                    f"⚡ [{tag}] Loaded {len(cached_agreements)} agreement(s) from local Chair Notes cache."
                )
                return cached_agenda, cached_agreements

        chair_notes_dir = agenda_dir / "ChairNotes"
        os.makedirs(to_extended_path(chair_notes_dir), exist_ok=True)

        # Check if there are already local files in Agenda/ChairNotes/
        local_files: List[Path] = []
        if not self.bypass_cache and os.path.exists(to_extended_path(chair_notes_dir)):
            for f in chair_notes_dir.iterdir():
                if f.is_file() and f.suffix.lower() in (".docx", ".doc", ".zip", ".htm", ".html") and not f.name.startswith("~$"):
                    local_files.append(f)

        # If no local files (or bypass_cache=True) and auto_download_missing is enabled, fetch from FTP!
        if (not local_files or self.bypass_cache) and self.auto_download_missing and not self.is_cancelled:
            session = NetworkSession.get_instance()
            candidate_urls = self._build_meeting_candidate_urls(mtg)
            base_url, html_text = ChairNotesDownloaderThread.locate_chair_notes_dir(
                session,
                candidate_urls,
                progress_callback=lambda m: self._log(f"   [{tag}] {m}"),
                cancel_check=lambda: self.is_cancelled,
            )
            if base_url and html_text and not self.is_cancelled:
                all_remote = ChairNotesDownloaderThread.parse_file_list_static(html_text, base_url)
                selected_remote = ChairNotesDownloaderThread.select_best_chair_notes_files(
                    all_remote, only_final_or_latest=True
                )
                if selected_remote:
                    self._log(
                        f"📥 [{tag}] Downloading {len(selected_remote)} final/latest Chair Notes file(s)..."
                    )
                    downloaded_names = ChairNotesDownloaderThread.download_files_sync(
                        session=session,
                        files_to_download=selected_remote,
                        target_dir=chair_notes_dir,
                        progress_callback=lambda m: self._log(f"   [{tag}] {m}"),
                        cancel_check=lambda: self.is_cancelled,
                    )
                    if downloaded_names:
                        local_files = [chair_notes_dir / fname for fname in downloaded_names]

        if not local_files:
            return {}, []

        # Parse all local files (unzipping .zip or converting .doc if needed)
        merged_agenda: Dict[str, dict] = {}
        merged_agreements: List[dict] = []

        for f_path in local_files:
            if self.is_cancelled:
                break
            ext = f_path.suffix.lower()
            parse_targets: List[Path] = []
            if ext == ".zip":
                extracted_names = ChairNotesDownloaderThread.extract_chair_notes_zip(f_path, chair_notes_dir)
                parse_targets.extend([chair_notes_dir / fn for fn in extracted_names])
            elif ext == ".doc":
                converted = self._convert_doc_to_docx(f_path, chair_notes_dir)
                if converted:
                    parse_targets.append(converted)
            else:
                parse_targets.append(f_path)

            for target in parse_targets:
                if self.is_cancelled:
                    break
                t_ext = target.suffix.lower()
                if t_ext == ".docx":
                    ag_data, ag_list = ChairNotesAgreementsParser.parse_docx_with_agreements(
                        str(target), mtg_info=mtg
                    )
                elif t_ext in (".htm", ".html"):
                    ag_data = TDocsParser.parse_tdocs_by_agenda(str(target))
                    ag_list = []
                else:
                    continue

                for tid, info in ag_data.items():
                    merged_agenda[tid] = info
                merged_agreements.extend(ag_list)

        if merged_agenda or merged_agreements:
            ChairNotesAgreementsParser.save_cached_chair_notes_data(
                local_mtg_dir, merged_agenda, merged_agreements
            )

        return merged_agenda, merged_agreements

    @staticmethod
    def _build_meeting_candidate_urls(mtg: dict) -> List[str]:
        urls = []
        raw_url = str(mtg.get("url_key", "")).strip()
        if raw_url:
            full = raw_url if raw_url.startswith("http") else f"https://www.3gpp.org/ftp/{raw_url.lstrip('/')}"
            urls.append(full.rstrip("/"))
        docs_url = str(mtg.get("docs_folder_url", "")).strip()
        if docs_url:
            full_docs = docs_url if docs_url.startswith("http") else f"https://www.3gpp.org/ftp/{docs_url.lstrip('/')}"
            base = re.sub(r'/Docs/?$', '', full_docs.rstrip("/"), flags=re.IGNORECASE)
            if base not in urls:
                urls.append(base)
        return urls

    def _convert_doc_to_docx(self, doc_path: Path, out_dir: Path) -> Optional[Path]:
        target_docx = out_dir / f"{doc_path.stem}.docx"
        if os.path.exists(to_extended_path(target_docx)):
            return target_docx
        if not is_libreoffice_available():
            return None
        try:
            converted = convert_doc_to_docx_libreoffice(str(doc_path), str(out_dir))
            if converted and os.path.exists(to_extended_path(converted)):
                return Path(converted)
        except Exception as e:
            logger.warning(f"[AgreementTrackerWorker] Failed converting .doc '{doc_path.name}': {e}")
        return None
