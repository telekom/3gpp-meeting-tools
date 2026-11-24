"""Topic rule evaluation, live meeting detection, TDoc loading, and Chair's Notes extraction mixin."""

from __future__ import annotations

import datetime
import json as _json
import logging
import os
import re
import shutil
import time
from pathlib import Path
from typing import Any, Dict, List, Optional, Set

from core.utils.paths import get_project_root, to_extended_path
from modules.meetings.core.cohort_resolver import get_effective_meeting_dates
from modules.meetings.core.exporters.export_manifest import ExportManifest
from modules.meetings.core.tdocs.chair_notes_agreements_parser import ChairNotesAgreementsParser
from modules.meetings.core.tdocs.tdocs_parser import TDocsParser

logger = logging.getLogger(__name__)


class TopicDossierDiscoveryMixin:
    """Evaluates topic rules, loads meeting TDoc tables/agreements, and prunes superseded files."""

    def _matches_topic_rule(self, wg: str, ai: str, text: str, wi: str) -> bool:
        """Evaluates whether an agreement item matches the topic rules."""
        if self.topic.rules.wi_tokens and wi:
            tokens = {t.lower() for t in re.split(r'[,;\s/&]+', wi) if t.strip()}
            if tokens.intersection(set(self.topic.rules.wi_tokens)):
                return True

        norm_ai = str(ai or "").strip().lower()
        if norm_ai and wg in self.topic.rules.ai_patterns:
            patterns = self.topic.rules.ai_patterns[wg]
            for pat in patterns:
                clean_pat = pat.lower().strip().rstrip("*")
                if norm_ai.startswith(clean_pat):
                    return True

        if self.topic.rules.keywords and text:
            text_lower = text.lower()
            for kw in self.topic.rules.keywords:
                if kw.lower() in text_lower:
                    return True

        return False

    def _resolve_category_folder(self, tdoc_type: str, status: str) -> str:
        t = (tdoc_type or "").lower()
        if "ls in" in t:
            return "LS_IN"
        elif "ls out" in t or "liaison" in t:
            return "LS_OUT"
        elif "pcr" in t or "pseudo" in t:
            return "pCR"
        elif "cr" in t:
            return "CR"
        elif "wf" in t or "way forward" in t:
            return "WF"
        return "Discussion"

    def _get_meeting_dir(self, folder_name: str) -> Optional[Path]:
        if not folder_name:
            return None
        from modules.meetings.core.settings import MeetingsSettings
        cache_dir = Path(MeetingsSettings().cache_dir)
        mtg_dir = cache_dir / folder_name
        if mtg_dir.is_dir():
            return mtg_dir
        from core.utils.paths import get_cache_root
        c_root = Path(get_cache_root())
        if (c_root / folder_name).is_dir():
            return c_root / folder_name
        root = get_project_root()
        if (root / "Data" / "Meetings" / folder_name).is_dir():
            return root / "Data" / "Meetings" / folder_name
        return mtg_dir

    @staticmethod
    def _is_meeting_live_for_refresh(mtg: Dict[str, Any], prev_mtg_state: Dict[str, Any]) -> bool:
        """Determines if an incremental run should probe FTP for fresher TDocs_by_Agenda / ChairNotes."""
        if not prev_mtg_state:
            return True
        if prev_mtg_state.get("lifecycle") in ("upcoming", "no_activity"):
            return True
        if int(prev_mtg_state.get("stub_count", 0) or 0) > 0:
            return True
        if int(prev_mtg_state.get("decision_count", 0) or 0) == 0:
            return True
        start_str = str(mtg.get("start_date", "") or "").strip()[:10]
        if start_str:
            try:
                m_date = datetime.date.fromisoformat(start_str)
                if (datetime.date.today() - m_date).days <= 45:
                    return True
            except Exception:
                pass
        return False

    @staticmethod
    def is_meeting_cache_stale_for_preview(
        mtg: Dict[str, Any],
        mtg_dir: Optional[Path],
        reference_date: Optional[datetime.date] = None,
        recent_days: int = 45,
        max_cache_age_hours: float = 6.0,
    ) -> bool:
        """
        Determines whether a meeting's local TDoc/ChairNotes cache should be live-refreshed
        for `View Considered Inputs...` (e.g. concluded meeting still holding a pre-meeting cache).
        """
        ref_date = reference_date or datetime.date.today()
        eff_start, eff_end = get_effective_meeting_dates(mtg)
        m_start: Optional[datetime.date] = None
        m_end: Optional[datetime.date] = None
        try:
            if eff_start:
                m_start = datetime.date.fromisoformat(eff_start[:10])
        except ValueError:
            pass
        try:
            if eff_end:
                m_end = datetime.date.fromisoformat(eff_end[:10])
        except ValueError:
            pass

        if m_start is None and m_end is None:
            return False
        if m_start is None:
            m_start = m_end
        if m_end is None:
            m_end = m_start

        if (ref_date - m_end).days > recent_days or (m_start - ref_date).days > recent_days:
            return False

        if not mtg_dir:
            return True

        newest_cache_mtime: Optional[float] = None
        agenda_dir = mtg_dir / "Agenda"
        if agenda_dir.is_dir():
            json_files = sorted(
                [jf for jf in agenda_dir.glob("*.xlsx.json") if jf.is_file()],
                key=lambda p: p.stat().st_mtime,
                reverse=True,
            )
            for jf in json_files:
                try:
                    if jf.stat().st_size <= 2:
                        continue
                    with open(jf, "r", encoding="utf-8") as f:
                        if _json.load(f):
                            newest_cache_mtime = jf.stat().st_mtime
                            break
                except Exception:
                    continue

        if newest_cache_mtime is None:
            return True

        cache_date = datetime.datetime.fromtimestamp(newest_cache_mtime).date()
        if ref_date > m_end and cache_date <= m_end:
            return True

        if m_start <= ref_date and (ref_date - m_end).days <= 14:
            now_ts = (
                time.time()
                if reference_date is None or ref_date == datetime.date.today()
                else datetime.datetime.combine(ref_date, datetime.time(12, 0)).timestamp()
            )
            if (now_ts - newest_cache_mtime) > max_cache_age_hours * 3600.0:
                return True

        return False

    @staticmethod
    def _prune_superseded_topic_docs(
        target_dir: Path,
        manifest: ExportManifest,
        meeting_rel_root: str,
        active_tids: Set[str],
        superseded_or_withdrawn: Set[str],
    ) -> int:
        """
        Removes obsolete topic-folder Markdown files and companion `{stem}_media/`
        directories for TDocs explicitly superseded by a newer revision or withdrawn,
        while preserving `_company_positions_analysis.md` and shared canonical caches.
        """
        prefix = f"{meeting_rel_root}/"
        active_upper = {str(t).strip().upper() for t in active_tids}
        superseded_upper = {str(t).strip().upper() for t in superseded_or_withdrawn}
        removed = 0
        for tid in list(manifest.tdocs.keys()):
            tid_upper = str(tid).strip().upper()
            if tid_upper in active_upper or tid_upper not in superseded_upper:
                continue
            info = manifest.tdocs.get(tid) or {}
            rel_path = str(info.get("rel_path", "")).replace("\\", "/")
            if not rel_path.startswith(prefix):
                continue
            abs_file = target_dir / rel_path
            ext_file = Path(to_extended_path(abs_file))
            try:
                if ext_file.is_file():
                    os.remove(ext_file)
                media_dir = abs_file.parent / f"{abs_file.stem}_media"
                ext_media = Path(to_extended_path(media_dir))
                if ext_media.is_dir():
                    shutil.rmtree(ext_media, ignore_errors=True)
            except Exception as e:
                logger.debug(f"[TopicDossierExporter] Prune cleanup warning for {rel_path}: {e}")
            manifest.tdocs.pop(tid, None)
            removed += 1
        return removed

    def _download_and_parse_tdocs_excel(self, mtg: Dict[str, Any], mtg_dir: Path) -> List[Dict[str, Any]]:
        """Downloads `TDocs_by_Agenda.xlsx` via `TDocsDownloaderThread` and parses the newest Excel."""
        # Must use the 3GPP portal `mtg_id` (NOT SQLite row `id` or `meeting_number`).
        mtg_id = str(mtg.get("mtg_id") or "").strip()
        if not mtg_id:
            logger.debug(f"[TDocsDownloader] Skipping auto-download for {mtg_dir.name}: missing 3GPP portal mtg_id")
            return []
        try:
            from modules.meetings.core.tdocs.tdocs_downloader import TDocsDownloaderThread
            dl = TDocsDownloaderThread(mtg_id, mtg_dir, cancel_check=lambda: self.is_cancelled)
            dl.run()
            agenda_dir = mtg_dir / "Agenda"
            if agenda_dir.is_dir():
                excels = sorted(
                    [ef for ef in agenda_dir.glob("*.xlsx") if ef.is_file() and not ef.name.startswith("~$")],
                    key=lambda p: p.stat().st_mtime,
                    reverse=True,
                )
                for ef in excels:
                    res = TDocsParser.parse_tdocs_excel(str(ef))
                    if res:
                        return res
        except Exception as e:
            logger.debug(f"[TDocsDownloader] Auto-download failed for {mtg_dir.name}: {e}")
        return []

    def _load_meeting_tdocs(
        self,
        mtg: Dict[str, Any],
        force_refresh: Optional[bool] = None,
    ) -> List[Dict[str, Any]]:
        folder_name = mtg.get("folder_name", "")
        mtg_dir = self._get_meeting_dir(folder_name)
        if not mtg_dir:
            return []

        eff_refresh = force_refresh if force_refresh is not None else getattr(self, "_force_live_refresh", False)
        if eff_refresh and self.allow_network_download and not self.is_cancelled:
            live_rows = self._download_and_parse_tdocs_excel(mtg, mtg_dir)
            if live_rows:
                return live_rows

        agenda_dir = mtg_dir / "Agenda"

        newest_valid_xlsx_mtime = 0.0
        if agenda_dir.is_dir():
            for ef in agenda_dir.glob("*.xlsx"):
                if not ef.is_file() or ef.name.startswith("~$"):
                    continue
                try:
                    with open(ef, "rb") as xf:
                        if xf.read(4) == b"PK\x03\x04":
                            newest_valid_xlsx_mtime = max(newest_valid_xlsx_mtime, ef.stat().st_mtime)
                except OSError:
                    pass

        if agenda_dir.is_dir():
            json_caches = sorted(
                [jf for jf in agenda_dir.glob("*.xlsx.json") if jf.is_file()],
                key=lambda p: p.stat().st_mtime,
                reverse=True,
            )
            for jf in json_caches:
                try:
                    if newest_valid_xlsx_mtime > 0.0 and jf.stat().st_mtime < newest_valid_xlsx_mtime:
                        continue
                    with open(jf, "r", encoding="utf-8") as f:
                        cached_rows = _json.load(f)
                    if cached_rows:
                        return cached_rows
                except Exception:
                    pass

        excel_candidates = []
        if agenda_dir.is_dir():
            excel_candidates.extend(list(agenda_dir.glob("*.xlsx")))
        if mtg_dir.is_dir():
            excel_candidates.extend(list(mtg_dir.glob("*.xlsx")))
        excel_candidates.sort(
            key=lambda p: p.stat().st_mtime if p.is_file() else 0.0,
            reverse=True,
        )

        for ef in excel_candidates:
            if ef.is_file() and not ef.name.startswith("~$"):
                try:
                    res = TDocsParser.parse_tdocs_excel(str(ef))
                    if res:
                        return res
                except Exception as e:
                    logger.debug(f"[TDocsParser] Excel parse failed for {ef}: {e}")

        if not eff_refresh and self.allow_network_download and not self.is_cancelled:
            live_rows = self._download_and_parse_tdocs_excel(mtg, mtg_dir)
            if live_rows:
                return live_rows

        return []

    def _download_and_scan_chair_notes(self, mtg: Dict[str, Any], mtg_dir: Path) -> List[Dict[str, Any]]:
        """Downloads latest RAN Chairman's Notes from 3GPP FTP and rescans agreements."""
        try:
            from modules.meetings.core.chair_notes_downloader import ChairNotesDownloaderThread
            from core.network.session import NetworkSession

            candidate_urls = []
            raw_url = str(mtg.get("url_key", "")).strip()
            if raw_url:
                full = raw_url if raw_url.startswith("http") else f"https://www.3gpp.org/ftp/{raw_url.lstrip('/')}"
                candidate_urls.append(full.rstrip("/"))
            docs_url = str(mtg.get("docs_folder_url", "")).strip()
            if docs_url:
                full_docs = docs_url if docs_url.startswith("http") else f"https://www.3gpp.org/ftp/{docs_url.lstrip('/')}"
                base = re.sub(r'/Docs/?$', '', full_docs.rstrip("/"), flags=re.IGNORECASE)
                if base not in candidate_urls:
                    candidate_urls.append(base)

            if candidate_urls:
                session = NetworkSession.get_instance()
                base_url, html_text = ChairNotesDownloaderThread.locate_chair_notes_dir(
                    session, candidate_urls, cancel_check=lambda: self.is_cancelled
                )
                if base_url and html_text and not self.is_cancelled:
                    all_remote = ChairNotesDownloaderThread.parse_file_list_static(html_text, base_url)
                    selected_remote = ChairNotesDownloaderThread.select_best_chair_notes_files(
                        all_remote, only_final_or_latest=True
                    )
                    if selected_remote:
                        cn_dir = mtg_dir / "Agenda" / "ChairNotes"
                        cn_dir.mkdir(parents=True, exist_ok=True)
                        downloaded_names = ChairNotesDownloaderThread.download_files_sync(
                            session=session,
                            files_to_download=selected_remote,
                            target_dir=cn_dir,
                            cancel_check=lambda: self.is_cancelled,
                        )
                        if downloaded_names:
                            for fname in downloaded_names:
                                f_path = cn_dir / fname
                                if f_path.suffix.lower() == ".zip":
                                    ChairNotesDownloaderThread.extract_chair_notes_zip(f_path, cn_dir)
                            _, agreements, _ = ChairNotesAgreementsParser.scan_and_cache_chair_notes(
                                meeting_dir=mtg_dir,
                                mtg_info=mtg,
                                force_rescan=True,
                            )
                            if agreements:
                                return agreements
        except Exception as e:
            logger.debug(f"[TopicDossierExporter] Auto-download ChairNotes failed for {mtg_dir.name}: {e}")
        return []

    def _load_meeting_agreements(
        self,
        mtg: Dict[str, Any],
        force_refresh: Optional[bool] = None,
    ) -> List[Dict[str, Any]]:
        # Chairman's Notes text agreements are strictly RAN-only; SA/CT triage notes must not become Text Agreements
        wg_upper = str(mtg.get("wg_name", "")).strip().upper()
        if "RAN" not in wg_upper:
            return []

        folder_name = mtg.get("folder_name", "")
        mtg_dir = self._get_meeting_dir(folder_name)
        if not mtg_dir:
            return []

        eff_refresh = force_refresh if force_refresh is not None else getattr(self, "_force_live_refresh", False)
        if eff_refresh and self.allow_network_download and not self.is_cancelled:
            live_agreements = self._download_and_scan_chair_notes(mtg, mtg_dir)
            if live_agreements:
                return live_agreements

        from modules.meetings.core.tdocs.chair_notes_agreements_parser import (
            CHAIR_NOTES_CACHE_VERSION,
        )
        import json as _json

        agenda_dir = mtg_dir / "Agenda"
        cn_cache = agenda_dir / "chair_notes_cache.json"
        if cn_cache.is_file():
            try:
                with open(cn_cache, "r", encoding="utf-8") as f:
                    data = _json.load(f)
                if data.get("version", 1) >= CHAIR_NOTES_CACHE_VERSION:
                    return data.get("agreements", [])
            except Exception:
                pass

        try:
            _, agreements, _ = ChairNotesAgreementsParser.scan_and_cache_chair_notes(
                meeting_dir=mtg_dir,
                mtg_info=mtg,
                force_rescan=True,
            )
            if agreements:
                return agreements
        except Exception as e:
            logger.debug(f"[TopicDossierExporter] scan_and_cache_chair_notes error: {e}")

        cn_dir = agenda_dir / "ChairNotes"
        if cn_dir.is_dir():
            for docx_f in cn_dir.glob("*.docx"):
                if not docx_f.name.startswith("~$"):
                    try:
                        res = ChairNotesAgreementsParser.parse_docx_with_agreements(
                            docx_path=str(docx_f),
                            mtg_info=mtg,
                        )
                        if res and len(res) >= 2 and res[1]:
                            return res[1]
                    except Exception:
                        pass

        if not eff_refresh and self.allow_network_download and not self.is_cancelled:
            live_agreements = self._download_and_scan_chair_notes(mtg, mtg_dir)
            if live_agreements:
                return live_agreements

        return []
