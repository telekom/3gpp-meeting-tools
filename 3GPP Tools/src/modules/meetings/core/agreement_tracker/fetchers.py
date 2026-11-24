"""Local cache loading, FTP download, .doc conversion, and Chairman's Notes fetchers."""

from __future__ import annotations

import logging
import os
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from core.network.session import NetworkSession
from core.utils.paths import to_extended_path
from modules.meetings.core.chair_notes_downloader import ChairNotesDownloaderThread
from modules.meetings.core.tdocs.chair_notes_agreements_parser import (
    ChairNotesAgreementsParser,
)
from modules.meetings.core.tdocs.tdocs_downloader import TDocsDownloaderThread
from modules.meetings.core.tdocs.tdocs_parser import TDocsParser
from modules.word_tools.core.libreoffice_converter import (
    convert_doc_to_docx_libreoffice,
    is_libreoffice_available,
)

logger = logging.getLogger(__name__)


class AgreementFetchersMixin:
    """Mixin fetching/loading TDocs workbooks and Chairman's Notes packages per meeting."""

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

        local_files: List[Path] = []
        if not self.bypass_cache and os.path.exists(to_extended_path(chair_notes_dir)):
            for f in chair_notes_dir.iterdir():
                if f.is_file() and f.suffix.lower() in (".docx", ".doc", ".zip", ".htm", ".html") and not f.name.startswith("~$"):
                    local_files.append(f)

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
