"""Persistence cache and meeting-folder scanner mixin for Chairman's Notes."""

from __future__ import annotations

import json
import logging
import os
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from core.utils.paths import to_extended_path
from .text_and_results import CHAIR_NOTES_CACHE_VERSION

logger = logging.getLogger(__name__)


class ChairNotesCacheAndScanMixin:
    """Manages Agenda/chair_notes_cache.json persistence and multi-format scanning."""

    @classmethod
    def save_cached_chair_notes_data(
        cls,
        meeting_dir: Path,
        agenda_data: Dict[str, dict],
        agreements: List[dict],
        overwrite: bool = False,
    ):
        """
        Persists merged `agenda_data` and `agreements` into `Agenda/chair_notes_cache.json`
        inside `meeting_dir` so opening `TDocsWindow` or running `PostMeetingExporterThread`
        or `AgreementTrackerWorker` can load them instantaneously.
        """
        if not meeting_dir:
            return
        try:
            agenda_dir = Path(meeting_dir) / "Agenda"
            os.makedirs(to_extended_path(agenda_dir), exist_ok=True)
            cache_path = agenda_dir / "chair_notes_cache.json"

            existing_agenda = {}
            existing_agreements = []
            if not overwrite and os.path.exists(to_extended_path(cache_path)):
                try:
                    with open(to_extended_path(cache_path), "r", encoding="utf-8") as f:
                        loaded = json.load(f)
                        # Invalidate older cache versions so upgraded parsing replaces truncated records
                        if loaded.get("version", 1) >= CHAIR_NOTES_CACHE_VERSION:
                            existing_agenda = loaded.get("agenda_data", {})
                            existing_agreements = loaded.get("agreements", [])
                except Exception:
                    pass

            # Merge agenda_data
            for tid, info in (agenda_data or {}).items():
                if tid in existing_agenda:
                    old = existing_agenda[tid]
                    for k, v in info.items():
                        if v:
                            old[k] = v
                else:
                    existing_agenda[tid] = info

            # Merge agreements (deduplicate by (ai, kind, text[:120]))
            seen_keys = {
                (a.get("ai", ""), a.get("kind", ""), a.get("text", "")[:120])
                for a in existing_agreements
            }
            for ag in (agreements or []):
                key = (ag.get("ai", ""), ag.get("kind", ""), ag.get("text", "")[:120])
                if key not in seen_keys:
                    seen_keys.add(key)
                    existing_agreements.append(ag)

            payload = {
                "version": CHAIR_NOTES_CACHE_VERSION,
                "agenda_data": existing_agenda,
                "agreements": existing_agreements,
            }
            with open(to_extended_path(cache_path), "w", encoding="utf-8") as f:
                json.dump(payload, f, indent=2, ensure_ascii=False)
        except Exception as e:
            logger.warning(f"[ChairNotesAgreementsParser] Failed to save chair notes cache: {e}")

    @staticmethod
    def load_cached_chair_notes_data(meeting_dir: Path, min_version: int = 1) -> Tuple[Dict[str, dict], List[dict]]:
        """Loads cached `(agenda_data, agreements)` from `Agenda/chair_notes_cache.json` if available."""
        if not meeting_dir:
            return {}, []
        cache_path = Path(meeting_dir) / "Agenda" / "chair_notes_cache.json"
        if not os.path.exists(to_extended_path(cache_path)):
            return {}, []
        try:
            with open(to_extended_path(cache_path), "r", encoding="utf-8") as f:
                loaded = json.load(f)
                if min_version > 1 and loaded.get("version", 1) < min_version:
                    return {}, []
                return loaded.get("agenda_data", {}), loaded.get("agreements", [])
        except Exception as e:
            logger.warning(f"[ChairNotesAgreementsParser] Failed to load chair notes cache: {e}")
            return {}, []

    @classmethod
    def parse_chair_notes_file(
        cls,
        file_path: Path,
        mtg_info: Optional[dict] = None,
        agenda_map: Optional[dict] = None,
    ) -> List[dict]:
        """Convenience method parsing a single Chairman's Notes file and returning extracted agreements."""
        file_path = Path(file_path)
        if not file_path.exists() or file_path.name.startswith("~$"):
            return []
        ext = file_path.suffix.lower()
        if ext == ".docx":
            _, agreements = cls.parse_docx_with_agreements(str(file_path), mtg_info=mtg_info, agenda_map=agenda_map)
            return agreements
        return []

    @classmethod
    def scan_and_cache_chair_notes(
        cls,
        meeting_dir: Path,
        mtg_info: Optional[dict] = None,
        agenda_map: Optional[dict] = None,
        force_rescan: bool = False,
    ) -> Tuple[Dict[str, dict], List[dict], int]:
        """
        Scans local Chairman's Notes files (.docx, .doc, .zip, .htm, .html) across
        `meeting_dir/Agenda/ChairNotes/` and `meeting_dir/Agenda/`, parses them for
        TDoc remarks and structured agreements, persists them to `Agenda/chair_notes_cache.json`,
        and returns `(agenda_data, agreements, files_scanned_count)`.

        If `force_rescan` is False and `chair_notes_cache.json` exists and is newer than
        all candidate notes files, returns the cached data immediately.
        """
        if not meeting_dir:
            return {}, [], 0

        meeting_dir = Path(meeting_dir)
        agenda_dir = meeting_dir / "Agenda"
        chair_notes_dir = agenda_dir / "ChairNotes"

        candidate_dirs = []
        if chair_notes_dir.is_dir():
            candidate_dirs.append(chair_notes_dir)
        if agenda_dir.is_dir():
            candidate_dirs.append(agenda_dir)

        # 1. Discover all candidate files
        candidate_files: List[Path] = []
        seen_names = set()

        for d in candidate_dirs:
            for ext in ("*.docx", "*.doc", "*.zip", "*.htm", "*.html"):
                for p in d.glob(ext):
                    if p.name.startswith("~$"):
                        continue
                    if p.name not in seen_names:
                        seen_names.add(p.name)
                        candidate_files.append(p)

        # 2. Extract any ZIP archives if needed
        unpacked_files: List[Path] = []
        for f in list(candidate_files):
            if f.suffix.lower() == ".zip":
                try:
                    import zipfile
                    target_extract = chair_notes_dir if chair_notes_dir.is_dir() else agenda_dir
                    os.makedirs(to_extended_path(target_extract), exist_ok=True)
                    with zipfile.ZipFile(to_extended_path(f), "r") as zf:
                        for zinfo in zf.infolist():
                            zname = zinfo.filename
                            if zname.startswith("~$") or "/" in zname or "\\" in zname:
                                continue
                            zext = Path(zname).suffix.lower()
                            if zext in (".docx", ".doc", ".htm", ".html"):
                                out_p = target_extract / zname
                                if not out_p.exists():
                                    zf.extract(zinfo, path=to_extended_path(target_extract))
                                if out_p.name not in seen_names:
                                    seen_names.add(out_p.name)
                                    unpacked_files.append(out_p)
                except Exception as e:
                    logger.warning(f"[ChairNotesAgreementsParser] Failed extracting ZIP '{f.name}': {e}")

        all_files = [f for f in (candidate_files + unpacked_files) if f.suffix.lower() != ".zip"]

        if not all_files:
            cached_ag, cached_list = cls.load_cached_chair_notes_data(meeting_dir)
            return cached_ag, cached_list, 0

        # 3. Check cache freshness if not forcing rescan
        cache_path = agenda_dir / "chair_notes_cache.json"
        if not force_rescan and os.path.exists(to_extended_path(cache_path)):
            try:
                cache_mtime = os.path.getmtime(to_extended_path(cache_path))
                latest_mtime = max(os.path.getmtime(to_extended_path(p)) for p in all_files)
                with open(to_extended_path(cache_path), "r", encoding="utf-8") as f:
                    loaded_meta = json.load(f)
                cache_ver = loaded_meta.get("version", 1) if isinstance(loaded_meta, dict) else 1
                if cache_ver >= CHAIR_NOTES_CACHE_VERSION and cache_mtime >= latest_mtime:
                    cached_ag, cached_list = cls.load_cached_chair_notes_data(meeting_dir, min_version=CHAIR_NOTES_CACHE_VERSION)
                    if cached_ag or cached_list:
                        return cached_ag, cached_list, len(all_files)
            except Exception:
                pass

        # 4. Parse files and merge
        merged_agenda: Dict[str, dict] = {}
        merged_agreements: List[dict] = []
        scanned_count = 0

        # Sort so clean docx files take precedence over tracked changes or interim files
        def _file_sort_key(p: Path):
            name_lower = p.name.lower()
            is_clean = 1 if "clean" in name_lower else 0
            is_final = 1 if "final" in name_lower or "eom" in name_lower else 0
            return (is_clean, is_final, os.path.getmtime(to_extended_path(p)))

        sorted_files = sorted(all_files, key=_file_sort_key, reverse=True)

        for p in sorted_files:
            ext = p.suffix.lower()
            try:
                if ext == ".docx":
                    ad, ag = cls.parse_docx_with_agreements(str(p), mtg_info=mtg_info, agenda_map=agenda_map)
                    scanned_count += 1
                    merged_agenda.update(ad)
                    merged_agreements.extend(ag)
                elif ext in (".htm", ".html"):
                    from modules.meetings.core.tdocs.tdocs_parser import TDocsParser
                    ad = TDocsParser.parse_tdocs_by_agenda(str(p))
                    if ad:
                        scanned_count += 1
                        merged_agenda.update(ad)
                elif ext == ".doc":
                    # Check if converted .docx already exists
                    docx_candidate = p.with_suffix(".docx")
                    if docx_candidate.exists() and docx_candidate in sorted_files:
                        continue
                    try:
                        from modules.word_tools.word_comparator import is_libreoffice_available, convert_doc_to_docx_libreoffice
                        if is_libreoffice_available():
                            out_docx = convert_doc_to_docx_libreoffice(p, docx_candidate)
                            if out_docx and out_docx.exists():
                                ad, ag = cls.parse_docx_with_agreements(str(out_docx), mtg_info=mtg_info, agenda_map=agenda_map)
                                scanned_count += 1
                                merged_agenda.update(ad)
                                merged_agreements.extend(ag)
                    except Exception as e_lo:
                        logger.debug(f"[ChairNotesAgreementsParser] Could not convert legacy .doc '{p.name}': {e_lo}")
            except Exception as e_parse:
                logger.warning(f"[ChairNotesAgreementsParser] Error parsing '{p.name}': {e_parse}", exc_info=True)

        # 5. Persist merged cache if anything was extracted
        if merged_agenda or merged_agreements:
            cls.save_cached_chair_notes_data(meeting_dir, merged_agenda, merged_agreements, overwrite=True)

        return merged_agenda, merged_agreements, scanned_count
