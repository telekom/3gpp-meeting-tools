# --- File: src/modules/meetings/core/exporters/topic_dossier_exporter.py ---
"""
Cross-WG, Multi-Meeting Topic Dossier & Briefing Exporter Thread.

Performs parallel multi-WG, multi-meeting document conversion, extracts unified
decisions (RAN Chairman's Notes text agreements + SA/CT approved pCRs/CRs),
threads cross-WG Liaison Statements, and maintains an incremental export manifest
with dedicated LLM delta feeds.
"""

from concurrent.futures import ThreadPoolExecutor, as_completed
import datetime
import json
import logging
import os
import re
import shutil
import urllib.parse
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple

from PyQt5.QtCore import QThread, pyqtSignal

from core.utils.company_sanitizer import CompanySanitizer
from core.utils.paths import to_extended_path, get_project_root
from modules.meetings.core.cohort_resolver import (
    MeetingCohort,
    MeetingCohortResolver,
    format_cohort_label,
)
from modules.meetings.core.cross_wg_ls_tracker import CrossWgLiaisonTracker
from modules.meetings.core.exporters.export_manifest import (
    ExportManifest,
    get_manifest_path,
    load_manifest,
    is_export_complete,
)
from modules.meetings.core.meetings_db import MeetingsDatabase
from modules.meetings.core.tdocs.chair_notes_agreements_parser import (
    ChairNotesAgreementsParser,
    clean_chair_notes_text,
)
from modules.meetings.core.tdocs.tdoc_file_handler import TDocFileHandler
from modules.meetings.core.tdocs.tdoc_markdown_converter import (
    convert_document_to_markdown,
    sanitize_filename,
    is_conversion_stale,
    resolve_revision_mode_for_type,
    CONVERTER_VERSION,
)
from modules.meetings.core.tdocs.tdoc_url_resolver import resolve_tdoc_download_url
from modules.meetings.core.tdocs.tdocs_parser import TDocsParser
from modules.meetings.core.topic_decision import (
    TopicDecision,
    classify_tdoc_decision_type,
    normalize_chair_notes_decision_type,
)
from modules.meetings.core.topic_manager import MetaTopicDefinition

logger = logging.getLogger(__name__)


def sanitize_folder_name(name: str) -> str:
    """Sanitizes directory names, removing illegal Windows/filesystem characters."""
    clean = re.sub(r'[\\/*?:"<>|#%]', "_", str(name or "")).strip()
    return re.sub(r'_+', '_', clean).strip("._ ") or "General"


def safe_relpath(target: Path, base: Path) -> str:
    """Computes clean relative path between two paths, safe for Windows extended paths."""
    try:
        t_str = str(target).replace("\\\\?\\", "").replace("//?/", "")
        b_str = str(base).replace("\\\\?\\", "").replace("//?/", "")
        t_abs = Path(t_str).resolve()
        b_abs = Path(b_str).resolve()
        return os.path.relpath(t_abs, b_abs).replace("\\", "/")
    except Exception:
        return str(target).replace("\\", "/")


def safe_markdown_rel_link(target: Path, base: Path) -> str:
    """Computes URL-encoded relative path for Markdown links."""
    rel = safe_relpath(target, base)
    clean_path = str(rel).replace("\\", "/")
    if "#" in clean_path:
        p_part, frag = clean_path.split("#", 1)
        return f"{urllib.parse.quote(p_part, safe='/')}#{urllib.parse.quote(frag, safe='')}"
    return urllib.parse.quote(clean_path, safe="/")


def get_default_topic_export_base() -> Path:
    """
    Returns the default base directory for cross-WG topic exports,
    located in the user's meeting cache directory rather than the code repository.
    """
    try:
        from modules.meetings.core.settings import MeetingsSettings
        settings = MeetingsSettings()
        if settings.cache_dir and Path(settings.cache_dir).is_dir():
            return Path(settings.cache_dir) / "Export"
    except Exception:
        pass
    from core.utils.paths import get_cache_root
    cache_root = get_cache_root()
    if cache_root and Path(cache_root).is_dir():
        return Path(cache_root) / "Export"
    return Path.home() / "3GPP_Delegate_Helper" / "cache" / "Export"



class TopicDossierExporterThread(QThread):
    progress_msg = pyqtSignal(str)
    progress_step = pyqtSignal(int, int)
    session_finished = pyqtSignal(bool, str, str)  # success, target_dir, summary

    def __init__(
        self,
        meetings_db: MeetingsDatabase,
        topic: MetaTopicDefinition,
        selected_cohorts: List[MeetingCohort],
        export_scope_mode: str = "full_dossier",  # "pre_meeting", "post_meeting", "full_dossier"
        output_base_dir: Optional[Path] = None,
        target_folder_name: Optional[str] = None,
        include_discussions: bool = True,
        deduplicate_revisions: bool = True,
        generate_consolidated_md: bool = True,
        allow_network_download: bool = True,
        parent=None,
    ):
        super().__init__(parent)
        self.meetings_db = meetings_db
        self.topic = topic
        self.cohorts = list(selected_cohorts)
        self.export_scope_mode = export_scope_mode
        self.output_base_dir = Path(output_base_dir) if output_base_dir else get_default_topic_export_base()
        self.target_folder_name = target_folder_name
        self.include_discussions = bool(include_discussions)
        self.deduplicate_revisions = bool(deduplicate_revisions)
        self.generate_consolidated_md = bool(generate_consolidated_md)
        self.allow_network_download = bool(allow_network_download)

        self._is_cancelled = False
        self._executor: Optional[ThreadPoolExecutor] = None

    def cancel(self):
        self._is_cancelled = True
        self.requestInterruption()
        if self._executor:
            try:
                self._executor.shutdown(wait=False, cancel_futures=True)
            except Exception:
                pass

    @property
    def is_cancelled(self) -> bool:
        return self._is_cancelled or self.isInterruptionRequested()

    def run(self):
        try:
            self._run_export()
        except Exception as e:
            logger.exception(f"[TopicDossierExporter] Export failed: {e}")
            self.session_finished.emit(False, "", f"Export failed: {str(e)}")

    def _run_export(self):
        self.progress_msg.emit("Initializing Topic Dossier Studio...")
        self.progress_step.emit(0, 100)

        # 1. Resolve Target Directory
        folder_stem = self.target_folder_name or f"Topic_{sanitize_folder_name(self.topic.name)}"
        target_dir = self.output_base_dir / folder_stem
        target_dir_ext = to_extended_path(target_dir)
        os.makedirs(target_dir_ext, exist_ok=True)

        # 2. Load or Initialize Manifest
        manifest = load_manifest(target_dir)
        is_incremental = manifest is not None
        if not manifest:
            manifest = ExportManifest(
                meeting=self.topic.name,
                export_type=f"topic_{self.export_scope_mode}",
            )
        session_id = len(manifest.sessions) + 1
        session_timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        self.progress_msg.emit(f"Export target: {folder_stem} (Session #{session_id})")

        # 3. Discovery across Cohorts & WGs
        total_cohorts = len(self.cohorts)
        ls_tracker = CrossWgLiaisonTracker()
        all_decisions: List[TopicDecision] = []
        new_decisions_this_session: List[TopicDecision] = []
        candidate_docs_to_export: List[Dict[str, Any]] = []

        self.progress_msg.emit("Scanning meetings and Chairman's Notes across cohorts...")

        for c_idx, cohort in enumerate(self.cohorts):
            if self.is_cancelled:
                return

            self.progress_msg.emit(f"Analyzing {cohort.label} ({c_idx + 1}/{total_cohorts})...")
            self.progress_step.emit(int((c_idx / total_cohorts) * 30), 100)

            for mtg in cohort.meetings:
                wg_name = str(mtg.get("wg_name", "")).strip().upper()
                mtg_num = str(mtg.get("meeting_number", "")).strip()
                folder_name = str(mtg.get("folder_name", "")).strip()
                mtg_date = str(mtg.get("start_date", "")).strip()

                # Parse TDoc list if present locally
                tdocs_rows = self._load_meeting_tdocs(mtg)
                # Parse Chairman's Notes agreements if available
                chair_agreements = self._load_meeting_agreements(mtg)

                # Process Chairman's Notes text agreements
                for ca in chair_agreements:
                    # Check if agreement matches topic
                    wi_str = str(ca.get("wi", "") or ca.get("study_item", "")).strip()
                    ai_str = str(ca.get("ai", "") or ca.get("agenda_item", "")).strip()
                    text_str = clean_chair_notes_text(ca.get("text", "") or ca.get("summary", ""))

                    # Match against topic definition
                    if self._matches_topic_rule(wg_name, ai_str, text_str, wi_str):
                        dec_type = normalize_chair_notes_decision_type(ca.get("decision_type", "Agreement"))
                        dec = TopicDecision(
                            cohort_key=cohort.key,
                            wg=wg_name,
                            meeting_number=mtg_num,
                            date=mtg_date,
                            decision_type=dec_type,
                            title_or_summary=text_str,
                            tdoc_id=ca.get("tdoc_id"),
                            source=ca.get("source") or f"Chairman's Notes AI {ai_str}",
                            target_spec=ca.get("spec", ""),
                            ai=ai_str,
                            wi=wi_str,
                            status=ca.get("status", "Agreed"),
                            notes=ca.get("notes", ""),
                        )
                        all_decisions.append(dec)
                        new_decisions_this_session.append(dec)

                # Process TDoc rows
                for row in tdocs_rows:
                    tid = str(row.get("TDoc", "") or row.get("tdoc_id", "") or row.get("TD#", "")).strip().upper()
                    if not tid:
                        continue
                    title = str(row.get("Title", "") or row.get("title", "")).strip()
                    ai_val = str(
                        row.get("Agenda Item", "") or 
                        row.get("Agenda item", "") or 
                        row.get("AI", "") or 
                        row.get("ai", "") or 
                        row.get("AI#", "")
                    ).strip()
                    wi_val = str(
                        row.get("Related WIs", "") or 
                        row.get("wi", "") or 
                        row.get("WI", "") or 
                        row.get("Work Item", "")
                    ).strip()
                    t_type = str(row.get("Type", "") or row.get("type", "") or row.get("TDoc Type", "")).strip()
                    status = str(
                        row.get("TDoc Status", "") or 
                        row.get("Status", "") or 
                        row.get("status", "") or 
                        row.get("Result", "")
                    ).strip()
                    source = str(row.get("Source", "") or row.get("source", "")).strip()
                    abstract = str(row.get("Abstract", "") or row.get("abstract", "") or row.get("Secretary Remarks", "")).strip()
                    ai_desc = str(row.get("Agenda item description", "") or row.get("AI Description", "")).strip()

                    if not self.topic.matches_tdoc(tid, wg_name, ai_val, title, wi_val, abstract, ai_description=ai_desc):
                        continue

                    # Filter based on scope mode
                    is_agreed = any(s in status.lower() for s in ("approved", "agreed", "endorsed"))
                    is_ls = "ls" in t_type.lower() or "liaison" in t_type.lower()

                    if self.export_scope_mode == "post_meeting" and not (is_agreed or is_ls):
                        continue

                    # Check decision classification
                    dec_type = classify_tdoc_decision_type(t_type, title, status)
                    if dec_type:
                        dec = TopicDecision(
                            cohort_key=cohort.key,
                            wg=wg_name,
                            meeting_number=mtg_num,
                            date=mtg_date,
                            decision_type=dec_type,
                            title_or_summary=title,
                            tdoc_id=tid,
                            source=source,
                            target_spec=str(row.get("Spec", "") or row.get("spec", "")).strip(),
                            ai=ai_val,
                            wi=wi_val,
                            status=status,
                            source_url=resolve_tdoc_download_url(tid, folder_name),
                        )
                        all_decisions.append(dec)

                    # Prepare relative destination path
                    cat_folder = self._resolve_category_folder(t_type, status)
                    rel_md_path = f"{cohort.key}/{wg_name}_{sanitize_folder_name(mtg_num)}/{cat_folder}/{tid}.md"
                    abs_md_path = target_dir / rel_md_path

                    # Track LS in dialogue tracker
                    if is_ls:
                        ls_tracker.ingest_ls_tdoc(
                            tdoc_id=tid,
                            wg=wg_name,
                            meeting_number=mtg_num,
                            cycle=cohort.key,
                            tdoc_type=t_type,
                            title=title,
                            source=source,
                            status=status,
                            reply_to=row.get("Reply to", ""),
                            reply_in=row.get("Reply in", ""),
                            dossier_rel_path=rel_md_path,
                        )

                    # Queue candidate for document conversion
                    candidate_docs_to_export.append({
                        "tdoc_id": tid,
                        "wg": wg_name,
                        "meeting_number": mtg_num,
                        "folder_name": folder_name,
                        "cycle": cohort.key,
                        "title": title,
                        "ai": ai_val,
                        "wi": wi_val,
                        "type": t_type,
                        "status": status,
                        "source": source,
                        "category": cat_folder,
                        "rel_path": rel_md_path,
                        "abs_path": abs_md_path,
                        "meeting_dict": mtg,
                        "row_data": row,
                    })

        if self.is_cancelled:
            return

        # 4. Filter already exported documents via Manifest
        docs_to_convert: List[Dict[str, Any]] = []
        already_complete_count = 0

        for doc in candidate_docs_to_export:
            tid = doc["tdoc_id"]
            abs_p = doc["abs_path"]
            if manifest.is_tdoc_exported(tid) and is_export_complete(abs_p, doc_info=doc):
                already_complete_count += 1
                # Attach dossier path to matching decisions (for formal doc decisions)
                for d in all_decisions:
                    if d.tdoc_id == tid and d.decision_type in ("Approved pCR", "Approved CR"):
                        d.dossier_rel_path = doc["rel_path"]
            else:
                docs_to_convert.append(doc)

        self.progress_msg.emit(
            f"Corpus check: {already_complete_count} already exported. {len(docs_to_convert)} new documents to process."
        )

        # 5. Two-Phase Parallel Pre-Download & Conversion
        if docs_to_convert:
            self._convert_docs_pipeline(docs_to_convert, target_dir)

        # Ensure all formal document decisions (pCR / CR) have valid dossier links if converted
        for d in all_decisions:
            if d.tdoc_id and not d.dossier_rel_path and d.decision_type in ("Approved pCR", "Approved CR"):
                for doc in candidate_docs_to_export:
                    if doc["tdoc_id"] == d.tdoc_id and doc["abs_path"].is_file():
                        d.dossier_rel_path = doc["rel_path"]
                        break

        # 5b. Export Dedicated Standalone Markdown Files for Text Agreements (Chairman's Notes)
        self.progress_msg.emit("Exporting standalone agreements to Markdown...")
        seq_counters: Dict[str, int] = {}
        for dec in all_decisions:
            if not dec.dossier_rel_path:
                key = f"{dec.cohort_key}_{dec.wg}_{dec.meeting_number}_{dec.ai}"
                seq_counters[key] = seq_counters.get(key, 0) + 1
                dec.dossier_rel_path = self._write_agreement_markdown(
                    target_dir=target_dir,
                    dec=dec,
                    candidate_docs=candidate_docs_to_export,
                    seq=seq_counters[key],
                )

        # 6. Synthesize Reports
        self.progress_msg.emit("Synthesizing Master Indices and LLM Delta Feed...")
        self.progress_step.emit(85, 100)

        # A. Master agreements_timeline.md
        self._write_agreements_timeline(target_dir, all_decisions)

        # B. Cross-WG ls_dialogue.md
        ls_report_md = ls_tracker.generate_markdown_report()
        with open(target_dir / "ls_dialogue.md", "w", encoding="utf-8") as f:
            f.write(ls_report_md)

        # C. Incremental Delta Feed for LLMs (deltas/delta_{cohort}.md)
        newly_exported_tids = [d["tdoc_id"] for d in docs_to_convert if d["abs_path"].is_file()]
        successfully_converted_docs = [d for d in docs_to_convert if d["abs_path"].is_file()]
        self._write_llm_delta_feed(
            target_dir=target_dir,
            session_id=session_id,
            session_timestamp=session_timestamp,
            new_decisions=new_decisions_this_session if is_incremental else all_decisions,
            new_tdocs=successfully_converted_docs if successfully_converted_docs else docs_to_convert,
            ls_tracker=ls_tracker,
        )

        # D. Master index.md
        self._write_master_index(
            target_dir=target_dir,
            manifest=manifest,
            all_decisions=all_decisions,
            all_candidates=candidate_docs_to_export,
            session_id=session_id,
        )

        # E. Optional Consolidated Single Markdown File
        if self.generate_consolidated_md:
            self._write_consolidated_briefing(target_dir, all_decisions, candidate_docs_to_export)

        # 7. Record Session in Manifest
        manifest.record_session(
            session_id=session_id,
            timestamp=session_timestamp,
            scope_mode=self.export_scope_mode,
            added_tdocs=newly_exported_tids,
            updated_tdocs=[],
            added_entries=successfully_converted_docs,
        )
        with open(get_manifest_path(target_dir), "w", encoding="utf-8") as f:
            json.dump(manifest.to_dict(), f, indent=2, ensure_ascii=False)

        self.progress_step.emit(100, 100)
        summary = (
            f"Export completed for '{self.topic.name}'!\n"
            f"• {len(all_decisions)} Unified Decisions (Agreements, WAs, pCRs, CRs)\n"
            f"• {len(candidate_docs_to_export)} Total Contributions across {len(self.cohorts)} Meeting Round(s)\n"
            f"• {len(newly_exported_tids)} Documents converted in Session #{session_id} ({already_complete_count} skipped)\n"
            f"• LLM Delta feed written to deltas/delta_session_{session_id}.md"
        )
        self.session_finished.emit(True, str(target_dir), summary)

    def _find_or_download_tdoc_file(self, tid: str, folder_name: str, meeting_dict: Dict[str, Any]) -> Optional[Path]:
        mtg_dir = self._get_meeting_dir(folder_name)
        if not mtg_dir:
            return None

        base_match = re.search(r'^(.*?)-?(?:r|rev)\d{1,2}[a-zA-Z]?$', tid, re.IGNORECASE)
        base_tdoc = base_match.group(1).upper() if base_match else tid.upper()

        candidate_folders = [
            mtg_dir / tid,
            mtg_dir / base_tdoc,
            mtg_dir,
            mtg_dir / "Docs",
            mtg_dir / "Revisions",
        ]

        # 1. Fast direct check if already extracted in meeting folder
        for folder in candidate_folders:
            if folder.is_dir():
                for ext in (".docx", ".docm", ".pptx", ".doc", ".ppt", ".pdf"):
                    direct_f = folder / f"{tid}{ext}"
                    if direct_f.is_file() and direct_f.stat().st_size > 0:
                        return direct_f
                    if base_tdoc != tid:
                        base_f = folder / f"{base_tdoc}{ext}"
                        if base_f.is_file() and base_f.stat().st_size > 0:
                            return base_f
                    for f in folder.glob(f"*{tid}*{ext}"):
                        if not f.name.startswith("~$") and f.is_file() and f.stat().st_size > 0:
                            return f

        # 2. Check if ZIP archive already exists in tdoc_dir or mtg_dir
        candidate_zips = [
            mtg_dir / tid / f"{tid}.zip",
            mtg_dir / base_tdoc / f"{base_tdoc}.zip",
            mtg_dir / f"{tid}.zip",
            mtg_dir / f"{base_tdoc}.zip",
            mtg_dir / "Docs" / f"{tid}.zip",
            mtg_dir / "Docs" / f"{base_tdoc}.zip",
            mtg_dir / "Revisions" / f"{tid}.zip",
            mtg_dir / "Revisions" / f"{base_tdoc}.zip",
        ]
        target_extract_dir = mtg_dir / base_tdoc
        for zip_path in candidate_zips:
            if zip_path.is_file() and zip_path.stat().st_size > 0:
                try:
                    import zipfile
                    with zipfile.ZipFile(zip_path, 'r') as zf:
                        for member in zf.namelist():
                            if member.lower().endswith(('.docx', '.docm', '.pptx', '.doc', '.ppt', '.pdf')) and not os.path.basename(member).startswith('~$'):
                                zf.extract(member, target_extract_dir)
                                return target_extract_dir / member
                except Exception:
                    pass

        # 3. If missing and network download is not enabled, return None (fast metadata stub fallback)
        if not self.allow_network_download or self.is_cancelled:
            return None

        from modules.meetings.core.tdocs.tdoc_url_resolver import resolve_tdoc_download_url, normalize_3gpp_url

        candidate_urls = []
        is_rev = bool(re.search(r'^(.*?)-?(?:r|rev)\d{1,2}[a-zA-Z]?$', tid, re.IGNORECASE))
        docs_url = meeting_dict.get("docs_folder_url")
        url_key = meeting_dict.get("url_key")

        # 1. Primary official candidate from URL router
        resolved_primary = resolve_tdoc_download_url(
            tid,
            docs_url=docs_url,
            main_url=url_key,
            meta=meeting_dict,
        )
        if resolved_primary:
            candidate_urls.append(resolved_primary.rsplit('/', 1)[0])

        # 2. Derive base folders from docs_url and url_key
        bases = []
        if docs_url:
            norm_docs = normalize_3gpp_url(docs_url)
            m_base = re.sub(r'/[Dd]ocs/?$', '', norm_docs, flags=re.IGNORECASE)
            bases.append(m_base)
        if url_key:
            norm_key = normalize_3gpp_url(url_key)
            bases.append(norm_key)

        for b in bases:
            if is_rev:
                candidate_urls.append(f"{b}/INBOX/Revisions")
                candidate_urls.append(f"{b}/Inbox/Revisions")
                candidate_urls.append(f"{b}/Revisions")
                candidate_urls.append(f"{b}/Docs")
            else:
                candidate_urls.append(f"{b}/Docs")
                candidate_urls.append(f"{b}/INBOX")
                candidate_urls.append(f"{b}/Inbox")

        # Deduplicate while preserving order
        dedup_urls = list(dict.fromkeys(u.rstrip('/') for u in candidate_urls if u))

        # Attempt download for tid (and fallback to base_tdoc if revision is not found)
        targets_to_try = [tid]
        if is_rev and base_tdoc != tid:
            targets_to_try.append(base_tdoc)

        for target_id in targets_to_try:
            for base_url in dedup_urls:
                if self.is_cancelled:
                    return None
                try:
                    os.makedirs(target_extract_dir, exist_ok=True)
                    docs = TDocFileHandler.download_and_extract_tdoc(
                        target_filename=target_id,
                        base_url=base_url,
                        tdoc_dir=target_extract_dir,
                        timeout=12,
                        cancel_check=lambda: self.is_cancelled,
                    )
                    if docs:
                        logger.info(f"📥 [TopicDossier] Downloaded {target_id} from {base_url}")
                        return docs[0]
                except Exception as e:
                    logger.debug(f"[Download] Failed {target_id} from {base_url}: {e}")

        return None

    def _convert_docs_pipeline(self, docs_to_convert: List[Dict[str, Any]], target_dir: Path):
        """Two-phase parallel pre-download and sequential in-memory conversion."""
        total = len(docs_to_convert)

        # Phase 1: Polite Concurrent Pre-Download (max_workers=5)
        self.progress_msg.emit(f"Phase 1/2: Resolving source documents ({total} files)...")
        with ThreadPoolExecutor(max_workers=5) as executor:
            self._executor = executor
            future_to_doc = {
                executor.submit(self._find_or_download_tdoc_file, doc["tdoc_id"], doc["folder_name"], doc.get("meeting_dict", {})): doc
                for doc in docs_to_convert
            }
            done_count = 0
            for fut in as_completed(future_to_doc):
                if self.is_cancelled:
                    return
                done_count += 1
                self.progress_step.emit(30 + int((done_count / total) * 30), 100)
                try:
                    res = fut.result()
                    doc = future_to_doc[fut]
                    doc["source_file"] = res
                except Exception as e:
                    logger.debug(f"[PreDownload] Failed item: {e}")

        # Phase 2: In-Memory Markdown Conversion
        self.progress_msg.emit(f"Phase 2/2: Converting documents to Markdown ({total} files)...")
        for idx, doc in enumerate(docs_to_convert):
            if self.is_cancelled:
                return
            tid = doc["tdoc_id"]
            abs_p = doc["abs_path"]
            src_file = doc.get("source_file")

            self.progress_msg.emit(f"Converting ({idx + 1}/{total}): {tid}...")
            self.progress_step.emit(60 + int((idx / total) * 25), 100)

            converted_ok = False
            # Check if local meeting cache already had a converted markdown file
            mtg_dir = self._get_meeting_dir(doc.get("folder_name", ""))
            if mtg_dir:
                base_match = re.search(r'^(.*?)-?(?:r|rev)\d{1,2}[a-zA-Z]?$', tid, re.IGNORECASE)
                base_tid = base_match.group(1).upper() if base_match else tid.upper()
                for cand_dir in [mtg_dir / tid, mtg_dir / base_tid]:
                    if cand_dir.is_dir():
                        for mdf in [cand_dir / f"{tid}.md", cand_dir / f"{sanitize_filename(tid)}.md"]:
                            if mdf.is_file() and mdf.stat().st_size > 120:
                                try:
                                    os.makedirs(abs_p.parent, exist_ok=True)
                                    shutil.copy2(mdf, abs_p)
                                    loc_media = cand_dir / "media"
                                    if loc_media.is_dir():
                                        tgt_media = abs_p.parent / f"{abs_p.stem}_media"
                                        os.makedirs(tgt_media, exist_ok=True)
                                        for mf in loc_media.glob("*.*"):
                                            shutil.copy2(mf, tgt_media / mf.name)
                                    converted_ok = abs_p.is_file() and abs_p.stat().st_size > 0
                                    break
                                except Exception:
                                    pass
                    if converted_ok:
                        break

            if not converted_ok and src_file and os.path.exists(src_file):
                os.makedirs(abs_p.parent, exist_ok=True)
                media_dir = abs_p.parent / f"{abs_p.stem}_media"
                try:
                    mode = resolve_revision_mode_for_type(doc.get("type", ""))
                    convert_document_to_markdown(
                        doc_path=Path(src_file),
                        out_dir=abs_p.parent,
                        tdoc_id=tid,
                        meta=doc.get("row_data", {}),
                        revisions=mode,
                        media_dir=media_dir,
                        media_rel=f"{abs_p.stem}_media",
                        tdoc_url=resolve_tdoc_download_url(tid, doc["folder_name"]),
                        cancellation_check=lambda: self.is_cancelled,
                    )
                    converted_ok = abs_p.is_file() and abs_p.stat().st_size > 0
                except Exception as e:
                    logger.warning(f"[Convert] Failed to convert {tid}: {e}")

            if not converted_ok and not abs_p.is_file():
                # Fallback: if document binary is unavailable, create a structured metadata stub
                try:
                    os.makedirs(abs_p.parent, exist_ok=True)
                    row = doc.get("row_data", {})
                    src_url = resolve_tdoc_download_url(tid, doc["folder_name"])
                    source_lines = f"- **Source URL:** [{src_url}]({src_url})" if src_url else ""
                    stub_content = (
                        f"# {tid}: {doc.get('title', '')}\n\n"
                        f"- **Working Group:** {doc.get('wg', '')}\n"
                        f"- **Meeting:** {doc.get('meeting_number', '')}\n"
                        f"- **Agenda Item:** {doc.get('ai', '')}\n"
                        f"- **Work Item:** {doc.get('wi', '')}\n"
                        f"- **Type:** {doc.get('type', '')}\n"
                        f"- **Status:** {doc.get('status', '')}\n"
                        f"- **Source / Company:** {doc.get('source', '')}\n"
                        f"{source_lines}\n\n"
                        f"> [!NOTE]\n"
                        f"> *Document conversion was not available in local cache at export time.*\n\n"
                        f"### Abstract / Remarks\n"
                        f"{row.get('Abstract', '') or row.get('Secretary Remarks', 'No abstract provided.')}\n"
                    )
                    with open(abs_p, "w", encoding="utf-8") as sf:
                        sf.write(stub_content)
                except Exception as stub_err:
                    logger.warning(f"[Stub] Failed to create stub for {tid}: {stub_err}")

    def _matches_topic_rule(self, wg: str, ai: str, text: str, wi: str) -> bool:
        """Evaluates whether an agreement item matches the topic rules."""
        # Work Item token match
        if self.topic.rules.wi_tokens and wi:
            tokens = {t.lower() for t in re.split(r'[,;\s/&]+', wi) if t.strip()}
            if tokens.intersection(set(self.topic.rules.wi_tokens)):
                return True

        # AI pattern match
        norm_ai = str(ai or "").strip().lower()
        if norm_ai and wg in self.topic.rules.ai_patterns:
            patterns = self.topic.rules.ai_patterns[wg]
            for pat in patterns:
                clean_pat = pat.lower().strip().rstrip("*")
                if norm_ai.startswith(clean_pat):
                    return True

        # Keyword match
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

    def _load_meeting_tdocs(self, mtg: Dict[str, Any]) -> List[Dict[str, Any]]:
        folder_name = mtg.get("folder_name", "")
        mtg_dir = self._get_meeting_dir(folder_name)
        if not mtg_dir:
            return []

        agenda_dir = mtg_dir / "Agenda"

        # 1. Fast path: check for pre-parsed JSON cache (loads in ~2ms)
        if agenda_dir.is_dir():
            for jf in agenda_dir.glob("*.xlsx.json"):
                try:
                    with open(jf, "r", encoding="utf-8") as f:
                        return json.load(f)
                except Exception:
                    pass

        # 2. Check for Excel workbook in Agenda/ or root mtg_dir
        excel_candidates = []
        if agenda_dir.is_dir():
            excel_candidates.extend(list(agenda_dir.glob("*.xlsx")))
        if mtg_dir.is_dir():
            excel_candidates.extend(list(mtg_dir.glob("*.xlsx")))

        for ef in excel_candidates:
            if ef.is_file() and not ef.name.startswith("~$"):
                try:
                    res = TDocsParser.parse_tdocs_excel(str(ef))
                    if res:
                        return res
                except Exception as e:
                    logger.debug(f"[TDocsParser] Excel parse failed for {ef}: {e}")

        # 3. If missing and allowed, auto-download via TDocsDownloaderThread
        if self.allow_network_download and not self.is_cancelled:
            mtg_id = mtg.get("id") or mtg.get("meeting_number")
            if mtg_id:
                try:
                    from modules.meetings.core.tdocs.tdocs_downloader import TDocsDownloaderThread
                    dl = TDocsDownloaderThread(str(mtg_id), mtg_dir, cancel_check=lambda: self.is_cancelled)
                    dl.run()
                    if agenda_dir.is_dir():
                        for ef in agenda_dir.glob("*.xlsx"):
                            if not ef.name.startswith("~$"):
                                res = TDocsParser.parse_tdocs_excel(str(ef))
                                if res:
                                    return res
                except Exception as e:
                    logger.debug(f"[TDocsDownloader] Auto-download failed: {e}")

        return []

    def _load_meeting_agreements(self, mtg: Dict[str, Any]) -> List[Dict[str, Any]]:
        folder_name = mtg.get("folder_name", "")
        mtg_dir = self._get_meeting_dir(folder_name)
        if not mtg_dir:
            return []

        from modules.meetings.core.tdocs.chair_notes_agreements_parser import (
            ChairNotesAgreementsParser, CHAIR_NOTES_CACHE_VERSION
        )

        agenda_dir = mtg_dir / "Agenda"
        cn_cache = agenda_dir / "chair_notes_cache.json"
        if cn_cache.is_file():
            try:
                with open(cn_cache, "r", encoding="utf-8") as f:
                    data = json.load(f)
                if data.get("version", 1) >= CHAIR_NOTES_CACHE_VERSION:
                    return data.get("agreements", [])
            except Exception:
                pass

        # If cache missing or outdated version, run scan_and_cache_chair_notes
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

        # Fallback: check ChairNotes subdirectory for docx
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
                            return res[1]  # agreements list
                    except Exception:
                        pass

        # If still no agreements and network download is allowed, auto-fetch Chairman's Notes for RAN WGs
        if self.allow_network_download and not self.is_cancelled:
            wg_upper = str(mtg.get("wg_name", "")).strip().upper()
            if "RAN" in wg_upper:
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
                                cn_dir = agenda_dir / "ChairNotes"
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
                    logger.debug(f"[TopicDossierExporter] Auto-download ChairNotes failed for {folder_name}: {e}")

        return []

    def _write_agreement_markdown(
        self,
        target_dir: Path,
        dec: TopicDecision,
        candidate_docs: List[Dict[str, Any]],
        seq: int = 1,
    ) -> str:
        """Exports a standalone Markdown document for a text agreement or working assumption from Chairman's Notes."""
        wg_clean = sanitize_folder_name(dec.wg)
        mtg_clean = sanitize_folder_name(dec.meeting_number)
        cohort_clean = sanitize_folder_name(dec.cohort_key)

        agreements_dir = target_dir / cohort_clean / f"{wg_clean}_{mtg_clean}" / "Agreements"
        agreements_dir.mkdir(parents=True, exist_ok=True)

        if dec.tdoc_id:
            tid_clean = sanitize_folder_name(dec.tdoc_id)
            if seq > 1:
                filename = f"Agreement_{tid_clean}_{seq}.md"
            else:
                filename = f"Agreement_{tid_clean}.md"
        else:
            ai_clean = sanitize_folder_name(dec.ai or "General")
            filename = f"Agreement_AI_{ai_clean}_{seq:02d}.md"

        out_file = agreements_dir / filename
        rel_path = f"{cohort_clean}/{wg_clean}_{mtg_clean}/Agreements/{filename}"

        # Resolve background discussion TDoc link if present in candidate_docs
        bg_link = "—"
        if dec.tdoc_id:
            bg_doc = next((doc for doc in candidate_docs if doc.get("tdoc_id") == dec.tdoc_id), None)
            if bg_doc and bg_doc.get("rel_path"):
                doc_rel = bg_doc["rel_path"]
                try:
                    rel_to_doc = os.path.relpath(target_dir / doc_rel, agreements_dir).replace("\\", "/")
                    bg_link = f"[{dec.tdoc_id}]({rel_to_doc}) *(Noted Proposal)*"
                except Exception:
                    bg_link = f"[{dec.tdoc_id}](../../{doc_rel}) *(Noted Proposal)*"
            else:
                bg_link = f"`{dec.tdoc_id}` *(Noted / Discussed)*"

        portal_url = dec.source_url or (resolve_tdoc_download_url(dec.tdoc_id, "") if dec.tdoc_id else "")
        portal_link = f"[{dec.tdoc_id}]({portal_url})" if portal_url else "—"

        lines = [
            f"# 📜 3GPP Text Agreement: {dec.wg} #{dec.meeting_number} (AI {dec.ai or 'N/A'})",
            "",
            "| Metadata Attribute | Value |",
            "| :--- | :--- |",
            f"| **Working Group** | **{dec.wg}** |",
            f"| **Meeting** | #{dec.meeting_number} ({dec.date or 'N/A'}) |",
            f"| **Agenda Item (AI)** | `{dec.ai or 'N/A'}` |",
            f"| **Decision Type** | **{dec.badge_icon} {dec.decision_type}** |",
            f"| **Status** | {dec.status or 'Agreed'} |",
            f"| **Target Specification** | {dec.target_spec or '—'} |",
            f"| **Related Work Item** | {dec.wi or '—'} |",
            f"| **Chairman's Notes Source** | {dec.source or 'Chairman Notes'} |",
            f"| **Background Contribution** | {bg_link} |",
            f"| **Official 3GPP Portal** | {portal_link} |",
            "",
            "---",
            "",
            "## 📝 Verbatim Consensus Text (Chairman's Notes)",
            "",
            "> [!NOTE] 3GPP RAN Decision Nature",
            "> In 3GPP RAN working groups (RAN1, RAN2, RAN3), the **verbatim text below represents the formal, binding consensus** reached by the delegation.",
            "> The cited contribution was \"Noted\" in the meeting report and serves as technical provenance/context.",
            "",
            "```text",
            dec.title_or_summary.strip(),
            "```",
            "",
            "### Formatted Excerpt",
            "",
        ]

        for paragraph in dec.title_or_summary.strip().split("\n"):
            p_clean = paragraph.strip()
            if p_clean:
                lines.append(f"> {p_clean}")
            else:
                lines.append(">")

        lines.extend([
            "",
            "---",
            "",
            "## 🔍 Context & Topic Traceability",
            f"- **Meta-Topic:** {self.topic.name}",
            f"- **Meeting Cycle:** {dec.cohort_key}",
            f"- **Meeting Record:** Documented under Agenda Item **{dec.ai or 'N/A'}** in the Chairman's Notes of 3GPP **{dec.wg}** Meeting **#{dec.meeting_number}**.",
            "",
        ])

        try:
            with open(out_file, "w", encoding="utf-8") as f:
                f.write("\n".join(lines))
        except Exception as e:
            logger.warning(f"Failed to write agreement markdown {out_file}: {e}")

        return rel_path

    def _write_agreements_timeline(self, target_dir: Path, decisions: List[TopicDecision]):
        """Generates agreements_timeline.md unifying RAN text agreements and SA/CT pCRs."""
        lines = [
            f"# ⚖️ Unified Technical Decisions & Baseline: {self.topic.name}",
            "",
            "> [!NOTE]",
            "> This timeline unites technical decisions across Working Groups at an equal level:",
            "> - **RAN Working Groups (RAN1/2/3):** Substantive text agreements, working assumptions, and endorsed Way Forwards from Chairman's Notes.",
            "> - **SA & CT Groups (SA2, CT1, etc.):** Approved Pseudo-CRs (pCRs) to TRs and normative Change Requests (CRs) to Technical Specifications.",
            "",
            "| Meeting Round / Date | WG | Type | Substantive Decision / Scope | Contribution | Target Spec |",
            "| :--- | :--- | :--- | :--- | :--- | :--- |",
        ]

        # Sort chronologically by cohort, date, WG
        sorted_decisions = sorted(decisions, key=lambda d: (d.cohort_key, d.date or "", d.wg, d.decision_type))
        for dec in sorted_decisions:
            lines.append(dec.to_markdown_row())

        lines.append("")
        with open(target_dir / "agreements_timeline.md", "w", encoding="utf-8") as f:
            f.write("\n".join(lines))

    def _write_llm_delta_feed(
        self,
        target_dir: Path,
        session_id: int,
        session_timestamp: str,
        new_decisions: List[TopicDecision],
        new_tdocs: List[Dict[str, Any]],
        ls_tracker: CrossWgLiaisonTracker,
    ):
        """Writes an incremental delta feed optimized for LLM / AI digestion."""
        delta_dir = target_dir / "deltas"
        os.makedirs(delta_dir, exist_ok=True)
        delta_file = delta_dir / f"delta_session_{session_id}.md"

        lines = [
            f"# 🧠 Incremental AI Delta Feed — Session #{session_id}",
            f"**Generated:** {session_timestamp} | **Topic:** {self.topic.name}",
            "",
            "> [!IMPORTANT] LLM Prompting Guidance",
            "> This document contains strictly the **new decisions and contributions** added during this update session.",
            "> Feed this file to your local LLM (Ollama, Claude, NotebookLM, or Gemini) alongside `agreements_timeline.md` to analyze what shifted in the latest meeting cycle.",
            "",
            "## 🆕 Newly Added Technical Decisions",
            "",
        ]

        if new_decisions:
            lines.extend([
                "| Meeting Round | WG | Meeting | Decision Type | Summary / Scope | Contribution |",
                "| :--- | :--- | :--- | :--- | :--- | :--- |",
            ])
            for d in new_decisions:
                doc_link = f"[{d.tdoc_id}]({d.dossier_rel_path})" if d.dossier_rel_path else (d.tdoc_id or "—")
                summary_clean = re.sub(r'[\r\n\t]+', ' ', d.title_or_summary).replace("|", "\\|")
                lines.append(f"| {d.cohort_key} | **{d.wg}** | {d.meeting_number} | {d.badge_icon} {d.decision_type} | {summary_clean} | {doc_link} |")
        else:
            lines.append("*No new decisions recorded in this session.*")

        lines.extend([
            "",
            "## 📄 Newly Converted Document Dossiers",
            "",
        ])
        if new_tdocs:
            lines.extend([
                "| TDoc | WG | Meeting | Type | Title | Relative Dossier Path |",
                "| :--- | :--- | :--- | :--- | :--- | :--- |",
            ])
            for doc in new_tdocs:
                tid = doc["tdoc_id"]
                rel_p = doc["rel_path"]
                title_c = re.sub(r'[\r\n\t]+', ' ', doc.get("title", "")).replace("|", "\\|")
                lines.append(f"| **{tid}** | {doc.get('wg')} | {doc.get('meeting_number')} | {doc.get('type')} | {title_c} | [{rel_p}]({rel_p}) |")
        else:
            lines.append("*All relevant documents were already converted in earlier sessions.*")

        lines.append("")
        with open(delta_file, "w", encoding="utf-8") as f:
            f.write("\n".join(lines))

    def _write_master_index(
        self,
        target_dir: Path,
        manifest: ExportManifest,
        all_decisions: List[TopicDecision],
        all_candidates: List[Dict[str, Any]],
        session_id: int,
    ):
        """Generates master index.md."""
        now_str = datetime.datetime.now().strftime("%Y-%m-%d")
        lines = [
            f"# 🌐 Cross-WG Topic Dossier: {self.topic.name}",
            f"**Last Updated:** {now_str} (Session #{session_id}) | **Scope Mode:** `{self.export_scope_mode}`",
            "",
            "> [!NOTE]",
            f"> {self.topic.description or 'Consolidated multi-WG tracking across 3GPP Working Groups and quarterly cycles.'}",
            "",
            "## 📊 Executive KPI Strip",
            "",
            f"- **Participating Working Groups:** {', '.join(sorted(self.topic.target_wgs))}",
            f"- **Analyzed Meeting Scope:** {len(self.cohorts)} synchronized rounds ({self.cohorts[0].key} → {self.cohorts[-1].key})" if self.cohorts else "- **Analyzed Meeting Scope:** None",
            f"- **Unified Technical Decisions:** {len(all_decisions)} (Agreements, WAs, pCRs, CRs, WFs)",
            f"- **Total Scoped Contributions:** {len(all_candidates)} documents",
            "",
            "## 🧭 Navigation & Core Summaries",
            "",
            "- ⚖️ **[Unified Technical Decisions Baseline](agreements_timeline.md)** — Chronological matrix of all RAN agreements and SA/CT pCRs/CRs.",
            "- 📬 **[Cross-WG Liaison Statement Dialogue](ls_dialogue.md)** — Inter-group question and reply threads.",
            f"- 🧠 **[Latest LLM Delta Feed](deltas/delta_session_{session_id}.md)** — Incremental diff for feeding into AI prompts.",
            "",
            "## 🗓️ Synchronized Meeting Rounds Overview",
            "",
            "| Cycle | Covered WGs | Meetings | Decisions | Dossiers |",
            "| :--- | :--- | :--- | :--- | :--- |",
        ]

        for cohort in self.cohorts:
            c_decisions = [d for d in all_decisions if d.cohort_key == cohort.key]
            c_docs = [doc for doc in all_candidates if doc.get("cycle") == cohort.key]
            wgs_str = ", ".join(sorted(cohort.wgs_present))
            mtg_nums = ", ".join([f"{m.get('wg_name')} #{m.get('meeting_number')}" for m in cohort.meetings])
            lines.append(f"| **{cohort.key}** | {wgs_str} | {mtg_nums} | {len(c_decisions)} | {len(c_docs)} |")

        lines.extend([
            "",
            "---",
            "### 🔄 Incremental Export History",
            "",
            "| Session | Timestamp | Scope Mode | New Documents | Total Registered |",
            "| :--- | :--- | :--- | :--- | :--- |",
        ])
        for sess in manifest.sessions:
            s_id = sess.get("session_id")
            s_time = sess.get("timestamp")
            s_mode = sess.get("scope_mode")
            added_cnt = len(sess.get("added_tdocs", []))
            tot_cnt = sess.get("total_tdocs", 0)
            lines.append(f"| #{s_id} | {s_time} | `{s_mode}` | +{added_cnt} | {tot_cnt} |")

        lines.append("")
        with open(target_dir / "index.md", "w", encoding="utf-8") as f:
            f.write("\n".join(lines))

    def _write_consolidated_briefing(
        self,
        target_dir: Path,
        all_decisions: List[TopicDecision],
        all_candidates: List[Dict[str, Any]],
    ):
        """Generates single-file consolidated Markdown briefing for NotebookLM / LLM ingestion."""
        lines = [
            f"# Consolidated Executive Briefing: {self.topic.name}",
            "",
            "## 1. Executive Summary & Baseline",
            f"This briefing consolidates decisions across {', '.join(sorted(self.topic.target_wgs))} spanning {len(self.cohorts)} synchronized meeting rounds.",
            "",
            "## 2. Chronological Technical Decisions",
            "",
        ]
        for dec in sorted(all_decisions, key=lambda d: (d.cohort_key, d.date or "", d.wg)):
            lines.append(f"- **[{dec.cohort_key} | {dec.wg} {dec.meeting_number}]** ({dec.decision_type}): {dec.title_or_summary}")

        lines.append("")
        with open(target_dir / "consolidated_briefing.md", "w", encoding="utf-8") as f:
            f.write("\n".join(lines))
