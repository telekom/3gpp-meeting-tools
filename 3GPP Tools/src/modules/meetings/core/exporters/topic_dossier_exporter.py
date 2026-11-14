# --- File: src/modules/meetings/core/exporters/topic_dossier_exporter.py ---
"""
Cross-WG, Multi-Meeting Topic Dossier & Briefing Exporter Thread.

Performs chronological per-meeting discovery, shared canonical Markdown cache
read/write-back, revision lineage deduplication, unified decision extraction
(RAN Chairman's Notes text agreements + SA/CT approved pCRs/CRs), cross-WG
Liaison Statement threading, deterministic per-meeting LLM digests, and
atomic per-meeting checkpointing in `.export_manifest.json`.
"""

from concurrent.futures import ThreadPoolExecutor, as_completed
import datetime
import hashlib
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
    format_round_label,
)
from modules.meetings.core.cross_wg_ls_tracker import CrossWgLiaisonTracker
from modules.meetings.core.exporters.export_manifest import (
    ExportManifest,
    is_export_complete,
    load_manifest,
    save_manifest,
)
from modules.meetings.core.exporters.topic_digest_builder import (
    classify_meeting_lifecycle,
    format_lifecycle_badge,
    resolve_meeting_folder_id,
    rewrite_media_links,
    write_company_positions_input,
    write_e2e_synthesis_prompt,
    write_meeting_brief,
)
from modules.meetings.core.meetings_db import MeetingsDatabase
from modules.meetings.core.tdocs.chair_notes_agreements_parser import (
    ChairNotesAgreementsParser,
    clean_chair_notes_text,
)
from modules.meetings.core.tdocs.tdoc_file_handler import TDocFileHandler
from modules.meetings.core.tdocs.tdoc_markdown_converter import (
    CONVERTER_VERSION,
    convert_document_to_markdown,
    is_conversion_stale,
    resolve_revision_mode_for_type,
    sanitize_filename,
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


def _normalize_tdoc_id(raw_tid: str) -> str:
    """Normalizes TDoc IDs to 3GPP convention: uppercase prefix/number + lowercase revision suffix ('R2-2600999r01')."""
    s = str(raw_tid or "").strip().upper()
    return re.sub(
        r'^([A-Z0-9]+-\d+)-?(?:R|REV)(\d{1,2}[A-Z]?)$',
        lambda m: f"{m.group(1)}r{m.group(2).lower()}",
        s,
    )


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


def _resolve_doc_server_url(
    tid: str,
    meeting_dict: Optional[Dict[str, Any]] = None,
    row_data: Optional[Dict[str, Any]] = None,
) -> str:
    """Builds the authoritative 3GPP HTTPS archive URL using meeting docs/main URLs."""
    mtg = meeting_dict or {}
    merged_meta: Dict[str, Any] = {}
    merged_meta.update(mtg)
    if row_data:
        merged_meta.update(row_data)
    return resolve_tdoc_download_url(
        tid,
        docs_url=mtg.get("docs_folder_url"),
        main_url=mtg.get("url_key"),
        meta=merged_meta,
    )


def _decision_checkpoint_key(meeting_folder_id: str, dec: TopicDecision) -> str:
    """Generates a deterministic manifest key for a TopicDecision across incremental sessions."""
    if dec.tdoc_id and not dec.is_text_agreement:
        return f"{meeting_folder_id}::{dec.decision_type}::{dec.tdoc_id}"
    digest = hashlib.sha1(dec.title_or_summary.strip().encode("utf-8")).hexdigest()[:10]
    return f"{meeting_folder_id}::{dec.decision_type}::{dec.tdoc_id or dec.ai or 'AI'}::{digest}"


def _copy_markdown_with_media(
    src_md: Path,
    dst_md: Path,
    src_media_dir: Optional[Path],
    src_media_rel: str,
    dst_media_dir: Path,
    dst_media_rel: str,
) -> bool:
    """
    Copies a Markdown file to `dst_md`, rewriting relative image/media links from
    `src_media_rel/` to `dst_media_rel/` and copying referenced media assets cleanly.
    """
    try:
        src_ext = Path(to_extended_path(src_md))
        dst_ext = Path(to_extended_path(dst_md))
        if not src_ext.is_file():
            return False

        raw_text = src_ext.read_text(encoding="utf-8", errors="replace")
        rewritten = rewrite_media_links(raw_text, src_media_rel, dst_media_rel)

        os.makedirs(dst_ext.parent, exist_ok=True)
        dst_ext.write_text(rewritten, encoding="utf-8")

        if src_media_dir and Path(to_extended_path(src_media_dir)).is_dir():
            src_m_ext = Path(to_extended_path(src_media_dir))
            # Determine which media filenames are referenced in rewritten markdown
            refs: Set[str] = set()
            for m in re.finditer(r'!\[[^\]]*\]\(([^)]+)\)|src=["\']([^"\']+)["\']', raw_text):
                ref = (m.group(1) or m.group(2) or "").strip().replace("\\", "/")
                if ref and not ref.startswith(("http://", "https://", "data:")):
                    refs.add(os.path.basename(ref))

            copied_any = False
            dst_m_ext = Path(to_extended_path(dst_media_dir))
            candidates = (
                [src_m_ext / name for name in refs if (src_m_ext / name).is_file()]
                if refs
                else [f for f in src_m_ext.glob("*.*") if f.is_file()]
            )
            if candidates:
                os.makedirs(dst_m_ext, exist_ok=True)
                for mf in candidates:
                    shutil.copy2(mf, dst_m_ext / mf.name)
                    copied_any = True
            elif not refs:
                # Even if markdown contained no media refs, avoid leaving empty folders
                pass
            _ = copied_any

        return dst_ext.is_file() and dst_ext.stat().st_size > 0
    except Exception as e:
        logger.debug(f"[CopyWithMedia] Failed copying {src_md} -> {dst_md}: {e}")
        return False


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
        os.makedirs(to_extended_path(target_dir), exist_ok=True)

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

        # 3. Build chronological flattened meeting list across Synchronized Rounds
        meeting_schedule: List[Tuple[MeetingCohort, Dict[str, Any]]] = []
        for cohort in self.cohorts:
            for mtg in cohort.meetings:
                meeting_schedule.append((cohort, mtg))

        meeting_schedule.sort(
            key=lambda item: (
                str(item[1].get("start_date", "") or ""),
                str(item[1].get("wg_name", "")).upper(),
                resolve_meeting_folder_id(item[1]),
            )
        )

        ls_tracker = CrossWgLiaisonTracker()
        all_decisions: List[TopicDecision] = []
        new_decisions_this_session: List[TopicDecision] = []
        all_candidates: List[Dict[str, Any]] = []
        newly_converted_docs: List[Dict[str, Any]] = []
        meeting_records: List[Dict[str, Any]] = []

        already_complete_count = 0
        reused_from_cache_count = 0
        total_meetings = len(meeting_schedule)

        # 4. Chronological Per-Meeting Processing & Atomic Checkpointing
        for m_idx, (cohort, mtg) in enumerate(meeting_schedule):
            if self.is_cancelled:
                save_manifest(target_dir, manifest)
                return

            wg_name = str(mtg.get("wg_name", "")).strip().upper()
            wg_clean = sanitize_folder_name(wg_name or "WG")
            mtg_num = str(mtg.get("meeting_number", "")).strip()
            folder_name = str(mtg.get("folder_name", "")).strip()
            mtg_date = str(mtg.get("start_date", "")).strip()
            meeting_folder_id = resolve_meeting_folder_id(mtg)
            meeting_rel_root = f"{wg_clean}/{meeting_folder_id}"
            meeting_dir = target_dir / wg_clean / meeting_folder_id

            base_pct = int((m_idx / max(1, total_meetings)) * 82)
            self.progress_msg.emit(
                f"Processing {wg_name} #{mtg_num} ({meeting_folder_id}) [{m_idx + 1}/{total_meetings}]..."
            )
            self.progress_step.emit(base_pct, 100)

            tdocs_rows = self._load_meeting_tdocs(mtg)
            chair_agreements = self._load_meeting_agreements(mtg)

            superseded_set, predecessors_map, _leaf_map = self._build_revision_lineage_map(tdocs_rows)
            decision_referenced_tids: Set[str] = set()
            meeting_decisions: List[TopicDecision] = []

            # A. Chairman's Notes Agreements (strictly RAN-scoped inside _load_meeting_agreements)
            for ca in chair_agreements:
                wi_str = str(ca.get("wi", "") or ca.get("study_item", "")).strip()
                ai_str = str(ca.get("ai", "") or ca.get("agenda_item", "")).strip()
                text_str = clean_chair_notes_text(ca.get("text", "") or ca.get("summary", ""))

                if self._matches_topic_rule(wg_name, ai_str, text_str, wi_str):
                    dec_type = normalize_chair_notes_decision_type(ca.get("decision_type", "Agreement"))
                    ca_tid = _normalize_tdoc_id(ca.get("tdoc_id") or "") or None
                    if ca_tid:
                        decision_referenced_tids.add(ca_tid)
                    dec = TopicDecision(
                        cohort_key=cohort.key,
                        wg=wg_name,
                        meeting_number=mtg_num,
                        date=mtg_date,
                        decision_type=dec_type,
                        title_or_summary=text_str,
                        tdoc_id=ca_tid,
                        source=ca.get("source") or f"Chairman's Notes AI {ai_str}",
                        target_spec=ca.get("spec", ""),
                        ai=ai_str,
                        wi=wi_str,
                        status=ca.get("status", "Agreed"),
                        notes=ca.get("notes", ""),
                        meeting_folder_id=meeting_folder_id,
                    )
                    meeting_decisions.append(dec)

            # B. Filter & Match Meeting TDocs
            deduped_skipped = 0
            meeting_candidates: List[Dict[str, Any]] = []

            for row in tdocs_rows:
                tid = _normalize_tdoc_id(row.get("TDoc", "") or row.get("tdoc_id", "") or row.get("TD#", ""))
                if not tid:
                    continue
                title = str(row.get("Title", "") or row.get("title", "")).strip()
                ai_val = str(
                    row.get("Agenda Item", "")
                    or row.get("Agenda item", "")
                    or row.get("AI", "")
                    or row.get("ai", "")
                    or row.get("AI#", "")
                ).strip()
                wi_val = str(
                    row.get("Related WIs", "")
                    or row.get("wi", "")
                    or row.get("WI", "")
                    or row.get("Work Item", "")
                ).strip()
                t_type = str(row.get("Type", "") or row.get("type", "") or row.get("TDoc Type", "")).strip()
                status = str(
                    row.get("TDoc Status", "")
                    or row.get("Status", "")
                    or row.get("status", "")
                    or row.get("Result", "")
                ).strip()
                source = str(row.get("Source", "") or row.get("source", "")).strip()
                abstract = str(row.get("Abstract", "") or row.get("abstract", "") or row.get("Secretary Remarks", "")).strip()
                ai_desc = str(row.get("Agenda item description", "") or row.get("AI Description", "")).strip()

                if not self.topic.matches_tdoc(tid, wg_name, ai_val, title, wi_val, abstract, ai_description=ai_desc):
                    continue

                is_agreed = any(s in status.lower() for s in ("approved", "agreed", "endorsed"))
                is_withdrawn = "withdrawn" in status.lower()
                is_ls = "ls" in t_type.lower() or "liaison" in t_type.lower()
                is_referenced = tid in decision_referenced_tids or is_agreed

                if self.export_scope_mode == "post_meeting" and not (is_agreed or is_ls or is_referenced):
                    continue

                if self.deduplicate_revisions and not is_referenced:
                    if is_withdrawn or tid in superseded_set:
                        deduped_skipped += 1
                        continue

                cat_folder = self._resolve_category_folder(t_type, status)
                if not self.include_discussions and cat_folder == "Discussion" and not (is_agreed or is_ls or is_referenced):
                    continue

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
                        source_url=_resolve_doc_server_url(tid, mtg, row),
                        meeting_folder_id=meeting_folder_id,
                    )
                    meeting_decisions.append(dec)

                safe_tid = sanitize_filename(tid)
                rel_md_path = f"{meeting_rel_root}/{cat_folder}/{safe_tid}.md"
                abs_md_path = target_dir / rel_md_path

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

                meeting_candidates.append({
                    "tdoc_id": tid,
                    "wg": wg_name,
                    "meeting_number": mtg_num,
                    "meeting_folder_id": meeting_folder_id,
                    "folder_name": folder_name,
                    "cycle": cohort.key,
                    "date": mtg_date,
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
                    "supersedes": predecessors_map.get(tid, []),
                })

            # C. Resolve Already Exported / Canonical Cache / Legacy Migration BEFORE FTP Download
            meeting_docs_to_convert: List[Dict[str, Any]] = []
            meeting_newly_converted: List[Dict[str, Any]] = []

            for doc in meeting_candidates:
                tid = doc["tdoc_id"]
                abs_p = doc["abs_path"]
                expected_mode = resolve_revision_mode_for_type(doc.get("type", ""), default_mode="markup")
                local_bin = self._find_local_binary_fast(doc.get("folder_name", ""), tid)

                # 1. Check destination file in new topology
                if manifest.is_tdoc_exported(tid) and is_export_complete(abs_p, doc_info=doc):
                    if not is_conversion_stale(local_bin, abs_p, expected_version=CONVERTER_VERSION, expected_mode=expected_mode):
                        already_complete_count += 1
                        self._backfill_canonical_cache(doc, abs_p)
                        continue

                # 2. Check shared canonical meeting Markdown cache (before network download!)
                if self._restore_from_canonical_cache(doc, local_bin=local_bin, expected_mode=expected_mode):
                    reused_from_cache_count += 1
                    meeting_newly_converted.append(doc)
                    continue

                # 3. Check legacy YYYY-QN export path in target_dir (seamless layout migration)
                if self._migrate_legacy_export_file(doc, target_dir, manifest, local_bin=local_bin, expected_mode=expected_mode):
                    reused_from_cache_count += 1
                    meeting_newly_converted.append(doc)
                    continue

                meeting_docs_to_convert.append(doc)

            # D. Convert only truly missing / stale documents for this meeting
            if meeting_docs_to_convert:
                self._convert_docs_pipeline(
                    meeting_docs_to_convert,
                    target_dir,
                    base_pct=base_pct,
                    pct_span=max(1, int(82 / max(1, total_meetings))),
                )
                for doc in meeting_docs_to_convert:
                    if doc["abs_path"].is_file():
                        meeting_newly_converted.append(doc)

            if self.is_cancelled:
                save_manifest(target_dir, manifest)
                return

            # E. Link formal decisions & export standalone RAN agreements
            seq_counters: Dict[str, int] = {}
            for dec in meeting_decisions:
                if dec.tdoc_id and not dec.dossier_rel_path and dec.decision_type in ("Approved pCR", "Approved CR", "Endorsed WF", "Agreed LS"):
                    for doc in meeting_candidates:
                        if doc["tdoc_id"] == dec.tdoc_id and doc["abs_path"].is_file():
                            dec.dossier_rel_path = doc["rel_path"]
                            break

                if not dec.dossier_rel_path:
                    key = f"{meeting_folder_id}_{dec.ai}"
                    seq_counters[key] = seq_counters.get(key, 0) + 1
                    dec.dossier_rel_path = self._write_agreement_markdown(
                        target_dir=target_dir,
                        dec=dec,
                        candidate_docs=meeting_candidates,
                        seq=seq_counters[key],
                    )

                dec_key = _decision_checkpoint_key(meeting_folder_id, dec)
                if not manifest.is_decision_exported(dec_key):
                    new_decisions_this_session.append(dec)
                manifest.record_decision(dec_key, dec.dossier_rel_path)

            all_decisions.extend(meeting_decisions)
            all_candidates.extend(meeting_candidates)
            newly_converted_docs.extend(meeting_newly_converted)

            # F. Auto-classify meeting lifecycle & generate Brought->Decided LLM Digest & Meeting Brief
            lifecycle = classify_meeting_lifecycle(mtg, meeting_decisions, meeting_candidates)
            digest_paths = write_company_positions_input(
                meeting_dir=meeting_dir,
                topic_name=self.topic.name,
                mtg=mtg,
                meeting_folder_id=meeting_folder_id,
                decisions=meeting_decisions,
                candidates=meeting_candidates,
                lifecycle=lifecycle,
            )
            digest_rel_paths = [safe_relpath(p, target_dir) for p in digest_paths]

            brief_path = write_meeting_brief(
                meeting_dir=meeting_dir,
                topic_name=self.topic.name,
                cohort_label=format_round_label(cohort.key),
                mtg=mtg,
                meeting_folder_id=meeting_folder_id,
                decisions=meeting_decisions,
                candidates=meeting_candidates,
                deduped_count=deduped_skipped,
                digest_rel_paths=digest_rel_paths,
                lifecycle=lifecycle,
            )
            brief_rel = safe_relpath(brief_path, target_dir)

            # G. Atomic Per-Meeting Manifest Checkpoint (resumability boundary)
            valid_entries = [d for d in meeting_candidates if d["abs_path"].is_file()]
            manifest.record_meeting_checkpoint(
                meeting_key=meeting_rel_root,
                checkpoint_data={
                    "wg": wg_name,
                    "meeting_number": mtg_num,
                    "meeting_folder_id": meeting_folder_id,
                    "folder_name": folder_name,
                    "cohort_key": cohort.key,
                    "start_date": mtg_date,
                    "lifecycle": lifecycle,
                    "scope_mode": self.export_scope_mode,
                    "tdoc_count": len(meeting_candidates),
                    "decision_count": len(meeting_decisions),
                    "brief_rel_path": brief_rel,
                    "digest_rel_paths": digest_rel_paths,
                },
                entries=valid_entries,
                session_id=session_id,
                timestamp=session_timestamp,
            )
            save_manifest(target_dir, manifest)

            meeting_records.append({
                "cohort_key": cohort.key,
                "round_label": format_round_label(cohort.key),
                "wg": wg_name,
                "meeting_number": mtg_num,
                "meeting_folder_id": meeting_folder_id,
                "meeting_rel_root": meeting_rel_root,
                "start_date": mtg_date,
                "lifecycle": lifecycle,
                "tdoc_count": len(meeting_candidates),
                "decision_count": len(meeting_decisions),
                "ls_in_count": sum(1 for d in meeting_candidates if d.get("category") == "LS_IN"),
                "ls_out_count": sum(1 for d in meeting_candidates if d.get("category") == "LS_OUT"),
                "brief_rel_path": brief_rel,
                "digest_rel_paths": digest_rel_paths,
                "deduped_skipped": deduped_skipped,
            })

        if self.is_cancelled:
            return

        # 5. Synthesize Root Indices, Timelines, E2E Synthesis Guide & Delta Feed
        self.progress_msg.emit("Synthesizing Master Timeline Index, LS Tracker, and LLM Delta Feed...")
        self.progress_step.emit(88, 100)

        self._write_agreements_timeline(target_dir, all_decisions)

        ls_report_md = ls_tracker.generate_markdown_report()
        with open(to_extended_path(target_dir / "ls_dialogue.md"), "w", encoding="utf-8") as f:
            f.write(ls_report_md)

        write_e2e_synthesis_prompt(target_dir, self.topic, meeting_records)

        newly_exported_tids = [d["tdoc_id"] for d in newly_converted_docs if d["abs_path"].is_file()]
        self._write_llm_delta_feed(
            target_dir=target_dir,
            session_id=session_id,
            session_timestamp=session_timestamp,
            new_decisions=new_decisions_this_session if is_incremental else all_decisions,
            new_tdocs=newly_converted_docs,
            ls_tracker=ls_tracker,
        )

        # Record session before writing master index so session table includes this session
        manifest.record_session(
            session_id=session_id,
            timestamp=session_timestamp,
            scope_mode=self.export_scope_mode,
            added_tdocs=newly_exported_tids,
            updated_tdocs=[],
            added_entries=[d for d in newly_converted_docs if d["abs_path"].is_file()],
        )

        self._write_master_index(
            target_dir=target_dir,
            manifest=manifest,
            all_decisions=all_decisions,
            all_candidates=all_candidates,
            session_id=session_id,
            meeting_records=meeting_records,
        )

        if self.generate_consolidated_md:
            self._write_consolidated_briefing(target_dir, all_decisions, all_candidates, meeting_records)

        save_manifest(target_dir, manifest)

        self.progress_step.emit(100, 100)
        summary = (
            f"Export completed for '{self.topic.name}'!\n"
            f"• {len(all_decisions)} Unified Decisions (Agreements, WAs, pCRs, CRs)\n"
            f"• {len(all_candidates)} Active Contributions across {len(meeting_records)} Meetings ({len(self.cohorts)} Round(s))\n"
            f"• {len(newly_exported_tids)} Exported in Session #{session_id} "
            f"({reused_from_cache_count} reused from shared cache, {already_complete_count} already complete)\n"
            f"• Per-meeting briefs & LLM packets written to {{WG}}/{{Meeting}}/\n"
            f"• LLM Delta feed written to deltas/delta_session_{session_id}.md"
        )
        self.session_finished.emit(True, str(target_dir), summary)

    # ------------------------------------------------------------------
    # Revision Lineage & Shared Canonical Markdown Cache Helpers
    # ------------------------------------------------------------------

    @staticmethod
    def _build_revision_lineage_map(
        rows: List[Dict[str, Any]],
    ) -> Tuple[Set[str], Dict[str, List[str]], Dict[str, str]]:
        """
        Builds per-meeting revision lineage relationships:
        - superseded_set: TDoc IDs with a newer active (non-withdrawn) revision in `rows`.
        - predecessors_map: latest_leaf_tdoc_id -> sorted list of superseded predecessors.
        - leaf_map: tdoc_id -> latest_active_leaf_tdoc_id.
        """
        tdoc_map = {
            _normalize_tdoc_id(r.get("TDoc", "") or r.get("tdoc_id", "") or r.get("TD#", "")): r
            for r in rows
            if _normalize_tdoc_id(r.get("TDoc", "") or r.get("tdoc_id", "") or r.get("TD#", ""))
        }
        forward_edges: Dict[str, Set[str]] = {}

        for tdoc_id, row in tdoc_map.items():
            rev_to = str(row.get("Revised to", "") or "").strip()
            if rev_to and rev_to.lower() != "none":
                for cand in re.findall(r'[A-Za-z0-9]+-\d+(?:r\d+[a-zA-Z]?)?', rev_to):
                    cu = _normalize_tdoc_id(cand)
                    if cu in tdoc_map and cu != tdoc_id:
                        forward_edges.setdefault(tdoc_id, set()).add(cu)

            rev_of = str(row.get("Is revision of", "") or row.get("Revision of", "") or "").strip()
            if rev_of and rev_of.lower() != "none":
                for cand in re.findall(r'[A-Za-z0-9]+-\d+(?:r\d+[a-zA-Z]?)?', rev_of):
                    cu = _normalize_tdoc_id(cand)
                    if cu in tdoc_map and cu != tdoc_id:
                        forward_edges.setdefault(cu, set()).add(tdoc_id)

            base_match = re.search(r'^(.*?)-?(?:r|rev)\d{1,2}[a-zA-Z]?$', tdoc_id, re.IGNORECASE)
            if base_match:
                base_tdoc = base_match.group(1).upper()
                if base_tdoc in tdoc_map and base_tdoc != tdoc_id:
                    forward_edges.setdefault(base_tdoc, set()).add(tdoc_id)

        def _is_active(tid: str) -> bool:
            r = tdoc_map.get(tid)
            if not r:
                return False
            st = str(r.get("TDoc Status", "") or r.get("Status", "") or r.get("status", "")).lower()
            return "withdrawn" not in st

        leaf_map: Dict[str, str] = {}
        superseded_set: Set[str] = set()
        predecessors_map: Dict[str, List[str]] = {}

        for tdoc_id in tdoc_map:
            if not _is_active(tdoc_id):
                continue
            curr = tdoc_id
            visited = {curr}
            while True:
                children = [c for c in forward_edges.get(curr, set()) if c not in visited and _is_active(c)]
                if not children:
                    break
                children.sort(key=str)
                next_node = children[-1]
                visited.add(next_node)
                curr = next_node

            leaf_map[tdoc_id] = curr
            if curr != tdoc_id:
                superseded_set.add(tdoc_id)
                preds = predecessors_map.setdefault(curr, [])
                if tdoc_id not in preds:
                    preds.append(tdoc_id)

        for preds in predecessors_map.values():
            preds.sort()

        return superseded_set, predecessors_map, leaf_map

    def _get_canonical_tdoc_folder(self, folder_name: str, tid: str) -> Optional[Path]:
        mtg_dir = self._get_meeting_dir(folder_name)
        if not mtg_dir:
            return None
        base_match = re.search(r'^(.*?)-?(?:r|rev)\d{1,2}[a-zA-Z]?$', tid, re.IGNORECASE)
        base_tdoc = base_match.group(1).upper() if base_match else tid.upper()
        return mtg_dir / base_tdoc

    def _find_local_binary_fast(self, folder_name: str, tid: str) -> Optional[Path]:
        """Checks extracted office binaries already on disk (zero unzip, zero network I/O)."""
        mtg_dir = self._get_meeting_dir(folder_name)
        if not mtg_dir or not mtg_dir.is_dir():
            return None
        base_match = re.search(r'^(.*?)-?(?:r|rev)\d{1,2}[a-zA-Z]?$', tid, re.IGNORECASE)
        base_tdoc = base_match.group(1).upper() if base_match else tid.upper()
        for folder in [mtg_dir / tid, mtg_dir / base_tdoc, mtg_dir, mtg_dir / "Docs", mtg_dir / "Revisions"]:
            if not folder.is_dir():
                continue
            for ext in (".docx", ".docm", ".pptx", ".doc", ".ppt", ".pdf"):
                for name in (tid, base_tdoc):
                    f = folder / f"{name}{ext}"
                    if f.is_file() and f.stat().st_size > 0:
                        return f
        return None

    def _backfill_canonical_cache(self, doc: Dict[str, Any], topic_md: Path) -> None:
        """Seeds canonical `{cache_dir}/{folder_name}/{base_tdoc}/{tid}.md` if missing."""
        tid = doc["tdoc_id"]
        folder_name = doc.get("folder_name", "")
        if not folder_name:
            return
        tdoc_folder = self._get_canonical_tdoc_folder(folder_name, tid)
        if not tdoc_folder:
            return
        safe_tid = sanitize_filename(tid)
        canonical_md = tdoc_folder / f"{safe_tid}.md"
        if Path(to_extended_path(canonical_md)).is_file():
            return
        src_media = topic_md.parent / f"{topic_md.stem}_media"
        _copy_markdown_with_media(
            src_md=topic_md,
            dst_md=canonical_md,
            src_media_dir=src_media if src_media.is_dir() else None,
            src_media_rel=f"{topic_md.stem}_media",
            dst_media_dir=tdoc_folder / "media",
            dst_media_rel="media",
        )

    def _restore_from_canonical_cache(
        self,
        doc: Dict[str, Any],
        local_bin: Optional[Path],
        expected_mode: str,
    ) -> bool:
        """
        Checks canonical meeting Markdown cache before downloading/converting.
        Copies valid `.md` and rewrites `media/` links -> `{stem}_media/` into `doc["abs_path"]`.
        """
        folder_name = doc.get("folder_name", "")
        if not folder_name:
            return False
        mtg_dir = self._get_meeting_dir(folder_name)
        if not mtg_dir:
            return False

        tid = doc["tdoc_id"]
        safe_tid = sanitize_filename(tid)
        tdoc_folder = self._get_canonical_tdoc_folder(folder_name, tid)
        abs_p = doc["abs_path"]

        cand_files: List[Tuple[Path, Path, str]] = []
        if tdoc_folder and tdoc_folder.is_dir():
            cand_files.append((tdoc_folder / f"{safe_tid}.md", tdoc_folder / "media", "media"))
            if safe_tid != tid:
                cand_files.append((tdoc_folder / f"{tid}.md", tdoc_folder / "media", "media"))
            cand_files.append((tdoc_folder / f"{tid}_LLM_v{CONVERTER_VERSION}.md", tdoc_folder / "media", "media"))
        tid_dir = mtg_dir / tid
        if tid_dir.is_dir() and tid_dir != tdoc_folder:
            cand_files.append((tid_dir / f"{safe_tid}.md", tid_dir / "media", "media"))

        for mdf, m_dir, m_rel in cand_files:
            if not Path(to_extended_path(mdf)).is_file():
                continue
            if is_conversion_stale(local_bin, mdf, expected_version=CONVERTER_VERSION, expected_mode=expected_mode):
                continue
            ok = _copy_markdown_with_media(
                src_md=mdf,
                dst_md=abs_p,
                src_media_dir=m_dir,
                src_media_rel=m_rel,
                dst_media_dir=abs_p.parent / f"{abs_p.stem}_media",
                dst_media_rel=f"{abs_p.stem}_media",
            )
            if ok:
                # Ensure standard canonical `{safe_tid}.md` also exists if restored from legacy `_LLM_v...`
                if tdoc_folder and mdf.name != f"{safe_tid}.md":
                    canonical_md = tdoc_folder / f"{safe_tid}.md"
                    if not Path(to_extended_path(canonical_md)).is_file():
                        _copy_markdown_with_media(mdf, canonical_md, m_dir, m_rel, tdoc_folder / "media", "media")
                return True
        return False

    def _migrate_legacy_export_file(
        self,
        doc: Dict[str, Any],
        target_dir: Path,
        manifest: ExportManifest,
        local_bin: Optional[Path],
        expected_mode: str,
    ) -> bool:
        """Migrates valid Markdown files from legacy `YYYY-QN/` export layout into `{WG}/{MeetingFolder}/` + canonical cache."""
        tid = doc["tdoc_id"]
        safe_tid = sanitize_filename(tid)
        legacy_candidates: List[Path] = []

        m_entry = manifest.tdocs.get(tid)
        if m_entry and m_entry.get("rel_path") and m_entry["rel_path"] != doc["rel_path"]:
            legacy_candidates.append(target_dir / m_entry["rel_path"])

        if doc.get("cycle"):
            old_dir = f"{sanitize_folder_name(doc['wg'])}_{sanitize_folder_name(doc['meeting_number'])}"
            legacy_candidates.append(target_dir / sanitize_folder_name(doc["cycle"]) / old_dir / doc["category"] / f"{safe_tid}.md")

        abs_p = doc["abs_path"]
        for leg_p in legacy_candidates:
            if not Path(to_extended_path(leg_p)).is_file():
                continue
            if is_conversion_stale(local_bin, leg_p, expected_version=CONVERTER_VERSION, expected_mode=expected_mode):
                continue
            leg_media = leg_p.parent / f"{leg_p.stem}_media"
            ok = _copy_markdown_with_media(
                src_md=leg_p,
                dst_md=abs_p,
                src_media_dir=leg_media if leg_media.is_dir() else None,
                src_media_rel=f"{leg_p.stem}_media",
                dst_media_dir=abs_p.parent / f"{abs_p.stem}_media",
                dst_media_rel=f"{abs_p.stem}_media",
            )
            if ok:
                self._backfill_canonical_cache(doc, abs_p)
                return True
        return False

    # ------------------------------------------------------------------
    # Download & Two-Phase Conversion Pipeline
    # ------------------------------------------------------------------

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

        # 3. If missing and network download is not enabled, return None
        if not self.allow_network_download or self.is_cancelled:
            return None

        from modules.meetings.core.tdocs.tdoc_url_resolver import normalize_3gpp_url

        candidate_urls = []
        is_rev = bool(re.search(r'^(.*?)-?(?:r|rev)\d{1,2}[a-zA-Z]?$', tid, re.IGNORECASE))
        docs_url = meeting_dict.get("docs_folder_url")
        url_key = meeting_dict.get("url_key")

        resolved_primary = resolve_tdoc_download_url(
            tid,
            docs_url=docs_url,
            main_url=url_key,
            meta=meeting_dict,
        )
        if resolved_primary:
            candidate_urls.append(resolved_primary.rsplit('/', 1)[0])

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

        dedup_urls = list(dict.fromkeys(u.rstrip('/') for u in candidate_urls if u))

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

    def _convert_docs_pipeline(
        self,
        docs_to_convert: List[Dict[str, Any]],
        target_dir: Path,
        base_pct: int = 30,
        pct_span: int = 50,
    ):
        """Two-phase parallel pre-download and sequential conversion with canonical cache write-back."""
        total = len(docs_to_convert)
        if total == 0:
            return

        # Phase 1: Polite Concurrent Pre-Download (max_workers=5)
        self.progress_msg.emit(f"Resolving source documents ({total} files)...")
        with ThreadPoolExecutor(max_workers=5) as executor:
            self._executor = executor
            future_to_doc = {
                executor.submit(
                    self._find_or_download_tdoc_file,
                    doc["tdoc_id"],
                    doc["folder_name"],
                    doc.get("meeting_dict", {}),
                ): doc
                for doc in docs_to_convert
            }
            done_count = 0
            for fut in as_completed(future_to_doc):
                if self.is_cancelled:
                    return
                done_count += 1
                self.progress_step.emit(base_pct + int((done_count / total) * (pct_span * 0.45)), 100)
                try:
                    doc = future_to_doc[fut]
                    doc["source_file"] = fut.result()
                except Exception as e:
                    logger.debug(f"[PreDownload] Failed item: {e}")

        # Phase 2: Sequential Conversion into Canonical Cache + Write-Out to Topic Destination
        for idx, doc in enumerate(docs_to_convert):
            if self.is_cancelled:
                return
            tid = doc["tdoc_id"]
            abs_p = doc["abs_path"]
            src_file = doc.get("source_file")
            mtg_dict = doc.get("meeting_dict", {})
            row = doc.get("row_data", {})
            expected_mode = resolve_revision_mode_for_type(doc.get("type", ""), default_mode="markup")

            self.progress_msg.emit(f"Converting ({idx + 1}/{total}): {tid}...")
            self.progress_step.emit(base_pct + int((0.45 + 0.55 * (idx / total)) * pct_span), 100)

            local_bin = Path(src_file) if src_file else None
            converted_ok = self._restore_from_canonical_cache(doc, local_bin=local_bin, expected_mode=expected_mode)

            if not converted_ok and src_file and os.path.exists(src_file):
                tdoc_url = _resolve_doc_server_url(tid, mtg_dict, row)
                merged_meta = dict(row or {})
                merged_meta.setdefault("docs_folder_url", mtg_dict.get("docs_folder_url"))
                merged_meta.setdefault("url_key", mtg_dict.get("url_key"))
                if doc.get("supersedes"):
                    merged_meta["Supersedes"] = ", ".join(doc["supersedes"])

                tdoc_folder = self._get_canonical_tdoc_folder(doc.get("folder_name", ""), tid)
                try:
                    if tdoc_folder:
                        # Convert into shared canonical cache first, then copy with rewritten media links
                        os.makedirs(to_extended_path(tdoc_folder), exist_ok=True)
                        canon_media = tdoc_folder / "media"
                        res_doc = convert_document_to_markdown(
                            doc_path=Path(src_file),
                            out_dir=tdoc_folder,
                            tdoc_id=tid,
                            meta=merged_meta,
                            revisions=expected_mode,
                            media_dir=canon_media,
                            media_rel="media",
                            tdoc_url=tdoc_url,
                            cancellation_check=lambda: self.is_cancelled,
                        )
                        canon_md = Path(res_doc.md_path) if res_doc and getattr(res_doc, "md_path", None) else (tdoc_folder / f"{sanitize_filename(tid)}.md")
                        converted_ok = _copy_markdown_with_media(
                            src_md=canon_md,
                            dst_md=abs_p,
                            src_media_dir=canon_media,
                            src_media_rel="media",
                            dst_media_dir=abs_p.parent / f"{abs_p.stem}_media",
                            dst_media_rel=f"{abs_p.stem}_media",
                        )
                    else:
                        # Fallback when no meeting cache folder exists (e.g. unit test mocks)
                        os.makedirs(to_extended_path(abs_p.parent), exist_ok=True)
                        media_dir = abs_p.parent / f"{abs_p.stem}_media"
                        convert_document_to_markdown(
                            doc_path=Path(src_file),
                            out_dir=abs_p.parent,
                            tdoc_id=tid,
                            meta=merged_meta,
                            revisions=expected_mode,
                            media_dir=media_dir,
                            media_rel=f"{abs_p.stem}_media",
                            tdoc_url=tdoc_url,
                            cancellation_check=lambda: self.is_cancelled,
                        )
                        converted_ok = abs_p.is_file() and abs_p.stat().st_size > 0
                except Exception as e:
                    logger.warning(f"[Convert] Failed to convert {tid}: {e}")

            if not converted_ok and not abs_p.is_file():
                try:
                    os.makedirs(to_extended_path(abs_p.parent), exist_ok=True)
                    src_url = _resolve_doc_server_url(tid, mtg_dict, row)
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
                    with open(to_extended_path(abs_p), "w", encoding="utf-8") as sf:
                        sf.write(stub_content)
                except Exception as stub_err:
                    logger.warning(f"[Stub] Failed to create stub for {tid}: {stub_err}")

    # ------------------------------------------------------------------
    # Topic Matching, Cache Loading & Agreement Export
    # ------------------------------------------------------------------

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

    def _load_meeting_tdocs(self, mtg: Dict[str, Any]) -> List[Dict[str, Any]]:
        folder_name = mtg.get("folder_name", "")
        mtg_dir = self._get_meeting_dir(folder_name)
        if not mtg_dir:
            return []

        agenda_dir = mtg_dir / "Agenda"

        import json as _json
        if agenda_dir.is_dir():
            for jf in agenda_dir.glob("*.xlsx.json"):
                try:
                    with open(jf, "r", encoding="utf-8") as f:
                        return _json.load(f)
                except Exception:
                    pass

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
        # Chairman's Notes text agreements are strictly RAN-only; SA/CT triage notes must not become Text Agreements
        wg_upper = str(mtg.get("wg_name", "")).strip().upper()
        if "RAN" not in wg_upper:
            return []

        folder_name = mtg.get("folder_name", "")
        mtg_dir = self._get_meeting_dir(folder_name)
        if not mtg_dir:
            return []

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

        if self.allow_network_download and not self.is_cancelled:
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
        """Exports a standalone Markdown document for a RAN Chairman's Notes text agreement."""
        wg_clean = sanitize_folder_name(dec.wg)
        mtg_clean = sanitize_folder_name(dec.meeting_number)
        meeting_folder_id = dec.meeting_folder_id or f"{wg_clean}_{mtg_clean}"

        agreements_dir = target_dir / wg_clean / meeting_folder_id / "Agreements"
        os.makedirs(to_extended_path(agreements_dir), exist_ok=True)

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
        rel_path = f"{wg_clean}/{meeting_folder_id}/Agreements/{filename}"

        bg_link = "—"
        bg_doc = None
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

        portal_url = dec.source_url or (
            _resolve_doc_server_url(dec.tdoc_id, bg_doc.get("meeting_dict") if bg_doc else None, bg_doc.get("row_data") if bg_doc else None)
            if dec.tdoc_id
            else ""
        )
        portal_link = f"[{dec.tdoc_id}]({portal_url})" if portal_url else "—"

        lines = [
            f"# 📜 3GPP Text Agreement: {dec.wg} #{dec.meeting_number} (AI {dec.ai or 'N/A'})",
            "",
            "| Metadata Attribute | Value |",
            "| :--- | :--- |",
            f"| **Working Group** | **{dec.wg}** |",
            f"| **Meeting** | #{dec.meeting_number} (`{meeting_folder_id}`, {dec.date or 'N/A'}) |",
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
            f"- **Synchronized Round:** {format_round_label(dec.cohort_key)}",
            f"- **Meeting Record:** Documented under Agenda Item **{dec.ai or 'N/A'}** in the Chairman's Notes of 3GPP **{dec.wg}** Meeting **#{dec.meeting_number}**.",
            "",
        ])

        try:
            with open(to_extended_path(out_file), "w", encoding="utf-8") as f:
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
            "| Meeting Date | WG | Type | Substantive Decision / Scope | Contribution | Target Spec |",
            "| :--- | :--- | :--- | :--- | :--- | :--- |",
        ]

        sorted_decisions = sorted(decisions, key=lambda d: (d.date or d.cohort_key, d.wg, d.decision_type))
        for dec in sorted_decisions:
            lines.append(dec.to_markdown_row())

        lines.append("")
        with open(to_extended_path(target_dir / "agreements_timeline.md"), "w", encoding="utf-8") as f:
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
        os.makedirs(to_extended_path(delta_dir), exist_ok=True)
        delta_file = delta_dir / f"delta_session_{session_id}.md"

        lines = [
            f"# 🧠 Incremental AI Delta Feed — Session #{session_id}",
            f"**Generated:** {session_timestamp} | **Topic:** {self.topic.name}",
            "",
            "> [!IMPORTANT] LLM Prompting Guidance",
            "> This document contains strictly the **new decisions and contributions** added during this update session.",
            "> Feed this file alongside `company_positions_e2e_prompt.md` and per-meeting `_company_positions_input.md` files for token-efficient analysis.",
            "",
            "## 🆕 Newly Added Technical Decisions",
            "",
        ]

        if new_decisions:
            lines.extend([
                "| Date | WG | Meeting | Decision Type | Summary / Scope | Contribution |",
                "| :--- | :--- | :--- | :--- | :--- | :--- |",
            ])
            for d in new_decisions:
                doc_link = f"[{d.tdoc_id or 'Agreement'}](../{d.dossier_rel_path})" if d.dossier_rel_path else (d.tdoc_id or "—")
                summary_clean = re.sub(r'[\r\n\t]+', ' ', d.title_or_summary).replace("|", "\\|")
                m_lbl = d.meeting_folder_id or f"{d.wg}_{d.meeting_number}"
                lines.append(f"| {d.date or d.cohort_key} | **{d.wg}** | {m_lbl} | {d.badge_icon} {d.decision_type} | {summary_clean} | {doc_link} |")
        else:
            lines.append("*No new decisions recorded in this session.*")

        lines.extend([
            "",
            "## 📄 Newly Exported / Restored Document Dossiers",
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
                m_lbl = doc.get("meeting_folder_id") or doc.get("meeting_number", "")
                title_c = re.sub(r'[\r\n\t]+', ' ', doc.get("title", "")).replace("|", "\\|")
                lines.append(f"| **{tid}** | {doc.get('wg')} | {m_lbl} | {doc.get('type')} | {title_c} | [{rel_p}](../{rel_p}) |")
        else:
            lines.append("*All relevant documents were already exported in earlier sessions.*")

        lines.append("")
        with open(to_extended_path(delta_file), "w", encoding="utf-8") as f:
            f.write("\n".join(lines))

    def _write_master_index(
        self,
        target_dir: Path,
        manifest: ExportManifest,
        all_decisions: List[TopicDecision],
        all_candidates: List[Dict[str, Any]],
        session_id: int,
        meeting_records: Optional[List[Dict[str, Any]]] = None,
    ):
        """Generates master index.md with chronological cross-WG timeline and per-WG progression."""
        now_str = datetime.datetime.now().strftime("%Y-%m-%d")
        meeting_records = meeting_records or []

        lines = [
            f"# 🌐 Cross-WG Topic Dossier: {self.topic.name}",
            f"**Last Updated:** {now_str} (Session #{session_id}) | **Scope Mode:** `{self.export_scope_mode}`",
            "",
            "> [!NOTE]",
            f"> {self.topic.description or 'Consolidated multi-WG tracking organized by Working Group and Meeting.'}",
            "",
            "## 📊 Executive KPI Strip",
            "",
            f"- **Participating Working Groups:** {', '.join(sorted(self.topic.target_wgs))}",
            f"- **Analyzed Meetings:** {len(meeting_records)} meetings across {len(self.cohorts)} synchronized round(s)" if meeting_records else "- **Analyzed Meetings:** None",
            f"- **Unified Technical Decisions:** {len(all_decisions)} (Agreements, WAs, pCRs, CRs, WFs)",
            f"- **Active Scoped Contributions:** {len(all_candidates)} documents",
            "",
            "## 🧭 Navigation & LLM Workflows",
            "",
            "- ⚖️ **[Unified Technical Decisions Baseline](agreements_timeline.md)** — Chronological matrix of all RAN agreements and SA/CT pCRs/CRs.",
            "- 📬 **[Cross-WG Liaison Statement Dialogue](ls_dialogue.md)** — Inter-group question and reply threads.",
            "- 🧠 **[2-Stage E2E Company Position Prompt & Checklist](company_positions_e2e_prompt.md)** — Resumable, per-meeting LLM instructions designed for standard/non-frontier models.",
            f"- 🆕 **[Latest LLM Delta Feed](deltas/delta_session_{session_id}.md)** — Incremental additions in this session.",
            "",
            "## 🗓️ Cross-WG Chronological Meeting Timeline",
            "",
            "| Date | Synchronized Round | WG | Meeting | State | Decisions | Active Inputs | LS (In/Out) | Meeting Brief & LLM Packet |",
            "| :--- | :--- | :--- | :--- | :--- | :---: | :---: | :---: | :--- |",
        ]

        for rec in meeting_records:
            d_str = rec.get("start_date") or "N/A"
            r_lbl = rec.get("round_label") or rec.get("cohort_key", "")
            wg = rec["wg"]
            m_fid = rec["meeting_folder_id"]
            lc = rec.get("lifecycle") or ("concluded" if rec.get("decision_count", 0) else "upcoming")
            lc_badge = format_lifecycle_badge(lc)
            brief_rel = rec["brief_rel_path"]
            digests = rec.get("digest_rel_paths", [])
            digest_str = " · ".join(f"[{Path(p).name}]({p})" for p in digests) if digests else "—"
            ls_cell = f"{rec.get('ls_in_count', 0)} / {rec.get('ls_out_count', 0)}"
            lines.append(
                f"| `{d_str}` | {r_lbl} | **{wg}** | [{m_fid}]({brief_rel}) | {lc_badge} | {rec['decision_count']} | {rec['tdoc_count']} | {ls_cell} | [Brief]({brief_rel}) · {digest_str} |"
            )

        # Per-WG Progression Summary
        lines.extend([
            "",
            "## 🏛️ Per-Working-Group Evolution",
            "",
            "| Working Group | Chronological Meetings (Click to Open Brief) | Total Decisions | Total Active Docs |",
            "| :--- | :--- | :---: | :---: |",
        ])

        by_wg: Dict[str, List[Dict[str, Any]]] = {}
        for rec in meeting_records:
            by_wg.setdefault(rec["wg"], []).append(rec)

        for wg in sorted(by_wg.keys()):
            recs = by_wg[wg]
            links = " → ".join(f"[{r['meeting_folder_id']}]({r['brief_rel_path']})" for r in recs)
            tot_d = sum(r["decision_count"] for r in recs)
            tot_c = sum(r["tdoc_count"] for r in recs)
            lines.append(f"| **{wg}** | {links} | {tot_d} | {tot_c} |")

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
        with open(to_extended_path(target_dir / "index.md"), "w", encoding="utf-8") as f:
            f.write("\n".join(lines))

    def _write_consolidated_briefing(
        self,
        target_dir: Path,
        all_decisions: List[TopicDecision],
        all_candidates: List[Dict[str, Any]],
        meeting_records: Optional[List[Dict[str, Any]]] = None,
    ):
        """Generates single-file consolidated Markdown briefing for NotebookLM / LLM ingestion."""
        meeting_records = meeting_records or []
        lines = [
            f"# Consolidated Executive Briefing: {self.topic.name}",
            "",
            "## 1. Executive Summary & Scope",
            f"This briefing consolidates decisions and contribution briefs across {', '.join(sorted(self.topic.target_wgs))} "
            f"spanning {len(meeting_records)} meetings ({len(self.cohorts)} synchronized round(s)).",
            "",
            "## 2. Chronological Technical Decisions",
            "",
        ]
        for dec in sorted(all_decisions, key=lambda d: (d.date or d.cohort_key, d.wg, d.tdoc_id or "")):
            m_lbl = dec.meeting_folder_id or f"{dec.wg}_{dec.meeting_number}"
            link_suffix = f" ([{dec.tdoc_id or 'Agreement'}]({dec.dossier_rel_path}))" if dec.dossier_rel_path else ""
            lines.append(f"- **[{dec.date or dec.cohort_key} | {m_lbl}]** ({dec.decision_type}){link_suffix}: {dec.title_or_summary}")

        if meeting_records:
            lines.extend(["", "## 3. Per-Meeting Briefs & Company Position Digests", ""])
            for rec in meeting_records:
                digests = rec.get("digest_rel_paths", [])
                digest_links = ", ".join(f"[{Path(p).name}]({p})" for p in digests) if digests else "—"
                lines.append(
                    f"- **{rec.get('start_date') or 'N/A'} · {rec['meeting_folder_id']}**: "
                    f"[{Path(rec['brief_rel_path']).name}]({rec['brief_rel_path']}) · Digests: {digest_links}"
                )

        lines.append("")
        with open(to_extended_path(target_dir / "consolidated_briefing.md"), "w", encoding="utf-8") as f:
            f.write("\n".join(lines))
