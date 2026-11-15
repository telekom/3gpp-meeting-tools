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
import copy
import datetime
import hashlib
import json
import logging
import os
import re
import shutil
import time
import urllib.parse
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple

from PyQt5.QtCore import QThread, pyqtSignal

from core.utils.company_sanitizer import CompanySanitizer
from core.utils.office_printer_guard import temporary_virtual_default_printer
from core.utils.paths import to_extended_path, get_project_root
from modules.meetings.core.cohort_resolver import (
    MeetingCohort,
    format_round_label,
)
from modules.meetings.core.cross_wg_ls_tracker import CrossWgLiaisonTracker
from modules.meetings.core.exporters.export_manifest import (
    ExportManifest,
    compute_meeting_snapshot,
    diff_meeting_snapshots,
    is_export_complete,
    load_manifest,
    save_manifest,
)
from modules.meetings.core.exporters.topic_digest_builder import (
    STAGE2_REPORT_FILENAME,
    STAGE2_SYNTHESIS_FILENAME,
    build_delta_section_markdown,
    classify_meeting_lifecycle,
    format_lifecycle_badge,
    resolve_meeting_folder_id,
    rewrite_media_links,
    sync_meeting_brief_analysis_state,
    write_company_positions_input,
    write_consolidated_briefing_markdown,
    write_e2e_synthesis_prompt,
    write_master_index_markdown,
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
        excluded_tid_set = {str(t).strip().upper() for t in (self.topic.rules.excluded_tdocs or []) if str(t).strip()}

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

            prev_mtg_state = manifest.exported_meetings.get(meeting_rel_root, {})
            self._force_live_refresh = is_incremental and self._is_meeting_live_for_refresh(mtg, prev_mtg_state)
            try:
                tdocs_rows = self._load_meeting_tdocs(mtg)
                chair_agreements = self._load_meeting_agreements(mtg)
            finally:
                self._force_live_refresh = False

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
            withdrawn_tids: Set[str] = set()

            for row in tdocs_rows:
                tid = _normalize_tdoc_id(row.get("TDoc", "") or row.get("tdoc_id", "") or row.get("TD#", ""))
                if not tid:
                    continue
                if tid in excluded_tid_set:
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
                if "withdrawn" in status.lower():
                    withdrawn_tids.add(tid)

                source = str(row.get("Source", "") or row.get("source", "")).strip()
                abstract = str(row.get("Abstract", "") or row.get("abstract", "") or row.get("Secretary Remarks", "")).strip()
                ai_desc = str(row.get("Agenda item description", "") or row.get("AI Description", "")).strip()

                is_decision_ref = tid in decision_referenced_tids
                if not is_decision_ref and not self.topic.matches_tdoc(tid, wg_name, ai_val, title, wi_val, abstract, ai_description=ai_desc):
                    continue

                is_agreed = any(s in status.lower() for s in ("approved", "agreed", "endorsed"))
                is_withdrawn = "withdrawn" in status.lower()
                is_ls = "ls" in t_type.lower() or "liaison" in t_type.lower()
                is_referenced = is_decision_ref or is_agreed

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

            # B2. Prune superseded/withdrawn revisions from topic destination folder (keep global cache)
            if self.deduplicate_revisions:
                self._prune_superseded_topic_docs(
                    target_dir=target_dir,
                    manifest=manifest,
                    meeting_rel_root=meeting_rel_root,
                    active_tids={d["tdoc_id"] for d in meeting_candidates},
                    superseded_or_withdrawn=superseded_set | withdrawn_tids,
                )

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
            dec_snapshot_records: List[Dict[str, Any]] = []
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
                dec_snapshot_records.append({
                    "key": dec_key,
                    "type": dec.decision_type,
                    "tdoc_id": dec.tdoc_id or "",
                    "ai": dec.ai or "",
                    "rel_path": dec.dossier_rel_path or "",
                    "summary": dec.title_or_summary or "",
                })

            all_decisions.extend(meeting_decisions)
            all_candidates.extend(meeting_candidates)
            newly_converted_docs.extend(meeting_newly_converted)

            # F. Auto-classify meeting lifecycle, snapshot/delta check, LLM Digest & Meeting Brief
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
            max_digest_mtime = max(
                (os.path.getmtime(to_extended_path(p)) for p in digest_paths if Path(to_extended_path(p)).is_file()),
                default=time.time(),
            )

            active_tdocs_snap: Dict[str, Dict[str, Any]] = {}
            stub_count = 0
            for doc in meeting_candidates:
                abs_p = doc["abs_path"]
                is_stub = not is_export_complete(abs_p, doc_info=doc)
                if is_stub:
                    stub_count += 1
                doc["is_stub"] = is_stub
                active_tdocs_snap[doc["tdoc_id"]] = {
                    "status": doc.get("status", ""),
                    "is_stub": is_stub,
                    "supersedes": doc.get("supersedes", []),
                    "title": doc.get("title", ""),
                    "source": doc.get("source", ""),
                    "type": doc.get("type", ""),
                    "ai": doc.get("ai", ""),
                    "category": doc.get("category", ""),
                    "rel_path": doc.get("rel_path", ""),
                }

            prev_snap = prev_mtg_state.get("current_snapshot")
            analyzed_snap = copy.deepcopy(prev_mtg_state.get("analyzed_snapshot")) if prev_mtg_state.get("analyzed_snapshot") else None
            probe_snap = compute_meeting_snapshot(lifecycle, dec_snapshot_records, active_tdocs_snap, packet_mtime=max_digest_mtime)
            if prev_snap and prev_snap.get("fingerprint") == probe_snap["fingerprint"]:
                # Preserve original packet_mtime when exported content has not changed
                probe_snap["packet_mtime"] = float(prev_snap.get("packet_mtime", max_digest_mtime))

            analysis_path = meeting_dir / "_company_positions_analysis.md"
            analysis_exists = Path(to_extended_path(analysis_path)).is_file()
            analysis_mtime = os.path.getmtime(to_extended_path(analysis_path)) if analysis_exists else 0.0

            delta_info: Optional[Dict[str, Any]] = None
            if lifecycle == "no_activity":
                analysis_status = "no_activity"
            elif not analysis_exists:
                analyzed_snap = None
                analysis_status = "pending"
            else:
                if analyzed_snap is None and prev_snap:
                    prev_pkt_mtime = float(prev_snap.get("packet_mtime", 0.0) or 0.0)
                    if analysis_mtime >= prev_pkt_mtime - 5.0:
                        analyzed_snap = copy.deepcopy(prev_snap)
                elif analyzed_snap is None and not prev_snap:
                    if analysis_mtime >= probe_snap["packet_mtime"] - 5.0:
                        analyzed_snap = copy.deepcopy(probe_snap)

                if analyzed_snap and analyzed_snap.get("fingerprint") == probe_snap["fingerprint"]:
                    analysis_status = "analyzed"
                elif analyzed_snap:
                    if analysis_mtime > probe_snap["packet_mtime"]:
                        analyzed_snap = copy.deepcopy(probe_snap)
                        analysis_status = "analyzed"
                    else:
                        delta_info = diff_meeting_snapshots(analyzed_snap, probe_snap)
                        analysis_status = "stale" if delta_info.get("has_changes") else "analyzed"
                else:
                    analysis_status = "stale"

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
                analysis_status=analysis_status,
                analysis_exists=analysis_exists,
                delta_info=delta_info,
            )
            brief_rel = safe_relpath(brief_path, target_dir)

            ls_in_cnt = sum(1 for d in meeting_candidates if d.get("category") == "LS_IN")
            ls_out_cnt = sum(1 for d in meeting_candidates if d.get("category") == "LS_OUT")

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
                    "round_label": format_round_label(cohort.key),
                    "start_date": mtg_date,
                    "lifecycle": lifecycle,
                    "scope_mode": self.export_scope_mode,
                    "tdoc_count": len(meeting_candidates),
                    "decision_count": len(meeting_decisions),
                    "stub_count": stub_count,
                    "ls_in_count": ls_in_cnt,
                    "ls_out_count": ls_out_cnt,
                    "brief_rel_path": brief_rel,
                    "digest_rel_paths": digest_rel_paths,
                    "analysis_status": analysis_status,
                    "current_snapshot": probe_snap,
                    "analyzed_snapshot": analyzed_snap,
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
                "analysis_status": analysis_status,
                "tdoc_count": len(meeting_candidates),
                "decision_count": len(meeting_decisions),
                "stub_count": stub_count,
                "ls_in_count": ls_in_cnt,
                "ls_out_count": ls_out_cnt,
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

        synthesis_status = _evaluate_stage2_synthesis_state(target_dir, manifest)
        write_e2e_synthesis_prompt(
            target_dir,
            self.topic,
            meeting_records,
            synthesis_status=synthesis_status,
            synthesis_state=manifest.synthesis_state,
        )

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

        total_stubs = sum(r.get("stub_count", 0) for r in meeting_records)
        write_master_index_markdown(
            target_dir=target_dir,
            topic=self.topic,
            manifest=manifest,
            total_decisions=len(all_decisions),
            total_active_docs=len(all_candidates),
            stub_count=total_stubs,
            round_count=len(self.cohorts),
            session_id=session_id,
            scope_mode=self.export_scope_mode,
            meeting_records=meeting_records,
            synthesis_status=synthesis_status,
            synthesis_state=manifest.synthesis_state,
        )

        if self.generate_consolidated_md:
            write_consolidated_briefing_markdown(
                target_dir=target_dir,
                topic=self.topic,
                all_decisions=all_decisions,
                meeting_records=meeting_records,
                round_count=len(self.cohorts),
            )

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
        try:
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
        finally:
            self._executor = None

        # Phase 2: Sequential Conversion into Canonical Cache + Write-Out to Topic Destination
        # Wrap the entire batch once in temporary_virtual_default_printer so Windows default printer
        # isn't toggled repeatedly across 1000+ conversions.
        with temporary_virtual_default_printer(log=logger):
            for idx, doc in enumerate(docs_to_convert):
                if self.is_cancelled:
                    return
                self._convert_or_stub_single_doc(doc, idx, total, base_pct, pct_span)

    def _convert_or_stub_single_doc(
        self,
        doc: Dict[str, Any],
        idx: int,
        total: int,
        base_pct: int,
        pct_span: int,
    ) -> None:
        """Converts one downloaded/cached TDoc via shared canonical cache or writes a metadata stub."""
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
        mtg_id = mtg.get("id") or mtg.get("meeting_number")
        if not mtg_id:
            return []
        try:
            from modules.meetings.core.tdocs.tdocs_downloader import TDocsDownloaderThread
            dl = TDocsDownloaderThread(str(mtg_id), mtg_dir, cancel_check=lambda: self.is_cancelled)
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

        import json as _json
        if agenda_dir.is_dir():
            json_caches = sorted(
                [jf for jf in agenda_dir.glob("*.xlsx.json") if jf.is_file()],
                key=lambda p: p.stat().st_mtime,
                reverse=True,
            )
            for jf in json_caches:
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
        synthesis_status: Optional[str] = None,
    ):
        """Delegates master `index.md` generation to `write_master_index_markdown`."""
        recs = meeting_records or []
        total_stubs = sum(int(r.get("stub_count", 0) or 0) for r in recs)
        write_master_index_markdown(
            target_dir=target_dir,
            topic=self.topic,
            manifest=manifest,
            total_decisions=len(all_decisions),
            total_active_docs=len(all_candidates),
            stub_count=total_stubs,
            round_count=len(self.cohorts),
            session_id=session_id,
            scope_mode=self.export_scope_mode,
            meeting_records=recs,
            synthesis_status=synthesis_status,
            synthesis_state=manifest.synthesis_state,
        )

    def _write_consolidated_briefing(
        self,
        target_dir: Path,
        all_decisions: List[TopicDecision],
        all_candidates: List[Dict[str, Any]],
        meeting_records: Optional[List[Dict[str, Any]]] = None,
    ):
        """Delegates consolidated single-file briefing generation to `write_consolidated_briefing_markdown`."""
        write_consolidated_briefing_markdown(
            target_dir=target_dir,
            topic=self.topic,
            all_decisions=all_decisions,
            meeting_records=meeting_records or [],
            round_count=len(self.cohorts),
        )

    @staticmethod
    def check_and_refresh_dossier(
        target_dir: Path,
        topic: Optional[MetaTopicDefinition] = None,
    ) -> Dict[str, Any]:
        """Static alias for `check_and_refresh_dossier(target_dir, topic=topic)`."""
        return check_and_refresh_dossier(target_dir, topic=topic)


def _read_json_dict(path: Path) -> Dict[str, Any]:
    """Reads a UTF-8 JSON file into a dict, returning {} on missing/corrupt input."""
    try:
        if not path.is_file():
            return {}
        with open(to_extended_path(path), "r", encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def _extract_report_header_meta(report_path: Path) -> Dict[str, str]:
    """Extracts Package ID, Baseline ID, and cut-off date from a Stage-2 report header."""
    meta: Dict[str, str] = {}
    try:
        with open(to_extended_path(report_path), "r", encoding="utf-8", errors="replace") as f:
            head = f.read(3000)
    except Exception:
        return meta

    m_pkg = re.search(r"Package\s+ID:\s*`?([^`\r\n]+)`?", head, re.IGNORECASE)
    if m_pkg:
        meta["package_id"] = m_pkg.group(1).strip()
    m_base = re.search(r"Baseline\s+package\s+ID:\s*`?([^`\r\n]+)`?", head, re.IGNORECASE)
    if m_base:
        meta["baseline_package_id"] = m_base.group(1).strip()
    m_as_of = re.search(
        r"(?:Requested\s+cut-off(?:\s+and\s+review)?\s+date|As[- ]of\s+date):\s*`?(\d{4}-\d{2}-\d{2})`?",
        head,
        re.IGNORECASE,
    )
    if m_as_of:
        meta["as_of"] = m_as_of.group(1).strip()
    return meta


def _parse_stage2_subfolder(child_dir: Path) -> Optional[Dict[str, Any]]:
    """Inspects a candidate dated Stage-2 subfolder and returns normalized metadata."""
    ext_child = Path(to_extended_path(child_dir))
    if not ext_child.is_dir():
        return None

    folder_name = child_dir.name
    primary_rep = ext_child / STAGE2_REPORT_FILENAME
    legacy_rep = ext_child / STAGE2_SYNTHESIS_FILENAME
    working_dir = ext_child / "working"
    pkg_state_path = working_dir / "package_state.json"
    val_path = working_dir / "validation.json"
    cov_path = ext_child / "COVERAGE.md"

    is_stage2_named = folder_name.lower().startswith("stage2")
    has_artifacts = (
        primary_rep.is_file()
        or legacy_rep.is_file()
        or pkg_state_path.is_file()
        or (is_stage2_named and working_dir.is_dir())
    )
    if not has_artifacts:
        return None

    report_path: Optional[Path] = None
    if primary_rep.is_file():
        report_path = primary_rep
    elif legacy_rep.is_file():
        report_path = legacy_rep

    pkg_state = _read_json_dict(pkg_state_path)
    val_state = _read_json_dict(val_path)
    header_meta = _extract_report_header_meta(report_path) if report_path else {}

    pkg_id = str(pkg_state.get("package_id") or header_meta.get("package_id") or "").strip()
    baseline_id = str(
        pkg_state.get("baseline_package_id") or header_meta.get("baseline_package_id") or ""
    ).strip()

    as_of = str(pkg_state.get("as_of") or header_meta.get("as_of") or "").strip()
    if not as_of:
        m_date = re.search(r"(\d{4}-\d{2}-\d{2})", folder_name)
        if m_date:
            as_of = m_date.group(1)
        elif report_path:
            as_of = datetime.datetime.fromtimestamp(os.path.getmtime(report_path)).strftime("%Y-%m-%d")

    rev = ""
    for src in (pkg_id, folder_name):
        m_rev = re.search(r"(?:^|[-_])(r\d+)(?:$|[-_])", src, re.IGNORECASE)
        if m_rev:
            rev = m_rev.group(1).lower()
            break
    if not rev:
        rev = "r1"
    rev_num = int(rev[1:]) if rev.startswith("r") and rev[1:].isdigit() else 0

    report_mtime = float(os.path.getmtime(report_path)) if report_path else 0.0
    if report_path:
        status = "completed"
        report_rel = f"{folder_name}/{report_path.name}"
    elif working_dir.is_dir():
        status = "in_progress"
        report_rel = f"{folder_name}/{STAGE2_REPORT_FILENAME}"
    else:
        status = "pending"
        report_rel = f"{folder_name}/{STAGE2_REPORT_FILENAME}"

    coverage_rel = f"{folder_name}/COVERAGE.md" if cov_path.is_file() else ""
    zips = sorted(
        (z.name for z in ext_child.iterdir() if z.is_file() and z.suffix.lower() == ".zip"),
        key=lambda n: (0 if "handoff" in n.lower() else 1, n.lower()),
    )
    handoff_zip_rel = f"{folder_name}/{zips[0]}" if zips else ""

    return {
        "folder_name": folder_name,
        "status": status,
        "is_legacy_root": False,
        "package_id": pkg_id or folder_name,
        "baseline_package_id": baseline_id or "none supplied",
        "as_of": as_of,
        "revision": rev,
        "revision_num": rev_num,
        "report_rel_path": report_rel,
        "coverage_rel_path": coverage_rel,
        "handoff_zip_rel_path": handoff_zip_rel,
        "report_mtime": report_mtime,
        "validation_passed": val_state.get("passed") if "passed" in val_state else None,
        "claims": val_state.get("claims"),
        "synthesis_words": val_state.get("synthesis_words"),
        "actual_model": pkg_state.get("actual_model", ""),
        "input_manifest_sha256": pkg_state.get("input_manifest_sha256", ""),
    }


def discover_stage2_packages(
    target_dir: Path,
    excluded_dirs: Optional[Set[str]] = None,
) -> List[Dict[str, Any]]:
    """
    Discovers all Stage-2 synthesis packages inside `target_dir`:
      1. Dated subfolders (`stage2_YYYY-MM-DD_rN/`, `stage2*/`, or any non-WG subfolder
         containing `COMPANY_VIEWS_AND_PROGRESS.md`, `company_positions_e2e_synthesis.md`,
         or `working/package_state.json`).
      2. Legacy root `company_positions_e2e_synthesis.md` (if present).

    Returns packages sorted chronologically ascending (oldest first, latest last).
    """
    ext_target = Path(to_extended_path(target_dir))
    if not ext_target.is_dir():
        return []

    skip_names = {"deltas"}
    if excluded_dirs:
        skip_names.update(str(d).strip() for d in excluded_dirs if str(d).strip())

    packages: List[Dict[str, Any]] = []
    try:
        entries = list(ext_target.iterdir())
    except OSError:
        return []

    for child in entries:
        if not child.is_dir():
            continue
        name = child.name
        if name.startswith(".") or name in skip_names:
            continue
        parsed = _parse_stage2_subfolder(target_dir / name)
        if parsed:
            packages.append(parsed)

    root_legacy = ext_target / STAGE2_SYNTHESIS_FILENAME
    if root_legacy.is_file():
        mtime = float(os.path.getmtime(root_legacy))
        as_of = datetime.datetime.fromtimestamp(mtime).strftime("%Y-%m-%d")
        packages.append({
            "folder_name": "",
            "status": "completed",
            "is_legacy_root": True,
            "package_id": "legacy-root",
            "baseline_package_id": "none supplied",
            "as_of": as_of,
            "revision": "legacy",
            "revision_num": 0,
            "report_rel_path": STAGE2_SYNTHESIS_FILENAME,
            "coverage_rel_path": "",
            "handoff_zip_rel_path": "",
            "report_mtime": mtime,
            "validation_passed": None,
            "claims": None,
            "synthesis_words": None,
            "actual_model": "",
            "input_manifest_sha256": "",
        })

    packages.sort(
        key=lambda p: (
            0 if not p.get("is_legacy_root") else -1,
            str(p.get("as_of") or "9999-99-99"),
            int(p.get("revision_num", 0) or 0),
            float(p.get("report_mtime", 0.0) or 0.0),
            str(p.get("folder_name") or ""),
        )
    )
    return packages


def _evaluate_stage2_synthesis_state(target_dir: Path, manifest: ExportManifest) -> str:
    """
    Inspects dated Stage-2 synthesis subfolders and legacy root synthesis on disk against
    Stage-1 meeting snapshots/analyses and updates `manifest.synthesis_state` cleanly.
    Returns `"pending"`, `"in_progress"`, `"stale"`, or `"analyzed"`.
    """
    prev_state = dict(manifest.synthesis_state or {})
    prev_cov = dict(prev_state.get("covered_fingerprints") or {})

    excluded_wgs = {
        str(m_state.get("wg") or "").strip()
        for m_state in (manifest.exported_meetings or {}).values()
        if isinstance(m_state, dict) and m_state.get("wg")
    }
    packages = discover_stage2_packages(target_dir, excluded_dirs=excluded_wgs)
    completed_pkgs = [p for p in packages if p.get("status") == "completed"]
    in_progress_pkgs = [p for p in packages if p.get("status") == "in_progress"]

    if not completed_pkgs:
        status = "in_progress" if in_progress_pkgs else "pending"
        first_prog = in_progress_pkgs[-1] if in_progress_pkgs else {}
        manifest.synthesis_state = {
            "file": first_prog.get("report_rel_path") or STAGE2_SYNTHESIS_FILENAME,
            "folder_name": first_prog.get("folder_name", ""),
            "package_id": first_prog.get("package_id", ""),
            "baseline_package_id": first_prog.get("baseline_package_id", ""),
            "as_of": first_prog.get("as_of", ""),
            "revision": first_prog.get("revision", ""),
            "coverage_rel_path": first_prog.get("coverage_rel_path", ""),
            "handoff_zip_rel_path": first_prog.get("handoff_zip_rel_path", ""),
            "exists": False,
            "status": status,
            "analyzed_mtime": 0.0,
            "covered_fingerprints": prev_cov,
            "changed_meetings": [],
            "latest_package": None,
            "packages": packages,
        }
        return status

    latest_pkg = completed_pkgs[-1]
    syn_mtime = float(latest_pkg.get("report_mtime", 0.0) or 0.0)
    prev_folder = str(prev_state.get("folder_name", "") or "")
    prev_pkg_id = str(prev_state.get("package_id", "") or "")
    prev_syn_mtime = float(prev_state.get("analyzed_mtime", 0.0) or 0.0)

    # Compare covered_fingerprints only when inspecting the same Stage-2 report already tracked
    same_baseline_tracked = bool(
        prev_cov
        and prev_folder == str(latest_pkg.get("folder_name") or "")
        and (not prev_pkg_id or prev_pkg_id == str(latest_pkg.get("package_id") or ""))
        and syn_mtime <= prev_syn_mtime + 0.1
    )

    active_fps: Dict[str, str] = {}
    changed_meetings: Set[str] = set()

    for m_key, m_state in (manifest.exported_meetings or {}).items():
        if not isinstance(m_state, dict):
            continue
        lc = str(m_state.get("lifecycle", "")).strip().lower()
        if lc == "no_activity":
            continue

        snap = m_state.get("current_snapshot") or {}
        if snap.get("fingerprint"):
            active_fps[m_key] = str(snap["fingerprint"])

        if m_state.get("analysis_status") == "stale":
            changed_meetings.add(m_key)

        ext_mdir = Path(to_extended_path(target_dir / m_key))
        a_path = ext_mdir / "_company_positions_analysis.md"
        if a_path.is_file() and float(os.path.getmtime(a_path)) > syn_mtime + 0.5:
            changed_meetings.add(m_key)

        pkt_mtime = float(snap.get("packet_mtime", 0.0) or 0.0)
        if ext_mdir.is_dir():
            for pkt_file in ext_mdir.glob("_company_positions_input*.md"):
                pkt_mtime = max(pkt_mtime, float(os.path.getmtime(pkt_file)))
        if pkt_mtime > syn_mtime + 0.5:
            changed_meetings.add(m_key)

        if same_baseline_tracked and active_fps.get(m_key) != prev_cov.get(m_key):
            changed_meetings.add(m_key)

    if same_baseline_tracked:
        for old_k in set(prev_cov) - set(active_fps):
            changed_meetings.add(old_k)

    is_stale = bool(changed_meetings)
    status = "stale" if is_stale else "analyzed"
    keep_prev_cov = bool(is_stale and same_baseline_tracked)

    manifest.synthesis_state = {
        "file": latest_pkg["report_rel_path"],
        "folder_name": latest_pkg["folder_name"],
        "package_id": latest_pkg["package_id"],
        "baseline_package_id": latest_pkg["baseline_package_id"],
        "as_of": latest_pkg["as_of"],
        "revision": latest_pkg["revision"],
        "coverage_rel_path": latest_pkg["coverage_rel_path"],
        "handoff_zip_rel_path": latest_pkg["handoff_zip_rel_path"],
        "validation_passed": latest_pkg.get("validation_passed"),
        "claims": latest_pkg.get("claims"),
        "synthesis_words": latest_pkg.get("synthesis_words"),
        "exists": True,
        "status": status,
        "analyzed_mtime": prev_syn_mtime if keep_prev_cov and prev_syn_mtime > 0 else syn_mtime,
        "covered_fingerprints": prev_cov if keep_prev_cov else active_fps,
        "changed_meetings": sorted(changed_meetings),
        "latest_package": latest_pkg,
        "packages": packages,
    }
    return status


def _resolve_topic_for_checkup(
    target_dir: Path,
    manifest: ExportManifest,
    topic: Optional[MetaTopicDefinition] = None,
) -> MetaTopicDefinition:
    """Resolves a MetaTopicDefinition from caller, MetaTopicManager, or manifest fallback."""
    if topic is not None:
        return topic
    try:
        from modules.meetings.core.topic_manager import MetaTopicManager
        mgr = MetaTopicManager()
        matched = next((t for t in mgr.list_topics() if t.name == manifest.meeting), None)
        if matched is not None:
            return matched
    except Exception:
        pass
    wgs = sorted({str(m.get("wg", "")).upper() for m in (manifest.exported_meetings or {}).values() if isinstance(m, dict) and m.get("wg")})
    return MetaTopicDefinition(
        id="checkup",
        name=manifest.meeting or target_dir.name,
        description="",
        target_wgs=wgs or ["SA2", "RAN2", "RAN3", "CT1"],
    )


def _rebuild_meeting_state_for_checkup(
    target_dir: Path,
    manifest: ExportManifest,
    m_key: str,
    m_state: Dict[str, Any],
    tdocs_by_meeting: Dict[str, Dict[str, Dict[str, Any]]],
    decisions_by_fid: Dict[str, List[Dict[str, Any]]],
) -> Tuple[Dict[str, Any], Dict[str, Any]]:
    """
    Refreshes a single meeting's snapshot, 3-state lifecycle, Stage-1 analysis status,
    `_meeting_brief.md` delta block, and returns `(updated_m_state, meeting_record)`.
    """
    m_dir = target_dir / m_key
    m_fid = str(m_state.get("meeting_folder_id") or Path(m_key).name)
    wg = str(m_state.get("wg") or m_fid.split("_")[0]).upper()

    ext_mdir = Path(to_extended_path(m_dir))
    digest_files = (
        sorted([p for p in ext_mdir.glob("_company_positions_input*.md") if p.is_file()], key=lambda p: p.name)
        if ext_mdir.is_dir()
        else []
    )
    digest_rel_paths = (
        [safe_relpath(m_dir / p.name, target_dir) for p in digest_files]
        if digest_files
        else list(m_state.get("digest_rel_paths") or [])
    )
    packet_mtime = max((float(os.path.getmtime(p)) for p in digest_files), default=0.0)

    prev_snap = m_state.get("current_snapshot") if isinstance(m_state.get("current_snapshot"), dict) else {}
    prev_tdocs_snap = prev_snap.get("active_tdocs") if isinstance(prev_snap.get("active_tdocs"), dict) else {}
    raw_tdocs = tdocs_by_meeting.get(m_key, {})

    active_tdocs_snap: Dict[str, Dict[str, Any]] = {}
    stub_count = 0
    for raw_tid, t_info in raw_tdocs.items():
        tid = _normalize_tdoc_id(raw_tid)
        prev_t = (
            prev_tdocs_snap.get(tid)
            or prev_tdocs_snap.get(raw_tid)
            or prev_tdocs_snap.get(tid.upper())
            or {}
        ) if isinstance(prev_tdocs_snap, dict) else {}
        rel_p = str(t_info.get("rel_path") or prev_t.get("rel_path") or "")
        doc_probe = {
            "tdoc_id": tid,
            "title": t_info.get("title") or prev_t.get("title", ""),
            "type": t_info.get("type") or prev_t.get("type", ""),
        }
        is_stub = not is_export_complete(target_dir / rel_p, doc_info=doc_probe) if rel_p else bool(prev_t.get("is_stub", False))
        if is_stub:
            stub_count += 1
        active_tdocs_snap[tid] = {
            "status": prev_t.get("status") or t_info.get("status", ""),
            "is_stub": is_stub,
            "supersedes": list(prev_t.get("supersedes") or t_info.get("supersedes") or []),
            "title": doc_probe["title"],
            "source": t_info.get("source") or prev_t.get("source", ""),
            "type": doc_probe["type"],
            "ai": t_info.get("ai") or prev_t.get("ai", ""),
            "category": t_info.get("category") or prev_t.get("category", ""),
            "rel_path": rel_p,
        }

    raw_prev_decs = prev_snap.get("decisions")
    if isinstance(raw_prev_decs, dict):
        dec_records = [v for v in raw_prev_decs.values() if isinstance(v, dict)]
    elif isinstance(raw_prev_decs, list):
        dec_records = [v for v in raw_prev_decs if isinstance(v, dict)]
    else:
        dec_records = []
    if not dec_records:
        dec_records = list(decisions_by_fid.get(m_fid, []))
    tdoc_count = len(active_tdocs_snap)
    decision_count = len(dec_records) if dec_records else int(m_state.get("decision_count", 0) or 0)

    ls_in_cnt = (
        sum(1 for v in active_tdocs_snap.values() if v.get("category") == "LS_IN")
        if active_tdocs_snap
        else int(m_state.get("ls_in_count", 0) or 0)
    )
    ls_out_cnt = (
        sum(1 for v in active_tdocs_snap.values() if v.get("category") == "LS_OUT")
        if active_tdocs_snap
        else int(m_state.get("ls_out_count", 0) or 0)
    )

    if tdoc_count == 0 and decision_count == 0:
        lifecycle = "no_activity"
    else:
        raw_lc = str(m_state.get("lifecycle", "") or "").strip().lower()
        if raw_lc in ("", "no_activity"):
            lifecycle = "concluded" if decision_count > 0 else "upcoming"
        elif raw_lc == "upcoming" and decision_count > 0:
            lifecycle = "concluded"
        else:
            lifecycle = raw_lc

    probe_snap = compute_meeting_snapshot(lifecycle, dec_records, active_tdocs_snap, packet_mtime=packet_mtime)
    if prev_snap.get("fingerprint") == probe_snap["fingerprint"] and prev_snap.get("packet_mtime"):
        prev_pkt = float(prev_snap["packet_mtime"])
        probe_snap["packet_mtime"] = min(prev_pkt, packet_mtime) if packet_mtime > 0.0 else prev_pkt

    analysis_path = m_dir / "_company_positions_analysis.md"
    ext_analysis = Path(to_extended_path(analysis_path))
    analysis_exists = ext_analysis.is_file()
    analysis_mtime = float(os.path.getmtime(ext_analysis)) if analysis_exists else 0.0

    analyzed_snap = copy.deepcopy(m_state.get("analyzed_snapshot")) if isinstance(m_state.get("analyzed_snapshot"), dict) else None
    newly_promoted = False
    delta_md = ""

    if lifecycle == "no_activity":
        analysis_status = "no_activity"
    elif not analysis_exists:
        analyzed_snap = None
        analysis_status = "pending"
    else:
        if analyzed_snap is None:
            ref_pkt_mtime = float(prev_snap.get("packet_mtime") or probe_snap.get("packet_mtime") or 0.0)
            if ref_pkt_mtime <= 0.0 or analysis_mtime >= ref_pkt_mtime - 5.0:
                analyzed_snap = copy.deepcopy(probe_snap)
                newly_promoted = True
                analysis_status = "analyzed"
            else:
                analysis_status = "stale"
                delta_info = diff_meeting_snapshots({"lifecycle": "upcoming", "decisions": [], "active_tdocs": {}}, probe_snap)
                cand_docs = [
                    {**v, "tdoc_id": tid, "abs_path": target_dir / v["rel_path"]}
                    for tid, v in active_tdocs_snap.items()
                    if v.get("rel_path")
                ]
                delta_md = build_delta_section_markdown(
                    delta_info=delta_info,
                    wg=wg,
                    meeting_folder_id=m_fid,
                    eff_lifecycle=lifecycle,
                    candidates=cand_docs,
                    target_dir=target_dir,
                    digest_rel_paths=digest_rel_paths,
                )
        elif analyzed_snap.get("fingerprint") == probe_snap["fingerprint"]:
            analysis_status = "analyzed"
        elif analysis_mtime > float(probe_snap.get("packet_mtime", 0.0) or 0.0):
            analyzed_snap = copy.deepcopy(probe_snap)
            newly_promoted = True
            analysis_status = "analyzed"
        else:
            delta_info = diff_meeting_snapshots(analyzed_snap, probe_snap)
            if delta_info.get("has_changes"):
                analysis_status = "stale"
                cand_docs = [
                    {**v, "tdoc_id": tid, "abs_path": target_dir / v["rel_path"]}
                    for tid, v in active_tdocs_snap.items()
                    if tid in set(delta_info.get("highlight_tids") or []) and v.get("rel_path")
                ]
                delta_md = build_delta_section_markdown(
                    delta_info=delta_info,
                    wg=wg,
                    meeting_folder_id=m_fid,
                    eff_lifecycle=lifecycle,
                    candidates=cand_docs,
                    target_dir=target_dir,
                    digest_rel_paths=digest_rel_paths,
                )
            else:
                analysis_status = "analyzed"

    brief_rel = str(m_state.get("brief_rel_path") or f"{m_key}/_meeting_brief.md")
    sync_meeting_brief_analysis_state(
        brief_path=target_dir / brief_rel,
        lifecycle=lifecycle,
        analysis_status=analysis_status,
        analysis_exists=analysis_exists,
        delta_markdown=delta_md,
    )

    updated_state = dict(m_state)
    updated_state.update({
        "lifecycle": lifecycle,
        "tdoc_count": tdoc_count,
        "decision_count": decision_count,
        "stub_count": stub_count,
        "ls_in_count": ls_in_cnt,
        "ls_out_count": ls_out_cnt,
        "digest_rel_paths": digest_rel_paths,
        "analysis_status": analysis_status,
        "current_snapshot": probe_snap,
        "analyzed_snapshot": analyzed_snap,
    })

    rec = {
        "cohort_key": updated_state.get("cohort_key", ""),
        "round_label": updated_state.get("round_label") or format_round_label(updated_state.get("cohort_key", "")),
        "wg": wg,
        "meeting_number": updated_state.get("meeting_number", ""),
        "meeting_folder_id": m_fid,
        "meeting_rel_root": m_key,
        "start_date": updated_state.get("start_date", ""),
        "lifecycle": lifecycle,
        "analysis_status": analysis_status,
        "newly_promoted": newly_promoted,
        "tdoc_count": tdoc_count,
        "decision_count": decision_count,
        "stub_count": stub_count,
        "ls_in_count": ls_in_cnt,
        "ls_out_count": ls_out_cnt,
        "brief_rel_path": brief_rel,
        "digest_rel_paths": digest_rel_paths,
    }
    return updated_state, rec


def check_and_refresh_dossier(
    target_dir: Path,
    topic: Optional[MetaTopicDefinition] = None,
) -> Dict[str, Any]:
    """
    Fast offline checkup & non-destructive refresh for an existing Topic Dossier folder.

    Scans `.export_manifest.json` + on-disk `_company_positions_analysis.md`,
    `_company_positions_input*.md`, and `company_positions_e2e_synthesis.md` files
    to promote freshly completed analyses, preserve stale delta blocks, and
    regenerate `index.md`, `company_positions_e2e_prompt.md`, and `_meeting_brief.md`
    without overwriting any LLM analysis files or re-running document conversions.
    """
    target_dir = Path(target_dir)
    manifest = load_manifest(target_dir)
    if manifest is None:
        return {
            "ok": False,
            "target_dir": str(target_dir),
            "error": f"No .export_manifest.json found in {target_dir}",
        }

    eff_topic = _resolve_topic_for_checkup(target_dir, manifest, topic=topic)

    tdocs_by_meeting: Dict[str, Dict[str, Dict[str, Any]]] = {}
    meeting_keys = list((manifest.exported_meetings or {}).keys())
    for tid, t_info in (manifest.tdocs or {}).items():
        if not isinstance(t_info, dict):
            continue
        rel_p = str(t_info.get("rel_path", "")).replace("\\", "/")
        matched_key = next((mk for mk in meeting_keys if rel_p.startswith(f"{mk}/")), "")
        if matched_key:
            m_meta = manifest.exported_meetings.get(matched_key)
            if isinstance(m_meta, dict):
                if not t_info.get("wg") and m_meta.get("wg"):
                    t_info["wg"] = m_meta["wg"]
                if not t_info.get("meeting_number") and m_meta.get("meeting_number"):
                    t_info["meeting_number"] = m_meta["meeting_number"]
                if not t_info.get("meeting_folder_id") and m_meta.get("meeting_folder_id"):
                    t_info["meeting_folder_id"] = m_meta["meeting_folder_id"]
            tdocs_by_meeting.setdefault(matched_key, {})[tid] = t_info

    decisions_by_fid: Dict[str, List[Dict[str, Any]]] = {}
    for d_key, d_rel in (manifest.exported_decisions or {}).items():
        parts = str(d_key).split("::")
        if len(parts) >= 2:
            fid = parts[0]
            d_type = parts[1]
            tid_or_ai = parts[2] if len(parts) >= 3 else ""
            decisions_by_fid.setdefault(fid, []).append({
                "key": d_key,
                "type": d_type,
                "tdoc_id": tid_or_ai if "-" in tid_or_ai else "",
                "ai": "" if "-" in tid_or_ai else tid_or_ai,
                "rel_path": d_rel or "",
                "summary": "",
            })

    meeting_records: List[Dict[str, Any]] = []
    newly_promoted_count = 0
    for m_key in meeting_keys:
        m_state = manifest.exported_meetings[m_key]
        if not isinstance(m_state, dict):
            continue
        updated_state, rec = _rebuild_meeting_state_for_checkup(
            target_dir=target_dir,
            manifest=manifest,
            m_key=m_key,
            m_state=m_state,
            tdocs_by_meeting=tdocs_by_meeting,
            decisions_by_fid=decisions_by_fid,
        )
        manifest.exported_meetings[m_key] = updated_state
        if rec.pop("newly_promoted", False):
            newly_promoted_count += 1
        meeting_records.append(rec)

    meeting_records.sort(
        key=lambda r: (
            str(r.get("start_date") or ""),
            str(r.get("wg") or ""),
            str(r.get("meeting_folder_id") or ""),
        )
    )

    synthesis_status = _evaluate_stage2_synthesis_state(target_dir, manifest)
    write_e2e_synthesis_prompt(
        target_dir,
        eff_topic,
        meeting_records,
        synthesis_status=synthesis_status,
        synthesis_state=manifest.synthesis_state,
    )

    total_decisions = sum(int(r.get("decision_count", 0) or 0) for r in meeting_records)
    total_active_docs = sum(int(r.get("tdoc_count", 0) or 0) for r in meeting_records)
    total_stubs = sum(int(r.get("stub_count", 0) or 0) for r in meeting_records)
    round_count = len({r.get("cohort_key") for r in meeting_records if r.get("cohort_key")}) or 1
    latest_sess = manifest.sessions[-1] if manifest.sessions else {}
    session_id = int(latest_sess.get("session_id", len(manifest.sessions) or 1))
    scope_mode = str(latest_sess.get("scope_mode") or manifest.export_type.replace("topic_", "") or "full_dossier")

    write_master_index_markdown(
        target_dir=target_dir,
        topic=eff_topic,
        manifest=manifest,
        total_decisions=total_decisions,
        total_active_docs=total_active_docs,
        stub_count=total_stubs,
        round_count=round_count,
        session_id=session_id,
        scope_mode=scope_mode,
        meeting_records=meeting_records,
        synthesis_status=synthesis_status,
        synthesis_state=manifest.synthesis_state,
    )
    save_manifest(target_dir, manifest)

    analyzed_cnt = sum(1 for r in meeting_records if r.get("analysis_status") == "analyzed")
    stale_cnt = sum(1 for r in meeting_records if r.get("analysis_status") == "stale")
    pending_cnt = sum(1 for r in meeting_records if r.get("analysis_status") == "pending")
    no_act_cnt = sum(1 for r in meeting_records if r.get("analysis_status") == "no_activity")

    return {
        "ok": True,
        "target_dir": str(target_dir),
        "topic_name": eff_topic.name,
        "total_meetings": len(meeting_records),
        "analyzed": analyzed_cnt,
        "stale": stale_cnt,
        "pending": pending_cnt,
        "no_activity": no_act_cnt,
        "newly_promoted": newly_promoted_count,
        "synthesis_status": synthesis_status,
        "synthesis_state": manifest.synthesis_state,
        "total_decisions": total_decisions,
        "total_active_docs": total_active_docs,
        "total_stubs": total_stubs,
        "session_id": session_id,
    }
