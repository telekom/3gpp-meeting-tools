"""Main Topic Dossier multi-cohort export pipeline mixin."""

from __future__ import annotations

import datetime
import os
import time
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple

from core.utils.paths import to_extended_path
from modules.meetings.core.cohort_resolver import MeetingCohort
from modules.meetings.core.cross_wg_ls_tracker import CrossWgLiaisonTracker
from modules.meetings.core.exporters.export_manifest import (
    ExportManifest,
    is_export_complete,
    load_manifest,
    save_manifest,
)
from modules.meetings.core.exporters.topic_digest_builder import (
    resolve_meeting_folder_id,
    write_consolidated_briefing_markdown,
    write_e2e_synthesis_prompt,
    write_master_index_markdown,
)
from modules.meetings.core.tdocs.tdoc_markdown_converter import (
    CONVERTER_VERSION,
    is_conversion_stale,
    resolve_revision_mode_for_type,
)
from modules.meetings.core.topic_decision import TopicDecision

from ..path_utils import sanitize_folder_name
from ..stage2_discovery import _evaluate_stage2_synthesis_state
from .meeting_processor import TopicDossierMeetingProcessorMixin


class TopicDossierExportPipelineMixin(TopicDossierMeetingProcessorMixin):
    """Orchestrates incremental cohort discovery, document conversion, and dossier synthesis."""

    def _run_export(self):
        t_start = time.perf_counter()
        self.progress_msg.emit("Initializing Topic Dossier Studio...")
        self.progress_step.emit(0, 100)

        folder_stem = self.target_folder_name or f"Topic_{sanitize_folder_name(self.topic.name)}"
        target_dir = self.output_base_dir / folder_stem
        os.makedirs(to_extended_path(target_dir), exist_ok=True)

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

        meeting_schedule: List[Tuple[MeetingCohort, Dict[str, Any]]] = [
            (cohort, mtg) for cohort in self.cohorts for mtg in cohort.meetings
        ]
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
        excluded_tid_set: Set[str] = {
            str(t).strip().upper()
            for t in (self.topic.rules.excluded_tdocs or [])
            if str(t).strip()
        }

        already_complete_count = 0
        reused_from_cache_count = 0
        total_meetings = len(meeting_schedule)

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

            meeting_decisions, meeting_candidates, deduped_skipped = (
                self._collect_meeting_decisions_and_candidates(
                    cohort=cohort,
                    mtg=mtg,
                    wg_name=wg_name,
                    mtg_num=mtg_num,
                    mtg_date=mtg_date,
                    folder_name=folder_name,
                    meeting_folder_id=meeting_folder_id,
                    meeting_rel_root=meeting_rel_root,
                    target_dir=target_dir,
                    manifest=manifest,
                    tdocs_rows=tdocs_rows,
                    chair_agreements=chair_agreements,
                    excluded_tid_set=excluded_tid_set,
                    ls_tracker=ls_tracker,
                )
            )

            meeting_newly_converted, done_delta, reused_delta = self._resolve_and_convert_meeting_docs(
                meeting_candidates=meeting_candidates,
                target_dir=target_dir,
                manifest=manifest,
                base_pct=base_pct,
                total_meetings=total_meetings,
            )
            already_complete_count += done_delta
            reused_from_cache_count += reused_delta

            if self.is_cancelled:
                save_manifest(target_dir, manifest)
                return

            record = self._finalize_meeting_checkpoint(
                cohort=cohort,
                mtg=mtg,
                wg_name=wg_name,
                mtg_num=mtg_num,
                mtg_date=mtg_date,
                folder_name=folder_name,
                meeting_folder_id=meeting_folder_id,
                meeting_rel_root=meeting_rel_root,
                meeting_dir=meeting_dir,
                target_dir=target_dir,
                manifest=manifest,
                prev_mtg_state=prev_mtg_state,
                meeting_decisions=meeting_decisions,
                meeting_candidates=meeting_candidates,
                deduped_skipped=deduped_skipped,
                session_id=session_id,
                session_timestamp=session_timestamp,
                all_decisions=all_decisions,
                new_decisions_this_session=new_decisions_this_session,
            )

            all_candidates.extend(meeting_candidates)
            newly_converted_docs.extend(meeting_newly_converted)
            meeting_records.append(record)

        if self.is_cancelled:
            return

        self._synthesize_root_outputs(
            target_dir=target_dir,
            manifest=manifest,
            session_id=session_id,
            session_timestamp=session_timestamp,
            is_incremental=is_incremental,
            all_decisions=all_decisions,
            new_decisions_this_session=new_decisions_this_session,
            all_candidates=all_candidates,
            newly_converted_docs=newly_converted_docs,
            meeting_records=meeting_records,
            ls_tracker=ls_tracker,
            reused_from_cache_count=reused_from_cache_count,
            already_complete_count=already_complete_count,
            t_start=t_start,
        )

    def _resolve_and_convert_meeting_docs(
        self,
        *,
        meeting_candidates: List[Dict[str, Any]],
        target_dir: Path,
        manifest: ExportManifest,
        base_pct: int,
        total_meetings: int,
    ) -> Tuple[List[Dict[str, Any]], int, int]:
        """Resolve already exported/cached docs and run conversion for missing/stale items."""
        already_complete_delta = 0
        reused_delta = 0
        meeting_docs_to_convert: List[Dict[str, Any]] = []
        meeting_newly_converted: List[Dict[str, Any]] = []

        for doc in meeting_candidates:
            tid = doc["tdoc_id"]
            abs_p = doc["abs_path"]
            expected_mode = resolve_revision_mode_for_type(doc.get("type", ""), default_mode="markup")
            local_bin = self._find_local_binary_fast(doc.get("folder_name", ""), tid)

            if manifest.is_tdoc_exported(tid) and is_export_complete(abs_p, doc_info=doc):
                if not is_conversion_stale(
                    local_bin, abs_p, expected_version=CONVERTER_VERSION, expected_mode=expected_mode
                ):
                    already_complete_delta += 1
                    self._backfill_canonical_cache(doc, abs_p)
                    continue

            if self._restore_from_canonical_cache(doc, local_bin=local_bin, expected_mode=expected_mode):
                reused_delta += 1
                meeting_newly_converted.append(doc)
                continue

            if self._migrate_legacy_export_file(
                doc, target_dir, manifest, local_bin=local_bin, expected_mode=expected_mode
            ):
                reused_delta += 1
                meeting_newly_converted.append(doc)
                continue

            meeting_docs_to_convert.append(doc)

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

        return meeting_newly_converted, already_complete_delta, reused_delta

    def _synthesize_root_outputs(
        self,
        *,
        target_dir: Path,
        manifest: ExportManifest,
        session_id: int,
        session_timestamp: str,
        is_incremental: bool,
        all_decisions: List[TopicDecision],
        new_decisions_this_session: List[TopicDecision],
        all_candidates: List[Dict[str, Any]],
        newly_converted_docs: List[Dict[str, Any]],
        meeting_records: List[Dict[str, Any]],
        ls_tracker: CrossWgLiaisonTracker,
        reused_from_cache_count: int,
        already_complete_count: int,
        t_start: Optional[float] = None,
    ):
        """Generate root timeline/LS/synthesis/delta/index Markdown files and emit completion."""
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

        analyzed_cnt = sum(1 for r in meeting_records if r.get("analysis_status") == "analyzed")
        stale_cnt = sum(1 for r in meeting_records if r.get("analysis_status") == "stale")
        pending_cnt = sum(1 for r in meeting_records if r.get("analysis_status") == "pending")
        no_act_cnt = sum(1 for r in meeting_records if r.get("analysis_status") == "no_activity")
        newly_converted_count = max(0, len(newly_exported_tids) - reused_from_cache_count)
        elapsed_ms = int((time.perf_counter() - t_start) * 1000) if t_start is not None else 0

        self.last_export_report = {
            "ok": True,
            "mode": "export",
            "is_incremental": is_incremental,
            "target_dir": str(target_dir),
            "topic_name": self.topic.name,
            "session_id": session_id,
            "round_count": len(self.cohorts),
            "total_meetings": len(meeting_records),
            "analyzed": analyzed_cnt,
            "stale": stale_cnt,
            "pending": pending_cnt,
            "no_activity": no_act_cnt,
            "newly_promoted": 0,
            "newly_promoted_meetings": [],
            "meeting_records": meeting_records,
            "synthesis_status": synthesis_status,
            "synthesis_state": manifest.synthesis_state,
            "total_decisions": len(all_decisions),
            "total_active_docs": len(all_candidates),
            "total_stubs": total_stubs,
            "exported_count": len(newly_exported_tids),
            "newly_converted_count": newly_converted_count,
            "reused_from_cache_count": reused_from_cache_count,
            "already_complete_count": already_complete_count,
            "delta_rel_path": f"deltas/delta_session_{session_id}.md",
            "elapsed_ms": elapsed_ms,
        }

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
