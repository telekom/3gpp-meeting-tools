"""Main Topic Dossier multi-cohort export pipeline mixin."""

from __future__ import annotations

import copy
import datetime
import os
import time
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple

from core.utils.paths import to_extended_path
from modules.meetings.core.cohort_resolver import MeetingCohort, format_round_label
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
    classify_meeting_lifecycle,
    resolve_meeting_folder_id,
    write_company_positions_input,
    write_consolidated_briefing_markdown,
    write_e2e_synthesis_prompt,
    write_master_index_markdown,
    write_meeting_brief,
)
from modules.meetings.core.tdocs.chair_notes_agreements_parser import clean_chair_notes_text
from modules.meetings.core.tdocs.tdoc_markdown_converter import (
    CONVERTER_VERSION,
    is_conversion_stale,
    resolve_revision_mode_for_type,
    sanitize_filename,
)
from modules.meetings.core.topic_decision import (
    TopicDecision,
    classify_tdoc_decision_type,
    normalize_chair_notes_decision_type,
)
from ..path_utils import (
    _decision_checkpoint_key,
    _normalize_tdoc_id,
    _resolve_doc_server_url,
    safe_relpath,
    sanitize_folder_name,
)
from ..stage2_discovery import _evaluate_stage2_synthesis_state


class TopicDossierExportPipelineMixin:
    """Orchestrates incremental cohort discovery, document conversion, and dossier synthesis."""

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
