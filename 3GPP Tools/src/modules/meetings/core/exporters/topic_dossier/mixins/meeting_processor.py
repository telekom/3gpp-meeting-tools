"""Per-meeting TDoc discovery, cache/conversion resolution, and checkpointing mixin."""

from __future__ import annotations

import copy
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
    save_manifest,
)
from modules.meetings.core.exporters.topic_digest_builder import (
    classify_meeting_lifecycle,
    write_company_positions_input,
    write_meeting_brief,
)
from modules.meetings.core.tdocs.chair_notes_agreements_parser import clean_chair_notes_text
from modules.meetings.core.tdocs.tdoc_markdown_converter import sanitize_filename
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
)


class TopicDossierMeetingProcessorMixin:
    """Processes an individual meeting cohort entry inside `_run_export`."""

    def _collect_meeting_decisions_and_candidates(
        self,
        *,
        cohort: MeetingCohort,
        mtg: Dict[str, Any],
        wg_name: str,
        mtg_num: str,
        mtg_date: str,
        folder_name: str,
        meeting_folder_id: str,
        meeting_rel_root: str,
        target_dir: Path,
        manifest: ExportManifest,
        tdocs_rows: List[Dict[str, Any]],
        chair_agreements: List[Dict[str, Any]],
        excluded_tid_set: Set[str],
        ls_tracker: CrossWgLiaisonTracker,
    ) -> Tuple[List[TopicDecision], List[Dict[str, Any]], int]:
        """Extract Chair's Notes decisions, match TDocs, ingest LS items, and prune superseded docs."""
        superseded_set, predecessors_map, _leaf_map = self._build_revision_lineage_map(tdocs_rows)
        decision_referenced_tids: Set[str] = set()
        meeting_decisions: List[TopicDecision] = []

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

        deduped_skipped = 0
        meeting_candidates: List[Dict[str, Any]] = []
        withdrawn_tids: Set[str] = set()

        for row in tdocs_rows:
            tid = _normalize_tdoc_id(row.get("TDoc", "") or row.get("tdoc_id", "") or row.get("TD#", ""))
            if not tid or tid in excluded_tid_set:
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
            abstract = str(
                row.get("Abstract", "") or row.get("abstract", "") or row.get("Secretary Remarks", "")
            ).strip()
            ai_desc = str(row.get("Agenda item description", "") or row.get("AI Description", "")).strip()

            is_decision_ref = tid in decision_referenced_tids
            if not is_decision_ref and not self.topic.matches_tdoc(
                tid, wg_name, ai_val, title, wi_val, abstract, ai_description=ai_desc
            ):
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
            if not self.include_discussions and cat_folder == "Discussion" and not (
                is_agreed or is_ls or is_referenced
            ):
                continue

            dec_type = classify_tdoc_decision_type(t_type, title, status)
            if dec_type:
                meeting_decisions.append(
                    TopicDecision(
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
                )

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

        if self.deduplicate_revisions:
            self._prune_superseded_topic_docs(
                target_dir=target_dir,
                manifest=manifest,
                meeting_rel_root=meeting_rel_root,
                active_tids={d["tdoc_id"] for d in meeting_candidates},
                superseded_or_withdrawn=superseded_set | withdrawn_tids,
            )

        return meeting_decisions, meeting_candidates, deduped_skipped

    def _finalize_meeting_checkpoint(
        self,
        *,
        cohort: MeetingCohort,
        mtg: Dict[str, Any],
        wg_name: str,
        mtg_num: str,
        mtg_date: str,
        folder_name: str,
        meeting_folder_id: str,
        meeting_rel_root: str,
        meeting_dir: Path,
        target_dir: Path,
        manifest: ExportManifest,
        prev_mtg_state: Dict[str, Any],
        meeting_decisions: List[TopicDecision],
        meeting_candidates: List[Dict[str, Any]],
        deduped_skipped: int,
        session_id: int,
        session_timestamp: str,
        all_decisions: List[TopicDecision],
        new_decisions_this_session: List[TopicDecision],
    ) -> Dict[str, Any]:
        """Write agreement notes, digest packets, brief, and persist atomic meeting checkpoint."""
        seq_counters: Dict[str, int] = {}
        dec_snapshot_records: List[Dict[str, Any]] = []
        for dec in meeting_decisions:
            if (
                dec.tdoc_id
                and not dec.dossier_rel_path
                and dec.decision_type in ("Approved pCR", "Approved CR", "Endorsed WF", "Agreed LS")
            ):
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
        analyzed_snap = (
            copy.deepcopy(prev_mtg_state.get("analyzed_snapshot"))
            if prev_mtg_state.get("analyzed_snapshot")
            else None
        )
        probe_snap = compute_meeting_snapshot(
            lifecycle, dec_snapshot_records, active_tdocs_snap, packet_mtime=max_digest_mtime
        )
        if prev_snap and prev_snap.get("fingerprint") == probe_snap["fingerprint"]:
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

        return {
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
        }
