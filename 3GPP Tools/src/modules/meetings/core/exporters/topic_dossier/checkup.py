"""Standalone incremental checkup and meeting brief delta rebuild helpers."""

from __future__ import annotations

import copy
import os
import time
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Tuple

from core.utils.paths import to_extended_path
from modules.meetings.core.cohort_resolver import format_round_label
from modules.meetings.core.exporters.export_manifest import (
    ExportManifest,
    compute_meeting_snapshot,
    diff_meeting_snapshots,
    is_export_complete,
    load_manifest,
    save_manifest,
)
from modules.meetings.core.exporters.topic_digest_builder import (
    build_delta_section_markdown,
    sync_meeting_brief_analysis_state,
    write_e2e_synthesis_prompt,
    write_master_index_markdown,
)
from modules.meetings.core.topic_manager import MetaTopicDefinition
from .path_utils import _normalize_tdoc_id, safe_relpath
from .stage2_discovery import _evaluate_stage2_synthesis_state


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
    progress_cb: Optional[Callable[[int, int, str], None]] = None,
    cancel_check: Optional[Callable[[], bool]] = None,
) -> Dict[str, Any]:
    """
    Fast offline checkup & non-destructive refresh for an existing Topic Dossier folder.

    Scans `.export_manifest.json` + on-disk `_company_positions_analysis.md`,
    `_company_positions_input*.md`, and `company_positions_e2e_synthesis.md` files
    to promote freshly completed analyses, preserve stale delta blocks, and
    regenerate `index.md`, `company_positions_e2e_prompt.md`, and `_meeting_brief.md`
    without overwriting any LLM analysis files or re-running document conversions.
    """
    t_start = time.perf_counter()

    def _emit(pct: int, msg: str) -> None:
        if progress_cb is not None:
            try:
                progress_cb(pct, 100, msg)
            except Exception:
                pass

    target_dir = Path(target_dir)
    _emit(0, "Loading export manifest...")
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
    newly_promoted_meetings: List[str] = []
    total_keys = max(1, len(meeting_keys))
    for idx, m_key in enumerate(meeting_keys, start=1):
        if cancel_check is not None and cancel_check():
            return {
                "ok": False,
                "cancelled": True,
                "target_dir": str(target_dir),
                "error": "Checkup cancelled.",
            }
        m_state = manifest.exported_meetings[m_key]
        if not isinstance(m_state, dict):
            continue
        _emit(5 + int(((idx - 1) / total_keys) * 73), f"Scanning {m_key} ({idx}/{len(meeting_keys)})...")
        updated_state, rec = _rebuild_meeting_state_for_checkup(
            target_dir=target_dir,
            manifest=manifest,
            m_key=m_key,
            m_state=m_state,
            tdocs_by_meeting=tdocs_by_meeting,
            decisions_by_fid=decisions_by_fid,
        )
        manifest.exported_meetings[m_key] = updated_state
        if rec.get("newly_promoted"):
            newly_promoted_count += 1
            newly_promoted_meetings.append(m_key)
        meeting_records.append(rec)

    if cancel_check is not None and cancel_check():
        return {
            "ok": False,
            "cancelled": True,
            "target_dir": str(target_dir),
            "error": "Checkup cancelled.",
        }

    meeting_records.sort(
        key=lambda r: (
            str(r.get("start_date") or ""),
            str(r.get("wg") or ""),
            str(r.get("meeting_folder_id") or ""),
        )
    )

    _emit(82, "Evaluating Stage-2 synthesis packages...")
    synthesis_status = _evaluate_stage2_synthesis_state(target_dir, manifest)

    _emit(90, "Regenerating company_positions_e2e_prompt.md & index.md...")
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

    _emit(96, "Saving .export_manifest.json...")
    save_manifest(target_dir, manifest)
    elapsed_ms = int((time.perf_counter() - t_start) * 1000)
    _emit(100, "Checkup complete")

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
        "newly_promoted_meetings": newly_promoted_meetings,
        "meeting_records": meeting_records,
        "synthesis_status": synthesis_status,
        "synthesis_state": manifest.synthesis_state,
        "total_decisions": total_decisions,
        "total_active_docs": total_active_docs,
        "total_stubs": total_stubs,
        "session_id": session_id,
        "elapsed_ms": elapsed_ms,
    }
