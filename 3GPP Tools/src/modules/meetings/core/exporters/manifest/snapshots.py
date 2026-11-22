# --- File: src/modules/meetings/core/exporters/manifest/snapshots.py ---
"""Deterministic meeting snapshot computation and incremental delta diffing."""

import hashlib
import json
import re
import time
from typing import Any, Dict, List, Optional, Set


def extract_base_tdoc_id(tid: str) -> str:
    """Strips trailing revision suffix ('R2-2600999r01' -> 'R2-2600999') in uppercase."""
    s = str(tid or "").strip().upper()
    m = re.match(r"^([A-Z0-9]+-\d+)-?(?:R|REV)\d{1,2}[A-Z]?$", s, re.IGNORECASE)
    return m.group(1) if m else s


def compute_meeting_snapshot(
    lifecycle: str,
    decision_records: List[Dict[str, Any]],
    active_tdocs: Dict[str, Dict[str, Any]],
    packet_mtime: Optional[float] = None,
) -> Dict[str, Any]:
    """
    Computes a deterministic content snapshot & SHA-1 fingerprint for a single meeting.
    Used to detect stale pre-meeting `_company_positions_analysis.md` files and generate
    token-minimal incremental LLM deltas without creating extra files on disk.
    """
    norm_decisions: Dict[str, Dict[str, Any]] = {}
    for dec in decision_records or []:
        d_key = str(dec.get("key") or "").strip()
        if not d_key:
            continue
        norm_decisions[d_key] = {
            "key": d_key,
            "type": str(dec.get("type", "")).strip(),
            "tdoc_id": str(dec.get("tdoc_id", "") or "").strip().upper(),
            "ai": str(dec.get("ai", "") or "").strip(),
            "rel_path": str(dec.get("rel_path", "") or "").replace("\\", "/"),
            "summary": str(dec.get("summary", "") or "").strip()[:300],
        }

    norm_tdocs: Dict[str, Dict[str, Any]] = {}
    for tid_raw, info in (active_tdocs or {}).items():
        tid = str(tid_raw or "").strip().upper()
        if not tid:
            continue
        norm_tdocs[tid] = {
            "tdoc_id": tid,
            "base_tdoc": extract_base_tdoc_id(tid),
            "status": str(info.get("status", "") or "").strip(),
            "is_stub": bool(info.get("is_stub", False)),
            "supersedes": [str(x).strip().upper() for x in (info.get("supersedes") or []) if str(x).strip()],
            "title": str(info.get("title", "") or "").strip(),
            "source": str(info.get("source", "") or "").strip(),
            "type": str(info.get("type", "") or "").strip(),
            "ai": str(info.get("ai", "") or "").strip(),
            "category": str(info.get("category", "") or "").strip(),
            "rel_path": str(info.get("rel_path", "") or "").replace("\\", "/"),
        }

    fp_payload = {
        "lifecycle": str(lifecycle or "upcoming"),
        "decisions": sorted(norm_decisions.keys()),
        "tdocs": [
            (
                tid,
                norm_tdocs[tid]["status"],
                norm_tdocs[tid]["is_stub"],
                norm_tdocs[tid]["rel_path"],
            )
            for tid in sorted(norm_tdocs.keys())
        ],
    }
    fp_raw = json.dumps(fp_payload, sort_keys=True, ensure_ascii=False).encode("utf-8")
    fingerprint = hashlib.sha1(fp_raw).hexdigest()

    return {
        "fingerprint": fingerprint,
        "lifecycle": str(lifecycle or "upcoming"),
        "decisions": norm_decisions,
        "active_tdocs": norm_tdocs,
        "packet_mtime": float(packet_mtime if packet_mtime is not None else time.time()),
    }


def diff_meeting_snapshots(
    analyzed_snap: Optional[Dict[str, Any]],
    current_snap: Optional[Dict[str, Any]],
) -> Dict[str, Any]:
    """
    Computes the delta between `analyzed_snap` (state when `_company_positions_analysis.md`
    was generated) and `current_snap` (latest exported meeting state).
    """
    empty: Dict[str, Any] = {
        "has_changes": False,
        "previous_lifecycle": "",
        "current_lifecycle": str((current_snap or {}).get("lifecycle", "")),
        "added_decisions": [],
        "added_tids": [],
        "superseded_revisions": [],
        "unstubbed_tids": [],
        "status_changes": [],
        "highlight_tids": [],
    }
    if not analyzed_snap or not current_snap:
        return empty
    if analyzed_snap.get("fingerprint") == current_snap.get("fingerprint"):
        return empty

    prev_lc = str(analyzed_snap.get("lifecycle", ""))
    curr_lc = str(current_snap.get("lifecycle", ""))

    prev_decs = analyzed_snap.get("decisions", {}) or {}
    curr_decs = current_snap.get("decisions", {}) or {}
    added_decisions = [curr_decs[k] for k in sorted(curr_decs.keys()) if k not in prev_decs]

    prev_tdocs = analyzed_snap.get("active_tdocs", {}) or {}
    curr_tdocs = current_snap.get("active_tdocs", {}) or {}

    prev_by_base: Dict[str, List[str]] = {}
    for p_tid in sorted(prev_tdocs.keys()):
        b_id = prev_tdocs[p_tid].get("base_tdoc") or extract_base_tdoc_id(p_tid)
        prev_by_base.setdefault(b_id, []).append(p_tid)

    added_tids: List[str] = []
    superseded_revisions: List[Dict[str, str]] = []
    unstubbed_tids: List[str] = []
    status_changes: List[Dict[str, str]] = []
    highlight_set: Set[str] = set()

    for tid in sorted(curr_tdocs.keys()):
        c_info = curr_tdocs[tid]
        if tid in prev_tdocs:
            p_info = prev_tdocs[tid]
            if p_info.get("is_stub") and not c_info.get("is_stub"):
                unstubbed_tids.append(tid)
                highlight_set.add(tid)
            elif str(p_info.get("status", "")) != str(c_info.get("status", "")):
                status_changes.append(
                    {
                        "tdoc_id": tid,
                        "old_status": str(p_info.get("status", "") or "—"),
                        "new_status": str(c_info.get("status", "") or "—"),
                    }
                )
        else:
            preds = [p for p in (c_info.get("supersedes") or []) if p in prev_tdocs]
            if not preds:
                b_id = c_info.get("base_tdoc") or extract_base_tdoc_id(tid)
                preds = [p for p in prev_by_base.get(b_id, []) if p != tid]
            if preds:
                old_tid = preds[-1]
                old_status = str((prev_tdocs.get(old_tid) or {}).get("status", "") or "—")
                superseded_revisions.append(
                    {
                        "old_tid": old_tid,
                        "new_tid": tid,
                        "old_status": old_status,
                        "new_status": str(c_info.get("status", "") or "—"),
                    }
                )
                highlight_set.add(tid)
            else:
                added_tids.append(tid)
                highlight_set.add(tid)

    has_changes = bool(
        prev_lc != curr_lc
        or added_decisions
        or added_tids
        or superseded_revisions
        or unstubbed_tids
        or status_changes
    )
    return {
        "has_changes": has_changes,
        "previous_lifecycle": prev_lc,
        "current_lifecycle": curr_lc,
        "added_decisions": added_decisions,
        "added_tids": added_tids,
        "superseded_revisions": superseded_revisions,
        "unstubbed_tids": unstubbed_tids,
        "status_changes": status_changes,
        "highlight_tids": sorted(highlight_set),
    }
