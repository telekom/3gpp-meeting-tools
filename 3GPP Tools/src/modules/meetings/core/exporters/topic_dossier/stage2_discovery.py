"""Stage 2 synthesis package discovery and incremental freshness evaluation helpers."""

from __future__ import annotations

import datetime
import json
import os
import re
from pathlib import Path
from typing import Any, Dict, List, Optional, Set

from core.utils.paths import to_extended_path
from modules.meetings.core.exporters.export_manifest import ExportManifest
from modules.meetings.core.exporters.topic_digest_builder import (
    STAGE2_REPORT_FILENAME,
    STAGE2_SYNTHESIS_FILENAME,
)


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

        fp_unchanged = bool(
            same_baseline_tracked
            and active_fps.get(m_key)
            and prev_cov.get(m_key)
            and active_fps.get(m_key) == prev_cov.get(m_key)
        )

        if m_state.get("analysis_status") == "stale":
            changed_meetings.add(m_key)

        ext_mdir = Path(to_extended_path(target_dir / m_key))
        a_path = ext_mdir / "_company_positions_analysis.md"
        if a_path.is_file() and float(os.path.getmtime(a_path)) > syn_mtime + 0.5:
            changed_meetings.add(m_key)

        if not fp_unchanged:
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
