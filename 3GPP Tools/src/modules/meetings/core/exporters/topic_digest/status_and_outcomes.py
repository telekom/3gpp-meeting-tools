# --- File: src/modules/meetings/core/exporters/topic_digest/status_and_outcomes.py ---
"""Lifecycle badges, Stage-2 state resolution, TDoc outcomes, and digest card formatting."""

from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from core.utils.company_sanitizer import CompanySanitizer
from modules.meetings.core.topic_decision import TopicDecision

from .constants import STAGE2_SYNTHESIS_FILENAME, _TREATED_STATUS_TOKENS
from .extraction import _rel_to_meeting_dir, extract_tdoc_position_digest


def classify_meeting_lifecycle(
    mtg: Dict[str, Any],
    decisions: List[TopicDecision],
    candidates: List[Dict[str, Any]],
) -> str:
    """
    Classifies a meeting deterministically as:
    - 'concluded': extracted decisions exist OR any candidate has a treated status
    - 'no_activity': 0 decisions and 0 active candidate inputs
    - 'upcoming': 0 decisions, >= 1 candidates with pending/pre-meeting statuses
    """
    if decisions:
        return "concluded"
    if not candidates:
        return "no_activity"
    for doc in candidates:
        status_l = str(doc.get("status", "") or "").lower()
        if any(tok in status_l for tok in _TREATED_STATUS_TOKENS):
            return "concluded"
    return "upcoming"


def format_lifecycle_badge(lifecycle: str) -> str:
    """Returns a concise human-readable lifecycle badge for Markdown tables and briefs."""
    if lifecycle == "concluded":
        return "✅ Concluded"
    if lifecycle == "no_activity":
        return "⚪ No Activity (0 Inputs)"
    return "🧭 Upcoming / Pre-Meeting"


def format_analysis_badge(analysis_status: str) -> str:
    """Returns a concise human-readable Stage-1 or Stage-2 LLM analysis status badge."""
    st = str(analysis_status or "").strip().lower()
    if st == "analyzed":
        return "✅ Analyzed"
    if st == "stale":
        return "🔄 Stale (Delta Ready)"
    if st == "in_progress":
        return "🏗️ In Progress"
    if st == "no_activity":
        return "⚪ Skip (0 Inputs)"
    return "⬜ Pending"


def _resolve_effective_synthesis_state(
    target_dir: Path,
    synthesis_status: Optional[str],
    synthesis_state: Optional[Dict[str, Any]],
) -> Dict[str, Any]:
    """Normalizes Stage-2 synthesis state dict with safe fallback to root synthesis file."""
    if isinstance(synthesis_state, dict) and synthesis_state:
        eff = dict(synthesis_state)
        if synthesis_status and not eff.get("status"):
            eff["status"] = synthesis_status
        return eff

    s2_path = target_dir / STAGE2_SYNTHESIS_FILENAME
    s2_exists = s2_path.is_file()
    eff_status = synthesis_status or ("analyzed" if s2_exists else "pending")
    return {
        "file": STAGE2_SYNTHESIS_FILENAME,
        "folder_name": "",
        "package_id": "legacy-root" if s2_exists else "",
        "baseline_package_id": "none supplied",
        "as_of": "",
        "revision": "legacy" if s2_exists else "",
        "coverage_rel_path": "",
        "handoff_zip_rel_path": "",
        "exists": s2_exists,
        "status": eff_status,
        "changed_meetings": [],
        "latest_package": None,
        "packages": [],
    }


def format_stage2_summary_badge(
    synthesis_status: Optional[str],
    synthesis_state: Optional[Dict[str, Any]] = None,
    as_markdown_link: bool = True,
) -> str:
    """Formats a dated, revision-aware Stage-2 synthesis status string for Markdown."""
    state = dict(synthesis_state or {})
    st = str(synthesis_status or state.get("status") or "pending").strip().lower()
    exists = bool(state.get("exists"))
    folder_name = str(state.get("folder_name") or "").strip()
    report_rel = str(state.get("file") or STAGE2_SYNTHESIS_FILENAME).replace("\\", "/")
    coverage_rel = str(state.get("coverage_rel_path") or "").replace("\\", "/")
    as_of = str(state.get("as_of") or "").strip()
    rev = str(state.get("revision") or "").strip()
    changed = list(state.get("changed_meetings") or [])

    if not exists:
        if st == "in_progress" and folder_name:
            return f"🏗️ In Progress (`{folder_name}/`)"
        return f"⬜ Pending (`{STAGE2_SYNTHESIS_FILENAME}` or `stage2_YYYY-MM-DD_r1/` not yet created)"

    dated_parts = [p for p in (as_of, f"({rev})" if rev else "") if p]
    dated_str = " ".join(dated_parts)
    if st == "stale":
        delta_note = f"{len(changed)} meeting(s) updated" if changed else "inputs updated"
        status_label = f"🔄 Stale ({delta_note} since {dated_str or 'baseline'})"
    else:
        status_label = f"✅ Current ({dated_str})" if dated_str else "✅ Current"

    if not as_markdown_link:
        return f"{status_label} [{folder_name or report_rel}]"

    rep_name = Path(report_rel).name
    folder_suffix = f" in `{folder_name}/`" if folder_name else ""
    cov_suffix = f" · [COVERAGE.md]({coverage_rel})" if coverage_rel else ""
    return f"{status_label} — [{rep_name}]({report_rel}){cov_suffix}{folder_suffix}"


def build_tdoc_outcome_index(
    decisions: List[TopicDecision],
    candidates: List[Dict[str, Any]],
    lifecycle: str,
    wg: str,
    meeting_folder_id: str,
) -> Dict[str, Any]:
    """
    Reconciles incoming TDocs ('brought') against Chairman's Notes agreements and approved
    pCRs/CRs ('decided'), returning disposition summary counts and per-TDoc outcome records.
    """
    decisions_by_tdoc: Dict[str, List[TopicDecision]] = {}
    for dec in decisions:
        tid_key = str(dec.tdoc_id or "").strip().upper()
        if tid_key:
            decisions_by_tdoc.setdefault(tid_key, []).append(dec)

    counts = {
        "brought": len(candidates),
        "cited_in_agreement": 0,
        "approved_doc": 0,
        "noted_without_agreement": 0,
        "revised_withdrawn": 0,
        "pending": 0,
    }
    per_tdoc: Dict[str, Dict[str, Any]] = {}

    for doc in candidates:
        tid = str(doc.get("tdoc_id", "") or "").strip()
        tid_key = tid.upper()
        status_raw = str(doc.get("status", "") or "").strip()
        status_l = status_raw.lower()
        matched_decs = decisions_by_tdoc.get(tid_key, [])

        if matched_decs:
            text_agrs = [m for m in matched_decs if m.is_text_agreement]
            primary = text_agrs[0] if text_agrs else matched_decs[0]
            rel_link = _rel_to_meeting_dir(primary.dossier_rel_path, wg, meeting_folder_id)
            if text_agrs:
                counts["cited_in_agreement"] += 1
                link_md = f"[Cited in Agreement]({rel_link})" if rel_link else "Cited in Agreement"
                table_cell = f"✅ {link_md}"
                if "noted" in status_l:
                    table_cell += " *(Noted)*"
            else:
                counts["approved_doc"] += 1
                lbl = primary.decision_type or "Approved"
                link_md = f"[{lbl}]({rel_link})" if rel_link else lbl
                table_cell = f"🟢 {link_md}"
        elif any(w in status_l for w in ("agreed", "approved", "endorsed")):
            counts["approved_doc"] += 1
            table_cell = f"🟢 {status_raw}"
        elif any(w in status_l for w in ("revised", "merged", "withdrawn", "postponed")):
            counts["revised_withdrawn"] += 1
            table_cell = f"↩️ {status_raw}"
        elif "noted" in status_l or "handled" in status_l:
            counts["noted_without_agreement"] += 1
            table_cell = f"⚪ {status_raw} *(No agreement linked)*"
        else:
            counts["pending"] += 1
            if lifecycle == "upcoming":
                table_cell = f"🧭 Pending ({status_raw or 'Available'})"
            else:
                table_cell = f"⚪ {status_raw or 'Pending / Unknown'}"

        per_tdoc[tid_key] = {
            "table_cell": table_cell,
            "matched_decisions": matched_decs,
            "status_raw": status_raw,
        }

    return {"counts": counts, "per_tdoc": per_tdoc}


def _format_analysis_footprint_line(analysis_status: str, analysis_exists: bool) -> str:
    """Formats the Stage-1 LLM Analysis status line for `_meeting_brief.md`."""
    st = str(analysis_status or "").strip().lower()
    if st == "no_activity":
        return "- **Stage-1 LLM Analysis:** ⚪ Skip (0 active topic inputs)"
    badge = format_analysis_badge(st)
    if analysis_exists:
        return f"- **Stage-1 LLM Analysis:** [{badge}](_company_positions_analysis.md)"
    return f"- **Stage-1 LLM Analysis:** {badge} (`_company_positions_analysis.md` not yet generated)"


def _format_tdoc_outcome_lines(
    outcome_entry: Dict[str, Any],
    lifecycle: str,
    wg: str,
    meeting_folder_id: str,
) -> List[str]:
    """Formats concise 'Meeting Outcome (What Was Decided)' lines + verbatim agreement excerpt for a TDoc."""
    matched_decs: List[TopicDecision] = outcome_entry.get("matched_decisions", [])
    status_raw = str(outcome_entry.get("status_raw", "") or "").strip()
    lines: List[str] = []

    if matched_decs:
        text_agrs = [m for m in matched_decs if m.is_text_agreement]
        primary = text_agrs[0] if text_agrs else matched_decs[0]
        rel_link = _rel_to_meeting_dir(primary.dossier_rel_path, wg, meeting_folder_id)
        link_str = f" — [{Path(rel_link).name}]({rel_link})" if rel_link else ""
        if text_agrs:
            lines.append(
                f"- **Meeting Outcome (What Was Decided):** ✅ **Cited in Binding Chairman's Notes Agreement** "
                f"(TDoc status: `{status_raw or 'Noted'}`){link_str}"
            )
        else:
            lines.append(
                f"- **Meeting Outcome (What Was Decided):** 🟢 **{primary.decision_type}** "
                f"(Status: `{status_raw or primary.status}`){link_str}"
            )
        for dec in matched_decs[:2]:
            raw_excerpt = str(dec.title_or_summary or "").strip()
            if raw_excerpt:
                excerpt = raw_excerpt if len(raw_excerpt) <= 600 else (raw_excerpt[:597].rstrip() + "...")
                for qline in excerpt.splitlines():
                    lines.append(f"  > {qline.rstrip()}")
        return lines

    if lifecycle == "upcoming":
        lines.append(
            f"- **Meeting Outcome (What Was Decided):** 🧭 Pending review at upcoming meeting "
            f"(Current TDoc status: `{status_raw or 'Available'}`)."
        )
    else:
        lines.append(
            f"- **Meeting Outcome (What Was Decided):** `{status_raw or 'Noted / Unlinked'}` "
            f"*(No binding Chairman's Notes agreement or approved CR linked to this TDoc.)*"
        )
    return lines


def _format_tdoc_digest_card(
    d: Dict[str, Any],
    outcome_idx: Dict[str, Any],
    eff_lifecycle: str,
    wg: str,
    meeting_folder_id: str,
    target_dir: Optional[Path] = None,
) -> List[str]:
    """Formats a single TDoc structured digest card shared by full LLM packets and delta blocks."""
    tid = str(d.get("tdoc_id", "") or "").strip()
    title = str(d.get("title", "")).strip()
    preds = d.get("supersedes", []) or []
    lineage = f"Supersedes {', '.join(preds)}" if preds else "Initial / Standalone"
    src_ctx = CompanySanitizer.format_source_triage_context(d.get("source", ""))
    rel_local = _rel_to_meeting_dir(d.get("rel_path", ""), wg, meeting_folder_id)
    out_entry = (outcome_idx.get("per_tdoc") or {}).get(
        tid.strip().upper(),
        {"matched_decisions": [], "status_raw": d.get("status", "")},
    )

    c_lines = [
        f"### {tid}: {title}",
        f"- **Meta:** Type: `{d.get('type', '')}` | AI: `{d.get('ai', '')}` | Status: `{d.get('status', '')}` | Lineage: {lineage}",
        f"- **Attribution:** {src_ctx}",
    ]
    c_lines.extend(_format_tdoc_outcome_lines(out_entry, eff_lifecycle, wg, meeting_folder_id))
    if rel_local:
        c_lines.append(f"- **Full Document:** [{rel_local}]({rel_local})")

    abs_p = d.get("abs_path")
    if not abs_p and target_dir and d.get("rel_path"):
        abs_p = target_dir / str(d["rel_path"])
    digest = extract_tdoc_position_digest(abs_p)
    if digest["intro"]:
        c_lines.append(f"- **Summary / Context:**\n{digest['intro']}")
    if digest["proposals"]:
        c_lines.append(f"- **Proposals / Conclusions:**\n{digest['proposals']}")
    if digest["changes"]:
        c_lines.append(f"- **Specification Changes:**\n{digest['changes']}")
    if digest["fallback"]:
        c_lines.append(f"- **Extracted Content:**\n{digest['fallback']}")

    row = d.get("row_data", {}) or {}
    abstract = str(row.get("Abstract", "") or row.get("Secretary Remarks", "")).strip()
    if abstract and not digest["intro"] and not digest["proposals"] and not digest["fallback"]:
        c_lines.append(f"- **Abstract:** {abstract}")
    c_lines.append("")
    return c_lines


def _resolve_rec_analysis_status(rec: Dict[str, Any], target_dir: Path) -> Tuple[str, bool, str]:
    """Resolves (analysis_status, file_exists, rel_path) for a meeting record."""
    wg = rec["wg"]
    m_fid = rec["meeting_folder_id"]
    lc = rec.get("lifecycle") or ("concluded" if rec.get("decision_count", 0) else "upcoming")
    analysis_rel = f"{wg}/{m_fid}/_company_positions_analysis.md"
    exists = (target_dir / analysis_rel).is_file()
    raw_st = str(rec.get("analysis_status") or "").strip().lower()
    if lc == "no_activity":
        return "no_activity", exists, analysis_rel
    if raw_st in ("analyzed", "stale", "pending", "no_activity"):
        return raw_st, exists, analysis_rel
    return ("analyzed" if exists else "pending"), exists, analysis_rel
