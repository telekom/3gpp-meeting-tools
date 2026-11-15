# --- File: src/modules/meetings/core/exporters/topic_digest_builder.py ---
"""
Deterministic Per-Meeting Briefing & Token-Efficient LLM Digest Builder.

Produces compact, attribution-safe per-meeting packets (`_meeting_brief.md` and
`_company_positions_input.md`) plus root E2E synthesis instructions so non-frontier
models can extract company positions meeting-by-meeting without exceeding token budgets.
"""

import datetime
import os
import re
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from core.utils.company_sanitizer import CompanySanitizer
from core.utils.paths import to_extended_path
from modules.meetings.core.exporters.export_manifest import ExportManifest, sanitize_folder_name
from modules.meetings.core.topic_decision import TopicDecision
from modules.meetings.core.topic_manager import MetaTopicDefinition

STAGE1_DELTA_START = "<!-- STAGE1_DELTA_START -->"
STAGE1_DELTA_END = "<!-- STAGE1_DELTA_END -->"
STAGE2_SYNTHESIS_FILENAME = "company_positions_e2e_synthesis.md"

DEFAULT_MAX_DIGEST_CHARS = 140_000

_SUFFIX_IN_NUM_RE = re.compile(r'(?:bis|ahe|adhoc|ah\d*|-e\d*)$', re.IGNORECASE)


def resolve_meeting_folder_id(mtg: Dict[str, Any]) -> str:
    """
    Produces a collision-free, human-readable meeting folder identifier:
    - 'TSGR3_133'      -> 'RAN3_133'
    - 'TSGR3_133-bis'  -> 'RAN3_133-bis'
    - 'TSGR2_133bis'   -> 'RAN2_133bis' (or 'RAN2_133-bis')
    - 'TSGS2_175_AHE'  -> 'SA2_175-AHE'
    """
    wg = sanitize_folder_name(str(mtg.get("wg_name", "")).strip().upper() or "WG")
    raw_num = str(mtg.get("meeting_number", "")).strip()
    folder_name = str(mtg.get("folder_name", "")).strip()

    clean_num = sanitize_folder_name(raw_num or "0")
    if raw_num and not _SUFFIX_IN_NUM_RE.search(raw_num) and folder_name:
        pat = re.compile(
            r'(?:^|[_-])' + re.escape(raw_num) + r'([_-]?(?:bis|ahe|adhoc|ah\d*|e\d*)\b)',
            re.IGNORECASE,
        )
        m = pat.search(folder_name)
        if m:
            suffix = m.group(1).lstrip("_-")
            suffix_clean = sanitize_folder_name(suffix)
            if suffix_clean:
                return f"{wg}_{clean_num}-{suffix_clean}"

    return f"{wg}_{clean_num}"


def rewrite_media_links(md_text: str, old_media_rel: str, new_media_rel: str) -> str:
    """Rewrites relative Markdown & HTML image/media references when copying cached .md files."""
    if not md_text or not old_media_rel or not new_media_rel or old_media_rel == new_media_rel:
        return md_text

    pattern_md = re.compile(r'(\]\()' + re.escape(old_media_rel) + r'(/[^\)\s]+)')
    updated = pattern_md.sub(rf'\g<1>{new_media_rel}\g<2>', md_text)

    pattern_html = re.compile(r'(src=["\'])' + re.escape(old_media_rel) + r'(/[^"\']+)')
    updated = pattern_html.sub(rf'\g<1>{new_media_rel}\g<2>', updated)
    return updated


def build_contributor_matrix(
    candidates: List[Dict[str, Any]],
    decisions: Optional[List[TopicDecision]] = None,
) -> List[Dict[str, Any]]:
    """
    Aggregates company participation within a meeting using CompanySanitizer:
    separates Lead Proponents vs. Confirmed Co-Signers (ignoring tentative '?' / '[...]' carry-overs).
    """
    company_stats: Dict[str, Dict[str, Any]] = {}

    def _ensure(comp: str) -> Dict[str, Any]:
        if comp not in company_stats:
            company_stats[comp] = {
                "company": comp,
                "lead_count": 0,
                "cosigned_count": 0,
                "approved_count": 0,
                "discussion_count": 0,
                "ais": set(),
                "tdocs": [],
            }
        return company_stats[comp]

    for doc in candidates:
        tid = doc.get("tdoc_id", "")
        src = str(doc.get("source", "") or "")
        ai = str(doc.get("ai", "") or "").strip()
        status_l = str(doc.get("status", "") or "").lower()
        is_approved = any(w in status_l for w in ("approved", "agreed", "endorsed"))

        primaries = CompanySanitizer.get_primary_contributors(src, include_tentative=False)
        confirmed, _tentative = CompanySanitizer.parse_contributors_detailed(src)

        if not primaries and confirmed:
            primaries = [confirmed[0]]

        lead_set = set(primaries)
        for p in primaries:
            st = _ensure(p)
            st["lead_count"] += 1
            if is_approved:
                st["approved_count"] += 1
            else:
                st["discussion_count"] += 1
            if ai:
                st["ais"].add(ai)
            if tid and tid not in st["tdocs"]:
                st["tdocs"].append(tid)

        for c in confirmed:
            if c not in lead_set:
                st = _ensure(c)
                st["cosigned_count"] += 1
                if ai:
                    st["ais"].add(ai)
                if tid and tid not in st["tdocs"]:
                    st["tdocs"].append(tid)

    rows = list(company_stats.values())
    rows.sort(
        key=lambda r: (
            -(r["lead_count"] + r["cosigned_count"]),
            -r["lead_count"],
            r["company"],
        )
    )
    return rows


def _strip_markdown_frontmatter(text: str) -> str:
    """Removes tdoc_converter HTML comments and top bullet metadata block before the document body."""
    cleaned = re.sub(r'^<!---\s*tdoc_converter:.*?-->\s*', '', text, flags=re.IGNORECASE | re.DOTALL)
    # Find the separator line '---' that tdoc2md places after the metadata header
    lines = cleaned.splitlines()
    for idx, line in enumerate(lines[:45]):
        if line.strip() == "---":
            return "\n".join(lines[idx + 1:]).strip()
    # Fallback: skip leading '# TDoc' and '- **Key:**' bullets
    start_idx = 0
    for idx, line in enumerate(lines[:35]):
        s = line.strip()
        if idx == 0 and s.startswith("# "):
            start_idx = idx + 1
        elif s.startswith("- **") or not s:
            start_idx = idx + 1
        else:
            break
    return "\n".join(lines[start_idx:]).strip()


def extract_tdoc_position_digest(
    md_path: Optional[Path],
    max_intro_chars: int = 1200,
    max_proposal_chars: int = 2000,
    max_changes_chars: int = 1200,
    max_fallback_chars: int = 2200,
) -> Dict[str, str]:
    """
    Deterministically extracts high-signal company position blocks from a converted TDoc Markdown file,
    stripping verbose specification boilerplate while preserving intro, proposals, conclusions, and key CR edits.
    """
    result = {"intro": "", "proposals": "", "changes": "", "fallback": ""}
    if not md_path:
        return result

    ext_path = Path(to_extended_path(md_path))
    if not ext_path.is_file():
        return result

    try:
        raw_text = ext_path.read_text(encoding="utf-8", errors="replace")
    except Exception:
        return result

    if "Document conversion was not available" in raw_text[:2048]:
        return result

    body = _strip_markdown_frontmatter(raw_text)
    if not body:
        return result

    # Strip embedded base64 data URIs or massive markup legend headers
    body = re.sub(r'!\[[^\]]*\]\(data:image/[^)]+\)', '[Image]', body)
    body = re.sub(r'>\s*\*\*Markup Legend:\*\*.*?(?:\n\n|\Z)', '', body, flags=re.IGNORECASE | re.DOTALL).strip()

    # Split body into (heading, text) sections
    sections: List[Tuple[str, str]] = []
    cur_heading = ""
    cur_lines: List[str] = []
    for line in body.splitlines():
        if re.match(r'^#{1,5}\s+\S', line) or re.match(r'^\*\*(?:Proposal|Observation|Conclusion|Way Forward|Summary|Reason for change)[^*]*\*\*\s*:?$', line.strip(), re.IGNORECASE):
            if cur_lines or cur_heading:
                sections.append((cur_heading, "\n".join(cur_lines).strip()))
            cur_heading = line.strip().lstrip("#").strip("* ").strip()
            cur_lines = []
        else:
            cur_lines.append(line)
    if cur_lines or cur_heading:
        sections.append((cur_heading, "\n".join(cur_lines).strip()))

    intro_parts: List[str] = []
    proposal_parts: List[str] = []
    intro_kw = ("introduction", "abstract", "background", "scope", "reason for change", "overview", "general")
    prop_kw = ("proposal", "observation", "conclusion", "way forward", "recommendation", "solution", "summary", "discussion", "agreement")

    for h, content in sections:
        if not content:
            continue
        hl = h.lower()
        if any(k in hl for k in prop_kw):
            snippet = f"**{h}:**\n{content}" if h else content
            proposal_parts.append(snippet)
        elif not hl or any(k in hl for k in intro_kw):
            snippet = f"**{h}:**\n{content}" if h else content
            intro_parts.append(snippet)

    # Fallback intro from first non-empty section if nothing matched intro_kw
    if not intro_parts and sections:
        first_h, first_c = sections[0]
        if first_c:
            intro_parts.append(f"**{first_h}:**\n{first_c}" if first_h else first_c)

    intro_text = "\n\n".join(intro_parts).strip()
    prop_text = "\n\n".join(proposal_parts).strip()

    # Extract [INSERTED: ...] change highlights (for CR/pCRs)
    change_lines: List[str] = []
    for line in body.splitlines():
        if "[INSERTED:" in line or "[DELETED:" in line:
            s = line.strip()
            if s and len(s) > 12:
                change_lines.append(s)
    changes_text = "\n".join(change_lines[:25]).strip()

    if len(intro_text) > max_intro_chars:
        intro_text = intro_text[:max_intro_chars].rstrip() + " ..."
    if len(prop_text) > max_proposal_chars:
        prop_text = prop_text[:max_proposal_chars].rstrip() + " ..."
    if len(changes_text) > max_changes_chars:
        changes_text = changes_text[:max_changes_chars].rstrip() + " ..."

    result["intro"] = intro_text
    result["proposals"] = prop_text
    result["changes"] = changes_text

    if not intro_text and not prop_text and not changes_text:
        clean_body = re.sub(r'\n{3,}', '\n\n', body).strip()
        result["fallback"] = (clean_body[:max_fallback_chars].rstrip() + " ...") if len(clean_body) > max_fallback_chars else clean_body

    return result


_TREATED_STATUS_TOKENS = (
    "agreed",
    "approved",
    "endorsed",
    "noted",
    "revised",
    "merged",
    "handled",
    "postponed",
    "withdrawn",
)


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
    """Returns a concise human-readable Stage-1 LLM analysis status badge."""
    st = str(analysis_status or "").strip().lower()
    if st == "analyzed":
        return "✅ Analyzed"
    if st == "stale":
        return "🔄 Stale (Delta Ready)"
    if st == "no_activity":
        return "⚪ Skip (0 Inputs)"
    return "⬜ Pending"


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


def build_delta_section_markdown(
    delta_info: Optional[Dict[str, Any]],
    wg: str,
    meeting_folder_id: str,
    eff_lifecycle: str,
    candidates: Optional[List[Dict[str, Any]]] = None,
    decisions: Optional[List[TopicDecision]] = None,
    outcome_idx: Optional[Dict[str, Any]] = None,
    target_dir: Optional[Path] = None,
    digest_rel_paths: Optional[List[str]] = None,
    max_delta_digest_chars: int = 50_000,
) -> str:
    """
    Renders an in-place, token-minimal `## 🔄 Delta Since Last Analysis` block wrapped
    with HTML markers so Stage-1 LLMs can incrementally merge only new decisions, status
    shifts, and newly arrived/revised/un-stubbed TDocs into `_company_positions_analysis.md`.
    """
    if not delta_info or not delta_info.get("has_changes"):
        return ""

    prev_lc = format_lifecycle_badge(delta_info.get("previous_lifecycle") or "upcoming")
    curr_lc = format_lifecycle_badge(delta_info.get("current_lifecycle") or eff_lifecycle)
    added_decisions = delta_info.get("added_decisions") or []
    superseded_revs = delta_info.get("superseded_revisions") or []
    unstubbed_tids = delta_info.get("unstubbed_tids") or []
    status_changes = delta_info.get("status_changes") or []
    added_tids = delta_info.get("added_tids") or []
    highlight_tids = list(delta_info.get("highlight_tids") or [])

    lines = [
        STAGE1_DELTA_START,
        "",
        "## 🔄 Delta Since Last Analysis (Incremental Merge Packet)",
        "",
        "> [!IMPORTANT] Token-Efficient Incremental Update Mode",
        "> `_company_positions_analysis.md` already exists for an earlier state of this meeting.",
        "> **Do NOT re-read full `_company_positions_input*.md` files.** Read ONLY `_company_positions_analysis.md` and this delta section, then **merge** the new outcomes, revisions, and contributions directly into `_company_positions_analysis.md`.",
        "",
    ]

    if prev_lc != curr_lc:
        lines.append(f"- **Lifecycle Transition:** {prev_lc} → **{curr_lc}**")
    lines.append(
        f"- **Delta Summary:** +{len(added_decisions)} new decision(s), "
        f"+{len(added_tids)} new TDoc(s), {len(superseded_revs)} revised TDoc(s), "
        f"{len(unstubbed_tids)} un-stubbed upload(s), {len(status_changes)} status update(s)"
    )

    if added_decisions:
        lines.extend(["", "### 🆕 New Binding Decisions & Agreements", ""])
        dec_by_key = {}
        for d in decisions or []:
            k = f"{meeting_folder_id}::{d.decision_type}::{d.tdoc_id or d.ai or ''}"
            dec_by_key[k] = d
        for ad in added_decisions:
            d_type = ad.get("type", "Decision")
            tid_lbl = ad.get("tdoc_id") or f"AI {ad.get('ai') or 'N/A'}"
            summary = re.sub(r'\s+', ' ', str(ad.get("summary", ""))).strip()
            rel_d = _rel_to_meeting_dir(ad.get("rel_path", ""), wg, meeting_folder_id)
            link_suffix = f" ([View]({rel_d}))" if rel_d else ""
            lines.append(f"- **[{d_type}]** ({tid_lbl}){link_suffix}: {summary}")

    if superseded_revs or status_changes or unstubbed_tids:
        lines.extend([
            "",
            "### 🔄 Revision, Upload & Status Transitions",
            "",
            "| Previous TDoc / State | Updated TDoc / State | Change Type |",
            "| :--- | :--- | :--- |",
        ])
        for sr in superseded_revs:
            lines.append(
                f"| `{sr['old_tid']}` ({sr.get('old_status') or '—'}) | **`{sr['new_tid']}`** ({sr.get('new_status') or '—'}) | New Revision |"
            )
        for uid in unstubbed_tids:
            lines.append(f"| `{uid}` *(Placeholder Stub)* | **`{uid}`** *(Full Document Now Available)* | Un-Stubbed Upload |")
        for sc in status_changes:
            lines.append(
                f"| `{sc['tdoc_id']}` ({sc.get('old_status') or '—'}) | `{sc['tdoc_id']}` (**{sc.get('new_status') or '—'}**) | Status Update |"
            )

    cand_map = {str(c.get("tdoc_id", "")).strip().upper(): c for c in (candidates or [])}
    if highlight_tids and cand_map:
        eff_out_idx = outcome_idx or build_tdoc_outcome_index(decisions or [], candidates or [], eff_lifecycle, wg, meeting_folder_id)
        lines.extend(["", "### 📄 Digests for Newly Added / Revised / Un-Stubbed TDocs", ""])
        used_chars = 0
        omitted = 0
        for h_tid in highlight_tids:
            doc = cand_map.get(h_tid.upper())
            if not doc:
                continue
            card_text = "\n".join(
                _format_tdoc_digest_card(doc, eff_out_idx, eff_lifecycle, wg, meeting_folder_id, target_dir=target_dir)
            )
            if used_chars + len(card_text) > max_delta_digest_chars and used_chars > 0:
                omitted += 1
                continue
            lines.append(card_text)
            used_chars += len(card_text)
        if omitted:
            pkt_links = ", ".join(f"[{Path(p).name}]({Path(p).name})" for p in (digest_rel_paths or [])) or "`_company_positions_input.md`"
            lines.append(f"*({omitted} additional new/revised TDoc digest(s) omitted for brevity; inspect {pkt_links} if needed.)*")

    lines.extend(["", STAGE1_DELTA_END])
    return "\n".join(lines)


def sync_meeting_brief_analysis_state(
    brief_path: Path,
    lifecycle: str,
    analysis_status: str,
    analysis_exists: bool,
    delta_markdown: str = "",
) -> bool:
    """
    Non-destructively updates an existing `_meeting_brief.md` on disk to reflect the latest
    lifecycle badge (`no_activity` / `concluded` / `upcoming`), Stage-1 Analysis link/badge,
    and either inserts/updates or clears the `STAGE1_DELTA` block.
    """
    ext_brief = Path(to_extended_path(brief_path))
    if not ext_brief.is_file():
        return False

    try:
        raw = ext_brief.read_text(encoding="utf-8", errors="replace")
    except Exception:
        return False

    new_text = raw
    # 1. Keep State / Lifecycle badge accurate (e.g. reclassifying 0-input meetings)
    lc_badge = format_lifecycle_badge(lifecycle)
    new_text = re.sub(
        r'(\*\*State:\*\*\s*)[^\r\n]+',
        rf'\g<1>{lc_badge}',
        new_text,
        count=1,
    )
    new_text = re.sub(
        r'(-\s*\*\*Lifecycle State:\*\*\s*)[^\r\n]+',
        rf'\g<1>{lc_badge}',
        new_text,
        count=1,
    )

    # 2. Insert or update `- **Stage-1 LLM Analysis:** ...` inside Meeting Topic Footprint
    analysis_line = _format_analysis_footprint_line(analysis_status, analysis_exists)
    if re.search(r'^-\s*\*\*Stage-1 LLM Analysis:\*\*.*$', new_text, flags=re.MULTILINE):
        new_text = re.sub(
            r'^-\s*\*\*Stage-1 LLM Analysis:\*\*.*$',
            analysis_line,
            new_text,
            count=1,
            flags=re.MULTILINE,
        )
    elif re.search(r'^-\s*\*\*LLM Company Position Packet\(s\):.*$', new_text, flags=re.MULTILINE):
        new_text = re.sub(
            r'^(-\s*\*\*LLM Company Position Packet\(s\):.*)$',
            rf'\1\n{analysis_line}',
            new_text,
            count=1,
            flags=re.MULTILINE,
        )
    elif "## ⚖️ Binding Decisions" in new_text:
        new_text = new_text.replace("## ⚖️ Binding Decisions", f"{analysis_line}\n\n## ⚖️ Binding Decisions", 1)

    # 3. Strip any existing STAGE1_DELTA block
    new_text = re.sub(
        r'\n?' + re.escape(STAGE1_DELTA_START) + r'.*?' + re.escape(STAGE1_DELTA_END) + r'\n*',
        '\n\n',
        new_text,
        flags=re.DOTALL,
    )

    # 4. Insert delta_markdown right before Binding Decisions when stale
    if delta_markdown and delta_markdown.strip():
        marker = "## ⚖️ Binding Decisions & Agreements"
        if marker in new_text:
            new_text = new_text.replace(marker, f"{delta_markdown.strip()}\n\n{marker}", 1)
        else:
            new_text = new_text.rstrip() + "\n\n" + delta_markdown.strip() + "\n"

    if new_text != raw:
        try:
            ext_brief.write_text(new_text, encoding="utf-8")
            return True
        except Exception:
            return False
    return True


def write_meeting_brief(
    meeting_dir: Path,
    topic_name: str,
    cohort_label: str,
    mtg: Dict[str, Any],
    meeting_folder_id: str,
    decisions: List[TopicDecision],
    candidates: List[Dict[str, Any]],
    deduped_count: int = 0,
    digest_rel_paths: Optional[List[str]] = None,
    lifecycle: Optional[str] = None,
    analysis_status: str = "pending",
    analysis_exists: bool = False,
    delta_info: Optional[Dict[str, Any]] = None,
) -> Path:
    """Writes a deterministic executive brief (`_meeting_brief.md`) inside `{WG}/{MeetingFolder}/`."""
    meeting_dir_ext = Path(to_extended_path(meeting_dir))
    meeting_dir_ext.mkdir(parents=True, exist_ok=True)
    brief_path = meeting_dir / "_meeting_brief.md"

    wg = str(mtg.get("wg_name", "")).strip().upper()
    mtg_num = str(mtg.get("meeting_number", "")).strip()
    start_date = str(mtg.get("start_date", "") or "").strip()
    end_date = str(mtg.get("end_date", "") or "").strip()
    location = str(mtg.get("location", "") or "").strip()
    folder_name = str(mtg.get("folder_name", "") or "").strip()

    eff_lifecycle = lifecycle or classify_meeting_lifecycle(mtg, decisions, candidates)
    outcome_idx = build_tdoc_outcome_index(decisions, candidates, eff_lifecycle, wg, meeting_folder_id)
    disp = outcome_idx["counts"]

    eff_analysis_status = "no_activity" if eff_lifecycle == "no_activity" else analysis_status
    eff_analysis_exists = analysis_exists or (meeting_dir_ext / "_company_positions_analysis.md").is_file()

    lines = [
        f"# 📋 Meeting Brief: {wg} #{mtg_num} (`{meeting_folder_id}`)",
        f"**Meta-Topic:** {topic_name} | **Synchronized Round:** {cohort_label} | **State:** {format_lifecycle_badge(eff_lifecycle)}",
        f"**Dates:** {start_date or 'N/A'} → {end_date or 'N/A'}" + (f" | **Location:** {location}" if location else "") + (f" | **3GPP Folder:** `{folder_name}`" if folder_name else ""),
        "",
        "## 📊 Meeting Topic Footprint",
        f"- **Lifecycle State:** {format_lifecycle_badge(eff_lifecycle)}",
        f"- **Binding Decisions:** {len(decisions)}",
        f"- **Active Scoped Contributions:** {len(candidates)}" + (f" *(+{deduped_count} superseded/withdrawn revisions deduplicated)*" if deduped_count else ""),
        (
            f"- **Brought vs. Decided Disposition:** {disp['brought']} Brought → "
            f"{disp['cited_in_agreement']} Cited in Agreement, "
            f"{disp['approved_doc']} Approved / Agreed, "
            f"{disp['noted_without_agreement']} Noted (no linked agreement), "
            f"{disp['revised_withdrawn']} Revised/Withdrawn, "
            f"{disp['pending']} Pending"
        ),
    ]

    if digest_rel_paths:
        digest_links = ", ".join(f"[{Path(p).name}]({Path(p).name})" for p in digest_rel_paths)
        lines.append(f"- **LLM Company Position Packet(s):** {digest_links}")

    lines.append(_format_analysis_footprint_line(eff_analysis_status, eff_analysis_exists))

    if eff_analysis_status == "stale" and delta_info and delta_info.get("has_changes"):
        delta_block = build_delta_section_markdown(
            delta_info=delta_info,
            wg=wg,
            meeting_folder_id=meeting_folder_id,
            eff_lifecycle=eff_lifecycle,
            candidates=candidates,
            decisions=decisions,
            outcome_idx=outcome_idx,
            digest_rel_paths=digest_rel_paths,
        )
        if delta_block:
            lines.extend(["", delta_block])

    lines.extend(["", "## ⚖️ Binding Decisions & Agreements", ""])
    if decisions:
        lines.extend([
            "| Type | AI | Contribution / Reference | Summary / Substantive Text | Dossier |",
            "| :--- | :--- | :--- | :--- | :--- |",
        ])
        for dec in decisions:
            summary_clean = re.sub(r'[\r\n\t]+', ' ', dec.title_or_summary).replace("|", "\\|").strip()
            if len(summary_clean) > 220:
                summary_clean = summary_clean[:217] + "..."
            rel_link = _rel_to_meeting_dir(dec.dossier_rel_path, wg, meeting_folder_id)
            doc_link = f"[{dec.tdoc_id or 'Agreement'}]({rel_link})" if rel_link else (dec.tdoc_id or "—")
            lines.append(f"| {dec.badge_icon} **{dec.decision_type}** | `{dec.ai or '—'}` | {dec.tdoc_id or 'Chairman Notes'} | {summary_clean} | {doc_link} |")
    else:
        lines.append("*No binding topic decisions recorded for this meeting.*")

    # Liaison Statements in this meeting
    ls_docs = [d for d in candidates if d.get("category") in ("LS_IN", "LS_OUT")]
    lines.extend(["", "## 📬 Liaison Statements (In / Out)", ""])
    if ls_docs:
        lines.extend([
            "| TDoc | Dir | AI | Title | Source | Status | Dossier |",
            "| :--- | :--- | :--- | :--- | :--- | :--- | :--- |",
        ])
        for d in ls_docs:
            dir_lbl = "📥 IN" if d.get("category") == "LS_IN" else "📤 OUT"
            t_clean = re.sub(r'[\r\n\t]+', ' ', str(d.get("title", ""))).replace("|", "\\|")
            src_cell = CompanySanitizer.format_source_markdown_table_cell(d.get("source", ""))
            rel_local = _rel_to_meeting_dir(d.get("rel_path", ""), wg, meeting_folder_id)
            link_c = f"[{d['tdoc_id']}]({rel_local})" if rel_local else d["tdoc_id"]
            lines.append(f"| **{d['tdoc_id']}** | {dir_lbl} | `{d.get('ai', '')}` | {t_clean} | {src_cell} | {d.get('status', '')} | {link_c} |")
    else:
        lines.append("*No topic-related Liaison Statements in this meeting.*")

    # Normalized Contributor Matrix
    matrix = build_contributor_matrix(candidates, decisions)
    lines.extend([
        "",
        "## 🏢 Normalized Contributor Activity Matrix",
        "",
        "> [!NOTE]",
        "> Counts use confirmed attributions only (`CompanySanitizer`). Unconfirmed carry-over `[Company?]` co-signers in revised CRs/pCRs are excluded.",
        "",
    ])
    if matrix:
        lines.extend([
            "| Company | Lead TDocs | Confirmed Co-Signed | Approved / Agreed | Discussion / Noted | Agenda Items |",
            "| :--- | :---: | :---: | :---: | :---: | :--- |",
        ])
        for row in matrix:
            ais_str = ", ".join(sorted(row["ais"]))[:60]
            lines.append(
                f"| **{row['company']}** | {row['lead_count']} | {row['cosigned_count']} | {row['approved_count']} | {row['discussion_count']} | {ais_str} |"
            )
    else:
        lines.append("*No recognized company contributors.*")

    # Active Contribution Register
    lines.extend(["", "## 📄 Active Contribution Register", ""])
    if candidates:
        lines.extend([
            "| TDoc | Type | AI | Status | Decision / Outcome | Source | Lineage | Dossier |",
            "| :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- |",
        ])
        for d in sorted(candidates, key=lambda x: (str(x.get("ai", "")), str(x.get("tdoc_id", "")))):
            src_cell = CompanySanitizer.format_source_markdown_table_cell(d.get("source", ""))
            preds = d.get("supersedes", [])
            lineage_str = f"rev of {', '.join(preds[-2:])}" if preds else "—"
            rel_local = _rel_to_meeting_dir(d.get("rel_path", ""), wg, meeting_folder_id)
            link_c = f"[{d['tdoc_id']}]({rel_local})" if rel_local else d["tdoc_id"]
            tid_key = str(d.get("tdoc_id", "") or "").strip().upper()
            out_cell = outcome_idx["per_tdoc"].get(tid_key, {}).get("table_cell", "—")
            lines.append(
                f"| **{d['tdoc_id']}** | {d.get('type', '')} | `{d.get('ai', '')}` | {d.get('status', '')} | {out_cell} | {src_cell} | {lineage_str} | {link_c} |"
            )
    else:
        lines.append("*No scoped contributions matched this topic.*")

    lines.append("")
    with open(to_extended_path(brief_path), "w", encoding="utf-8") as f:
        f.write("\n".join(lines))

    return brief_path


def _rel_to_meeting_dir(rel_from_root: str, wg: str, meeting_folder_id: str) -> str:
    """Converts root-relative path '{wg}/{meeting_folder_id}/Cat/File.md' to meeting-relative 'Cat/File.md'."""
    if not rel_from_root:
        return ""
    norm = str(rel_from_root).replace("\\", "/")
    prefix = f"{sanitize_folder_name(wg)}/{meeting_folder_id}/"
    if norm.startswith(prefix):
        return norm[len(prefix):]
    return norm


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


def write_company_positions_input(
    meeting_dir: Path,
    topic_name: str,
    mtg: Dict[str, Any],
    meeting_folder_id: str,
    decisions: List[TopicDecision],
    candidates: List[Dict[str, Any]],
    max_chars_per_part: int = DEFAULT_MAX_DIGEST_CHARS,
    lifecycle: Optional[str] = None,
) -> List[Path]:
    """
    Generates token-compacted, company-grouped `_company_positions_input.md` (auto-split into `_partN.md`
    along company boundaries if exceeding `max_chars_per_part`) for reliable single-meeting analysis
    by non-frontier / lower-cost LLMs.
    """
    meeting_dir_ext = Path(to_extended_path(meeting_dir))
    meeting_dir_ext.mkdir(parents=True, exist_ok=True)

    wg = str(mtg.get("wg_name", "")).strip().upper()
    mtg_num = str(mtg.get("meeting_number", "")).strip()
    start_date = str(mtg.get("start_date", "") or "").strip()

    eff_lifecycle = lifecycle or classify_meeting_lifecycle(mtg, decisions, candidates)
    outcome_idx = build_tdoc_outcome_index(decisions, candidates, eff_lifecycle, wg, meeting_folder_id)

    if eff_lifecycle == "concluded":
        stage1_instructions = [
            "> [!IMPORTANT] Stage-1 Per-Meeting Extraction Instructions (Concluded Meeting — Brought vs. Decided)",
            f"> Save your output as `_company_positions_analysis.md` inside `{meeting_folder_id}/`.",
            "> 1. **Goal:** Compare what each company **brought** (`Summary`/`Proposals`) against what was **decided** (`Meeting Outcome` & `Meeting Decisions Reference`).",
            "> 2. **RAN TDoc Status Rule:** In 3GPP RAN, contributions that shaped a binding Chairman's Notes agreement often still carry `Status: Noted`. Only treat a contribution as adopted/influential if its `Meeting Outcome` cites a binding Agreement or Approved pCR/CR.",
            "> 3. **Attribution Rule:** Do NOT credit companies listed inside `[Note: ... carry-over co-sponsors]` — only credit Confirmed Lead and Confirmed Co-Signers.",
            "> 4. **Output Schema:**",
            ">    - Section A: **Decision & Consensus Summary (Brought vs. Decided)** — key architectural choices locked and whose proposals shaped them.",
            ">    - Section B: **Company Stance Matrix** (`| Company | Core Position | Key Proposals / TDocs | Allies (Confirmed) | Outcome (Adopted / Compromised / Noted) |`).",
            ">    - Section C: **Open Points / Disagreements / FFS** carried into the next meeting.",
        ]
    else:
        stage1_instructions = [
            "> [!IMPORTANT] Stage-1 Per-Meeting Extraction Instructions (Upcoming / Pre-Meeting — Positions Brought)",
            f"> Save your output as `_company_positions_analysis.md` inside `{meeting_folder_id}/`.",
            "> 1. **Goal:** Map **what companies are bringing** into this upcoming meeting (0 binding decisions yet). Do NOT invent decisions.",
            "> 2. **Attribution Rule:** Do NOT credit companies listed inside `[Note: ... carry-over co-sponsors]` — only credit Confirmed Lead and Confirmed Co-Signers.",
            "> 3. **Output Schema:**",
            ">    - Section A: **Agenda & Contested Issues Preview** — main architectural options competing in this meeting.",
            ">    - Section B: **Company Stance Matrix** (`| Company | Proposed Position | Key Proposals / TDocs | Allies (Confirmed) | Status |`).",
            ">    - Section C: **Key Fault Lines & Compromises to Watch** during the meeting.",
        ]

    header_lines = [
        f"# 🧠 LLM Company Position Input Packet: {wg} #{mtg_num} (`{meeting_folder_id}`)",
        f"**Meta-Topic:** {topic_name} | **State:** {format_lifecycle_badge(eff_lifecycle)} | **Date:** {start_date or 'N/A'} | **Active Docs:** {len(candidates)} | **Decisions:** {len(decisions)}",
        "",
        *stage1_instructions,
        "",
        "## ⚖️ Meeting Decisions Reference",
        "",
    ]
    if decisions:
        for dec in decisions:
            dec_txt = re.sub(r'\s+', ' ', dec.title_or_summary).strip()
            if len(dec_txt) > 350:
                dec_txt = dec_txt[:347] + "..."
            header_lines.append(f"- **[{dec.decision_type}]** ({dec.tdoc_id or f'AI {dec.ai}'}): {dec_txt}")
    else:
        header_lines.append("- *No formal agreements/approvals recorded in this meeting.*")
    header_lines.append("")
    base_header = "\n".join(header_lines)

    # Group candidates by Primary Company
    by_company: Dict[str, List[Dict[str, Any]]] = {}
    for doc in candidates:
        primaries = CompanySanitizer.get_primary_contributors(doc.get("source", ""), include_tentative=False)
        lead_company = primaries[0] if primaries else (CompanySanitizer.get_confirmed_source_text(doc.get("source", "")) or "Other / Multi-Company")
        if len(lead_company) > 50:
            lead_company = lead_company[:47] + "..."
        by_company.setdefault(lead_company, []).append(doc)

    # Sort company blocks by number of docs descending, then alphabetically
    ordered_companies = sorted(by_company.keys(), key=lambda c: (-len(by_company[c]), c))

    company_blocks: List[str] = []
    for comp in ordered_companies:
        docs = sorted(by_company[comp], key=lambda d: (str(d.get("ai", "")), str(d.get("tdoc_id", ""))))
        c_lines = [f"---", f"## 🏢 {comp} ({len(docs)} contribution{'s' if len(docs) != 1 else ''})", ""]
        for d in docs:
            c_lines.extend(_format_tdoc_digest_card(d, outcome_idx, eff_lifecycle, wg, meeting_folder_id))
        company_blocks.append("\n".join(c_lines))

    # Remove stale part files from previous runs before writing
    for old_f in meeting_dir_ext.glob("_company_positions_input*.md"):
        try:
            old_f.unlink()
        except OSError:
            pass

    if not company_blocks:
        out_p = meeting_dir / "_company_positions_input.md"
        with open(to_extended_path(out_p), "w", encoding="utf-8") as f:
            f.write(base_header + "\n*No scoped contributions in this meeting.*\n")
        return [out_p]

    # Partition into bounded parts along company boundaries
    parts: List[List[str]] = []
    current_part: List[str] = []
    current_len = len(base_header)

    for block in company_blocks:
        if current_part and (current_len + len(block) > max_chars_per_part):
            parts.append(current_part)
            current_part = [block]
            current_len = len(base_header) + len(block)
        else:
            current_part.append(block)
            current_len += len(block)
    if current_part:
        parts.append(current_part)

    written_paths: List[Path] = []
    if len(parts) == 1:
        out_p = meeting_dir / "_company_positions_input.md"
        with open(to_extended_path(out_p), "w", encoding="utf-8") as f:
            f.write(base_header + "\n" + "\n".join(parts[0]))
        written_paths.append(out_p)
    else:
        for idx, part_blocks in enumerate(parts, start=1):
            out_p = meeting_dir / f"_company_positions_input_part{idx}.md"
            part_banner = f"> **Part {idx} of {len(parts)}** — Merge extracted company stances across all parts into `{meeting_folder_id}/_company_positions_analysis.md`.\n\n"
            with open(to_extended_path(out_p), "w", encoding="utf-8") as f:
                f.write(base_header + part_banner + "\n".join(part_blocks))
            written_paths.append(out_p)

    return written_paths


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


def write_e2e_synthesis_prompt(
    target_dir: Path,
    topic: MetaTopicDefinition,
    meeting_records: List[Dict[str, Any]],
    synthesis_status: Optional[str] = None,
) -> Path:
    """
    Writes `company_positions_e2e_prompt.md` at the topic root containing copy-pasteable
    Stage-1 (per-meeting low-cost map + token-minimal delta merge) and Stage-2 (cross-WG E2E reduce)
    prompts + live status checklist.
    """
    out_p = target_dir / "company_positions_e2e_prompt.md"
    s2_path = target_dir / STAGE2_SYNTHESIS_FILENAME
    s2_exists = s2_path.is_file()
    eff_s2_status = synthesis_status or ("analyzed" if s2_exists else "pending")
    s2_badge = format_analysis_badge(eff_s2_status)

    pending_rows = []
    stale_rows = []
    done_rows = []
    skip_rows = []
    for rec in meeting_records:
        st, exists, a_rel = _resolve_rec_analysis_status(rec, target_dir)
        m_fid = rec["meeting_folder_id"]
        if st == "no_activity":
            skip_rows.append(m_fid)
        elif st == "stale":
            stale_rows.append(m_fid)
        elif st == "analyzed":
            done_rows.append(m_fid)
        else:
            pending_rows.append(m_fid)

    lines = [
        f"# 🌐 E2E Cross-WG Company Position Synthesis Guide: {topic.name}",
        "",
        "This workspace is structured for **token-efficient, resumable 2-stage AI analysis** across Working Groups:",
        "1. **Stage 1 (Map — Per-Meeting, Low/Medium Cost Model):** Analyze active meetings (`⬜ Pending`) once, and use token-minimal delta merges (`🔄 Stale`) whenever an already-analyzed meeting gains post-meeting outcomes or new revisions.",
        f"2. **Stage 2 (Reduce — Cross-WG Synthesis → [`{STAGE2_SYNTHESIS_FILENAME}`]({STAGE2_SYNTHESIS_FILENAME})):** Synthesize E2E company positions across WGs using only active `*/_company_positions_analysis.md` outputs + `agreements_timeline.md` + `ls_dialogue.md`.",
        "",
        f"**Current Status:** ✅ {len(done_rows)} Analyzed | 🔄 {len(stale_rows)} Stale (Delta Ready) | ⬜ {len(pending_rows)} Pending | ⚪ {len(skip_rows)} Skipped (0 Inputs) | **Stage-2 Report:** {s2_badge}"
        + (f" ([Open Report]({STAGE2_SYNTHESIS_FILENAME}))" if s2_exists else f" (`{STAGE2_SYNTHESIS_FILENAME}` not yet created)"),
        "",
        "---",
        "",
        "## 🔹 Stage 1 Prompt (Run Once Per Active Meeting — Supports Full & Delta Merge)",
        "```text",
        "Check each active meeting's status in the checklist below (skip rows marked '⚪ Skip (0 Inputs)'):",
        "",
        "- If Stage-1 Output is '⬜ Pending':",
        "  Read `<WG>/<MeetingFolder>/_meeting_brief.md` and `<WG>/<MeetingFolder>/_company_positions_input*.md`.",
        "  Write `<WG>/<MeetingFolder>/_company_positions_analysis.md` containing:",
        "  1. Key Technical Decisions & Outcomes in this meeting (and which proposals won/lost).",
        "  2. Company Stance Table: | Company | Architectural Position | Concrete Proposals (TDoc IDs) | Confirmed Allies | Outcome |",
        "     - Strict rule: ignore companies flagged as carry-over unconfirmed co-sponsors.",
        "  3. Open Issues / Contested Topics carrying forward.",
        "",
        "- If Stage-1 Output is '🔄 Stale (Delta Ready)':",
        "  TOKEN-SAVING RULE: Do NOT re-read `_company_positions_input*.md`.",
        "  Read ONLY `<WG>/<MeetingFolder>/_company_positions_analysis.md` and the `## 🔄 Delta Since Last Analysis` section in `<WG>/<MeetingFolder>/_meeting_brief.md`.",
        "  Merge the new decisions, revision transitions, and newly added/un-stubbed TDocs directly into `<WG>/<MeetingFolder>/_company_positions_analysis.md` in place (update stale stances/statuses and Lifecycle State; keep the single clean analysis file).",
        "",
        "Be concise, factual, and cite TDoc IDs for every claim.",
        "```",
        "",
        "### Meeting Analysis Checklist (Chronological)",
        "",
        "| Date | WG | Meeting | State | Active Docs | Decisions | Brief | Input Packet(s) | Stage-1 Output |",
        "| :--- | :--- | :--- | :--- | :---: | :---: | :--- | :--- | :--- |",
    ]

    for rec in meeting_records:
        wg = rec["wg"]
        m_fid = rec["meeting_folder_id"]
        d_str = rec.get("start_date", "") or "N/A"
        lc = rec.get("lifecycle") or ("concluded" if rec.get("decision_count", 0) else "upcoming")
        lc_badge = format_lifecycle_badge(lc)
        brief_rel = f"{wg}/{m_fid}/_meeting_brief.md"
        digests = rec.get("digest_rel_paths", [])
        digest_links = " · ".join(f"[{Path(p).name}]({p})" for p in digests) if digests else "—"
        st, exists, analysis_rel = _resolve_rec_analysis_status(rec, target_dir)
        badge = format_analysis_badge(st)
        if st == "no_activity":
            stage1_cell = f"{badge}" + (f" *([old file]({analysis_rel}))*" if exists else "")
        elif exists:
            stage1_cell = f"[{badge}]({analysis_rel})"
        else:
            stage1_cell = f"{badge} (`{analysis_rel}`)"
        lines.append(
            f"| {d_str} | **{wg}** | {m_fid} | {lc_badge} | {rec['tdoc_count']} | {rec['decision_count']} | [Brief]({brief_rel}) | {digest_links} | {stage1_cell} |"
        )

    lines.extend([
        "",
        "---",
        "",
        f"## 🔸 Stage 2 Prompt (Cross-WG End-to-End Reduce → `{STAGE2_SYNTHESIS_FILENAME}`)",
        "```text",
        f"You are analyzing the cross-WG evolution of '{topic.name}' across {', '.join(sorted(topic.target_wgs))}.",
        "Read:",
        "- `index.md`",
        "- `agreements_timeline.md`",
        "- `ls_dialogue.md`",
        "- Only active `*/*/_company_positions_analysis.md` files (skip any row marked '⚪ Skip (0 Inputs)' above; do NOT re-read raw TDocs unless verifying a specific quote).",
        f"- If `{STAGE2_SYNTHESIS_FILENAME}` already exists and only a few meetings were updated, read the existing `{STAGE2_SYNTHESIS_FILENAME}` and merge the updated meetings cleanly instead of regenerating unchanged sections from scratch.",
        "",
        f"Write/update `{STAGE2_SYNTHESIS_FILENAME}` at the topic root covering:",
        "1. Per-Company E2E Strategy Matrix (how each major vendor/operator aligns across WGs — e.g. consistency vs. contradictions between SA2 and RAN2 stances).",
        "2. Key Architectural Fault Lines & Alliances over time (chronological trajectory across meetings).",
        "3. Outcomes & Impact on Baseline Specifications (which company's vision shaped the final agreements/pCRs).",
        "4. Strategic Outlook & Open Conflicts entering the next meeting cycle.",
        "```",
        "",
    ])

    with open(to_extended_path(out_p), "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    return out_p


def write_master_index_markdown(
    target_dir: Path,
    topic: MetaTopicDefinition,
    manifest: ExportManifest,
    total_decisions: int,
    total_active_docs: int,
    stub_count: int,
    round_count: int,
    session_id: int,
    scope_mode: str,
    meeting_records: List[Dict[str, Any]],
    synthesis_status: Optional[str] = None,
) -> Path:
    """
    Generates root `index.md` using a single shared template for both full exports
    and fast offline checkup/sync passes.
    """
    now_str = datetime.datetime.now().strftime("%Y-%m-%d")
    s2_path = target_dir / STAGE2_SYNTHESIS_FILENAME
    s2_exists = s2_path.is_file()
    eff_s2_status = synthesis_status or ("analyzed" if s2_exists else "pending")

    concluded_cnt = sum(1 for r in meeting_records if r.get("lifecycle") == "concluded")
    upcoming_cnt = sum(1 for r in meeting_records if r.get("lifecycle") == "upcoming")
    no_act_cnt = sum(1 for r in meeting_records if r.get("lifecycle") == "no_activity")

    active_mtgs = 0
    analyzed_mtgs = 0
    stale_mtgs = 0
    pending_mtgs = 0
    for rec in meeting_records:
        st, _, _ = _resolve_rec_analysis_status(rec, target_dir)
        if st == "no_activity":
            continue
        active_mtgs += 1
        if st == "analyzed":
            analyzed_mtgs += 1
        elif st == "stale":
            stale_mtgs += 1
        else:
            pending_mtgs += 1

    stage1_summary = f"{analyzed_mtgs}/{active_mtgs} Active Meetings Analyzed"
    extra_s1 = []
    if stale_mtgs:
        extra_s1.append(f"🔄 {stale_mtgs} Stale (Delta Ready)")
    if pending_mtgs:
        extra_s1.append(f"⬜ {pending_mtgs} Pending")
    if no_act_cnt:
        extra_s1.append(f"⚪ {no_act_cnt} Empty Skipped")
    if extra_s1:
        stage1_summary += " (" + ", ".join(extra_s1) + ")"

    if s2_exists:
        s2_kpi = f"[{format_analysis_badge(eff_s2_status)}]({STAGE2_SYNTHESIS_FILENAME}) (`{STAGE2_SYNTHESIS_FILENAME}`)"
        s2_nav = f"- 🏆 **[Cross-WG E2E Company Position Synthesis]({STAGE2_SYNTHESIS_FILENAME})** — {format_analysis_badge(eff_s2_status)}."
    else:
        s2_kpi = f"⬜ Pending (`{STAGE2_SYNTHESIS_FILENAME}` not yet created)"
        s2_nav = f"- 🏆 **Cross-WG E2E Company Position Synthesis** — ⬜ Pending (generate `{STAGE2_SYNTHESIS_FILENAME}` using Stage 2 prompt)."

    contrib_kpi = f"{total_active_docs} documents"
    if stub_count > 0:
        contrib_kpi += f" *(including {stub_count} reserved/missing placeholder stub(s))*"

    delta_link_line = (
        f"- 🆕 **[Latest LLM Delta Feed](deltas/delta_session_{session_id}.md)** — Incremental additions in this session."
        if (target_dir / "deltas" / f"delta_session_{session_id}.md").is_file()
        else ""
    )

    lines = [
        f"# 🌐 Cross-WG Topic Dossier: {topic.name}",
        f"**Last Updated:** {now_str} (Session #{session_id}) | **Scope Mode:** `{scope_mode}`",
        "",
        "> [!NOTE]",
        f"> {topic.description or 'Consolidated multi-WG tracking organized by Working Group and Meeting.'}",
        "",
        "## 📊 Executive KPI Strip",
        "",
        f"- **Participating Working Groups:** {', '.join(sorted(topic.target_wgs))}",
        (
            f"- **Analyzed Meetings:** {len(meeting_records)} meetings "
            f"({concluded_cnt} Concluded, {upcoming_cnt} Upcoming, {no_act_cnt} No Activity) across {round_count} synchronized round(s)"
            if meeting_records
            else "- **Analyzed Meetings:** None"
        ),
        f"- **Unified Technical Decisions:** {total_decisions} (Agreements, WAs, pCRs, CRs, WFs)",
        f"- **Active Scoped Contributions:** {contrib_kpi}",
        f"- **Stage-1 Per-Meeting Analyses:** {stage1_summary}",
        f"- **Stage-2 Cross-WG Synthesis:** {s2_kpi}",
        "",
        "## 🧭 Navigation & LLM Workflows",
        "",
        "- ⚖️ **[Unified Technical Decisions Baseline](agreements_timeline.md)** — Chronological matrix of all RAN agreements and SA/CT pCRs/CRs.",
        "- 📬 **[Cross-WG Liaison Statement Dialogue](ls_dialogue.md)** — Inter-group question and reply threads.",
        "- 🧠 **[2-Stage E2E Company Position Prompt & Checklist](company_positions_e2e_prompt.md)** — Resumable, per-meeting LLM instructions designed for standard/non-frontier models.",
        s2_nav,
    ]
    if delta_link_line:
        lines.append(delta_link_line)

    lines.extend([
        "",
        "## 🗓️ Cross-WG Chronological Meeting Timeline",
        "",
        "| Date | Synchronized Round | WG | Meeting | State | Decisions | Active Inputs | LS (In/Out) | Meeting Brief & LLM Packet |",
        "| :--- | :--- | :--- | :--- | :--- | :---: | :---: | :---: | :--- |",
    ])

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
        st, exists, analysis_rel = _resolve_rec_analysis_status(rec, target_dir)
        a_badge = format_analysis_badge(st)
        if st == "no_activity":
            analysis_cell = "⚪ 0 Inputs"
        elif exists:
            analysis_cell = f"[{a_badge}]({analysis_rel})"
        else:
            analysis_cell = a_badge
        ls_cell = f"{rec.get('ls_in_count', 0)} / {rec.get('ls_out_count', 0)}"
        lines.append(
            f"| `{d_str}` | {r_lbl} | **{wg}** | [{m_fid}]({brief_rel}) | {lc_badge} | {rec['decision_count']} | {rec['tdoc_count']} | {ls_cell} | [Brief]({brief_rel}) · {digest_str} · {analysis_cell} |"
        )

    lines.extend([
        "",
        "## 🏛️ Per-Working-Group Evolution",
        "",
        "| Working Group | Chronological Meetings (Click to Open Brief) | Decisions | Active Docs | Stage-1 Analysis |",
        "| :--- | :--- | :---: | :---: | :--- |",
    ])

    by_wg: Dict[str, List[Dict[str, Any]]] = {}
    for rec in meeting_records:
        by_wg.setdefault(rec["wg"], []).append(rec)

    for wg in sorted(by_wg.keys()):
        recs = by_wg[wg]
        links = " → ".join(f"[{r['meeting_folder_id']}]({r['brief_rel_path']})" for r in recs)
        tot_d = sum(r["decision_count"] for r in recs)
        tot_c = sum(r["tdoc_count"] for r in recs)
        wg_statuses = [_resolve_rec_analysis_status(r, target_dir)[0] for r in recs]
        wg_active = [s for s in wg_statuses if s != "no_activity"]
        wg_done = sum(1 for s in wg_active if s == "analyzed")
        wg_stale = sum(1 for s in wg_active if s == "stale")
        if not wg_active:
            wg_s1 = f"⚪ 0 Active ({len(recs)} Empty)"
        else:
            wg_s1 = f"✅ {wg_done}/{len(wg_active)}"
            if wg_stale:
                wg_s1 += f" · 🔄 {wg_stale} Stale"
        lines.append(f"| **{wg}** | {links} | {tot_d} | {tot_c} | {wg_s1} |")

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
    out_p = target_dir / "index.md"
    with open(to_extended_path(out_p), "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    return out_p


def write_consolidated_briefing_markdown(
    target_dir: Path,
    topic: MetaTopicDefinition,
    all_decisions: List[TopicDecision],
    meeting_records: List[Dict[str, Any]],
    round_count: int,
) -> Path:
    """Generates single-file `consolidated_briefing.md` linking decisions, briefs, digests, and Stage-1 analyses."""
    lines = [
        f"# Consolidated Executive Briefing: {topic.name}",
        "",
        "## 1. Executive Summary & Scope",
        f"This briefing consolidates decisions and contribution briefs across {', '.join(sorted(topic.target_wgs))} "
        f"spanning {len(meeting_records)} meetings ({round_count} synchronized round(s)).",
        "",
        "## 2. Chronological Technical Decisions",
        "",
    ]
    for dec in sorted(all_decisions, key=lambda d: (d.date or d.cohort_key, d.wg, d.tdoc_id or "")):
        m_lbl = dec.meeting_folder_id or f"{dec.wg}_{dec.meeting_number}"
        link_suffix = f" ([{dec.tdoc_id or 'Agreement'}]({dec.dossier_rel_path}))" if dec.dossier_rel_path else ""
        lines.append(f"- **[{dec.date or dec.cohort_key} | {m_lbl}]** ({dec.decision_type}){link_suffix}: {dec.title_or_summary}")

    if meeting_records:
        lines.extend(["", "## 3. Per-Meeting Briefs, Digests & Stage-1 Analyses", ""])
        for rec in meeting_records:
            digests = rec.get("digest_rel_paths", [])
            digest_links = ", ".join(f"[{Path(p).name}]({p})" for p in digests) if digests else "—"
            st, exists, analysis_rel = _resolve_rec_analysis_status(rec, target_dir)
            a_suffix = f" · Analysis: [{format_analysis_badge(st)}]({analysis_rel})" if (exists and st != "no_activity") else ""
            lines.append(
                f"- **{rec.get('start_date') or 'N/A'} · {rec['meeting_folder_id']}**: "
                f"[{Path(rec['brief_rel_path']).name}]({rec['brief_rel_path']}) · Digests: {digest_links}{a_suffix}"
            )

    lines.append("")
    out_p = target_dir / "consolidated_briefing.md"
    with open(to_extended_path(out_p), "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    return out_p
