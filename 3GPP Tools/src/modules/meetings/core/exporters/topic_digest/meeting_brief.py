# --- File: src/modules/meetings/core/exporters/topic_digest/meeting_brief.py ---
"""Per-meeting executive brief generation and non-destructive Stage-1 delta synchronization."""

import re
from pathlib import Path
from typing import Any, Dict, List, Optional

from core.utils.company_sanitizer import CompanySanitizer
from core.utils.paths import to_extended_path
from modules.meetings.core.topic_decision import TopicDecision

from .constants import STAGE1_DELTA_END, STAGE1_DELTA_START
from .extraction import _rel_to_meeting_dir, build_contributor_matrix
from .status_and_outcomes import (
    _format_analysis_footprint_line,
    _format_tdoc_digest_card,
    build_tdoc_outcome_index,
    classify_meeting_lifecycle,
    format_lifecycle_badge,
)


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
