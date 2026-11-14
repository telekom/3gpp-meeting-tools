# --- File: src/modules/meetings/core/exporters/topic_digest_builder.py ---
"""
Deterministic Per-Meeting Briefing & Token-Efficient LLM Digest Builder.

Produces compact, attribution-safe per-meeting packets (`_meeting_brief.md` and
`_company_positions_input.md`) plus root E2E synthesis instructions so non-frontier
models can extract company positions meeting-by-meeting without exceeding token budgets.
"""

import os
import re
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from core.utils.company_sanitizer import CompanySanitizer
from core.utils.paths import to_extended_path
from modules.meetings.core.exporters.export_manifest import sanitize_folder_name
from modules.meetings.core.topic_decision import TopicDecision
from modules.meetings.core.topic_manager import MetaTopicDefinition

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

    lines = [
        f"# 📋 Meeting Brief: {wg} #{mtg_num} (`{meeting_folder_id}`)",
        f"**Meta-Topic:** {topic_name} | **Synchronized Round:** {cohort_label}",
        f"**Dates:** {start_date or 'N/A'} → {end_date or 'N/A'}" + (f" | **Location:** {location}" if location else "") + (f" | **3GPP Folder:** `{folder_name}`" if folder_name else ""),
        "",
        "## 📊 Meeting Topic Footprint",
        f"- **Binding Decisions:** {len(decisions)}",
        f"- **Active Scoped Contributions:** {len(candidates)}" + (f" *(+{deduped_count} superseded/withdrawn revisions deduplicated)*" if deduped_count else ""),
    ]

    if digest_rel_paths:
        digest_links = ", ".join(f"[{Path(p).name}]({Path(p).name})" for p in digest_rel_paths)
        lines.append(f"- **LLM Company Position Packet(s):** {digest_links}")

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
            "| TDoc | Type | AI | Status | Source | Lineage | Dossier |",
            "| :--- | :--- | :--- | :--- | :--- | :--- | :--- |",
        ])
        for d in sorted(candidates, key=lambda x: (str(x.get("ai", "")), str(x.get("tdoc_id", "")))):
            src_cell = CompanySanitizer.format_source_markdown_table_cell(d.get("source", ""))
            preds = d.get("supersedes", [])
            lineage_str = f"rev of {', '.join(preds[-2:])}" if preds else "—"
            rel_local = _rel_to_meeting_dir(d.get("rel_path", ""), wg, meeting_folder_id)
            link_c = f"[{d['tdoc_id']}]({rel_local})" if rel_local else d["tdoc_id"]
            lines.append(
                f"| **{d['tdoc_id']}** | {d.get('type', '')} | `{d.get('ai', '')}` | {d.get('status', '')} | {src_cell} | {lineage_str} | {link_c} |"
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


def write_company_positions_input(
    meeting_dir: Path,
    topic_name: str,
    mtg: Dict[str, Any],
    meeting_folder_id: str,
    decisions: List[TopicDecision],
    candidates: List[Dict[str, Any]],
    max_chars_per_part: int = DEFAULT_MAX_DIGEST_CHARS,
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

    header_lines = [
        f"# 🧠 LLM Company Position Input Packet: {wg} #{mtg_num} (`{meeting_folder_id}`)",
        f"**Meta-Topic:** {topic_name} | **Date:** {start_date or 'N/A'} | **Active Docs:** {len(candidates)} | **Decisions:** {len(decisions)}",
        "",
        "> [!IMPORTANT] Stage-1 Per-Meeting Extraction Instructions (For Standard / Flash / Local LLMs)",
        f"> Save your output as `_company_positions_analysis.md` inside `{meeting_folder_id}/`.",
        "> 1. **Goal:** Extract each company's technical position, concrete proposals, alliances (confirmed co-signers only), and outcome (`Agreed`/`Approved`/`Noted`/`Revised`) for this meeting.",
        "> 2. **Attribution Rule:** Do NOT credit companies listed inside `[Note: ... carry-over co-sponsors]` — only credit Confirmed Lead and Confirmed Co-Signers.",
        "> 3. **Output Schema:**",
        ">    - Section A: **Decision & Consensus Summary** (key architectural choices locked in this meeting).",
        ">    - Section B: **Company Stance Matrix** (`| Company | Core Position | Key Proposals / TDocs | Allies (Confirmed) | Outcome |`).",
        ">    - Section C: **Open Points / Disagreements / FFS** carried into the next meeting.",
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
            tid = d["tdoc_id"]
            title = str(d.get("title", "")).strip()
            preds = d.get("supersedes", [])
            lineage = f"Supersedes {', '.join(preds)}" if preds else "Initial / Standalone"
            src_ctx = CompanySanitizer.format_source_triage_context(d.get("source", ""))
            rel_local = _rel_to_meeting_dir(d.get("rel_path", ""), wg, meeting_folder_id)

            c_lines.append(f"### {tid}: {title}")
            c_lines.append(
                f"- **Meta:** Type: `{d.get('type', '')}` | AI: `{d.get('ai', '')}` | Status: `{d.get('status', '')}` | Lineage: {lineage}"
            )
            c_lines.append(f"- **Attribution:** {src_ctx}")
            if rel_local:
                c_lines.append(f"- **Full Document:** [{rel_local}]({rel_local})")

            digest = extract_tdoc_position_digest(d.get("abs_path"))
            if digest["intro"]:
                c_lines.append(f"- **Summary / Context:**\n{digest['intro']}")
            if digest["proposals"]:
                c_lines.append(f"- **Proposals / Conclusions:**\n{digest['proposals']}")
            if digest["changes"]:
                c_lines.append(f"- **Specification Changes:**\n{digest['changes']}")
            if digest["fallback"]:
                c_lines.append(f"- **Extracted Content:**\n{digest['fallback']}")

            row = d.get("row_data", {})
            abstract = str(row.get("Abstract", "") or row.get("Secretary Remarks", "")).strip()
            if abstract and not digest["intro"] and not digest["proposals"] and not digest["fallback"]:
                c_lines.append(f"- **Abstract:** {abstract}")
            c_lines.append("")
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


def write_e2e_synthesis_prompt(
    target_dir: Path,
    topic: MetaTopicDefinition,
    meeting_records: List[Dict[str, Any]],
) -> Path:
    """
    Writes `company_positions_e2e_prompt.md` at the topic root containing copy-pasteable
    Stage-1 (per-meeting low-cost map) and Stage-2 (cross-WG E2E reduce) prompts + checklist.
    """
    out_p = target_dir / "company_positions_e2e_prompt.md"
    lines = [
        f"# 🌐 E2E Cross-WG Company Position Synthesis Guide: {topic.name}",
        "",
        "This workspace is structured for **token-efficient, resumable 2-stage AI analysis** across Working Groups:",
        "1. **Stage 1 (Map — Per-Meeting, Low/Medium Cost Model):** Analyze one meeting at a time using its pre-compacted `_company_positions_input.md` and `_meeting_brief.md`.",
        "2. **Stage 2 (Reduce — Cross-WG Synthesis):** Synthesize the E2E company positions across all WGs using only the compact Stage-1 outputs + `agreements_timeline.md` + `ls_dialogue.md`.",
        "",
        "---",
        "",
        "## 🔹 Stage 1 Prompt (Run Once Per Meeting)",
        "```text",
        "Read `<WG>/<MeetingFolder>/_meeting_brief.md` and `<WG>/<MeetingFolder>/_company_positions_input*.md`.",
        "Write `<WG>/<MeetingFolder>/_company_positions_analysis.md` containing:",
        "1. Key Technical Decisions & Outcomes in this meeting (and which proposals won/lost).",
        "2. Company Stance Table: | Company | Architectural Position | Concrete Proposals (TDoc IDs) | Confirmed Allies | Outcome |",
        "   - Strict rule: ignore companies flagged as carry-over unconfirmed co-sponsors.",
        "3. Open Issues / Contested Topics carrying forward.",
        "Be concise, factual, and cite TDoc IDs for every claim.",
        "```",
        "",
        "### Meeting Analysis Checklist (Chronological)",
        "",
        "| Date | WG | Meeting | Active Docs | Decisions | Brief | Input Packet(s) | Stage-1 Output |",
        "| :--- | :--- | :--- | :---: | :---: | :--- | :--- | :--- |",
    ]

    for rec in meeting_records:
        wg = rec["wg"]
        m_fid = rec["meeting_folder_id"]
        d_str = rec.get("start_date", "") or "N/A"
        brief_rel = f"{wg}/{m_fid}/_meeting_brief.md"
        digests = rec.get("digest_rel_paths", [])
        digest_links = " · ".join(f"[{Path(p).name}]({p})" for p in digests) if digests else "—"
        analysis_rel = f"{wg}/{m_fid}/_company_positions_analysis.md"
        done_mark = "✅" if (target_dir / analysis_rel).is_file() else "⬜"
        lines.append(
            f"| {d_str} | **{wg}** | {m_fid} | {rec['tdoc_count']} | {rec['decision_count']} | [Brief]({brief_rel}) | {digest_links} | {done_mark} `{analysis_rel}` |"
        )

    lines.extend([
        "",
        "---",
        "",
        "## 🔸 Stage 2 Prompt (Cross-WG End-to-End Reduce)",
        "```text",
        f"You are analyzing the cross-WG evolution of '{topic.name}' across {', '.join(sorted(topic.target_wgs))}.",
        "Read:",
        "- `index.md`",
        "- `agreements_timeline.md`",
        "- `ls_dialogue.md`",
        "- All `*/_company_positions_analysis.md` files (do NOT re-read raw TDocs unless verifying a specific quote).",
        "",
        "Produce an End-to-End Cross-WG Company Position Report covering:",
        "1. Per-Company E2E Strategy Matrix (how each major vendor/operator aligns across SA2, RAN2, RAN3, SA5, CT4 — e.g. consistency vs. contradictions between SA2 and RAN2 stances).",
        "2. Key Architectural Fault Lines & Alliances over time (chronological trajectory across meetings).",
        "3. Outcomes & Impact on Baseline Specifications (which company's vision shaped the final agreements/pCRs).",
        "4. Strategic Outlook & Open Conflicts entering the next meeting cycle.",
        "```",
        "",
    ])

    with open(to_extended_path(out_p), "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    return out_p
