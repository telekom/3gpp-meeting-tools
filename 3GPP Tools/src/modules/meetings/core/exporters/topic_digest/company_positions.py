# --- File: src/modules/meetings/core/exporters/topic_digest/company_positions.py ---
"""Token-bounded per-meeting company position packet builder for Stage-1 LLM extraction."""

import re
from pathlib import Path
from typing import Any, Dict, List, Optional

from core.utils.company_sanitizer import CompanySanitizer
from core.utils.paths import to_extended_path
from modules.meetings.core.topic_decision import TopicDecision

from .constants import DEFAULT_MAX_DIGEST_CHARS
from .status_and_outcomes import (
    _format_tdoc_digest_card,
    build_tdoc_outcome_index,
    classify_meeting_lifecycle,
    format_lifecycle_badge,
)


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
        c_lines = ["---", f"## 🏢 {comp} ({len(docs)} contribution{'s' if len(docs) != 1 else ''})", ""]
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
