"""Companion AI citation map and evaluation prompt synthesizer."""

from __future__ import annotations

import datetime
import logging
from pathlib import Path
from typing import Any, Dict, List, Tuple

from core.utils.company_sanitizer import CompanySanitizer
from core.utils.paths import to_extended_path

logger = logging.getLogger(__name__)


def generate_ai_prompt(
    docs: List[Dict[str, Any]],
    page_spans: Dict[str, Tuple[int, int]],
    prompt_path: Path,
    mtg_info: Dict[str, Any],
) -> None:
    """
    Synthesizes a companion AI prompt detailing the exact page boundaries
    and structured evaluation directives for NotebookLM, Claude 3.5 Sonnet,
    Gemini 1.5 Pro, and Ollama.
    """
    wg_name = str(mtg_info.get("wg_name", "3GPP"))
    mtg_num = str(mtg_info.get("meeting_number", ""))
    timestamp_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")

    lines = [
        "# ================================================================================",
        "# 3GPP TECHNICAL DOSSIER — AI INGESTION & CITATION MAP",
        f"# Working Group : {wg_name} #{mtg_num}",
        f"# Generated     : {timestamp_str}",
        f"# Total TDocs   : {len(docs)}",
        "# ================================================================================",
        "",
        "## 1. DOCUMENT CITATION & EXACT PAGE MAP",
        "Use the exact page ranges below when answering questions or citing proposals:",
        "",
    ]

    for d in docs:
        tid = d["tdoc_id"]
        span = page_spans.get(tid.upper(), (1, 1))
        st, en = span
        span_text = f"pages {st} to {en}" if st != en else f"page {st}"
        title = d.get("title", "Unknown Title")
        source_raw = d.get("source", "Unknown Source")
        source_desc = CompanySanitizer.format_source_triage_context(source_raw)
        dtype = d.get("type", "Unknown Type")
        ai = d.get("ai", "Unknown AI")
        status = d.get("status", "")

        status_str = f" [Status: {status}]" if status else ""
        lines.append(
            f'- **{tid}**: "{title}"\n'
            f"  - Sourced by: {source_desc} | Type: {dtype} | AI: {ai}{status_str}\n"
            f"  - PDF Span: {span_text} ({en - st + 1} page(s))\n"
        )

    lines.extend([
        "",
        "# ================================================================================",
        "## 2. 3GPP TECHNICAL CONTRIBUTION EVALUATION DIRECTIVES",
        "# ================================================================================",
        "When analyzing the attached dossier, please follow these standardized 3GPP conventions:",
        "",
        "1. **Change Requests (CRs):**",
        "   - The first page contains the formal CR cover page. The 'Reason for change' and",
        "     'Summary of change' fields summarize the authoritative technical intent.",
        "   - The specification clauses that follow show proposed normative modifications.",
        "   - Focus on revision marks: text enclosed in '[INSERTED: ...]' (or green underline)",
        "     represents proposed additions; '[DELETED: ...]' (or red strikethrough) represents",
        "     proposed deletions.",
        "",
        "2. **Pseudo-CRs (pCRs):**",
        "   - The initial Discussion / Motivation section outlines the technical rationale.",
        "   - The subsequent clauses provide the proposed text for the Technical Report (TR)",
        "     or Technical Specification (TS). Both sections are relevant.",
        "",
        "3. **Discussion Papers & Study Item Contributions:**",
        "   - The entire document outlines technical positions, evaluation results, and architectural options.",
        "",
        "4. **Liaison Statements (LS IN / LS OUT):**",
        "   - Pay particular attention to the 'Actions Requested' and target deadlines.",
        "",
        "5. **Company Positions & Alliances:**",
        "   - Do not consider company names in brackets or with question marks as confirmed co-signers.",
        "   - Explicitly highlight which companies support or oppose specific solutions.",
        "",
        "# ================================================================================",
        "## 3. READY-TO-USE PROMPTS FOR NOTEBOOKLM / CLAUDE / GEMINI / CHATGPT",
        "# ================================================================================",
        "",
        "### [PROMPT A: Executive Agenda Item Summary]",
        '"Please analyze the attached 3GPP dossier using the page citation map above.',
        "For each Agenda Item included in the document:",
        "1. Provide a concise executive overview of the technical problem being addressed.",
        "2. Summarize each contribution, citing its TDoc ID and exact page numbers.",
        "3. Identify areas of consensus and points of technical disagreement between companies.",
        '4. Highlight any proposed changes to normative 3GPP specifications."',
        "",
        "### [PROMPT B: Company Stance & Conflict Analysis]",
        '"Based on the attached dossier, map out the positions taken by different companies.',
        "Which companies align on specific architectural choices, and which propose competing",
        'solutions? Reference each position with [TDoc ID, page X]."',
        "",
        "### [PROMPT C: Change Request & Normative Impact Audit]",
        '"Focus strictly on the CRs and pCRs in this dossier. For each:',
        "- Summarize the 'Reason for change' from the cover sheet.",
        "- Extract the exact normative alterations proposed in the revision-marked clauses.",
        '- Flag any potential backward compatibility risks or architectural side effects."',
        "",
    ])

    with open(to_extended_path(prompt_path), "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")
    logger.info(f"[DossierExporter] Saved companion AI prompt: {prompt_path}")
