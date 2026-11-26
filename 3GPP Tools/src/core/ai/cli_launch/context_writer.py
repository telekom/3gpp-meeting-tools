"""CLI workspace context (`_CLI_CONTEXT.md`) and standing/per-TDoc user notes helpers."""

from __future__ import annotations

import datetime
import logging
import os
import re
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence

from core.utils.paths import get_project_root, strip_extended_path, to_extended_path

logger = logging.getLogger(__name__)

CLI_CONTEXT_FILENAME = "_CLI_CONTEXT.md"
USER_NOTES_FILENAME = "_USER_NOTES.md"
STANDING_CONTEXT_PATH = get_project_root() / "config" / "cli_user_context.md"
DEFAULT_BOOTSTRAP_PROMPT = "Read _CLI_CONTEXT.md in this directory and follow its instructions."
_BASE_TDOC_RE = re.compile(r"^(.*?)-?(?:r|rev)\d{1,2}[a-zA-Z]?$", re.IGNORECASE)


def _resolve_standing_context_path() -> Path:
    """Resolve STANDING_CONTEXT_PATH via facade so patch.object(cli_launcher, ...) works."""
    facade = sys.modules.get("core.ai.cli_launcher")
    if facade is not None and hasattr(facade, "STANDING_CONTEXT_PATH"):
        return Path(facade.STANDING_CONTEXT_PATH)
    return STANDING_CONTEXT_PATH


def load_standing_cli_context() -> str:
    """Read global standing company / delegate guidance from ``src/config/cli_user_context.md``."""
    path = _resolve_standing_context_path()
    try:
        if path.is_file():
            return path.read_text(encoding="utf-8").strip()
    except Exception as exc:
        logger.debug("Could not read standing CLI context: %s", exc)
    return ""


def save_standing_cli_context(text: str) -> Path:
    """Persist global standing company / delegate guidance to ``src/config/cli_user_context.md``."""
    path = _resolve_standing_context_path()
    clean = (text or "").strip()
    path.parent.mkdir(parents=True, exist_ok=True)
    if clean:
        path.write_text(clean + "\n", encoding="utf-8")
    elif path.exists():
        try:
            path.unlink()
        except OSError:
            path.write_text("", encoding="utf-8")
    return path


def load_workspace_user_notes(target_dir: Path) -> str:
    """Read persistent workspace/TDoc user notes from ``{target_dir}/_USER_NOTES.md``."""
    notes_path = Path(strip_extended_path(target_dir)) / USER_NOTES_FILENAME
    try:
        ext_path = Path(to_extended_path(notes_path))
        if ext_path.is_file():
            return ext_path.read_text(encoding="utf-8").strip()
    except Exception as exc:
        logger.debug("Could not read workspace user notes in %s: %s", target_dir, exc)
    return ""


def save_workspace_user_notes(target_dir: Path, text: str) -> Path:
    """Write or clear ``{target_dir}/_USER_NOTES.md`` without disturbing converted artifacts."""
    clean_dir = Path(strip_extended_path(target_dir))
    notes_path = clean_dir / USER_NOTES_FILENAME
    ext_notes = Path(to_extended_path(notes_path))
    clean_text = (text or "").strip()
    if clean_text:
        os.makedirs(to_extended_path(clean_dir), exist_ok=True)
        ext_notes.write_text(clean_text + "\n", encoding="utf-8")
    elif ext_notes.exists():
        try:
            ext_notes.unlink()
        except OSError:
            ext_notes.write_text("", encoding="utf-8")
    return notes_path


def collect_per_tdoc_user_notes(
    meeting_dir: Optional[Path],
    tdocs_meta: Optional[Sequence[Dict[str, Any]]],
    exclude_dir: Optional[Path] = None,
) -> Dict[str, str]:
    """
    Collect saved ``_USER_NOTES.md`` files across canonical ``{meeting_dir}/{base_tdoc}/``
    folders for multi-TDoc batches so per-TDoc guidance automatically carries over.
    """
    if not meeting_dir or not tdocs_meta:
        return {}
    clean_meeting = Path(strip_extended_path(meeting_dir))
    clean_exclude = os.path.normcase(str(Path(strip_extended_path(exclude_dir)))) if exclude_dir else ""
    collected: Dict[str, str] = {}
    seen_bases: set = set()

    for row in tdocs_meta:
        tid = str((row or {}).get("TDoc", "")).strip()
        if not tid:
            continue
        m = _BASE_TDOC_RE.match(tid)
        base_tdoc = (m.group(1) if m else tid).upper()
        if base_tdoc in seen_bases:
            continue
        seen_bases.add(base_tdoc)

        tdoc_dir = clean_meeting / base_tdoc
        if clean_exclude and os.path.normcase(str(tdoc_dir)) == clean_exclude:
            continue
        notes = load_workspace_user_notes(tdoc_dir)
        if notes:
            collected[base_tdoc] = notes
    return collected


def _resolve_cli_tools_script() -> Path:
    """Resolve the path to `src/cli_tools.py` regardless of launch layout."""
    proj = get_project_root()
    direct = proj / "cli_tools.py"
    if direct.is_file():
        return direct
    nested = proj / "src" / "cli_tools.py"
    if nested.is_file():
        return nested
    return direct


def _build_task_instruction(
    task_mode: str,
    custom_question: str = "",
    topic_name: str = "",
    tdoc_ids: Optional[List[str]] = None,
) -> str:
    """Build the primary task instruction block for `_CLI_CONTEXT.md`."""
    clean_q = (custom_question or "").strip()
    if clean_q:
        return clean_q

    ids_str = ", ".join(tdoc_ids or [])
    mode = (task_mode or "interactive").strip().lower()

    if mode == "summarize":
        return (
            f"Read the converted Markdown document(s) ({ids_str or 'in this workspace'}), "
            "provide a crisp technical summary (problem statement, proposed mechanism, "
            "affected specifications/clauses, and potential impacts/open issues), "
            "and stand by for follow-up questions."
        )
    if mode == "compare":
        return (
            f"Read and compare all converted Markdown documents in this workspace ({ids_str or 'listed below'}). "
            "Highlight technical differences, revision deltas or competing company positions, "
            "trade-offs, and areas of alignment vs. conflict. Stand by for follow-up questions."
        )
    if mode == "topic_stage1":
        return (
            f"Perform Stage-1 meeting-by-meeting analysis for topic '{topic_name or 'this topic'}' "
            "focusing on the pending and stale meeting folders listed below. For each target meeting, "
            "read `_meeting_brief.md` and `_company_positions_input.md` (plus referenced TDoc `.md` files "
            "when claims need verification) and produce/update `_company_positions_analysis.md` inside "
            "that meeting folder."
        )
    if mode == "topic_stage2":
        return (
            f"Perform Stage-2 End-to-End Cross-WG synthesis for topic '{topic_name or 'this topic'}'. "
            "Read `company_positions_e2e_prompt.md`, `index.md`, `agreements_timeline.md`, `ls_dialogue.md`, "
            "and each meeting's `_company_positions_analysis.md` to produce/update "
            "`company_positions_e2e_synthesis.md`."
        )
    if mode in ("topic_idle", "topic_qa"):
        return (
            f"You have been launched in the Cross-WG Topic Dossier for '{topic_name or 'this 3GPP topic'}'. "
            "Do NOT automatically read or analyze all files or meetings across the folder. "
            "Wait for the user's question first, and inspect only `index.md`, `agreements_timeline.md`, "
            "`ls_dialogue.md`, or specific meetings/TDocs as needed to answer what the user asks."
        )
    if mode == "briefing_qa":
        return (
            "You have been launched in a Pre/Post-Meeting Dossier workspace. "
            "Do NOT automatically read or analyze all files in this directory. "
            "Wait for the user's question first, starting with `index.md` only when asked."
        )

    return (
        "Read the converted 3GPP Markdown document(s) listed below, summarize the key technical "
        "points and specification/protocol impact concisely, and stand by for follow-up questions "
        "in this interactive session."
    )


def write_cli_context_file(
    target_dir: Path,
    *,
    task_mode: str = "interactive",
    custom_question: str = "",
    tdocs_meta: Optional[Sequence[Dict[str, Any]]] = None,
    markdown_files: Optional[Sequence[Path]] = None,
    local_triage_summary: str = "",
    topic_name: str = "",
    manifest_summary: Optional[Dict[str, Any]] = None,
    briefing_title: str = "",
    meeting_dir: Optional[Path] = None,
    user_notes_override: Optional[str] = None,
    standing_context_override: Optional[str] = None,
) -> Path:
    """
    Generate `_CLI_CONTEXT.md` inside `target_dir` with workspace files,
    TDoc/Topic metadata, persistent company/TDoc `_USER_NOTES.md` priorities,
    optional local Ollama triage output, and 3GPP domain rules.
    """
    clean_dir = Path(strip_extended_path(target_dir))
    os.makedirs(to_extended_path(clean_dir), exist_ok=True)

    tdocs_list = list(tdocs_meta or [])
    tdoc_ids = [str(r.get("TDoc", "")).strip() for r in tdocs_list if str(r.get("TDoc", "")).strip()]

    lines: List[str] = [
        "# 3GPP Tools — AI Agent Context & Briefing",
        "",
        f"- **Generated:** {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        f"- **Workspace:** `{clean_dir}`",
    ]
    if topic_name:
        lines.append(f"- **Cross-WG Topic:** {topic_name}")
    if briefing_title:
        lines.append(f"- **Briefing:** {briefing_title}")
    if tdoc_ids:
        lines.append(f"- **Target TDoc(s):** {', '.join(tdoc_ids)}")

    lines.extend(
        [
            "",
            "## 1. Current Task & User Instruction",
            "",
            _build_task_instruction(
                task_mode=task_mode,
                custom_question=custom_question,
                topic_name=topic_name,
                tdoc_ids=tdoc_ids,
            ),
            "",
        ]
    )

    standing_ctx = (
        standing_context_override.strip()
        if standing_context_override is not None
        else load_standing_cli_context()
    )
    workspace_notes = (
        user_notes_override.strip()
        if user_notes_override is not None
        else load_workspace_user_notes(clean_dir)
    )
    per_tdoc_notes = collect_per_tdoc_user_notes(meeting_dir, tdocs_list, exclude_dir=clean_dir)

    if standing_ctx or workspace_notes or per_tdoc_notes:
        lines.extend(["## 2. Delegate & Company Priorities (Must Follow)", ""])
        if standing_ctx:
            lines.extend(["### Standing Company / Delegate Context (`cli_user_context.md`)", "", standing_ctx, ""])
        if workspace_notes:
            lines.extend(
                [
                    f"### Workspace Notes (`{USER_NOTES_FILENAME}` — re-read if edited mid-session)",
                    "",
                    workspace_notes,
                    "",
                ]
            )
        if per_tdoc_notes:
            lines.extend(["### Per-TDoc Saved Delegate Notes (from canonical TDoc folders)", ""])
            for b_tid, b_notes in per_tdoc_notes.items():
                lines.append(f"#### {b_tid}")
                lines.append(b_notes)
                lines.append("")

    lines.extend(["## Workspace Documents & Metadata", ""])

    excluded_md_names = {CLI_CONTEXT_FILENAME, USER_NOTES_FILENAME}
    rel_files: List[str] = []
    if markdown_files:
        for p in markdown_files:
            p_clean = Path(strip_extended_path(p))
            try:
                rel = os.path.relpath(p_clean, clean_dir).replace("\\", "/")
            except ValueError:
                rel = str(p_clean)
            if rel not in rel_files:
                rel_files.append(rel)
    else:
        for candidate in (
            "index.md",
            "agreements_timeline.md",
            "ls_dialogue.md",
            "company_positions_e2e_prompt.md",
            "company_positions_e2e_synthesis.md",
            "COMPANY_VIEWS_AND_PROGRESS.md",
        ):
            if (clean_dir / candidate).is_file():
                rel_files.append(candidate)
        for md_path in sorted(clean_dir.glob("*.md")):
            if md_path.name in excluded_md_names:
                continue
            if md_path.name not in rel_files:
                rel_files.append(md_path.name)
        for md_path in sorted(clean_dir.glob("*/*.md")):
            if md_path.name in excluded_md_names:
                continue
            rel = os.path.relpath(md_path, clean_dir).replace("\\", "/")
            if rel not in rel_files and len(rel_files) < 60:
                rel_files.append(rel)

    if rel_files:
        lines.append("### Available Markdown Files")
        for rf in rel_files[:60]:
            lines.append(f"- `{rf}`")
        if len(rel_files) > 60:
            lines.append(f"- *(... plus {len(rel_files) - 60} more `.md` files; inspect `index.md` for full navigation)*")
        lines.append("")

    if tdocs_list:
        lines.append("### TDoc Metadata")
        for row in tdocs_list:
            tid = str(row.get("TDoc", "")).strip()
            if not tid:
                continue
            title = str(row.get("Title", "")).strip()
            src = str(row.get("Source", "")).strip()
            dtype = str(row.get("Type", "")).strip()
            ai = str(row.get("Agenda Item", "")).strip()
            wi = str(row.get("Related WIs", "")).strip()
            status = str(row.get("TDoc Status", "")).strip()
            remarks = str(row.get("Secretary Remarks", "")).strip()
            my_status = str(row.get("My Status", "")).strip()
            my_notes = str(row.get("My Notes", "")).strip()

            lines.append(f"#### {tid}" + (f": {title}" if title else ""))
            if dtype or ai or wi:
                lines.append(f"- **Type:** {dtype or 'N/A'} | **Agenda Item:** {ai or 'N/A'} | **Work Item(s):** {wi or 'N/A'}")
            if src:
                lines.append(f"- **Source (Raw):** {src}")
            if status:
                lines.append(f"- **TDoc Status:** {status}")
            if remarks:
                lines.append(f"- **Secretary Remarks:** {remarks}")
            if my_status and my_status != "⚪ Neutral":
                lines.append(f"- **Delegate Stance:** {my_status}")
            if my_notes:
                lines.append(f"- **Delegate Notes:** {my_notes}")
            lines.append("")

    if manifest_summary:
        lines.append("### Topic Corpus Status")
        pending_mtgs = manifest_summary.get("pending_meetings") or []
        stale_mtgs = manifest_summary.get("stale_meetings") or []
        analyzed_mtgs = manifest_summary.get("analyzed_meetings") or []
        lines.append(
            f"- **Analyzed:** {len(analyzed_mtgs)} | **Stale:** {len(stale_mtgs)} | **Pending:** {len(pending_mtgs)}"
        )
        if pending_mtgs:
            lines.append(f"- **Pending Meetings:** {', '.join(pending_mtgs)}")
        if stale_mtgs:
            lines.append(f"- **Stale Meetings (Delta Ready):** {', '.join(stale_mtgs)}")
        if manifest_summary.get("synthesis_status"):
            lines.append(f"- **Stage-2 E2E Synthesis Status:** {manifest_summary['synthesis_status']}")
        lines.append("")

    if local_triage_summary and local_triage_summary.strip():
        lines.extend(["### Prior Local AI Triage Summary (Ollama)", "", local_triage_summary.strip(), ""])

    cli_tools_path = _resolve_cli_tools_script()
    py_exe = sys.executable or "python"
    lines.extend(
        [
            "## 3. Critical 3GPP Domain Guardrails & Tools",
            "",
            "1. **Tentative / Carry-Over Co-Signers (`?`, `[...]`):** Companies marked with `?`, `(?)`, or `[...]` "
            "(other than role tags like `[Rapporteur]` or `[pen holder]`) are unconfirmed carry-over co-signers from an "
            "earlier revision. **Never** attribute a proposal to a tentative company as a confirmed supporter.",
            "2. **Revision Markup Semantics (`markup` vs `accept`):** In CRs and pCRs, `[INSERTED: ...]` and `[DELETED: ...]` "
            "denote proposed changes relative to the 3GPP specification baseline. Liaison Statements (`LS in`, `LS out`) "
            "are exported in clean `accept` mode.",
            "3. **3GPP Specifications & Cross-Meeting Lookups:** If `3gpp-tools` MCP tools are attached in this CLI, "
            "use `spec_text_and_clauses`, `protocol_inspector`, `meetings_and_tdocs`, and `spec_history_and_attribution` "
            "to inspect baseline clauses or related contributions. Otherwise, you can query the local 3GPP Tools databases "
            "directly via the headless CLI:",
            "```powershell",
            f'& "{py_exe}" "{cli_tools_path}" --profile all --read-only list',
            f'& "{py_exe}" "{cli_tools_path}" --profile all --read-only call spec_text_and_clauses --action search_text --arg spec_number=23.501 --arg query="..." --arg limit=5',
            "```",
            "",
        ]
    )

    ctx_path = clean_dir / CLI_CONTEXT_FILENAME
    with open(to_extended_path(ctx_path), "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    return ctx_path
