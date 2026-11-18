# --- File: src/core/ai/cli_launcher.py ---
"""
External LLM CLI Context Bootstrapper & Terminal Launcher.

Writes `_CLI_CONTEXT.md` briefings (TDocs, revision families, local Ollama
triage carry-over, Pre/Post-Meeting exports, or Cross-WG Topic Dossiers)
and opens an interactive terminal session (`wt.exe` / `powershell.exe`)
for Codex CLI, Claude CLI, or Gemini CLI.
"""

from __future__ import annotations

import base64
import datetime
import logging
import os
import re
import shutil
import subprocess
import sys
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Sequence

from core.ai.cli_detector import (
    CLI_AGENT_SPECS,
    CliAgentStatus,
    build_interactive_cli_args,
    detect_cli_agents,
    discover_cli_models,
    format_active_model_badge,
    get_cli_agent_model_config,
    set_cli_agent_model_config,
)
from core.utils.paths import get_project_root, strip_extended_path, to_extended_path

logger = logging.getLogger(__name__)

CLI_CONTEXT_FILENAME = "_CLI_CONTEXT.md"
USER_NOTES_FILENAME = "_USER_NOTES.md"
STANDING_CONTEXT_PATH = get_project_root() / "config" / "cli_user_context.md"
DEFAULT_BOOTSTRAP_PROMPT = "Read _CLI_CONTEXT.md in this directory and follow its instructions."
_BASE_TDOC_RE = re.compile(r"^(.*?)-?(?:r|rev)\d{1,2}[a-zA-Z]?$", re.IGNORECASE)


def load_standing_cli_context() -> str:
    """Read global standing company / delegate guidance from ``src/config/cli_user_context.md``."""
    try:
        if STANDING_CONTEXT_PATH.is_file():
            return STANDING_CONTEXT_PATH.read_text(encoding="utf-8").strip()
    except Exception as exc:
        logger.debug("Could not read standing CLI context: %s", exc)
    return ""


def save_standing_cli_context(text: str) -> Path:
    """Persist global standing company / delegate guidance to ``src/config/cli_user_context.md``."""
    clean = (text or "").strip()
    STANDING_CONTEXT_PATH.parent.mkdir(parents=True, exist_ok=True)
    if clean:
        STANDING_CONTEXT_PATH.write_text(clean + "\n", encoding="utf-8")
    elif STANDING_CONTEXT_PATH.exists():
        try:
            STANDING_CONTEXT_PATH.unlink()
        except OSError:
            STANDING_CONTEXT_PATH.write_text("", encoding="utf-8")
    return STANDING_CONTEXT_PATH


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


def _ps_quote(value: str) -> str:
    """Single-quote a string safely for PowerShell `-Command`."""
    return "'" + str(value).replace("'", "''") + "'"


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

    lines.extend([
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
    ])

    # Layer persistent Delegate & Company Priorities right after the primary instruction
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
        lines.extend([
            "## 2. Delegate & Company Priorities (Must Follow)",
            "",
        ])
        if standing_ctx:
            lines.extend([
                "### Standing Company / Delegate Context (`cli_user_context.md`)",
                "",
                standing_ctx,
                "",
            ])
        if workspace_notes:
            lines.extend([
                f"### Workspace Notes (`{USER_NOTES_FILENAME}` — re-read if edited mid-session)",
                "",
                workspace_notes,
                "",
            ])
        if per_tdoc_notes:
            lines.extend([
                "### Per-TDoc Saved Delegate Notes (from canonical TDoc folders)",
                "",
            ])
            for b_tid, b_notes in per_tdoc_notes.items():
                lines.append(f"#### {b_tid}")
                lines.append(b_notes)
                lines.append("")

    lines.extend([
        "## Workspace Documents & Metadata",
        "",
    ])

    # Workspace Markdown files
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

    # TDoc metadata table/cards
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

    # Topic manifest status (for Cross-WG Topic Dossiers)
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

    # Prior Local Ollama Triage carry-over
    if local_triage_summary and local_triage_summary.strip():
        lines.extend([
            "### Prior Local AI Triage Summary (Ollama)",
            "",
            local_triage_summary.strip(),
            "",
        ])

    # Critical 3GPP Domain Guardrails & MCP/CLI instructions
    cli_tools_path = _resolve_cli_tools_script()
    py_exe = sys.executable or "python"
    lines.extend([
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
    ])

    ctx_path = clean_dir / CLI_CONTEXT_FILENAME
    with open(to_extended_path(ctx_path), "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    return ctx_path


def build_powershell_launch_command(
    agent_id: str,
    executable: str,
    working_dir: Path,
    bootstrap_prompt: str = DEFAULT_BOOTSTRAP_PROMPT,
    *,
    model: Optional[str] = None,
    reasoning_effort: Optional[str] = None,
    stage_clipboard_prompt: Optional[str] = None,
) -> str:
    """
    Construct a quote-safe PowerShell command string that moves into `working_dir`,
    prepends `executable`'s parent directory to `$env:PATH`, optionally stages a prompt
    on the Windows clipboard, and invokes the CLI in interactive mode.
    """
    clean_dir = strip_extended_path(working_dir)
    clean_exe = strip_extended_path(executable)
    exe_dir = str(Path(clean_exe).parent)
    cli_args = build_interactive_cli_args(
        agent_id,
        clean_exe,
        bootstrap_prompt,
        model=model,
        reasoning_effort=reasoning_effort,
    )

    quoted_exe = _ps_quote(cli_args[0])
    quoted_rest = " ".join(_ps_quote(a) for a in cli_args[1:])
    invoke_expr = f"& {quoted_exe}" + (f" {quoted_rest}" if quoted_rest else "")

    parts = [
        f"Set-Location -LiteralPath {_ps_quote(clean_dir)}",
    ]
    if exe_dir:
        parts.append(f"$env:PATH = {_ps_quote(exe_dir)} + ';' + $env:PATH")
    staged = (stage_clipboard_prompt or "").strip()
    if staged:
        preview = staged.replace("\r", " ").replace("\n", " ")
        if len(preview) > 140:
            preview = preview[:137] + "..."
        parts.append(f"try {{ Set-Clipboard -Value {_ps_quote(staged)} }} catch {{}}")
        parts.append("Write-Host '[3GPP Tools] Prompt copied to clipboard (press Ctrl+V to paste when ready):' -ForegroundColor Cyan")
        parts.append(f"Write-Host {_ps_quote('  ' + preview)} -ForegroundColor DarkGray")
    parts.append(invoke_expr)
    return "; ".join(parts)


def encode_powershell_command(ps_command: str) -> str:
    """
    Encode a PowerShell command string as UTF-16LE Base64 for `-EncodedCommand`.

    Using `-EncodedCommand` eliminates all semicolons (`;`), quotes, spaces, and
    `&` operators from the outer command line so `wt.exe` (which treats unescaped
    `;` as a Windows Terminal subcommand separator) never splits the command.
    """
    return base64.b64encode(ps_command.encode("utf-16-le")).decode("ascii")


def launch_cli_in_terminal(
    agent_id: str,
    working_dir: Path,
    bootstrap_prompt: Optional[str] = None,
    parent: Optional[Any] = None,
    *,
    model: Optional[str] = None,
    reasoning_effort: Optional[str] = None,
    auto_execute: bool = False,
    statuses: Optional[Dict[str, CliAgentStatus]] = None,
    popen_fn: Optional[Callable[..., Any]] = None,
    which_fn: Optional[Callable[[str], Optional[str]]] = None,
) -> bool:
    """
    Open Windows Terminal (`wt.exe`) or PowerShell (`powershell.exe`) in `working_dir`
    and launch the requested interactive LLM CLI (`codex`, `claude`, or `gemini`).

    When ``auto_execute`` is ``False`` (default), the CLI starts idle with zero
    automatic prompt submission (preventing unintended token burn), and any
    prepared prompt is staged onto the Windows clipboard for ``Ctrl+V``.
    """
    clean_dir = Path(strip_extended_path(working_dir))
    if not clean_dir.is_dir():
        _show_warning(parent, "Directory Not Found", f"The target directory does not exist:\n{clean_dir}")
        return False

    resolved_statuses = statuses if statuses is not None else detect_cli_agents()
    agent_key = (agent_id or "").strip().lower()
    st = resolved_statuses.get(agent_key)

    if not st or not st.installed or not st.executable_path:
        spec = CLI_AGENT_SPECS.get(agent_key)
        hint = f"\n\nInstall hint:\n{spec.install_hint}" if spec else ""
        _show_warning(
            parent,
            "CLI Agent Not Installed",
            f"{spec.label if spec else agent_id} was not found on this machine.{hint}",
        )
        return False

    raw_prompt = bootstrap_prompt if bootstrap_prompt is not None else DEFAULT_BOOTSTRAP_PROMPT
    if auto_execute:
        exec_prompt = raw_prompt
        staged_prompt = None
    else:
        exec_prompt = ""
        staged_prompt = raw_prompt.strip() or None

    ps_command = build_powershell_launch_command(
        agent_id=agent_key,
        executable=st.executable_path,
        working_dir=clean_dir,
        bootstrap_prompt=exec_prompt,
        model=model,
        reasoning_effort=reasoning_effort,
        stage_clipboard_prompt=staged_prompt,
    )
    encoded_cmd = encode_powershell_command(ps_command)

    spawn = popen_fn or subprocess.Popen
    which_lookup = which_fn or shutil.which

    wt_exe = which_lookup("wt.exe") or which_lookup("wt")
    if wt_exe:
        try:
            spawn(
                [wt_exe, "-d", str(clean_dir), "powershell.exe", "-NoExit", "-EncodedCommand", encoded_cmd],
                cwd=str(clean_dir),
            )
            logger.info("Launched %s via Windows Terminal in %s", st.label, clean_dir)
            return True
        except FileNotFoundError:
            pass
        except OSError as exc:
            logger.warning("wt.exe launch failed (%s); falling back to powershell.exe", exc)

    try:
        kwargs: Dict[str, Any] = {"cwd": str(clean_dir)}
        if os.name == "nt" and hasattr(subprocess, "CREATE_NEW_CONSOLE"):
            kwargs["creationflags"] = subprocess.CREATE_NEW_CONSOLE
        spawn(
            ["powershell.exe", "-NoExit", "-EncodedCommand", encoded_cmd],
            **kwargs,
        )
        logger.info("Launched %s via PowerShell in %s", st.label, clean_dir)
        return True
    except OSError as exc:
        _show_critical(parent, "Terminal Launch Error", f"Could not launch {st.label}:\n{exc}")
        return False


def _show_warning(parent: Any, title: str, msg: str) -> None:
    try:
        from PyQt5.QtWidgets import QApplication, QMessageBox
        if QApplication.instance() is not None:
            QMessageBox.warning(parent, title, msg)
            return
    except Exception:
        pass
    logger.warning("%s: %s", title, msg)


def _show_critical(parent: Any, title: str, msg: str) -> None:
    try:
        from PyQt5.QtWidgets import QApplication, QMessageBox
        if QApplication.instance() is not None:
            QMessageBox.critical(parent, title, msg)
            return
    except Exception:
        pass
    logger.error("%s: %s", title, msg)


def open_mcp_agent_setup_dialog(parent: Any = None) -> None:
    """Open `McpMonitorDialog` directly on the AI Agent Client Setup tab."""
    try:
        from core.ui.mcp_monitor_dialog import McpMonitorDialog
        dlg = McpMonitorDialog(parent=parent)
        if hasattr(dlg, "tabs") and hasattr(dlg, "setup_tab"):
            dlg.tabs.setCurrentWidget(dlg.setup_tab)
        dlg.show()
        dlg.raise_()
        dlg.activateWindow()
    except Exception as exc:
        logger.warning("Could not open MCP & AI Agent Cockpit: %s", exc)


def add_model_selector_submenu(
    agent_submenu: Any,
    agent_id: str,
    *,
    home_dir: Optional[Path] = None,
) -> Any:
    """
    Append a nested ``🧠 Model / Reasoning Effort ▸`` selector menu inside an agent's QMenu.
    Selecting an item updates ``cli_agents_config.json`` immediately via ``set_cli_agent_model_config``.
    """
    active_model, active_effort = get_cli_agent_model_config(agent_id)
    models, cli_def_model, cli_def_effort = discover_cli_models(agent_id, home_dir=home_dir)
    badge = format_active_model_badge(agent_id, home_dir=home_dir)

    model_menu = agent_submenu.addMenu(f"🧠 Model / Effort ({badge})")
    model_menu.setToolTipsVisible(True)

    def_hint = (
        f" ({cli_def_model}" + (f" · {cli_def_effort}" if cli_def_effort else "") + ")"
        if cli_def_model
        else ""
    )
    act_default = model_menu.addAction(f"🔄 Use CLI Default{def_hint}")
    act_default.setCheckable(True)
    act_default.setChecked(not active_model and not active_effort)
    act_default.setToolTip("Clear any 3GPP Tools override and inherit the CLI's own configured default.")
    act_default.triggered.connect(
        lambda _=False, aid=agent_id: set_cli_agent_model_config(aid, "", "")
    )
    model_menu.addSeparator()

    for opt in models:
        is_this_model = active_model == opt.slug
        if opt.reasoning_levels:
            label = f"{opt.display_name} ({opt.slug})" if opt.display_name != opt.slug else opt.slug
            if is_this_model and active_effort:
                label = f"✓ {label} [{active_effort}]"
            m_sub = model_menu.addMenu(label)
            if opt.description:
                m_sub.setToolTip(opt.description)

            act_m_def = m_sub.addAction(
                f"Default Effort" + (f" ({opt.default_reasoning_level})" if opt.default_reasoning_level else "")
            )
            act_m_def.setCheckable(True)
            act_m_def.setChecked(is_this_model and not active_effort)
            act_m_def.triggered.connect(
                lambda _=False, aid=agent_id, slug=opt.slug: set_cli_agent_model_config(aid, slug, "")
            )
            m_sub.addSeparator()
            for lvl in opt.reasoning_levels:
                act_lvl = m_sub.addAction(lvl)
                act_lvl.setCheckable(True)
                act_lvl.setChecked(is_this_model and active_effort == lvl)
                act_lvl.triggered.connect(
                    lambda _=False, aid=agent_id, slug=opt.slug, eff=lvl: set_cli_agent_model_config(aid, slug, eff)
                )
        else:
            label = f"{opt.display_name} ({opt.slug})" if opt.display_name != opt.slug else opt.slug
            act_m = model_menu.addAction(label)
            act_m.setCheckable(True)
            act_m.setChecked(is_this_model)
            act_m.triggered.connect(
                lambda _=False, aid=agent_id, slug=opt.slug: set_cli_agent_model_config(aid, slug, "")
            )

    return model_menu


def populate_cli_agent_menu(
    menu: Any,
    on_launch_callback: Callable[[str, str], None],
    *,
    has_multiple_or_revs: bool = False,
    multi_count: int = 0,
    parent: Any = None,
    statuses: Optional[Dict[str, CliAgentStatus]] = None,
) -> Any:
    """
    Populate a QMenu with detected CLI agents (`Codex CLI`, `Claude CLI`, `Gemini CLI`)
    and task presets (`interactive`, `summarize`, optional `compare`, `custom`, `edit_context`),
    quick Model / Reasoning Effort switchers, disabled entries for missing CLIs, and Cockpit link.
    """
    menu.clear()
    menu.setToolTipsVisible(True)

    resolved = statuses if statuses is not None else detect_cli_agents()
    installed_agents = [st for st in resolved.values() if st.installed]
    missing_agents = [st for st in resolved.values() if not st.installed]

    if installed_agents:
        for st in installed_agents:
            badge = format_active_model_badge(st.agent_id)
            sub = menu.addMenu(f"{st.icon} {st.label}  [{badge}]")
            sub.setToolTipsVisible(True)
            mcp_note = "MCP configured" if st.mcp_configured else "MCP not yet registered (local Markdown + cli_tools.py ready)"
            sub.setToolTip(f"{st.executable_path}\nModel: {badge}\n{mcp_note}")

            act_inter = sub.addAction("💬 Interactive Session (Read && Stand By)")
            act_inter.setToolTip(f"Prepare Markdown and start an interactive {st.label} session.")
            act_inter.triggered.connect(lambda _=False, aid=st.agent_id: on_launch_callback(aid, "interactive"))

            act_sum = sub.addAction("⚡ Summarize && Check Spec Impact")
            act_sum.setToolTip(f"Launch {st.label} with an immediate prompt to summarize mechanisms and spec impact.")
            act_sum.triggered.connect(lambda _=False, aid=st.agent_id: on_launch_callback(aid, "summarize"))

            if has_multiple_or_revs:
                cmp_label = (
                    f"⚖️ Compare {multi_count} Selected TDocs"
                    if multi_count > 1
                    else "⚖️ Compare Base vs. Revisions"
                )
                act_cmp = sub.addAction(cmp_label)
                act_cmp.setToolTip(f"Stage all relevant Markdown files and ask {st.label} to compare them.")
                act_cmp.triggered.connect(lambda _=False, aid=st.agent_id: on_launch_callback(aid, "compare"))

            sub.addSeparator()
            act_custom = sub.addAction("✏️ Ask Custom Question...")
            act_custom.setToolTip(f"Enter a custom question before launching {st.label}.")
            act_custom.triggered.connect(lambda _=False, aid=st.agent_id: on_launch_callback(aid, "custom"))

            act_ctx = sub.addAction("📝 Edit Context / Company Notes && Launch...")
            act_ctx.setToolTip(
                "Edit persistent _USER_NOTES.md / standing company guidance and select Model / Reasoning Effort before launching."
            )
            act_ctx.triggered.connect(lambda _=False, aid=st.agent_id: on_launch_callback(aid, "edit_context"))

            sub.addSeparator()
            add_model_selector_submenu(sub, st.agent_id)
    else:
        no_cli = menu.addAction("⚪ No LLM CLI Detected (codex / claude / gemini)")
        no_cli.setEnabled(False)
        no_cli.setToolTip("Install Codex CLI, Claude CLI, or Gemini CLI and click Re-scan in MCP Agent Setup.")

    if missing_agents:
        menu.addSeparator()
        for st in missing_agents:
            act_miss = menu.addAction(f"⚪ {st.label} (Not Found)")
            act_miss.setEnabled(False)
            act_miss.setToolTip(f"Not detected in PATH or user directories.\nInstall: {st.install_hint}")

    menu.addSeparator()
    act_cfg = menu.addAction("⚙️ Configure CLI Agents && MCP...")
    act_cfg.setToolTip("Open the MCP Server & External AI Agent Cockpit to inspect CLI detection, default models, and MCP setup.")
    act_cfg.triggered.connect(lambda _=False: open_mcp_agent_setup_dialog(parent))

    return menu
