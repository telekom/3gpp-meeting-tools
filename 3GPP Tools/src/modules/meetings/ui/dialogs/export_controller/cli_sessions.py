"""Interactive LLM CLI session staging and terminal launch controllers."""

from __future__ import annotations

import datetime
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional

from PyQt5.QtWidgets import QDialog, QInputDialog, QMessageBox

from core.ai.cli_detector import CLI_AGENT_SPECS
from core.ai.cli_launcher import (
    collect_per_tdoc_user_notes,
    launch_cli_in_terminal,
    save_workspace_user_notes,
    write_cli_context_file,
)
from core.ui.cli_context_dialog import CliContextEditorDialog, DEFAULT_TDOC_TASK_OPTIONS
from modules.meetings.core.exporters.llm_exporter import TDocMarkdownExporterThread
from modules.meetings.core.tdocs.tdoc_markdown_converter import (
    is_canonical_tdoc_markdown_ready,
    resolve_canonical_tdoc_paths,
)
from modules.meetings.ui.dialogs.export_controller.markdown_and_dossiers import _resolve_facade


def _expand_revision_rows(window: Any, rows_data: List[dict]) -> List[dict]:
    """Expand a single-TDoc row to include its known revisions for side-by-side comparison."""
    if len(rows_data) != 1:
        return list(rows_data)

    base_row = rows_data[0]
    tdoc_id = str(base_row.get("TDoc", "")).strip()
    if not tdoc_id:
        return list(rows_data)

    model = getattr(window, "model", None)
    revs_map = getattr(model, "revisions", {}) if model else {}
    revisions = list(revs_map.get(tdoc_id, []) or [])
    if not revisions:
        return list(rows_data)

    expanded = [base_row]
    seen = {tdoc_id.upper()}
    for rev in revisions:
        rev_str = str(rev).strip()
        if not rev_str or rev_str.upper() in seen:
            continue
        seen.add(rev_str.upper())
        rev_row = dict(base_row)
        rev_row["TDoc"] = rev_str
        expanded.append(rev_row)
    return expanded


def prepare_and_launch_cli_session(
    *,
    parent: Any,
    meeting_dir: Path,
    rows_data: List[dict],
    agent_id: str,
    task_mode: str = "interactive",
    custom_question: str = "",
    docs_ftp_url: str = "",
    revisions_url: str = "",
    local_triage_summary: str = "",
    briefing_title: str = "",
    status_button: Optional[Any] = None,
    idle_button_text: str = "📤 Export ▾",
    busy_button_text: str = "⏳ Preparing CLI Markdown...",
    progress_slot: Optional[Callable[[str], None]] = None,
    thread_owner: Optional[Any] = None,
    thread_attr: str = "cli_md_export_thread",
    terminal_launcher: Optional[Callable[..., bool]] = None,
    expand_compare_fn: Optional[Callable[[List[dict]], List[dict]]] = None,
) -> None:
    """
    Prepare Markdown + `_CLI_CONTEXT.md` for one or more TDocs and launch an interactive LLM CLI.

    - **Single TDoc:** Works directly inside the canonical ``{meeting_dir}/{base_tdoc}/``
      folder (launching immediately if ``{tdoc_id}.md`` is already ready, or converting in-place
      with ``stage_copies=False`` without creating an ``Export/`` subfolder).
    - **Multiple TDocs / Revision Comparison:** Stages documents side-by-side under
      ``{meeting_dir}/Export/CLI_Batch_{timestamp}/`` with an ``index.md`` manifest,
      automatically carrying over saved per-TDoc ``_USER_NOTES.md`` notes.
    """
    rows = [dict(r) for r in (rows_data or []) if r and str(r.get("TDoc", "")).strip()]
    if not rows:
        QMessageBox.warning(parent, "No TDocs", "There are no TDocs selected to send to the CLI agent.")
        return

    batch_notes_override: Optional[str] = None
    model_override: Optional[str] = None
    effort_override: Optional[str] = None

    if task_mode == "edit_context":
        if len(rows) == 1:
            single_tid = str(rows[0].get("TDoc", "")).strip()
            canonical_dir, _ = resolve_canonical_tdoc_paths(meeting_dir, single_tid)
            dlg_target_dir: Optional[Path] = canonical_dir
            scope_lbl = single_tid
            per_tdoc_notes: Dict[str, str] = {}
        else:
            dlg_target_dir = None
            scope_lbl = f"{len(rows)} Selected TDocs"
            per_tdoc_notes = collect_per_tdoc_user_notes(meeting_dir, rows)

        task_opts = list(DEFAULT_TDOC_TASK_OPTIONS)
        if len(rows) == 1 and expand_compare_fn is None:
            task_opts = [opt for opt in task_opts if opt[0] != "compare"]

        dlg = CliContextEditorDialog(
            agent_id=agent_id,
            scope_label=scope_lbl,
            target_dir=dlg_target_dir,
            task_options=task_opts,
            initial_task_mode="interactive",
            initial_custom_question=custom_question,
            per_tdoc_notes=per_tdoc_notes,
            parent=parent,
        )
        if dlg.exec_() != QDialog.Accepted:
            return
        res = dlg.get_result()
        if not res.launch_requested:
            return
        task_mode = res.task_mode
        custom_question = res.custom_question
        model_override = res.model
        effort_override = res.reasoning_effort
        if dlg_target_dir is None:
            batch_notes_override = res.workspace_notes

    if task_mode == "compare" and expand_compare_fn is not None:
        rows = expand_compare_fn(rows)

    spec = CLI_AGENT_SPECS.get((agent_id or "").strip().lower())
    agent_label = spec.label if spec else agent_id

    if task_mode == "custom" and not (custom_question or "").strip():
        tdoc_preview = ", ".join(str(r.get("TDoc", "")).strip() for r in rows[:5])
        if len(rows) > 5:
            tdoc_preview += f" (+{len(rows) - 5} more)"
        question, ok = QInputDialog.getMultiLineText(
            parent,
            f"Ask {agent_label}",
            f"Enter your question or analysis instruction for {tdoc_preview}:",
            "",
        )
        if not ok or not question.strip():
            return
        custom_question = question.strip()

    default_launcher = _resolve_facade("launch_cli_in_terminal", launch_cli_in_terminal)
    launcher = terminal_launcher or default_launcher
    is_single = len(rows) == 1

    if is_single:
        row = rows[0]
        has_urls = bool((docs_ftp_url or "").strip() or (revisions_url or "").strip() or row.get("3gpp_url"))
        target_dir, local_md, md_ready = is_canonical_tdoc_markdown_ready(
            meeting_dir,
            row,
            has_download_urls=has_urls,
        )
        if md_ready:
            write_cli_context_file(
                target_dir,
                task_mode=task_mode,
                custom_question=custom_question,
                tdocs_meta=rows,
                markdown_files=[local_md],
                local_triage_summary=local_triage_summary,
                briefing_title=briefing_title,
                meeting_dir=meeting_dir,
            )
            launcher(
                agent_id,
                target_dir,
                parent=parent,
                model=model_override,
                reasoning_effort=effort_override,
            )
            return
        stage_copies = False
        is_bulk = False
    else:
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H%M%S")
        target_dir = Path(meeting_dir) / "Export" / f"CLI_Batch_{timestamp}"
        local_md = None
        stage_copies = True
        is_bulk = True

    if status_button is not None:
        status_button.setText(busy_button_text)
        status_button.setEnabled(False)

    thread_cls = _resolve_facade("TDocMarkdownExporterThread", TDocMarkdownExporterThread)
    thread = thread_cls(
        meeting_dir,
        rows,
        docs_ftp_url,
        revisions_url,
        is_bulk=is_bulk,
        output_dir=target_dir,
        stage_copies=stage_copies,
    )
    if thread_owner is not None:
        setattr(thread_owner, thread_attr, thread)
    if progress_slot is not None:
        thread.progress.connect(progress_slot)

    def _on_finished(success: bool, msg: str) -> None:
        if status_button is not None:
            status_button.setText(idle_button_text)
            status_button.setEnabled(True)
        if not success or (local_md is not None and not local_md.is_file()):
            QMessageBox.warning(parent, "CLI Pre-Conversion Failed", msg)
            return

        if batch_notes_override:
            save_workspace_user_notes(target_dir, batch_notes_override)

        if local_md is not None:
            md_files = [local_md]
        else:
            md_files = [
                p for p in sorted(target_dir.glob("*/*.md"))
                if p.name not in {"_CLI_CONTEXT.md", "_USER_NOTES.md"}
            ]
            if (target_dir / "index.md").is_file():
                md_files.insert(0, target_dir / "index.md")

        write_cli_context_file(
            target_dir,
            task_mode=task_mode,
            custom_question=custom_question,
            tdocs_meta=rows,
            markdown_files=md_files,
            local_triage_summary=local_triage_summary,
            briefing_title=briefing_title,
            meeting_dir=meeting_dir,
        )
        launcher(
            agent_id,
            target_dir,
            parent=parent,
            model=model_override,
            reasoning_effort=effort_override,
        )

    thread.finished.connect(_on_finished)
    thread.start()


def launch_cli_for_tdocs(
    window: Any,
    rows_data: List[dict],
    agent_id: str,
    task_mode: str = "interactive",
    custom_question: str = "",
) -> None:
    """
    Prepare Markdown + `_CLI_CONTEXT.md` for TDoc row(s) from `TDocsWindow` and launch
    the requested interactive LLM CLI terminal.
    """
    rows = [dict(r) for r in (rows_data or []) if r and str(r.get("TDoc", "")).strip()]
    if not rows:
        QMessageBox.warning(window, "No TDocs", "There are no TDocs selected to send to the CLI agent.")
        return

    has_export_btn = hasattr(window, "export_btn") and window.export_btn is not None
    wg = str(window.mtg_info.get("wg_name", "")).strip()
    mtg = str(window.mtg_info.get("meeting_number", "")).strip()

    prepare_and_launch_cli_session(
        parent=window,
        meeting_dir=Path(window.meeting_dir),
        rows_data=rows,
        agent_id=agent_id,
        task_mode=task_mode,
        custom_question=custom_question,
        docs_ftp_url=getattr(window, "docs_ftp_url", ""),
        revisions_url=getattr(window, "revisions_url", ""),
        briefing_title=f"{wg} {mtg} TDoc CLI Session".strip(),
        status_button=window.export_btn if has_export_btn else None,
        idle_button_text="📤 Export ▾",
        busy_button_text="⏳ Preparing CLI Markdown...",
        progress_slot=window._on_md_progress if has_export_btn else None,
        thread_owner=window,
        thread_attr="cli_md_export_thread",
        expand_compare_fn=lambda r: _expand_revision_rows(window, r),
    )
