# --- File: src/modules/meetings/ui/tdocs_export_controller.py ---
import datetime
import logging
import re
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional

from PyQt5.QtGui import QCursor
from PyQt5.QtWidgets import QDialog, QFileDialog, QInputDialog, QMenu, QMessageBox

from core.ai.cli_detector import CLI_AGENT_SPECS, get_installed_cli_agents
from core.ai.cli_launcher import (
    collect_per_tdoc_user_notes,
    launch_cli_in_terminal,
    save_workspace_user_notes,
    write_cli_context_file,
)
from core.ui.cli_context_dialog import CliContextEditorDialog, DEFAULT_TDOC_TASK_OPTIONS
from modules.meetings.core.exporters.excel_exporter import ExcelExporterThread
from modules.meetings.core.exporters.tdocs_package_exporter import TDocsPackageExporterThread
from modules.meetings.core.exporters.llm_exporter import TDocMarkdownExporterThread
from modules.meetings.core.exporters.dossier_exporter import DossierExporterThread
from modules.meetings.core.tdocs.tdoc_markdown_converter import (
    is_canonical_tdoc_markdown_ready,
    resolve_canonical_tdoc_paths,
)
from modules.meetings.ui.dialogs.excel_export_dialog import ALL_AVAILABLE_COLUMNS, ExcelExportDialog
from modules.meetings.ui.dialogs.dossier_export_dialog import DossierExportDialog
from core.utils.folder_opener import open_folder as _open_folder

logger = logging.getLogger(__name__)


def get_package_export_rows(window: Any) -> List[dict]:
    """Return selected TDoc rows, or all currently visible rows when nothing is selected."""
    selection_model = window.table.selectionModel()
    selected_indexes = selection_model.selectedIndexes() if selection_model else []

    if selected_indexes:
        proxy_rows = sorted({idx.row() for idx in selected_indexes if idx.isValid()})
    else:
        proxy_rows = list(range(window.proxy.rowCount()))

    rows = []
    seen_tdocs = set()
    for proxy_row in proxy_rows:
        source_idx = window.proxy.mapToSource(window.proxy.index(proxy_row, 0))
        if not source_idx.isValid():
            continue
        row_data = window.model._data[source_idx.row()]
        tdoc_id = str(row_data.get("TDoc", "")).strip().upper()
        if not tdoc_id or tdoc_id in seen_tdocs:
            continue
        seen_tdocs.add(tdoc_id)
        rows.append(row_data)
    return rows


def export_tdocs_package(window: Any) -> None:
    """Packages all selected or visible TDocs and writes an Excel manifest."""
    rows_to_export = get_package_export_rows(window)
    if not rows_to_export:
        QMessageBox.warning(window, "No TDocs", "There are no selected or visible TDocs to export.")
        return

    selected_count = (
        len(window.table.selectionModel().selectedIndexes())
        if window.table.selectionModel()
        else 0
    )
    scope_text = "selected" if selected_count else "visible"

    window.export_btn.setText(f"⏳ Exporting {len(rows_to_export)} TDocs...")
    window.export_btn.setEnabled(False)

    window.package_export_thread = TDocsPackageExporterThread(
        meeting_dir=window.meeting_dir,
        rows_data=rows_to_export,
        mtg_info=window.mtg_info,
        main_ftp_url=window.main_ftp_url,
        scope_label=scope_text,
        parent=window,
    )
    window.package_export_thread.progress.connect(window._on_package_export_progress)
    window.package_export_thread.finished.connect(
        lambda success, export_dir, failures, rows=rows_to_export: on_package_files_finished(
            window, success, export_dir, failures, rows
        )
    )
    window.package_export_thread.start()


def on_package_files_finished(
    window: Any, success: bool, export_dir: str, failures: object, rows_to_export: list
) -> None:
    if not success:
        window.export_btn.setText("📤 Export ▾")
        window.export_btn.setEnabled(True)
        QMessageBox.warning(window, "Package Export Failed", str(export_dir))
        return

    export_path = Path(export_dir)
    wg = str(window.mtg_info.get("wg_name", "3GPP")).strip().replace(" ", "_")
    mtg = str(window.mtg_info.get("meeting_number", "")).strip()
    excel_path = export_path / f"TDocs_{wg}_{mtg}.xlsx"

    window.package_excel_exporter_thread = ExcelExporterThread(
        output_path=excel_path,
        rows_data=rows_to_export,
        selected_columns=list(ALL_AVAILABLE_COLUMNS),
        mtg_info=window.mtg_info,
        docs_ftp_url=window.docs_ftp_url,
        auto_open=False,
    )
    window.package_excel_exporter_thread.progress.connect(window._on_package_export_progress)
    window.package_excel_exporter_thread.finished.connect(
        lambda excel_success, msg, d=export_path, f=list(failures or []), n=len(rows_to_export): on_package_export_finished(
            window, excel_success, msg, d, f, n
        )
    )
    window.package_excel_exporter_thread.start()


def on_package_export_finished(
    window: Any,
    success: bool,
    msg: str,
    export_dir: Path,
    failures: list,
    requested_count: int,
) -> None:
    window.export_btn.setText("📤 Export ▾")
    window.export_btn.setEnabled(True)

    if not success:
        QMessageBox.warning(
            window,
            "Package Export Incomplete",
            f"The TDoc files were packaged, but the Excel manifest could not be created:\n{msg}\n\nPackage folder:\n{export_dir}",
        )
        _open_folder(export_dir)
        return

    exported_count = requested_count - len(failures)
    if failures:
        failed_ids = "\n".join(f"• {item}" for item in failures)
        QMessageBox.warning(
            window,
            "Package Export Completed with Warnings",
            f"Exported {exported_count}/{requested_count} TDocs.\n"
            f"The Excel manifest contains all {requested_count} requested rows.\n\n"
            f"Could not retrieve {len(failures)} TDoc(s):\n{failed_ids}\n\n"
            f"Saved to:\n{export_dir}",
        )
    else:
        QMessageBox.information(
            window,
            "Package Export Complete",
            f"Successfully exported {exported_count} TDocs and the Excel manifest to:\n{export_dir}",
        )

    _open_folder(export_dir)


def open_excel_export_dialog(window: Any) -> None:
    """Launches the column & scope configuration dialog and executes Excel export."""
    visible_count = window.proxy.rowCount()
    total_count = window.model.rowCount()

    dialog = ExcelExportDialog(window, visible_count, total_count)
    if dialog.exec_() != QDialog.Accepted:
        return

    selected_columns = dialog.get_selected_columns()
    is_visible_only = dialog.is_visible_only()
    auto_open = dialog.should_auto_open()
    column_headers = dialog.get_column_headers()
    company_name = dialog.get_company_name()

    if is_visible_only:
        rows_to_export = []
        for r in range(window.proxy.rowCount()):
            source_idx = window.proxy.mapToSource(window.proxy.index(r, 0))
            rows_to_export.append(window.model._data[source_idx.row()])
    else:
        rows_to_export = list(window.model._data)

    if not rows_to_export:
        QMessageBox.warning(window, "No Data", "There are no rows available to export.")
        return

    wg = str(window.mtg_info.get("wg_name", "3GPP")).strip().replace(" ", "_")
    mtg = str(window.mtg_info.get("meeting_number", "")).strip()
    scope_str = "Filtered" if is_visible_only else "All"

    is_review_sheet = "My Notes" in selected_columns or "My Status" in selected_columns
    if is_review_sheet and dialog.is_company_attribution_enabled() and company_name:
        comp_clean = re.sub(r'[^\w\-]', '_', company_name).strip('_')
        default_filename = f"Comments_{comp_clean}_{wg}_{mtg}_{scope_str}.xlsx"
    else:
        default_filename = f"TDocs_{wg}_{mtg}_{scope_str}.xlsx"

    default_target = window.meeting_dir / "Export" / default_filename

    save_path, _ = QFileDialog.getSaveFileName(
        window,
        "Save TDocs Excel Export",
        str(default_target),
        "Excel Workbook (*.xlsx)",
    )
    if not save_path:
        return

    window.export_btn.setText("⏳ Exporting Excel...")
    window.export_btn.setEnabled(False)

    window.excel_exporter_thread = ExcelExporterThread(
        Path(save_path),
        rows_to_export,
        selected_columns,
        window.mtg_info,
        docs_ftp_url=window.docs_ftp_url,
        auto_open=auto_open,
        column_headers=column_headers,
        company_name=company_name,
    )
    window.excel_exporter_thread.progress.connect(window._on_excel_export_progress)
    window.excel_exporter_thread.finished.connect(
        lambda success, msg, auto=auto_open: on_excel_export_finished(window, success, msg, auto)
    )
    window.excel_exporter_thread.start()


def on_excel_export_finished(window: Any, success: bool, msg: str, auto_open: bool) -> None:
    window.export_btn.setText("📤 Export ▾")
    window.export_btn.setEnabled(True)

    if success:
        QMessageBox.information(
            window,
            "Export Complete",
            f"Successfully exported TDocs to:\n{msg}",
        )
        if auto_open:
            _open_folder(Path(msg))
    else:
        QMessageBox.warning(window, "Export Failed", f"Failed to export Excel:\n{msg}")


def export_markdown_visible(window: Any) -> None:
    """Exports all currently visible TDocs in the table to Markdown dossier."""
    visible_tdocs = []
    for r in range(window.proxy.rowCount()):
        index = window.proxy.index(r, 0)
        source_index = window.proxy.mapToSource(index)
        row_data = window.model._data[source_index.row()]
        visible_tdocs.append(row_data)

    if not visible_tdocs:
        QMessageBox.warning(
            window,
            "Empty View",
            "No TDocs currently visible in the table to export.",
        )
        return

    window.export_btn.setText("⏳ Exporting Markdown...")
    window.export_btn.setEnabled(False)

    window.md_export_thread = TDocMarkdownExporterThread(
        window.meeting_dir,
        visible_tdocs,
        window.docs_ftp_url,
        window.revisions_url,
        is_bulk=True,
    )
    window.llm_thread = window.md_export_thread
    window.md_export_thread.progress.connect(window._on_md_progress)
    window.md_export_thread.finished.connect(window._on_md_export_finished)
    window.md_export_thread.start()


def export_markdown_single(window: Any, tdoc_id: str) -> None:
    """Exports a single TDoc by ID to Markdown dossier."""
    row_data = next((r for r in window.model._data if r.get("TDoc") == tdoc_id), None)
    if not row_data:
        return

    window.export_btn.setText("⏳ Exporting Markdown...")
    window.export_btn.setEnabled(False)

    window.md_export_thread = TDocMarkdownExporterThread(
        window.meeting_dir,
        [row_data],
        window.docs_ftp_url,
        window.revisions_url,
        is_bulk=False,
    )
    window.llm_thread = window.md_export_thread
    window.md_export_thread.progress.connect(window._on_md_progress)
    window.md_export_thread.finished.connect(window._on_md_single_finished)
    window.md_export_thread.start()


def open_post_meeting_export_dialog(window: Any) -> None:
    """Opens PostMeetingExportDialog and triggers PostMeetingExporterThread."""
    from modules.meetings.ui.dialogs.post_meeting_export_dialog import PostMeetingExportDialog
    from modules.meetings.core.exporters.post_meeting_exporter import PostMeetingExporterThread

    visible_tdocs = []
    for r in range(window.proxy.rowCount()):
        index = window.proxy.index(r, 0)
        source_index = window.proxy.mapToSource(index)
        if source_index.isValid() and source_index.row() < len(window.model._data):
            visible_tdocs.append(window.model._data[source_index.row()])

    followed_tokens = getattr(window.proxy, "followed_wi_tokens", set())
    starred_tokens = getattr(window.proxy, "starred_wi_tokens", set())

    dialog = PostMeetingExportDialog(
        parent=window,
        meeting_dir=window.meeting_dir,
        all_rows=window.model._data,
        visible_rows=visible_tdocs,
        agenda_map=window.agenda_map,
        mtg_info=window.mtg_info,
        followed_tokens=followed_tokens,
        starred_tokens=starred_tokens,
    )

    if dialog.exec_() != QDialog.Accepted:
        return

    cfg = dialog.get_export_config()

    window.export_btn.setText("⏳ Preparing Post-Meeting Export...")
    window.export_btn.setEnabled(False)

    thread = PostMeetingExporterThread(
        meeting_dir=window.meeting_dir,
        all_rows=window.model._data,
        visible_rows=visible_tdocs,
        agenda_map=window.agenda_map,
        mtg_info=window.mtg_info,
        docs_ftp_url=window.docs_ftp_url,
        revisions_url=window.revisions_url,
        scope_mode=cfg["scope_mode"],
        custom_wis=cfg["custom_wis"],
        custom_ais=cfg.get("custom_ais"),
        custom_wi_ais=cfg.get("custom_wi_ais"),
        followed_tokens=followed_tokens,
        starred_tokens=starred_tokens,
        include_ls_in=cfg["include_ls_in"],
        include_pcr=cfg["include_pcr"],
        include_cr=cfg["include_cr"],
        include_ls_out=cfg["include_ls_out"],
        include_disc=cfg.get("include_disc", False),
        include_chair_agreements=cfg.get("include_chair_agreements", True),
        include_general_ls=cfg["include_general_ls"],
        generate_local_indices=cfg["generate_local_indices"],
        download_missing=cfg["download_missing"],
        force_reconvert=cfg["force_reconvert"],
        compile_dossier=cfg.get("compile_dossier", False),
        output_dir=cfg["output_dir"],
        parent=window,
    )
    window.md_export_thread = thread
    thread.progress.connect(window._on_md_progress)
    thread.finished.connect(
        lambda success, msg, auto_open=cfg["auto_open"]: on_post_meeting_export_finished(
            window, success, msg, auto_open
        )
    )
    thread.start()


def _show_export_complete_with_cli_prompt(
    parent: Any,
    title: str,
    msg: str,
    target_dir: Optional[Path],
    briefing_title: str,
    auto_open: bool = True,
) -> None:
    """Show export completion message and optionally launch an installed LLM CLI in `target_dir`."""
    installed = get_installed_cli_agents() if (target_dir and target_dir.is_dir()) else []
    if not installed:
        QMessageBox.information(parent, title, msg)
        if auto_open and target_dir and target_dir.is_dir():
            _open_folder(target_dir)
        return

    box = QMessageBox(parent)
    box.setIcon(QMessageBox.Information)
    box.setWindowTitle(title)
    box.setText(msg)
    btn_ok = box.addButton(QMessageBox.Ok)
    btn_cli = box.addButton("💬 Launch CLI Agent...", QMessageBox.ActionRole)
    box.setDefaultButton(btn_ok)
    box.exec_()

    if box.clickedButton() == btn_cli and target_dir and target_dir.is_dir():
        def _launch(agent_id: str) -> None:
            write_cli_context_file(
                target_dir,
                task_mode="briefing_qa",
                briefing_title=briefing_title,
            )
            # Launch cleanly at interactive prompt without auto-scanning the full meeting dossier
            launch_cli_in_terminal(agent_id, target_dir, bootstrap_prompt="", parent=parent)

        if len(installed) == 1:
            _launch(installed[0].agent_id)
        else:
            menu = QMenu(parent)
            for st in installed:
                act = menu.addAction(f"{st.icon} {st.label}")
                act.triggered.connect(lambda _=False, aid=st.agent_id: _launch(aid))
            menu.exec_(QCursor.pos())
    elif auto_open and target_dir and target_dir.is_dir():
        _open_folder(target_dir)


def on_post_meeting_export_finished(window: Any, success: bool, msg: str, auto_open: bool = True) -> None:
    window.export_btn.setText("📤 Export ▾")
    window.export_btn.setEnabled(True)
    if success:
        target = getattr(window.md_export_thread, "export_dir", None)
        wg = str(window.mtg_info.get("wg_name", "")).strip()
        mtg = str(window.mtg_info.get("meeting_number", "")).strip()
        _show_export_complete_with_cli_prompt(
            window,
            "Post-Meeting Export Complete",
            msg,
            target,
            briefing_title=f"Post-Meeting Dossier ({wg} {mtg})".strip(),
            auto_open=auto_open,
        )
    else:
        QMessageBox.warning(window, "Post-Meeting Export Failed", msg)


def open_pre_meeting_export_dialog(window: Any) -> None:
    """Opens PreMeetingExportDialog and triggers PreMeetingExporterThread."""
    from modules.meetings.ui.dialogs.pre_meeting_export_dialog import PreMeetingExportDialog
    from modules.meetings.core.exporters.pre_meeting_exporter import PreMeetingExporterThread

    visible_tdocs = []
    for r in range(window.proxy.rowCount()):
        index = window.proxy.index(r, 0)
        source_index = window.proxy.mapToSource(index)
        if source_index.isValid() and source_index.row() < len(window.model._data):
            visible_tdocs.append(window.model._data[source_index.row()])

    followed_tokens = getattr(window.proxy, "followed_wi_tokens", set())
    starred_tokens = getattr(window.proxy, "starred_wi_tokens", set())

    dialog = PreMeetingExportDialog(
        parent=window,
        meeting_dir=window.meeting_dir,
        all_rows=window.model._data,
        visible_rows=visible_tdocs,
        agenda_map=window.agenda_map,
        mtg_info=window.mtg_info,
        followed_tokens=followed_tokens,
        starred_tokens=starred_tokens,
    )

    if dialog.exec_() != QDialog.Accepted:
        return

    cfg = dialog.get_export_config()

    window.export_btn.setText("⏳ Preparing Pre-Meeting Export...")
    window.export_btn.setEnabled(False)

    thread = PreMeetingExporterThread(
        meeting_dir=window.meeting_dir,
        all_rows=window.model._data,
        visible_rows=visible_tdocs,
        agenda_map=window.agenda_map,
        mtg_info=window.mtg_info,
        docs_ftp_url=window.docs_ftp_url,
        revisions_url=window.revisions_url,
        scope_mode=cfg["scope_mode"],
        custom_wis=cfg["custom_wis"],
        custom_ais=cfg.get("custom_ais"),
        custom_wi_ais=cfg.get("custom_wi_ais"),
        followed_tokens=followed_tokens,
        starred_tokens=starred_tokens,
        include_ls_in=cfg["include_ls_in"],
        include_pcr=cfg["include_pcr"],
        include_cr=cfg["include_cr"],
        include_ls_out=cfg["include_ls_out"],
        include_disc=cfg.get("include_disc", True),
        include_general_ls=cfg["include_general_ls"],
        deduplicate_revisions=cfg.get("deduplicate_revisions", True),
        include_reserved_in_index=cfg.get("include_reserved_in_index", True),
        company_filter=cfg.get("company_filter"),
        include_user_notes=cfg.get("include_user_notes", True),
        generate_local_indices=cfg["generate_local_indices"],
        download_missing=cfg["download_missing"],
        force_reconvert=cfg["force_reconvert"],
        compile_dossier=cfg.get("compile_dossier", False),
        output_dir=cfg["output_dir"],
        parent=window,
    )
    window.md_export_thread = thread
    thread.progress.connect(window._on_md_progress)
    thread.finished.connect(
        lambda success, msg, auto_open=cfg["auto_open"]: on_pre_meeting_export_finished(
            window, success, msg, auto_open
        )
    )
    thread.start()


def on_pre_meeting_export_finished(window: Any, success: bool, msg: str, auto_open: bool = True) -> None:
    window.export_btn.setText("📤 Export ▾")
    window.export_btn.setEnabled(True)
    if success:
        target = getattr(window.md_export_thread, "export_dir", None)
        wg = str(window.mtg_info.get("wg_name", "")).strip()
        mtg = str(window.mtg_info.get("meeting_number", "")).strip()
        _show_export_complete_with_cli_prompt(
            window,
            "Pre-Meeting Export Complete",
            msg,
            target,
            briefing_title=f"Pre-Meeting Briefing ({wg} {mtg})".strip(),
            auto_open=auto_open,
        )
    else:
        QMessageBox.warning(window, "Pre-Meeting Export Failed", msg)


def open_dossier_export_dialog(window: Any, target_rows: Optional[List[dict]] = None) -> None:
    """
    Opens DossierExportDialog and triggers DossierExporterThread to generate
    an interactive HTML dossier and bookmarked vector PDF (F-04).
    """
    if target_rows is not None:
        rows_to_export = target_rows
        scope_text = "selected"
    else:
        rows_to_export = get_package_export_rows(window)
        selected_count = (
            len(window.table.selectionModel().selectedIndexes())
            if window.table.selectionModel()
            else 0
        )
        scope_text = "selected" if selected_count else "visible"

    if not rows_to_export:
        QMessageBox.warning(window, "No TDocs", "There are no selected or visible TDocs to export to dossier.")
        return

    dialog = DossierExportDialog(
        parent=window,
        meeting_dir=window.meeting_dir,
        tdocs_count=len(rows_to_export),
        mtg_info=window.mtg_info,
        scope_label=scope_text,
    )

    if dialog.exec_() != QDialog.Accepted:
        return

    cfg = dialog.get_config()

    window.export_btn.setText("⏳ Compiling Dossier...")
    window.export_btn.setEnabled(False)

    thread = DossierExporterThread(
        meeting_dir=window.meeting_dir,
        tdocs_list=rows_to_export,
        docs_ftp_url=window.docs_ftp_url,
        revisions_url=window.revisions_url,
        mtg_info=window.mtg_info,
        export_format=cfg["export_format"],
        include_prompt=cfg["include_prompt"],
        embed_base64_images=cfg["embed_base64_images"],
        download_missing=cfg["download_missing"],
        force_reconvert=cfg["force_reconvert"],
        compile_dossier=cfg.get("compile_dossier", False),
        output_dir=cfg["output_dir"],
        parent=window,
    )
    window.dossier_thread = thread
    thread.progress.connect(window._on_md_progress)
    thread.finished.connect(
        lambda success, msg, files, auto_open=cfg["auto_open"]: on_dossier_export_finished(
            window, success, msg, files, auto_open
        )
    )
    thread.start()


def on_dossier_export_finished(
    window: Any, success: bool, msg: str, files: dict, auto_open: bool = True
) -> None:
    window.export_btn.setText("📤 Export ▾")
    window.export_btn.setEnabled(True)
    if success:
        target_dir = getattr(window.dossier_thread, "export_dir", None)
        QMessageBox.information(window, "Navigable Dossier Complete", msg)
        if auto_open and target_dir and target_dir.is_dir():
            _open_folder(target_dir)
    else:
        QMessageBox.warning(window, "Dossier Export Failed", msg)


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

    launcher = terminal_launcher or launch_cli_in_terminal
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

    thread = TDocMarkdownExporterThread(
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


