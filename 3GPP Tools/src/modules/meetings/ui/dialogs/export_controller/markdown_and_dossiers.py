"""Markdown, Pre-/Post-Meeting briefing, and HTML/PDF dossier export controllers."""

from __future__ import annotations

import sys
from pathlib import Path
from typing import Any, List, Optional

from PyQt5.QtGui import QCursor
from PyQt5.QtWidgets import QDialog, QMenu, QMessageBox

from core.ai.cli_detector import get_installed_cli_agents
from core.ai.cli_launcher import launch_cli_in_terminal, write_cli_context_file
from core.utils.folder_opener import open_folder as _open_folder
from modules.meetings.core.exporters.dossier_exporter import DossierExporterThread
from modules.meetings.core.exporters.llm_exporter import TDocMarkdownExporterThread
from modules.meetings.ui.dialogs.dossier_export_dialog import DossierExportDialog
from modules.meetings.ui.dialogs.export_controller.package_and_excel import get_package_export_rows

_SHIM_MODULE_NAME = "modules.meetings.ui.dialogs.tdocs_export_controller"


def _resolve_facade(attr_name: str, default: Any) -> Any:
    """Resolve patchable collaborators through the legacy facade module when loaded."""
    facade = sys.modules.get(_SHIM_MODULE_NAME)
    if facade is not None and hasattr(facade, attr_name):
        return getattr(facade, attr_name)
    return default


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

    thread_cls = _resolve_facade("TDocMarkdownExporterThread", TDocMarkdownExporterThread)
    window.md_export_thread = thread_cls(
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

    thread_cls = _resolve_facade("TDocMarkdownExporterThread", TDocMarkdownExporterThread)
    window.md_export_thread = thread_cls(
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
        launcher = _resolve_facade("launch_cli_in_terminal", launch_cli_in_terminal)

        def _launch(agent_id: str) -> None:
            write_cli_context_file(
                target_dir,
                task_mode="briefing_qa",
                briefing_title=briefing_title,
            )
            # Launch cleanly at interactive prompt without auto-scanning the full meeting dossier
            launcher(agent_id, target_dir, bootstrap_prompt="", parent=parent)

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
