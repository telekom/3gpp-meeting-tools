# --- File: src/modules/meetings/ui/tdocs_export_controller.py ---
import logging
import re
from pathlib import Path
from typing import Any, List, Optional

from PyQt5.QtWidgets import QDialog, QFileDialog, QMessageBox

from modules.meetings.core.excel_exporter import ExcelExporterThread
from modules.meetings.core.tdocs_package_exporter import TDocsPackageExporterThread
from modules.meetings.core.llm_exporter import TDocMarkdownExporterThread
from modules.meetings.core.dossier_exporter import DossierExporterThread
from modules.meetings.ui.excel_export_dialog import ALL_AVAILABLE_COLUMNS, ExcelExportDialog
from modules.meetings.ui.dossier_export_dialog import DossierExportDialog
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
    from modules.meetings.ui.post_meeting_export_dialog import PostMeetingExportDialog
    from modules.meetings.core.post_meeting_exporter import PostMeetingExporterThread

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
        include_general_ls=cfg["include_general_ls"],
        generate_local_indices=cfg["generate_local_indices"],
        download_missing=cfg["download_missing"],
        force_reconvert=cfg["force_reconvert"],
        compile_dossier=cfg.get("compile_dossier", True),
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


def on_post_meeting_export_finished(window: Any, success: bool, msg: str, auto_open: bool = True) -> None:
    window.export_btn.setText("📤 Export ▾")
    window.export_btn.setEnabled(True)
    if success:
        QMessageBox.information(window, "Post-Meeting Export Complete", msg)
        target = getattr(window.md_export_thread, "export_dir", None)
        if auto_open and target and target.is_dir():
            _open_folder(target)
    else:
        QMessageBox.warning(window, "Post-Meeting Export Failed", msg)


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
        compile_dossier=cfg.get("compile_dossier", True),
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

