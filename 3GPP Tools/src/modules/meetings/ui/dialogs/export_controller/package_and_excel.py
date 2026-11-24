"""Package bundle and Excel table export controllers for TDocsWindow."""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any, List

from PyQt5.QtWidgets import QDialog, QFileDialog, QMessageBox

from core.utils.folder_opener import open_folder as _open_folder
from modules.meetings.core.exporters.excel_exporter import ExcelExporterThread
from modules.meetings.core.exporters.tdocs_package_exporter import TDocsPackageExporterThread
from modules.meetings.ui.dialogs.excel_export_dialog import ALL_AVAILABLE_COLUMNS, ExcelExportDialog


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
