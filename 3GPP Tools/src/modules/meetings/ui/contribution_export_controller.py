# --- File: src/modules/meetings/ui/contribution_export_controller.py ---
"""
Export and statistics dashboard controller for ContributionReportDialog.
Handles asynchronous Excel export generation and HTML analytics dashboard creation.
"""
import logging
import os
import webbrowser
from pathlib import Path

from PyQt5.QtCore import QDate
from PyQt5.QtWidgets import QFileDialog, QMessageBox

from modules.meetings.core.excel_exporter import ExcelExporterThread
from modules.meetings.core.stats.contribution_stats import ContributionStatsExporterThread
from modules.meetings.ui.contribution_models import ContributionResultsTableModel
from modules.meetings.ui.tdocs_dialogs import StatisticsSettingsDialog

logger = logging.getLogger(__name__)


def export_to_excel(dialog):
    """Exports visible contribution results to an Excel spreadsheet."""
    data = dialog.table_model._data
    if not data:
        return

    default_name = f"Contributions_Report_{QDate.currentDate().toString('yyyy-MM-dd')}.xlsx"
    save_path, _ = QFileDialog.getSaveFileName(
        dialog,
        "Save Contributions Report",
        str(Path.home() / "Desktop" / default_name),
        "Excel Files (*.xlsx)",
    )
    if not save_path:
        return

    dialog.btn_export.setEnabled(False)
    dialog.btn_export.setText("⏳ Exporting...")

    mtg_info = {"wg_name": "Contributions", "meeting_number": "Audit"}
    selected_cols = ContributionResultsTableModel.COLUMNS

    dialog.exporter_thread = ExcelExporterThread(
        output_path=Path(save_path),
        rows_data=data,
        selected_columns=selected_cols,
        mtg_info=mtg_info,
        docs_ftp_url="",
        auto_open=True,
    )

    dialog.exporter_thread.finished.connect(
        lambda success, msg: on_export_finished(dialog, success, msg)
    )
    dialog.exporter_thread.start()


def on_export_finished(dialog, success: bool, msg: str):
    """Handles completion of the Excel report export thread."""
    dialog.btn_export.setEnabled(True)
    dialog.btn_export.setText("📥 Export to Excel...")

    if success:
        reply = QMessageBox.question(
            dialog,
            "Export Complete",
            f"Report saved successfully:\n{msg}\n\nOpen spreadsheet now?",
            QMessageBox.Yes | QMessageBox.No,
        )
        if reply == QMessageBox.Yes:
            try:
                os.startfile(msg) if hasattr(os, "startfile") else webbrowser.open(f"file:///{msg}")
            except Exception as e:
                logger.error(f"Failed to open Excel report: {e}")
    else:
        QMessageBox.critical(dialog, "Export Failed", f"Could not generate Excel report:\n{msg}")


def generate_statistics(dialog):
    """Generates an HTML statistics dashboard for the searched contributions."""
    data = dialog.table_model._data
    if not data:
        return

    dialog.btn_stats.setEnabled(False)
    dialog.btn_stats.setText("⏳ Generating...")

    config = StatisticsSettingsDialog().load_config()
    current_cache = dialog.settings.cache_dir
    export_dir = Path(current_cache) / "Contributions_Audit" / "Export"

    dialog.stats_thread = ContributionStatsExporterThread(
        export_dir=export_dir,
        tdocs_data=data,
        target_companies=dialog.company_widget.get_selected_companies(),
        target_wis=dialog.wi_token_widget.get_tokens(),
        config=config,
        parent=dialog,
    )
    dialog.stats_thread.finished.connect(
        lambda success, msg: on_stats_finished(dialog, success, msg)
    )
    dialog.stats_thread.start()


def on_stats_finished(dialog, success: bool, msg: str):
    """Handles completion of the contribution statistics generation thread."""
    dialog.btn_stats.setEnabled(True)
    dialog.btn_stats.setText("📊 Statistics Dashboard...")

    if success:
        QMessageBox.information(
            dialog,
            "Dashboard Ready",
            f"Analytics dashboard generated successfully!\n\nSaved to:\n{msg}",
        )
        try:
            os.startfile(msg) if hasattr(os, "startfile") else webbrowser.open(f"file:///{msg}")
        except Exception as e:
            logger.error(f"Could not open dashboard: {e}")
    else:
        QMessageBox.warning(
            dialog, "Generation Failed", f"Could not generate statistics dashboard:\n{msg}"
        )
