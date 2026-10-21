# --- File: src/modules/meetings/ui/meetings_actions.py ---
"""
Meeting action workflows and export controllers for MeetingsTab.
Handles opening recent meetings, launching contribution reports, caching
TDocs, and exporting merged multi-meeting TDoc Excel workbooks.
"""
import logging
import os
import re
import webbrowser
from pathlib import Path

from PyQt5.QtCore import QDate
from PyQt5.QtWidgets import QFileDialog, QMessageBox

from modules.meetings.core.tdocs_merger import TDocsMergerThread
from modules.meetings.ui.contribution_dialog import ContributionReportDialog

logger = logging.getLogger(__name__)


def open_last_meeting(tab):
    """Reopens the most recently viewed meeting from stored user settings."""
    try:
        last_id, last_num, last_wg = tab.settings.get_last_meeting()

        if not last_id or not last_num:
            QMessageBox.information(
                tab, "No History", "No recent meeting history found. Please open a meeting first."
            )
            return

        wg_filter = [last_wg] if last_wg else None
        results = tab.db.search_meetings(wg_name=wg_filter, search_term=last_num)
        target_meeting = next((m for m in results if m.get("mtg_id") == last_id), None)

        if not target_meeting:
            results = tab.db.search_meetings(search_term=last_num)
            target_meeting = next((m for m in results if m.get("mtg_id") == last_id), None)

        display_tag = f"{last_wg}#{last_num}" if last_wg else f"'{last_num}'"

        if not target_meeting:
            QMessageBox.warning(
                tab,
                "Not Found",
                f"Meeting {display_tag} could not be found in the database.\nIt may have been cleared or the database was updated.",
            )
            return

        filepath = tab._get_tdoc_list_path(target_meeting)

        if filepath and filepath.exists():
            tab._open_tdocs_window(target_meeting, str(filepath))
        else:
            tab._download_and_open_tdocs(target_meeting)

    except Exception as e:
        QMessageBox.critical(tab, "Launch Error", f"Could not open last meeting:\n{e}")


def start_tdocs_caching(tab, docs_url: str, local_path: Path):
    """Initiates asynchronous caching of TDoc files for an entire meeting."""
    from modules.meetings.core.tdocs_cacher import TDocsCacherThread

    tab.update_btn.setText("⏳ Caching TDocs...")
    tab.update_btn.setEnabled(False)
    tab.cacher_thread = TDocsCacherThread(docs_url, local_path, tab)
    tab.cacher_thread.finished.connect(lambda success, msg: on_tdocs_caching_finished(tab, success, msg))
    tab.cacher_thread.start()


def on_tdocs_caching_finished(tab, success: bool, msg: str):
    """Resets UI and displays result of meeting TDoc caching."""
    tab.update_btn.setText("🔄 Sync All Meetings")
    tab.update_btn.setEnabled(True)
    if success:
        QMessageBox.information(tab, "Caching Complete", msg)
    else:
        QMessageBox.warning(tab, "Caching Failed", msg)


def export_merged_tdocs(tab):
    """Exports visible meetings into a unified merged Excel spreadsheet."""
    meetings_data = getattr(tab.table_model, "_data", [])
    if not meetings_data:
        QMessageBox.warning(
            tab, "No Data", "There are no meetings currently visible in the table to export."
        )
        return

    default_name = f"Merged_TDocs_{QDate.currentDate().toString('yyyy-MM-dd')}.xlsx"
    save_path, _ = QFileDialog.getSaveFileName(
        tab,
        "Save Merged TDocs",
        str(Path.home() / "Desktop" / default_name),
        "Excel Files (*.xlsx)",
    )
    if not save_path:
        return

    reply = QMessageBox.question(
        tab,
        "Force Download?",
        "Do you want to force a fresh download of all Excel files from 3GPP?\n\n"
        "• Select 'Yes' to fetch the absolute latest updates.\n"
        "• Select 'No' to use your local cache for files you've already downloaded.",
        QMessageBox.Yes | QMessageBox.No,
        QMessageBox.No,
    )
    force_download = reply == QMessageBox.Yes

    tab.btn_export_merged.setText("⏳ Merging...")
    tab.btn_export_merged.setEnabled(False)

    current_cache = (
        tab.dl_dir_input.text().strip() if hasattr(tab, "dl_dir_input") else tab.settings.cache_dir
    )
    tab.merger_thread = TDocsMergerThread(meetings_data, force_download, save_path, current_cache, tab)
    tab.merger_thread.progress.connect(lambda msg: tab.btn_export_merged.setText(f"⏳ {msg[:20]}..."))
    tab.merger_thread.finished.connect(lambda success, msg: on_merger_finished(tab, success, msg))
    tab.merger_thread.start()


def on_merger_finished(tab, success: bool, msg: str):
    """Handles completion of merged TDocs Excel file generation."""
    tab.btn_export_merged.setText("📥 Export Merged TDocs (Excel)")
    tab.btn_export_merged.setEnabled(True)

    if success:
        QMessageBox.information(tab, "Export Complete", msg)
        if (
            QMessageBox.question(
                tab,
                "Open File",
                "Would you like to open the generated Excel file now?",
                QMessageBox.Yes | QMessageBox.No,
            )
            == QMessageBox.Yes
        ):
            path_match = re.search(r"Saved to:\n(.*)", msg)
            if path_match:
                filepath = path_match.group(1).strip()
                try:
                    os.startfile(filepath) if hasattr(os, "startfile") else webbrowser.open(
                        f"file:///{filepath}"
                    )
                except Exception as e:
                    logger.error(f"Could not open merged Excel: {e}")
    else:
        QMessageBox.warning(tab, "Export Failed", msg)


def open_contribution_report(tab):
    """Opens or activates the independent Contribution Report dialog window."""
    if hasattr(tab, "contribution_dialog") and tab.contribution_dialog is not None:
        if tab.contribution_dialog.isVisible():
            tab.contribution_dialog.raise_()
            tab.contribution_dialog.activateWindow()
            return
        else:
            tab.contribution_dialog.deleteLater()
            tab.contribution_dialog = None

    tab.contribution_dialog = ContributionReportDialog(tab.db, parent=None)
    tab.contribution_dialog.open_meeting_requested.connect(
        lambda wg, num: open_meeting_from_report(tab, wg, num)
    )

    if tab.enable_dates_cb.isChecked():
        tab.contribution_dialog.date_from.setDate(tab.date_from.date())
        tab.contribution_dialog.date_to.setDate(tab.date_to.date())
    active_wgs = tab.wg_filter.getCheckedItems()
    if active_wgs:
        tab.contribution_dialog.set_active_wgs(active_wgs)

    tab.contribution_dialog.show()


def open_meeting_from_report(tab, wg_name: str, meeting_number: str):
    """Opens a meeting window triggered from a row selection in the contributions report."""
    results = tab.db.search_meetings(
        wg_name=[wg_name] if wg_name else None, search_term=meeting_number
    )
    target_meeting = next(
        (
            m
            for m in results
            if m.get("meeting_number") == meeting_number and (not wg_name or m.get("wg_name") == wg_name)
        ),
        None,
    )
    if not target_meeting and results:
        target_meeting = results[0]

    if target_meeting:
        filepath = tab._get_tdoc_list_path(target_meeting)
        if filepath and filepath.exists():
            tab._open_tdocs_window(target_meeting, str(filepath))
        else:
            tab._download_and_open_tdocs(target_meeting)
    else:
        QMessageBox.warning(
            tab,
            "Meeting Not Found",
            f"Could not find meeting {wg_name} #{meeting_number} in the database.",
        )
