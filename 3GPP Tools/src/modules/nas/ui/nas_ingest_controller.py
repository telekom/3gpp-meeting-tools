# --- File: src/modules/nas/ui/nas_ingest_controller.py ---
"""
Ingestion and database maintenance controller for NASTab.
Handles fetching from specifications database, importing local Word documents,
managing asynchronous parsing/ingestion threads, version removal, and database wiping.
"""
import logging
import re
from pathlib import Path
from typing import Any, Dict, List

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QDialog, QFileDialog, QMessageBox

from modules.nas.core.nas_db import parse_version_tuple
from modules.nas.core.nas_threads import NASFetchAndImportThread
from modules.nas.core.parsing.protocol_parser_common import ProtocolDocxDispatcher
from modules.nas.ui.nas_dialogs import NASVersionSelectDialog

logger = logging.getLogger(__name__)


def fetch_from_specs_db_clicked(tab):
    """Opens version selection dialog to ingest protocol specifications from 3gpp_data.db."""
    if not tab.specs_db:
        QMessageBox.warning(
            tab,
            "Specs DB Unavailable",
            "The 3GPP Specifications database (3gpp_data.db) is not configured or reachable.",
        )
        return

    dialog = NASVersionSelectDialog(tab.specs_db, tab.db, tab.cache_dir, tab)
    if dialog.exec_() == QDialog.Accepted and dialog.selected_files_info:
        start_batch_ingestion(tab, dialog.selected_files_info)


def import_local_file_clicked(tab):
    """Prompts file dialog to select local .docx specification files and queues ingestion."""
    file_paths, _ = QFileDialog.getOpenFileNames(
        tab,
        "Select 3GPP Specification(s) (.docx)",
        "",
        "Word Files (*.docx)",
    )
    if not file_paths:
        return

    grouped_tasks: Dict[str, List[Path]] = {}
    for fp in file_paths:
        p = Path(fp)
        base_key = re.sub(r"_\d+_.*$", "", p.stem)
        grouped_tasks.setdefault(base_key, []).append(p)

    tasks = []
    for base_key, paths in grouped_tasks.items():
        parser_temp = ProtocolDocxDispatcher(paths)
        spec_num = parser_temp.extract_spec_number()
        version = parser_temp.extract_version_from_filename()
        tasks.append({
            "spec_number": spec_num,
            "version": version,
            "filename": f"{base_key}.docx" if len(paths) == 1 else f"{base_key}.zip",
            "file_url": "",
            "local_docx_paths": paths,
        })

    start_batch_ingestion(tab, tasks)


def start_batch_ingestion(tab, tasks: List[Dict[str, Any]]):
    """Spawns background worker thread to parse and ingest specifications."""
    tab.progress_bar.setVisible(True)
    tab.progress_bar.setValue(0)
    tab.fetch_btn.setEnabled(False)
    tab.import_file_btn.setEnabled(False)

    tab.thread = NASFetchAndImportThread(
        nas_db_path=tab.nas_db_path,
        tasks=tasks,
        cache_dir=tab.cache_dir,
    )
    tab.thread.progress.connect(lambda msg, val: on_import_progress(tab, msg, val))
    tab.thread.finished_success.connect(lambda s_cnt, m_cnt: on_import_success(tab, s_cnt, m_cnt))
    tab.thread.error.connect(lambda err: on_import_error(tab, err))
    tab.thread.start()


def on_import_error(tab, err_msg: str):
    """Handles error during background specification ingestion."""
    tab.progress_bar.setVisible(False)
    tab.fetch_btn.setEnabled(True)
    tab.import_file_btn.setEnabled(True)
    logger.log(logging.ERROR, f"❌ Import failed: {err_msg}")
    QMessageBox.critical(
        tab,
        "Import Error",
        f"An error occurred during specification ingestion:\n{err_msg}",
    )


def on_import_progress(tab, msg: str, val: int):
    """Updates progress bar and diagnostic log during ingestion."""
    tab.progress_bar.setValue(val)
    logger.log(logging.INFO, msg)


def on_import_success(tab, spec_count: int, msg_count: int):
    """Handles successful completion of specification ingestion batch."""
    tab.progress_bar.setVisible(False)
    tab.fetch_btn.setEnabled(True)
    tab.import_file_btn.setEnabled(True)
    logger.info(
        f"✅ Successfully ingested {spec_count} specification(s) ({msg_count} total messages/PDUs)."
    )
    tab.refresh_versions()


def clear_version_clicked(tab):
    """Prompts to delete all checked specification versions from the database."""
    checked_versions = tab.version_tree.get_checked_versions_info()
    if not checked_versions:
        QMessageBox.warning(tab, "Select Version", "Please check at least one specific version to clear.")
        return

    reply = QMessageBox.question(
        tab,
        "Confirm Delete",
        f"Delete {len(checked_versions)} checked specification version(s)?",
        QMessageBox.Yes | QMessageBox.No,
    )
    if reply == QMessageBox.Yes:
        for item in checked_versions:
            tab.db.clear_version(item["spec_number"], item["version"])
        tab.current_selected_message_name = None
        tab.refresh_versions()
        tab._save_config()


def wipe_db_clicked(tab):
    """Prompts confirmation and wipes the entire protocol SQLite database."""
    reply = QMessageBox.critical(
        tab,
        "Confirm Wipe",
        "This will delete ALL imported specifications and tables. Continue?",
        QMessageBox.Yes | QMessageBox.No,
    )
    if reply == QMessageBox.Yes:
        tab.db.wipe_database()
        tab.current_selected_message_name = None
        if tab.config_path.exists():
            try:
                tab.config_path.unlink()
            except Exception:
                pass
        tab.refresh_versions()
        tab.msg_list.clear()
        tab.matrix_table.setModel(None)
        tab.inspector.clear_display()
        logger.log(logging.INFO, "🧹 Protocol Database wiped.")


def import_cross_ref_spec_requested(tab, spec_number: str, major_version: int):
    """Opens version selection dialog pre-filtered to a specific referenced spec."""
    if not tab.specs_db:
        QMessageBox.warning(
            tab,
            "Specs DB Unavailable",
            "The 3GPP Specifications database (3gpp_data.db) is not configured or reachable.",
        )
        return

    dialog = NASVersionSelectDialog(tab.specs_db, tab.db, tab.cache_dir, tab)

    for idx in range(dialog.spec_combo.count()):
        if dialog.spec_combo.itemData(idx) == spec_number:
            dialog.spec_combo.setCurrentIndex(idx)
            break

    if major_version > 0:
        dialog.table.clearSelection()
        for row in range(dialog.table.rowCount()):
            item = dialog.table.item(row, 0)
            if item:
                data = item.data(Qt.UserRole) or {}
                ver_tuple = parse_version_tuple(data.get("version", ""))
                if ver_tuple and (ver_tuple[0][1] if isinstance(ver_tuple[0], tuple) else ver_tuple[0]) == major_version:
                    dialog.table.selectRow(row)
                    dialog.table.scrollToItem(item)
                    break

    if dialog.exec_() == QDialog.Accepted and dialog.selected_files_info:
        start_batch_ingestion(tab, dialog.selected_files_info)


def delete_single_version(tab, spec_number: str, version: str):
    """Deletes all data for a single specification version."""
    reply = QMessageBox.question(
        tab,
        "Confirm Delete",
        f"Are you sure you want to delete all stored data for TS {spec_number} v{version}?",
        QMessageBox.Yes | QMessageBox.No,
    )
    if reply == QMessageBox.Yes:
        if tab.db.clear_version(spec_number, version):
            tab.current_selected_message_name = None
            tab.refresh_versions()
            tab._save_config()
            logger.log(logging.INFO, f"🗑️ Deleted TS {spec_number} v{version} from database.")
        else:
            QMessageBox.warning(tab, "Error", f"Failed to delete TS {spec_number} v{version}.")


def delete_spec_group(tab, spec_number: str, child_count: int):
    """Deletes all stored versions of a specification group."""
    reply = QMessageBox.question(
        tab,
        "Confirm Delete",
        f"Are you sure you want to delete all {child_count} version(s) of TS {spec_number} from the database?",
        QMessageBox.Yes | QMessageBox.No,
    )
    if reply == QMessageBox.Yes:
        for item in tab.version_tree.get_all_versions_for_spec(spec_number):
            if item.get("version"):
                tab.db.clear_version(spec_number, item["version"])
        tab.current_selected_message_name = None
        tab.refresh_versions()
        tab._save_config()
        logger.log(logging.INFO, f"🗑️ Deleted all versions of TS {spec_number} from database.")
