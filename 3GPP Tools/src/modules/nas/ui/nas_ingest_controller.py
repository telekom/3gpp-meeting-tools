# --- File: src/modules/nas/ui/nas_ingest_controller.py ---
"""
Ingestion and database maintenance controller for NASTab.
Handles fetching from specifications database, importing local Word documents,
managing asynchronous parsing/ingestion threads, version removal, and database wiping.
"""
import logging
import re
from pathlib import Path
from typing import Any, Dict, List, Optional

from PyQt5.QtCore import Qt, QThread
from PyQt5.QtWidgets import QDialog, QFileDialog, QMessageBox

from modules.nas.core.nas_db import parse_version_tuple
from modules.nas.core.nas_threads import NASFetchAndImportThread, find_cached_spec_file
from modules.nas.core.parsing.protocol_parser_common import ProtocolDocxDispatcher
from modules.nas.ui.nas_dialogs import NASVersionSelectDialog
from modules.spec_search.core.spec_doc_resolver import SpecDocResolver

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


def is_ingestion_running(tab) -> bool:
    """Safely checks if a background specification ingestion thread is currently active.

    Guards against QObject.thread() bound method collision where hasattr(tab, 'thread')
    evaluates to True on uninitialized QWidget instances.
    """
    th = getattr(tab, "thread", None)
    if th is None or not hasattr(th, "isRunning"):
        if "_ingest_thread" in getattr(tab, "__dict__", {}):
            th = tab.__dict__["_ingest_thread"]
        else:
            return False

    if th is not None and hasattr(th, "isRunning") and callable(getattr(th, "isRunning", None)):
        try:
            return bool(th.isRunning())
        except Exception:
            return False
    return False


def start_batch_ingestion(tab, tasks: List[Dict[str, Any]]):
    """Spawns background worker thread to parse and ingest specifications."""
    if is_ingestion_running(tab):
        QMessageBox.warning(
            tab,
            "Ingestion In Progress",
            "A specification ingestion task is currently running. Please wait for it to finish.",
        )
        return

    if not tasks:
        return

    tab.progress_bar.setVisible(True)
    tab.progress_bar.setValue(0)
    tab.fetch_btn.setEnabled(False)
    tab.import_file_btn.setEnabled(False)
    if hasattr(tab, "manage_btn"):
        tab.manage_btn.setEnabled(False)

    worker = NASFetchAndImportThread(
        nas_db_path=tab.nas_db_path,
        tasks=tasks,
        cache_dir=tab.cache_dir,
    )
    tab.thread = worker
    tab._ingest_thread = worker
    worker.progress.connect(lambda msg, val: on_import_progress(tab, msg, val))
    worker.finished_success.connect(lambda s_cnt, m_cnt: on_import_success(tab, s_cnt, m_cnt))
    worker.error.connect(lambda err: on_import_error(tab, err))
    worker.start()


def on_import_error(tab, err_msg: str):
    """Handles error during background specification ingestion."""
    tab.progress_bar.setVisible(False)
    tab.fetch_btn.setEnabled(True)
    tab.import_file_btn.setEnabled(True)
    if hasattr(tab, "manage_btn"):
        tab.manage_btn.setEnabled(True)
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
    if hasattr(tab, "manage_btn"):
        tab.manage_btn.setEnabled(True)
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
    elif dialog.table.rowCount() > 0:
        dialog.table.selectRow(0)
        dialog.table.scrollToItem(dialog.table.item(0, 0))

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


def build_spec_reload_task(
    tab,
    spec_number: str,
    version: str,
    prompt_on_missing: bool = True,
) -> Optional[Dict[str, Any]]:
    """Resolves local cached files or FTP metadata for a given spec number and version.

    If cached Word files or archive are found, constructs a background ingestion task.
    If not found and prompt_on_missing is True, allows the user to browse for the .docx file.
    """
    clean_spec = re.sub(r"^(?:TS|TR)\s*", "", str(spec_number or "").strip(), flags=re.IGNORECASE)
    clean_ver = str(version or "").lstrip("vV").strip()

    if not clean_spec or not clean_ver:
        return None

    filename = ""
    file_url = ""

    # 1. Query specs_db for official archive metadata if available
    if getattr(tab, "specs_db", None):
        try:
            rows = tab.specs_db.search_files(spec_number=clean_spec, release_version=clean_ver)
            if rows:
                filename = rows[0][4]
                file_url = rows[0][6]
        except Exception as e:
            logger.warning(f"Could not query specs_db for TS {clean_spec} v{clean_ver}: {e}")

    # 2. Resolve cached files or download directory contents via SpecDocResolver
    info = SpecDocResolver.resolve_spec_files(
        clean_spec,
        clean_ver,
        cache_dir=tab.cache_dir,
        specs_db=getattr(tab, "specs_db", None),
    )

    if not filename and info.get("filename"):
        filename = info["filename"]
    if not file_url and info.get("url"):
        file_url = info["url"]

    local_docs = info.get("word_docs") or []
    if local_docs:
        local_docs.sort(key=ProtocolDocxDispatcher._extract_part_index)

    # 3. Check for cached zip archive if Word documents are not yet extracted
    if not local_docs:
        archive_name = filename or f"{clean_spec}_{clean_ver}.zip"
        cached_zip = find_cached_spec_file(archive_name, clean_spec)
        if not cached_zip and info.get("zip_file"):
            cached_zip = info["zip_file"]

        if cached_zip and cached_zip.exists():
            filename = cached_zip.name

    # 4. If neither local documents nor archive nor remote URL can be resolved
    if not local_docs and not filename and not file_url:
        if prompt_on_missing:
            file_paths, _ = QFileDialog.getOpenFileNames(
                tab,
                f"Select Specification Document(s) for TS {clean_spec} v{clean_ver} (.docx)",
                "",
                "Word Files (*.docx)",
            )
            if file_paths:
                p_list = [Path(p) for p in file_paths]
                p_list.sort(key=ProtocolDocxDispatcher._extract_part_index)
                return {
                    "spec_number": clean_spec,
                    "version": clean_ver,
                    "filename": f"{clean_spec}_{clean_ver}.docx" if len(p_list) == 1 else f"{clean_spec}_{clean_ver}.zip",
                    "file_url": "",
                    "local_docx_paths": p_list,
                }
        return None

    task: Dict[str, Any] = {
        "spec_number": clean_spec,
        "version": clean_ver,
        "filename": filename or f"{clean_spec}_{clean_ver}.zip",
        "file_url": file_url,
    }
    if local_docs:
        task["local_docx_paths"] = local_docs

    return task


def reload_single_version(tab, spec_number: str, version: str):
    """Re-parses and re-ingests a single specification version into the database."""
    if is_ingestion_running(tab):
        QMessageBox.warning(
            tab,
            "Ingestion In Progress",
            "A specification ingestion task is currently running. Please wait for it to finish.",
        )
        return

    reply = QMessageBox.question(
        tab,
        "Confirm Reload",
        f"Reload TS {spec_number} v{version}?\n\nThis will re-parse the specification and refresh all protocol messages and Information Elements.",
        QMessageBox.Yes | QMessageBox.No,
    )
    if reply != QMessageBox.Yes:
        return

    task = build_spec_reload_task(tab, spec_number, version, prompt_on_missing=True)
    if not task:
        QMessageBox.warning(
            tab,
            "Reload Cancelled",
            f"Could not locate cached files or download metadata for TS {spec_number} v{version}.",
        )
        return

    tab._save_config()
    start_batch_ingestion(tab, [task])


def reload_spec_group(tab, spec_number: str, child_count: int):
    """Re-parses and re-ingests all stored versions of a specification group."""
    if is_ingestion_running(tab):
        QMessageBox.warning(
            tab,
            "Ingestion In Progress",
            "A specification ingestion task is currently running. Please wait for it to finish.",
        )
        return

    reply = QMessageBox.question(
        tab,
        "Confirm Reload",
        f"Reload all {child_count} version(s) of TS {spec_number}?\n\nThis will re-parse all releases in the background and refresh the protocol library.",
        QMessageBox.Yes | QMessageBox.No,
    )
    if reply != QMessageBox.Yes:
        return

    all_versions = tab.version_tree.get_all_versions_for_spec(spec_number)
    tasks = []
    skipped = []

    for item in all_versions:
        ver = item.get("version")
        if ver:
            task = build_spec_reload_task(tab, spec_number, ver, prompt_on_missing=False)
            if task:
                tasks.append(task)
            else:
                skipped.append(ver)

    if not tasks:
        QMessageBox.warning(
            tab,
            "Reload Failed",
            f"Could not locate cached documents or 3GPP archive metadata for TS {spec_number}.",
        )
        return

    if skipped:
        logger.warning(
            f"Could not automatically resolve {len(skipped)} version(s) for TS {spec_number}: {skipped}"
        )

    tab._save_config()
    start_batch_ingestion(tab, tasks)


def reload_checked_versions(tab):
    """Re-parses and re-ingests all currently checked specification versions."""
    if is_ingestion_running(tab):
        QMessageBox.warning(
            tab,
            "Ingestion In Progress",
            "A specification ingestion task is currently running. Please wait for it to finish.",
        )
        return

    checked_versions = tab.version_tree.get_checked_versions_info()
    if not checked_versions:
        QMessageBox.warning(tab, "Select Version", "Please check at least one specification version to reload.")
        return

    reply = QMessageBox.question(
        tab,
        "Confirm Reload",
        f"Reload {len(checked_versions)} checked specification version(s)?\n\nThis will re-parse each specification and refresh the protocol library.",
        QMessageBox.Yes | QMessageBox.No,
    )
    if reply != QMessageBox.Yes:
        return

    tasks = []
    skipped = []

    for item in checked_versions:
        spec_num = item.get("spec_number")
        ver = item.get("version")
        if spec_num and ver:
            task = build_spec_reload_task(tab, spec_num, ver, prompt_on_missing=False)
            if task:
                tasks.append(task)
            else:
                skipped.append(f"TS {spec_num} v{ver}")

    if not tasks:
        QMessageBox.warning(
            tab,
            "Reload Failed",
            "Could not locate files or archive metadata for the checked specifications.",
        )
        return

    if skipped:
        logger.warning(
            f"Could not automatically resolve {len(skipped)} version(s): {', '.join(skipped)}"
        )

    tab._save_config()
    start_batch_ingestion(tab, tasks)
