"""Specification indexing, smart import, retention pruning, and DB maintenance mixin."""

from __future__ import annotations

import logging
import re
from pathlib import Path
from typing import Any, Dict, List, Optional

from PyQt5.QtCore import QThread
from PyQt5.QtWidgets import QDialog, QFileDialog, QMessageBox

from core.ui.toast import notify_info
from modules.nas.core.parsing.protocol_parser_common import ProtocolDocxDispatcher
from modules.spec_search.core.spec_search_threads import (
    SpecSearchImportThread,
    SpecSearchWipeWorker,
    is_change_mark_file,
)
from modules.spec_search.ui.spec_search_dialogs import (
    SpecSearchVersionSelectDialog,
    SpecSmartImportDialog,
)

logger = logging.getLogger(__name__)


class SpecSearchTabIngestionMixin:
    """Handles version card refresh, FTP/local batch indexing, pruning, and DB wipes."""

    def refresh_versions(self):
        versions = self.db.get_imported_versions()
        saved_checked = self.version_tree.get_checked_versions_info()
        if not saved_checked:
            cfg = self._load_config()
            saved_checked = cfg.get("checked_versions") if cfg else None

        self.version_tree.populate(versions, saved_checked=saved_checked, spec_titles=self._spec_titles)
        self.selected_version_ids = self.version_tree.get_selected_version_ids()
        self._update_db_stats()

    def _open_smart_import_dialog(self, initial_spec: Optional[str] = None):
        """Opens the targeted Smart Import dialog to check and ingest latest versions of all releases."""
        if not self.specs_db:
            QMessageBox.warning(self, "Specs DB Unavailable", "The specifications database (3gpp_data.db) is not configured.")
            return

        if not initial_spec:
            checked_info = self.version_tree.get_checked_versions_info()
            if checked_info:
                initial_spec = checked_info[0].get("spec_number")

        dialog = SpecSmartImportDialog(
            specs_db=self.specs_db,
            search_db=self.db,
            cache_dir=self.cache_dir,
            initial_spec=initial_spec,
            parent=self,
        )
        if hasattr(dialog, "chk_auto_supersede"):
            dialog.chk_auto_supersede.setChecked(getattr(self, "auto_replace_older_versions", True))

        if dialog.exec_() == dialog.Accepted and dialog.selected_tasks:
            if hasattr(dialog, "auto_replace_older"):
                self.auto_replace_older_versions = dialog.auto_replace_older
                if hasattr(self, "act_auto_replace"):
                    self.act_auto_replace.setChecked(self.auto_replace_older_versions)
                self._save_config()
            self._start_batch_ingestion(dialog.selected_tasks)

    def _on_fetch_clicked(self):
        if not self.specs_db:
            QMessageBox.warning(self, "Specs DB Unavailable", "The specifications database (3gpp_data.db) is not configured.")
            return

        dialog = SpecSearchVersionSelectDialog(self.specs_db, self.db, self.cache_dir, self)
        if hasattr(dialog, "chk_auto_supersede"):
            dialog.chk_auto_supersede.setChecked(getattr(self, "auto_replace_older_versions", True))
        if dialog.exec_() == dialog.Accepted and dialog.selected_files_info:
            if hasattr(dialog, "auto_replace_older"):
                self.auto_replace_older_versions = dialog.auto_replace_older
                if hasattr(self, "act_auto_replace"):
                    self.act_auto_replace.setChecked(self.auto_replace_older_versions)
                self._save_config()
            self._start_batch_ingestion(dialog.selected_files_info)

    def _on_import_local_clicked(self):
        paths, _ = QFileDialog.getOpenFileNames(
            self, "Select Specification Document(s) (.docx)", "", "Word Files (*.docx)"
        )
        if not paths:
            return

        clean_paths = [p for p in paths if not is_change_mark_file(p)]
        if len(clean_paths) < len(paths):
            skipped = len(paths) - len(clean_paths)
            logger.log(logging.INFO, f"ℹ️ Skipped {skipped} revision mark (-rm) file(s).")

        grouped: Dict[str, List[Path]] = {}
        for fp in clean_paths:
            p = Path(fp)
            base_key = re.sub(r"_\d+_.*$", "", p.stem)
            base_key = re.sub(r"[-_]cl$", "", base_key, flags=re.IGNORECASE)
            grouped.setdefault(base_key, []).append(p)

        tasks = []
        for base_key, p_list in grouped.items():
            dispatcher = ProtocolDocxDispatcher(p_list)
            spec_num = dispatcher.extract_spec_number()
            version = dispatcher.extract_version_from_filename()
            tasks.append({
                "spec_number": spec_num,
                "version": version,
                "filename": f"{base_key}.docx",
                "file_url": "",
                "release_date": "",
                "local_docx_paths": p_list,
            })

        self._start_batch_ingestion(tasks)

    def _start_batch_ingestion(self, tasks: List[Dict[str, Any]]):
        active_th = getattr(self, "_import_thread", None) or getattr(self, "thread", None)
        if isinstance(active_th, QThread) and active_th.isRunning():
            QMessageBox.warning(
                self,
                "Ingestion In Progress",
                "A specification indexing task is currently running. Please wait for it to finish.",
            )
            return

        self._current_ingesting_tasks = tasks
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.fetch_btn.setEnabled(False)
        self.import_local_btn.setEnabled(False)

        self.thread = SpecSearchImportThread(
            search_db_path=self.search_db_path,
            tasks=tasks,
            cache_dir=self.cache_dir,
            auto_replace_older=getattr(self, "auto_replace_older_versions", True),
        )
        self._import_thread = self.thread
        self.thread.progress.connect(self._on_import_progress)
        self.thread.finished_success.connect(self._on_import_success)
        self.thread.error.connect(self._on_import_error)
        self.thread.start()

    def _on_import_progress(self, msg: str, val: int):
        self.progress_bar.setValue(val)
        logger.log(logging.INFO, msg)

    def _on_import_success(self, specs_count: int, clauses_count: int):
        self.progress_bar.setVisible(False)
        self.fetch_btn.setEnabled(True)
        self.import_local_btn.setEnabled(True)
        logger.log(logging.INFO, f"✅ Indexed {specs_count} specification(s) ({clauses_count} total clauses).")

        # Ensure newly ingested versions are preserved in saved_checked
        tasks = getattr(self, "_current_ingesting_tasks", [])
        saved_checked = self.version_tree.get_checked_versions_info()
        if saved_checked is None:
            saved_checked = []
        for t in tasks:
            s_num = str(t.get("spec_number", "")).strip()
            ver = str(t.get("version", "")).lstrip("vV").strip()
            if s_num and ver:
                if not any(
                    str(item.get("spec_number", "")).strip() == s_num
                    and str(item.get("version", "")).lstrip("vV").strip() == ver
                    for item in saved_checked
                ):
                    saved_checked.append({"spec_number": s_num, "version": ver})

        versions = self.db.get_imported_versions()
        self.version_tree.populate(versions, saved_checked=saved_checked, spec_titles=self._spec_titles)
        self.selected_version_ids = self.version_tree.get_selected_version_ids()
        self._save_config()

        # Focus newly indexed version if single spec
        if len(tasks) == 1:
            t = tasks[0]
            self.focus_spec(t.get("spec_number", ""), t.get("version", ""))

    def _on_import_error(self, err: str):
        self.progress_bar.setVisible(False)
        self.fetch_btn.setEnabled(True)
        self.import_local_btn.setEnabled(True)
        QMessageBox.critical(self, "Indexing Error", err)
        logger.log(logging.ERROR, f"❌ Indexing error: {err}")


    def _on_toggle_auto_replace(self, checked: bool):
        self.auto_replace_older_versions = checked
        self._save_config()
        logger.log(logging.INFO, f"Auto-supersede older versions on import set to: {checked}")

    def _on_prune_superseded_clicked(self):
        superseded = self.db.get_superseded_versions()
        if not superseded:
            notify_info(
                self,
                "No Superseded Versions",
                "All indexed specifications are already at their latest version per release.",
            )
            return

        from modules.spec_search.ui.spec_search_dialogs import SpecSearchPruneDialog
        dialog = SpecSearchPruneDialog(self.db, superseded, self)
        if dialog.exec_() == QDialog.Accepted and dialog.pruned_count > 0:
            self.refresh_versions()
            self._save_config()
            self._execute_search()
            logger.log(logging.INFO, f"🧹 Pruned {dialog.pruned_count} superseded version(s) from search index.")

    def _prune_single_spec(self, spec_number: str):
        superseded = [
            v for v in self.db.get_superseded_versions()
            if v["spec_number"] == spec_number
        ]
        if not superseded:
            notify_info(
                self,
                "No Superseded Versions",
                f"All indexed versions of TS {spec_number} are already the latest version for each release.",
            )
            return

        from modules.spec_search.ui.spec_search_dialogs import SpecSearchPruneDialog
        dialog = SpecSearchPruneDialog(self.db, superseded, self)
        if dialog.exec_() == QDialog.Accepted and dialog.pruned_count > 0:
            self.refresh_versions()
            self._save_config()
            self._execute_search()
            logger.log(logging.INFO, f"🧹 Pruned {dialog.pruned_count} superseded version(s) of TS {spec_number}.")

    def _delete_single_version(self, spec_number: str, version: str):
        if QMessageBox.question(self, "Confirm Delete", f"Delete TS {spec_number} v{version}?", QMessageBox.Yes | QMessageBox.No) == QMessageBox.Yes:
            self.db.clear_version(spec_number, version)
            self.refresh_versions()
            self._save_config()
            self._execute_search()

    def _delete_spec_group(self, spec_number: str, count: int):
        if QMessageBox.question(self, "Confirm Delete", f"Delete all {count} version(s) of TS {spec_number}?", QMessageBox.Yes | QMessageBox.No) == QMessageBox.Yes:
            for item in self.version_tree.get_all_versions_for_spec(spec_number):
                if item.get("version"):
                    self.db.clear_version(spec_number, item["version"])
            self.refresh_versions()
            self._save_config()
            self._execute_search()

    def _on_clear_version_clicked(self):
        checked = self.version_tree.get_checked_versions_info()
        if not checked:
            QMessageBox.warning(self, "Select Version", "Please check at least one version to clear.")
            return
        if QMessageBox.question(self, "Confirm Delete", f"Delete {len(checked)} checked version(s)?", QMessageBox.Yes | QMessageBox.No) == QMessageBox.Yes:
            for item in checked:
                self.db.clear_version(item["spec_number"], item["version"])
            self.refresh_versions()
            self._save_config()
            self._execute_search()

    def _on_wipe_db_clicked(self):
        if QMessageBox.critical(
            self,
            "Confirm Wipe",
            "Delete all indexed specification clauses and full-text indexes?",
            QMessageBox.Yes | QMessageBox.No,
        ) != QMessageBox.Yes:
            return

        for w in list(self._active_query_workers.values()):
            if w.isRunning():
                w.requestInterruption()
                w.wait(1500)

        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, 0)
        self.wipe_db_btn.setEnabled(False)
        self.fetch_btn.setEnabled(False)
        self.import_local_btn.setEnabled(False)
        self.clear_ver_btn.setEnabled(False)

        self._wipe_worker = SpecSearchWipeWorker(self.db)
        self._wipe_worker.finished_success.connect(self._on_wipe_db_success)
        self._wipe_worker.error.connect(self._on_wipe_db_error)
        self._wipe_worker.start()

    def _on_wipe_db_success(self):
        self.progress_bar.setVisible(False)
        self.progress_bar.setRange(0, 100)
        self.wipe_db_btn.setEnabled(True)
        self.fetch_btn.setEnabled(True)
        self.import_local_btn.setEnabled(True)
        self.clear_ver_btn.setEnabled(True)

        self.refresh_versions()
        self.spec_results_tabs.clear()
        self.inspector.clear_display()
        self.matrix_title.setText("Type a query above to see Release Evolution Matrix")
        self._save_config()
        logger.log(logging.INFO, "🧹 Specification Search DB wiped successfully.")

    def _on_wipe_db_error(self, err: str):
        self.progress_bar.setVisible(False)
        self.progress_bar.setRange(0, 100)
        self.wipe_db_btn.setEnabled(True)
        self.fetch_btn.setEnabled(True)
        self.import_local_btn.setEnabled(True)
        self.clear_ver_btn.setEnabled(True)

        QMessageBox.critical(self, "Wipe Error", f"Failed to wipe search database: {err}")
        logger.log(logging.ERROR, f"❌ Wipe error: {err}")
