"""User actions, context menu, dialogs, and DB maintenance mixin for SpecificationsTab."""

from __future__ import annotations

import logging
import os
import re
import webbrowser
from pathlib import Path
from typing import Optional

from PyQt5 import sip
from PyQt5.QtWidgets import QComboBox, QDialog, QLabel, QMessageBox, QPushButton

from core.ui.toast import notify_info
from modules.specifications.ui.dialogs import (
    AddSpecDialog,
    AdvancedSyncDialog,
    SpecClauseExcerptDialog,
    SpecInfoDialog,
    SpecsConfigDialog,
    TableFilterDialog,
    TargetedSyncDialog,
)
from modules.specifications.ui.spec_action_controller import (
    handle_document_action,
    handle_zip_action,
)
from modules.specifications.ui.spec_annotations_dialog import (
    BatchTagDialog,
    EditSpecAnnotationsDialog,
)
from modules.specifications.ui.spec_history_dialog import SpecVersionHistoryDialog
from modules.specifications.ui.spec_menu import copy_specs_default, show_spec_context_menu
from modules.specifications.utils.utils import build_dynareport_url

logger = logging.getLogger(__name__)


class SpecsTabActionsMixin:
    """Implements context menus, Quick Copy, dialog launchers, document actions, and deletions."""

    def _open_settings_dialog(self):
        dialog = SpecsConfigDialog(self.download_dir, self)
        if dialog.exec_():
            new_path = dialog.get_download_path()
            if new_path and new_path != self.download_dir:
                self.download_dir = new_path
                self._save_settings()
                self.refresh_table()

    def _open_download_dir(self):
        target_dir = Path(self.download_dir)
        if not target_dir.exists():
            try:
                target_dir.mkdir(parents=True, exist_ok=True)
            except Exception as e:
                QMessageBox.warning(self, "Directory Error", f"Could not create directory:\n{e}")
                return
        try:
            os.startfile(str(target_dir))
        except Exception as e:
            QMessageBox.warning(self, "Explorer Error", f"Could not open directory:\n{e}")

    def _open_advanced_sync(self):
        dialog = AdvancedSyncDialog(self.db, self)
        if dialog.exec_():
            target_specs = dialog.matching_specs
            if target_specs:
                force_meta = self.force_meta_action.isChecked()
                self.update_specific_requested.emit(target_specs, force_meta)

    def _open_targeted_sync(self):
        dialog = TargetedSyncDialog(self)
        if dialog.exec_():
            targets = dialog.get_targets()
            if targets:
                force_meta = self.force_meta_action.isChecked()
                self.update_specific_requested.emit(targets, force_meta)

    def _open_table_filters(self):
        dialog = TableFilterDialog(self.db, self.table_filters, self)
        if dialog.exec_():
            self.table_filters = dialog.get_filters()
            self._sync_ui_to_filters()
            self.save_settings_timer.start()
            self.refresh_table()

    def _show_context_menu(self, position):
        show_spec_context_menu(self, position)

    def _copy_table_selection(self):
        """Copies formatted citation and rich link for selected specifications on Ctrl+C."""
        selected_rows = self.table.selectionModel().selectedRows()
        if not selected_rows:
            return

        target_items = []
        for index in selected_rows:
            row = index.row()
            widget = self.table.cellWidget(row, 0)
            if widget:
                num_label = widget.findChild(QLabel, "specNumberLabel")
                if num_label:
                    canonical = num_label.property("specNumber")
                    spec_num = ""
                    if canonical:
                        spec_num = str(canonical).strip()
                    else:
                        display_text = num_label.text().strip()
                        match = re.search(r"(\d{2}\.\d{2,3}(?:-[A-Za-z0-9]+)?)", display_text)
                        if match:
                            spec_num = match.group(1)

                    if spec_num:
                        cell_widget = self.table.cellWidget(row, 2)
                        v_info = None
                        if cell_widget:
                            combo = cell_widget.findChild(QComboBox)
                            if combo:
                                v_info = combo.currentData()
                        target_items.append((spec_num, v_info))

        if target_items:
            copy_specs_default(self, target_items, self.db)

    def _on_table_cell_double_clicked(self, row: int, column: int):
        if column in (0, 1):
            widget = self.table.cellWidget(row, 0)
            if widget:
                num_label = widget.findChild(QLabel, "specNumberLabel")
                if num_label:
                    canonical = num_label.property("specNumber")
                    if not canonical:
                        match = re.search(r"(\d{2}\.\d{2,3}(?:-[A-Za-z0-9]+)?)", num_label.text())
                        canonical = match.group(1) if match else num_label.text().strip().replace("⭐", "").strip()
                    self._open_edit_annotations(canonical)

    def _open_edit_annotations(self, spec_num: str):
        """Opens modal editor for specification tags and comments."""
        clean_num = str(spec_num).strip()
        if not clean_num or not self.db:
            return
        details = self.db.get_spec_details(clean_num)
        title = details.get("title", "")
        dlg = EditSpecAnnotationsDialog(clean_num, title, self.db, self)
        if dlg.exec_() == QDialog.Accepted:
            self._populate_tag_filter_combo()
            self.refresh_table()

    def _batch_add_tag(self, spec_nums: list):
        """Opens batch tag dialog for multiple selected specifications."""
        if not spec_nums or not self.db:
            return
        dlg = BatchTagDialog(spec_nums, self.db, self)
        if dlg.exec_() == QDialog.Accepted:
            self._populate_tag_filter_combo()
            self.refresh_table()

    def _show_spec_info(self, spec_num: str):
        details = self.db.get_spec_details(spec_num)
        if not details and "-" in spec_num:
            base_num = spec_num.split("-")[0]
            details = self.db.get_spec_details(base_num)
            if details:
                details["number"] = spec_num

        dialog = SpecInfoDialog(details or {}, self, db=self.db)
        dialog.exec_()
        self._populate_tag_filter_combo()
        self.refresh_table()

    def _open_spec_folder(self, spec_num: str):
        target_dir = Path(self.download_dir) / spec_num
        if target_dir.exists() and any(target_dir.iterdir()):
            try:
                os.startfile(str(target_dir))
            except Exception as e:
                QMessageBox.warning(self, "Explorer Error", f"Could not open directory:\n{e}")
        else:
            notify_info(
                self,
                "Folder Not Found",
                f"No downloaded files found for {spec_num}. Download a document first to create the local folder.",
            )

    def _open_web_report(self, spec_num: str):
        url = build_dynareport_url(spec_num)
        if url:
            webbrowser.open(url)

    def _get_row_version(self, row_idx: int) -> str:
        cell_widget = self.table.cellWidget(row_idx, 2)
        if cell_widget:
            combo = cell_widget.findChild(QComboBox)
            if combo:
                c_data = combo.currentData()
                if c_data and isinstance(c_data, dict):
                    return c_data.get("version", "")
        return ""

    def _handle_search_spec_action(self, spec_num: str, row_idx: int):
        ver = self._get_row_version(row_idx)
        self.search_in_spec_search_requested.emit(spec_num, ver)

    def _handle_index_spec_action(self, spec_num: str, row_idx: int):
        ver = self._get_row_version(row_idx)
        self.index_in_spec_search_requested.emit(spec_num, ver)

    def _batch_index_in_spec_search(self, target_items: list):
        spec_items = []
        for item in target_items:
            spec_num = item[0]
            v_info = item[1] if len(item) > 1 else None
            ver = v_info.get("version", "") if v_info else ""
            spec_items.append((spec_num, ver))
        self.batch_index_in_spec_search_requested.emit(spec_items)

    def _open_version_history_dialog(self, spec_number: str, version: Optional[str] = None):
        """Launches the dedicated Change History & Integrated TDocs dialog modelessly."""
        try:
            logger.info(f"📜 [Change History] Opening Change History & TDocs browser for TS {spec_number}{f' (v{version})' if version else ''}...")
            if hasattr(self, "_history_dialog") and self._history_dialog and not sip.isdeleted(self._history_dialog):
                self._history_dialog.load_spec(spec_number, version)
                self._history_dialog.show()
                self._history_dialog.raise_()
                self._history_dialog.activateWindow()
                return

            dialog = SpecVersionHistoryDialog(
                spec_number=spec_number,
                initial_version=version,
                parent=self,
                history_db=self.history_db,
                specs_db=self.db,
            )
            self._history_dialog = dialog
            dialog.show()
            dialog.raise_()
            dialog.activateWindow()
        except Exception as e:
            logger.error(f"Error opening SpecVersionHistoryDialog: {e}", exc_info=True)

    def _open_clause_excerpt_dialog(self, combo: QComboBox):
        if sip.isdeleted(combo):
            return
        c_data = combo.currentData()
        if not c_data:
            return

        spec_num = c_data["spec_num"]
        spec_dl_dir = Path(self.download_dir) / spec_num
        spec_type = "TS"
        details = self.db.get_spec_details(spec_num)
        if not details and "-" in spec_num:
            base_num = spec_num.split("-")[0]
            details = self.db.get_spec_details(base_num)
        if details and details.get("type"):
            spec_type = details["type"]

        dialog = SpecClauseExcerptDialog(
            spec_num=spec_num,
            spec_type=spec_type,
            version=c_data.get("version", ""),
            fname=c_data.get("fname", ""),
            url=c_data.get("url", ""),
            spec_dl_dir=spec_dl_dir,
            parent=self,
        )

        def _on_extraction_done(p: Path, c=combo):
            try:
                if not sip.isdeleted(c):
                    idx = c.currentIndex()
                    self._update_combo_item_badges(c, idx)
                    c.currentIndexChanged.emit(idx)
            except RuntimeError:
                pass

        dialog.extraction_completed.connect(_on_extraction_done)
        if not hasattr(self, "_active_excerpt_dialogs"):
            self._active_excerpt_dialogs = []
        self._active_excerpt_dialogs.append(dialog)
        dialog.destroyed.connect(
            lambda: self._active_excerpt_dialogs.remove(dialog)
            if hasattr(self, "_active_excerpt_dialogs") and dialog in self._active_excerpt_dialogs
            else None
        )
        dialog.show()
        dialog.raise_()
        dialog.activateWindow()

    def _open_cr_pcr_dialog(self, combo: QComboBox):
        if sip.isdeleted(combo):
            return
        c_data = combo.currentData()
        if not c_data:
            return

        from modules.word_tools.ui.cr_pcr_dialog import CrPcrDialog

        spec_num = c_data["spec_num"]
        spec_dl_dir = Path(self.download_dir) / spec_num
        spec_type = "TS"
        release = c_data.get("release", "")
        details = self.db.get_spec_details(spec_num)
        if not details and "-" in spec_num:
            base_num = spec_num.split("-")[0]
            details = self.db.get_spec_details(base_num)
        if details and details.get("type"):
            spec_type = details["type"]

        cr_dlg = CrPcrDialog(
            spec_num=spec_num,
            spec_type=spec_type,
            version=c_data.get("version", ""),
            release=release,
            fname=c_data.get("fname", ""),
            url=c_data.get("url", ""),
            spec_dl_dir=spec_dl_dir,
            parent=self,
        )
        if not hasattr(self, "_active_cr_dialogs"):
            self._active_cr_dialogs = []
        self._active_cr_dialogs.append(cr_dlg)
        cr_dlg.destroyed.connect(
            lambda: self._active_cr_dialogs.remove(cr_dlg)
            if hasattr(self, "_active_cr_dialogs") and cr_dlg in self._active_cr_dialogs
            else None
        )
        cr_dlg.show()
        cr_dlg.raise_()
        cr_dlg.activateWindow()

    def _handle_document_action(self, combo: QComboBox, doc_type: str, btn: QPushButton):
        handle_document_action(self, combo, doc_type, btn)

    def _handle_zip_action(self, combo: QComboBox, btn: QPushButton):
        handle_zip_action(self, combo, btn)

    def _open_add_spec_dialog(self):
        dialog = AddSpecDialog(self.db, self)
        if dialog.exec_() == QDialog.Accepted:
            self.refresh_table()
            if dialog.sync_requested and dialog.saved_spec_number:
                force_meta = self.force_meta_action.isChecked()
                self.update_specific_requested.emit([dialog.saved_spec_number], force_meta)

    def _delete_single_spec(self, spec_num: str):
        reply = QMessageBox.question(
            self,
            "Confirm Delete",
            f"Are you sure you want to delete specification TS {spec_num} from the database?",
            QMessageBox.Yes | QMessageBox.No,
        )
        if reply == QMessageBox.Yes:
            deleted = self.db.delete_specification(spec_num)
            if not deleted and "-" in spec_num:
                base_num = spec_num.split("-")[0]
                deleted = self.db.delete_specification(base_num)
                if deleted:
                    self.favorites.discard(base_num)
            if deleted:
                self.favorites.discard(spec_num)
                self._save_settings()
                self.refresh_table()
                logger.log(logging.INFO, f"🗑️ Deleted TS {spec_num} from specifications database.")
            else:
                QMessageBox.warning(self, "Delete Error", f"Could not delete TS {spec_num}.")

    def _delete_batch_specs(self, spec_nums: list):
        reply = QMessageBox.question(
            self,
            "Confirm Delete",
            f"Are you sure you want to delete {len(spec_nums)} selected specifications from the database?",
            QMessageBox.Yes | QMessageBox.No,
        )
        if reply == QMessageBox.Yes:
            deleted_count = 0
            for num in spec_nums:
                deleted = self.db.delete_specification(num)
                if not deleted and "-" in num:
                    base_num = num.split("-")[0]
                    deleted = self.db.delete_specification(base_num)
                    if deleted:
                        self.favorites.discard(base_num)
                if deleted:
                    self.favorites.discard(num)
                    deleted_count += 1
            self._save_settings()
            self.refresh_table()
            logger.log(logging.INFO, f"🗑️ Deleted {deleted_count} specification(s) from database.")

    def _on_wipe_db_clicked(self):
        reply = QMessageBox.critical(
            self,
            "Confirm Wipe Database",
            "This will delete ALL specification metadata, file versions, and related mappings from 3gpp_data.db.\n\nContinue?",
            QMessageBox.Yes | QMessageBox.No,
        )
        if reply == QMessageBox.Yes:
            if self.db.wipe_database():
                self.favorites.clear()
                self._save_settings()
                self.refresh_table()
                logger.log(logging.INFO, "🧹 Specifications database wiped successfully.")
            else:
                QMessageBox.critical(self, "Wipe Failed", "Failed to wipe specifications database.")
