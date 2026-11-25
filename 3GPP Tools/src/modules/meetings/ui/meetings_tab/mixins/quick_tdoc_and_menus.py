"""Quick TDoc metadata alerts, context menus, cart hooks, and meeting CRUD dialogs."""

from pathlib import Path

from PyQt5.QtCore import QPoint, Qt
from PyQt5.QtWidgets import QDialog, QFileDialog, QMenu, QMessageBox

from core.network.session import NetworkConfigDialog
from modules.meetings.ui.meeting_details_dialogs import AddMeetingDialog, MeetingInfoDialog
from modules.meetings.ui.meetings_menu import populate_meeting_context_menu


class MeetingsTabQuickTdocMenusMixin:
    """Quick TDoc jump menu/banners, right-click table menus, and comparison cart hooks."""

    def _browse_cache_dir(self):
        directory = QFileDialog.getExistingDirectory(self, "Select Cache Directory", self.dl_dir_input.text())
        if directory:
            normalized_dir = str(Path(directory))
            self.dl_dir_input.setText(normalized_dir)
            self.settings.save_settings(normalized_dir)

    def _on_table_double_clicked(self, index):
        if not index.isValid():
            return
        row_data = self.table_model.data(index, Qt.UserRole)
        if not row_data:
            return

        status = row_data.get('tdoc_btn_status', 'na')
        if status == 'open':
            self._open_tdocs_window(row_data, row_data.get('tdoc_filepath'))
        elif status == 'get':
            self._download_and_open_tdocs(row_data, index.row())
        else:
            self.show_meeting_info(row_data)

    def refresh_quick_tdoc_lookup(self):
        """Re-evaluate the current Quick TDoc Jump after meeting metadata changes."""
        text = self.global_tdoc_input.text().strip()
        if text:
            self.search_controller.on_tdoc_input_changed(text)

    def _on_tdoc_input_context_menu(self, pos: QPoint):
        """Context menu on the Quick TDoc Jump input box with standard edit actions + 3GPP Quick Copy."""
        menu = self.global_tdoc_input.createStandardContextMenu()
        tdoc_str = self.global_tdoc_input.text().strip()
        if tdoc_str and hasattr(self, "search_controller") and self.search_controller.current_found_meeting:
            from modules.meetings.ui.tdocs_menus import populate_quick_copy_menu
            menu.addSeparator()
            copy_sub = menu.addMenu("📋 Quick Copy")
            copy_sub.setToolTipsVisible(True)
            meeting = self.search_controller.current_found_meeting
            row_data = self.search_controller._resolve_tdoc_row_data(meeting, tdoc_str)
            populate_quick_copy_menu(
                menu=copy_sub,
                tdoc_id=tdoc_str,
                row_data=row_data,
                docs_ftp_url=meeting.get("docs_folder_url") or "",
                revisions_url=meeting.get("revisions_url") or "",
                main_url=meeting.get("url_key") or "",
                parent=self.global_tdoc_input,
            )
        menu.exec_(self.global_tdoc_input.mapToGlobal(pos))

    def _set_tdoc_metadata_warning(self, meetings: list):
        """Update the inline Quick TDoc Jump metadata-gap warning."""
        if not meetings:
            self.tdoc_metadata_warning.clear()
            self.tdoc_metadata_warning.setVisible(False)
            self.btn_update_tdoc_metadata.setVisible(False)
            self.btn_update_tdoc_metadata.setEnabled(True)
            self.btn_update_tdoc_metadata.setText("Update metadata")
            self.tdoc_metadata_banner.setVisible(False)
            self.ribbon_card.updateGeometry()
            return

        refreshable = [m for m in meetings if (m.get("url_key") or "").strip()]
        unavailable = len(meetings) - len(refreshable)

        wg_name = meetings[0].get("wg_name", "")
        start_date = meetings[0].get("start_date", "")
        year = start_date[:4] if len(start_date) >= 4 else ""

        count = len(meetings)
        noun = "meeting" if count == 1 else "meetings"
        message = (
            f"⚠ No indexed TDoc range matched. {count} {wg_name} {noun}"
            f"{f' from {year}' if year else ''} have incomplete TDoc metadata."
        )
        if refreshable:
            message += " A Docs/ metadata refresh may make this TDoc discoverable."
        if unavailable:
            missing_noun = "meeting has" if unavailable == 1 else "meetings have"
            message += f" {unavailable} {missing_noun} no FTP path and cannot be refreshed automatically."

        self.tdoc_metadata_warning.setText(message)
        self.tdoc_metadata_warning.setVisible(True)
        self.btn_update_tdoc_metadata.setVisible(bool(refreshable))
        self.btn_update_tdoc_metadata.setEnabled(bool(refreshable))
        self.btn_update_tdoc_metadata.setText("Update metadata")
        self.tdoc_metadata_banner.setVisible(True)
        self.tdoc_metadata_banner.updateGeometry()
        self.ribbon_card.updateGeometry()

    def _update_incomplete_tdoc_metadata(self):
        """
        Run only scraper Phase 2 for same-WG/same-year meetings whose TDoc
        range is incomplete.
        """
        meetings = getattr(self.search_controller, "current_incomplete_meetings", [])
        refreshable = [m for m in meetings if (m.get("url_key") or "").strip()]
        if not refreshable:
            return

        targets = [
            {"wg": meeting.get("wg_name"), "meeting": meeting.get("meeting_number")}
            for meeting in refreshable
            if meeting.get("wg_name") and meeting.get("meeting_number")
        ]
        if not targets:
            return

        self.tdoc_metadata_warning.setText(
            f"⏳ Requested Docs/ metadata refresh for {len(targets)} "
            f"{'meeting' if len(targets) == 1 else 'meetings'}. "
            "Re-run the Quick TDoc Jump after synchronization completes."
        )
        self.btn_update_tdoc_metadata.setEnabled(False)
        self.tdoc_metadata_banner.updateGeometry()
        self.ribbon_card.updateGeometry()
        self.update_specific_requested.emit(targets, False, True, False)

    def _open_add_meeting_dialog(self):
        dialog = AddMeetingDialog(self.db, self)
        if dialog.exec_() == QDialog.Accepted:
            self.refresh_table()

    def _open_network_config(self):
        NetworkConfigDialog(self).exec_()

    def _confirm_delete_all(self):
        if QMessageBox.question(self, 'Confirm Clear', "Delete ALL meetings? Cannot be undone.",
                                QMessageBox.Yes | QMessageBox.No) == QMessageBox.Yes:
            self.db.delete_all_meetings()
            self.refresh_table()

    def _confirm_delete_specific(self, targets: list):
        if QMessageBox.question(self, 'Confirm', f"Delete {len(targets)} meeting(s)?",
                                QMessageBox.Yes | QMessageBox.No) == QMessageBox.Yes:
            self.db.delete_specific_meetings(targets)
            self.refresh_table()

    def _emit_multi_delete(self, selected_rows):
        targets = [{"wg": self.table_model.data(r, Qt.UserRole).get("wg_name"),
                    "meeting": self.table_model.data(r, Qt.UserRole).get("meeting_number")} for r in selected_rows if
                   self.table_model.data(r, Qt.UserRole)]
        if targets:
            self._confirm_delete_specific(targets)

    def _update_cart_ui(self, slot_a: dict, slot_b: dict):
        if hasattr(self, "cart_frame") and hasattr(self.cart_frame, "update_cart_ui"):
            self.cart_frame.update_cart_ui(slot_a, slot_b)

    def _run_comparison(self):
        if hasattr(self, "cart_frame") and hasattr(self.cart_frame, "run_comparison"):
            self.cart_frame.run_comparison()
            self.cmp_thread = self.cart_frame.cmp_thread

    def _emit_multi_sync(self, selected_rows):
        targets = [{"wg": self.table_model.data(r, Qt.UserRole).get("wg_name"),
                    "meeting": self.table_model.data(r, Qt.UserRole).get("meeting_number")} for r in selected_rows if
                   self.table_model.data(r, Qt.UserRole)]
        if targets:
            self.update_specific_requested.emit(targets, True, True, True)

    def show_right_click_menu(self, pos: QPoint):
        index = self.table.indexAt(pos)
        if index.isValid():
            row_data = self.table_model.data(index, Qt.UserRole)
            if not row_data:
                return
            menu = QMenu(self)
            self._populate_dynamic_menu(menu, row_data, index.row())
            menu.exec_(self.table.viewport().mapToGlobal(pos))

    def show_meeting_info(self, data: dict):
        MeetingInfoDialog(data, self).exec_()

    def _populate_dynamic_menu(self, menu: QMenu, row_data: dict, row_idx: int):
        populate_meeting_context_menu(self, menu, row_data, row_idx)
