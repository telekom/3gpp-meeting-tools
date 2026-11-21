"""Modal dialogs, clipboard copy, and table context menu mixin for WorkItemsTab."""

import logging
import webbrowser

from PyQt5.QtWidgets import QApplication, QDialog, QMenu, QMessageBox

from modules.work_items.ui.work_items_dialogs import (
    AddWorkItemDialog,
    LinkedSpecsDialog,
    WorkItemInfoDialog,
    copy_wis_default,
    populate_wi_batch_copy_menu,
    populate_wi_copy_menu,
)

logger = logging.getLogger("modules.work_items.ui.ui_tabs")


class WorkItemsDialogsMenusMixin:
    """Row details dialogs, manual WI entry, Quick Copy, and batch context menu."""

    def _on_row_double_clicked(self, index):
        if not index.isValid():
            return
        row_data = self.table_model.get_row_data(index.row())
        wi_code = row_data.get('code')
        if wi_code:
            self._show_wi_info(wi_code)

    def _on_specs_clicked(self, row_idx: int):
        row_data = self.table_model.get_row_data(row_idx)
        wi_code = row_data.get('code')
        acronym = row_data.get('acronym', '')
        if wi_code:
            specs = self.db.get_linked_specs_for_wi(wi_code)
            dialog = LinkedSpecsDialog(wi_code, acronym, specs, self)
            dialog.exec_()

    def _show_wi_info(self, wi_code: str):
        details = self.db.get_work_item_details(wi_code)
        if details:
            code_str = str(wi_code).strip()
            dialog = WorkItemInfoDialog(
                details,
                parent=self,
                is_starred=code_str in self.starred,
                is_following=code_str in self.following,
                on_toggle_star=self._toggle_star,
                on_toggle_follow=self._toggle_follow,
            )
            dialog.exec_()

    def _open_add_wi_dialog(self, initial_query: str = ""):
        dialog = AddWorkItemDialog(self.db, initial_query=initial_query, parent=self)
        if dialog.exec_() == QDialog.Accepted and dialog.saved_wi_code:
            code = dialog.saved_wi_code

            # Update tracking if requested
            tracking_changed = False
            if dialog.is_starred_requested and code not in self.starred:
                self.starred.add(code)
                tracking_changed = True
            if dialog.is_following_requested and code not in self.following:
                self.following.add(code)
                tracking_changed = True

            if tracking_changed:
                self.settings.save_tracking_sets(self.starred, self.following)
                self.table_model.set_tracking_sets(self.starred, self.following)
                self._update_chip_styles()

            # Refresh dropdown filter options (new releases / WGs)
            self._populate_filters()

            # Clear search and active filter constraints so the new WI is clearly displayed
            self._clear_all_filters()
            self.search_input.setText(code)
            self.refresh_table(reset_pagination=True)
            self.status_lbl.setText(f"✅ Added Work Item #{code}")
            logger.info(f"✅ Added Work Item #{code} to database.")

    def _copy_table_selection(self):
        """Copies default rich-HTML and plain-text representation for selected Work Items on Ctrl+C."""
        selected_indexes = self.table.selectionModel().selectedRows()
        if not selected_indexes:
            return
        rows_data = [self.table_model.get_row_data(idx.row()) for idx in selected_indexes]
        copy_wis_default(rows_data, self)

    def _show_context_menu(self, position):
        selected_indexes = self.table.selectionModel().selectedRows()
        if not selected_indexes:
            menu = QMenu(self)
            add_wi_action = menu.addAction("➕ Add Work Item...")
            action = menu.exec_(self.table.viewport().mapToGlobal(position))
            if action == add_wi_action:
                self._open_add_wi_dialog()
            return

        menu = QMenu(self)
        len_indexes = len(selected_indexes)
        wi_code_list = [
            str(self.table_model.get_row_data(e.row()).get("code") or "").strip()
            for e in selected_indexes
        ]
        wi_code_list = [code for code in wi_code_list if code]

        if len_indexes == 1:
            row_idx = selected_indexes[0].row()
            row_data = self.table_model.get_row_data(row_idx)
            wi_code = str(row_data.get('code') or "").strip()
            acronym = row_data.get('acronym', '')
            name = row_data.get('name', '')
            wg = row_data.get('wg_names', '')
            rel = row_data.get('release', '')

            if not wi_code:
                return

            # Tracking Actions
            is_starred = wi_code in self.starred
            is_following = wi_code in self.following

            star_label = "⭐ Remove from Starred" if is_starred else "⭐ Add to Starred (Personal Interest)"
            follow_label = "🎯 Unfollow Work Item" if is_following else "🎯 Follow Work Item (Assigned Duty)"

            star_action = menu.addAction(star_label)
            follow_action = menu.addAction(follow_label)
            menu.addSeparator()

            info_action = menu.addAction("ℹ️ View WI Details")
            wi_page_action = menu.addAction("🌐 Open WI Page (Portal)")
            local_specs_action = menu.addAction("📑 View Linked Specifications in DB")
            specs_action = menu.addAction("📂 Specifications Resulting from this WI (Web)")
            crs_action = menu.addAction("📄 CRs Related to this WI (Web)")
            menu.addSeparator()
            copy_sub = menu.addMenu("📋 Quick Copy")
            populate_wi_copy_menu(copy_sub, row_data, self)
            citation_action = None
            menu.addSeparator()
            add_wi_action = menu.addAction("➕ Add Work Item...")
            update_action = menu.addAction("🔄 Update WI")
            delete_action = menu.addAction("🗑️ Delete this Work Item")
            batch_star_action = None
            batch_unstar_action = None
            batch_follow_action = None
            batch_unfollow_action = None
        else:
            star_action = None
            follow_action = None
            info_action = None
            wi_page_action = None
            local_specs_action = None
            specs_action = None
            crs_action = None
            citation_action = None
            wi_code = None

            batch_star_action = menu.addAction(f"⭐ Star selected ({len_indexes} WIs)")
            batch_unstar_action = menu.addAction(f"☆ Unstar selected ({len_indexes} WIs)")
            batch_follow_action = menu.addAction(f"🎯 Follow selected ({len_indexes} WIs)")
            batch_unfollow_action = menu.addAction(f"🚫 Unfollow selected ({len_indexes} WIs)")
            menu.addSeparator()

            selected_rows_data = [self.table_model.get_row_data(e.row()) for e in selected_indexes]
            copy_sub = menu.addMenu(f"📋 Quick Copy ({len_indexes} WIs)")
            populate_wi_batch_copy_menu(copy_sub, selected_rows_data, self)
            menu.addSeparator()

            add_wi_action = menu.addAction("➕ Add Work Item...")
            update_action = menu.addAction(f"🔄 Update WIs ({len_indexes} WIs)")
            delete_action = menu.addAction(f"🗑️ Delete selected Work Items ({len_indexes} WIs)")

        action = menu.exec_(self.table.viewport().mapToGlobal(position))

        if len_indexes == 1:
            if action == star_action:
                self._toggle_star(wi_code)
            elif action == follow_action:
                self._toggle_follow(wi_code)
            elif action == info_action:
                self._show_wi_info(wi_code)
            elif action == wi_page_action:
                webbrowser.open(f"https://portal.3gpp.org/desktopmodules/WorkItem/WorkItemDetails.aspx?workitemId={wi_code}")
            elif action == local_specs_action:
                specs = self.db.get_linked_specs_for_wi(wi_code)
                dialog = LinkedSpecsDialog(wi_code, acronym, specs, self)
                dialog.exec_()
            elif action == specs_action:
                webbrowser.open(f"https://portal.3gpp.org/Specifications.aspx?q=1&WiUid={wi_code}")
            elif action == crs_action:
                webbrowser.open(f"https://portal.3gpp.org/ChangeRequests.aspx?q=1&workitem={wi_code}")
            elif action == citation_action:
                citation = f"WI {wi_code}: {acronym} — {name} ({rel}, {wg})"
                QApplication.clipboard().setText(citation)
            elif action == add_wi_action:
                self._open_add_wi_dialog()
            elif action == delete_action:
                confirm = QMessageBox.question(
                    self, "Confirm Deletion",
                    f"Are you sure you want to delete Work Item '{wi_code}' from the database?",
                    QMessageBox.Yes | QMessageBox.No, QMessageBox.No
                )
                if confirm == QMessageBox.Yes:
                    self.db.delete_work_item(wi_code)
                    self.starred.discard(wi_code)
                    self.following.discard(wi_code)
                    self.settings.save_tracking_sets(self.starred, self.following)
                    self.table_model.set_tracking_sets(self.starred, self.following)
                    self._update_chip_styles()
                    self.refresh_table(reset_pagination=True)
            elif action == update_action:
                self._start_targeted_sync([wi_code])
        else:
            if action == batch_star_action:
                self._batch_set_starred(wi_code_list, True)
            elif action == batch_unstar_action:
                self._batch_set_starred(wi_code_list, False)
            elif action == batch_follow_action:
                self._batch_set_following(wi_code_list, True)
            elif action == batch_unfollow_action:
                self._batch_set_following(wi_code_list, False)
            elif action == add_wi_action:
                self._open_add_wi_dialog()
            elif action == delete_action:
                confirm = QMessageBox.question(
                    self, "Confirm Batch Deletion",
                    f"Are you sure you want to delete {len(wi_code_list)} selected Work Items from the database?",
                    QMessageBox.Yes | QMessageBox.No, QMessageBox.No
                )
                if confirm == QMessageBox.Yes:
                    self.db.delete_work_items(wi_code_list)
                    for c in wi_code_list:
                        self.starred.discard(c)
                        self.following.discard(c)
                    self.settings.save_tracking_sets(self.starred, self.following)
                    self.table_model.set_tracking_sets(self.starred, self.following)
                    self._update_chip_styles()
                    self.refresh_table(reset_pagination=True)
            elif action == update_action:
                self._start_targeted_sync(wi_code_list)
