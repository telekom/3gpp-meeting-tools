"""Table refresh, filtering, selection, reader pane, and click/menu handlers."""

from __future__ import annotations

import logging
import re
import sqlite3
import webbrowser

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QApplication, QDialog, QMenu

from core.ui.theme import C
from modules.emails.ui.email_dialogs import EditEmailDialog


class EmailManagerFilteringAndSelectionMixin:
    """Handles SQLite reloads, dropdown filters, selection changes, and table clicks."""

    def _refresh_table(self):
        logging.info("[EmailWindow] Executing full _refresh_table from database.")
        old_tdoc = self.email_proxy.target_tdoc

        old_email_id = None
        indexes = self.email_view.selectionModel().selectedRows()
        if indexes:
            source_idx = self.email_proxy.mapToSource(indexes[0])
            old_email_id = self.email_model.get_row_data(source_idx.row()).get("id")

        with self.db._get_connection() as conn:
            conn.row_factory = sqlite3.Row
            data = [dict(row) for row in conn.execute("SELECT * FROM emails ORDER BY date_received DESC").fetchall()]

        starred_tdocs = self.db.get_starred_tdocs()
        followed_ais = self.db.get_followed_ais()

        self.tdoc_model.update_data(data, starred_tdocs, followed_ais)
        self.email_model.update_data(data, starred_tdocs, followed_ais)

        self._refresh_comboboxes(data)

        if old_tdoc:
            for r in range(self.tdoc_proxy.rowCount()):
                idx = self.tdoc_proxy.index(r, 0)
                if self.tdoc_proxy.data(idx, Qt.UserRole) == old_tdoc:
                    self.tdoc_view.selectRow(r)
                    break

        if old_email_id:
            for r in range(self.email_proxy.rowCount()):
                idx = self.email_proxy.index(r, 0)
                src_idx = self.email_proxy.mapToSource(idx)
                if self.email_model.get_row_data(src_idx.row()).get("id") == old_email_id:
                    self.email_view.selectRow(r)
                    break

    def _refresh_comboboxes(self, data):
        logging.info("[EmailWindow] Updating Dropdown Filter Lists...")

        def clean(val):
            return str(val).strip() if val is not None else ""

        def natural_sort_key(s):
            return [int(text) if text.isdigit() else text.lower() for text in re.split("([0-9]+)", str(s))]

        unique_ais = sorted([ai for ai in set(clean(r.get("agenda_item")) for r in data) if ai], key=natural_sort_key)
        unique_companies = sorted([c for c in set(clean(r.get("company")) for r in data) if c])
        unique_senders = sorted([s for s in set(clean(r.get("sender_name")) for r in data) if s])

        self.cb_ai.updateItems(unique_ais)
        self.cb_company.updateItems(unique_companies)
        self.cb_sender.updateItems(unique_senders)

        logging.info(f"[EmailWindow] Dropdowns Ready: {len(unique_ais)} AIs, {len(unique_companies)} Companies.")
        self._apply_filters()

    def _on_date_range_changed(self, start_date: str, end_date: str):
        self.start_date = start_date
        self.end_date = end_date
        self._save_config(self.source_folder, self.target_folder, self.start_date, self.end_date)

    def _apply_filters(self, *args):
        self.tdoc_proxy.set_filters(
            self.btn_filter_star.isChecked(),
            self.btn_filter_follow.isChecked(),
            self.cb_ai.getCheckedItems(),
            self.tdoc_search_input.text(),
        )
        self.email_proxy.set_filters(
            self.search_input.text(),
            self.cb_company.getCheckedItems(),
            self.cb_sender.getCheckedItems(),
        )
        self._update_count_label()

    def _on_tdoc_selected(self, selected, deselected):
        indexes = self.tdoc_view.selectionModel().selectedRows()
        if not indexes:
            self.email_proxy.set_target_tdoc(None)
            self.lbl_right_title.setText("📧 Thread History (Select a TDoc from the left)")
            return

        proxy_index = indexes[0]
        tdoc_id = self.tdoc_proxy.data(proxy_index, Qt.UserRole)

        self.email_proxy.set_target_tdoc(tdoc_id)

        count = self.email_proxy.rowCount()
        self.lbl_right_title.setText(f"📧 Thread History: {tdoc_id} ({count} Emails)")
        self._update_count_label()

    def _on_email_selected(self, selected, deselected):
        indexes = self.email_view.selectionModel().selectedRows()
        if not indexes:
            self.reading_pane.clear()
            self.btn_open_msg.setEnabled(False)
            self.btn_open_folder.setEnabled(False)
            self.btn_toggle_star.setEnabled(False)
            self.btn_toggle_follow.setEnabled(False)
            return

        source_idx = self.email_proxy.mapToSource(indexes[0])
        row_data = self.email_model.get_row_data(source_idx.row())

        self.current_tdoc_id = row_data.get("tdoc_id", "")
        self.current_agenda_item = row_data.get("agenda_item", "")
        self.current_msg_path = row_data.get("msg_path", "")

        self.btn_open_msg.setEnabled(bool(self.current_msg_path))
        self.btn_open_folder.setEnabled(bool(self.current_msg_path))
        self.btn_toggle_star.setEnabled(bool(self.current_tdoc_id))
        self.btn_toggle_follow.setEnabled(bool(self.current_agenda_item))

        is_starred = self.current_tdoc_id in self.email_model.starred_tdocs
        self.btn_toggle_star.setText(
            f"❌ Unstar {self.current_tdoc_id}" if is_starred else f"⭐ Star {self.current_tdoc_id}"
        )

        is_followed = self.current_agenda_item in self.email_model.followed_ais
        ai_display = self.current_agenda_item or "Unknown AI"
        self.btn_toggle_follow.setText(f"❌ Unfollow {ai_display}" if is_followed else f"★ Follow {ai_display}")

        real_email = row_data.get("sender_email", "")
        html_body = f"""
        <h2 style='color:{C.OFFICE_BLUE}; margin-bottom: 2px;'>{row_data.get('subject', '')}</h2>
        <p style='color:#555; margin-top:0px;'><b>From:</b> {row_data.get('sender_name')} &lt;{real_email}&gt; ({row_data.get('company')}) | <b>TDoc:</b> {row_data.get('tdoc_id')} | <b>Rev:</b> {row_data.get('revisions_mentioned')}</p>
        <hr>
        <h3 style='color:{C.ORANGE_DARK};'>Short Text / Comments</h3>
        <p style='background-color:{C.WG_CT_BG}; padding:10px; border-left: 4px solid {C.ORANGE_BORDER};'>{row_data.get('short_text', '').replace(chr(10), '<br>')}</p>
        <h3>Free Text</h3>
        <p style='color:#333;'>{row_data.get('free_text', '').replace(chr(10), '<br>')}</p>
        """
        self.reading_pane.setHtml(html_body)

    def _update_count_label(self):
        self.lbl_count.setText(
            f"Showing {self.tdoc_proxy.rowCount()} Threads | {self.email_proxy.rowCount()} Emails"
        )

    def _on_table_clicked(self, index):
        if not index.isValid():
            return
        source_idx = self.email_proxy.mapToSource(index)
        col_name = self.email_model._headers[index.column()]
        row_data = self.email_model.get_row_data(source_idx.row())

        if col_name == "Sender":
            email = row_data.get("sender_email", "")
            if email:
                webbrowser.open(f"mailto:{email}")
        elif col_name == "Rev":
            revs = row_data.get("revisions_mentioned", "")
            if revs:
                first_rev = revs.split(",")[0].strip()
                self.tdoc_open_requested.emit(first_rev)
                self.lbl_status.setText(f"Requesting open for: {first_rev}")

    def _on_tdoc_double_clicked(self, index):
        if not index.isValid():
            return
        tdoc_id = self.tdoc_proxy.data(index, Qt.UserRole)
        if tdoc_id and tdoc_id != "General / Unlinked":
            self.lbl_status.setText(f"Jumping to TDoc: {tdoc_id}")
            self.tdoc_jump_requested.emit(tdoc_id)

    def _on_email_double_clicked(self, index):
        if not index.isValid():
            return
        source_idx = self.email_proxy.mapToSource(index)
        row_data = self.email_model.get_row_data(source_idx.row())
        tdoc_id = row_data.get("tdoc_id", "")
        if tdoc_id:
            self.lbl_status.setText(f"Jumping to TDoc: {tdoc_id}")
            self.tdoc_jump_requested.emit(tdoc_id)

    def _show_context_menu(self, pos):
        index = self.email_view.indexAt(pos)
        if not index.isValid():
            return

        source_idx = self.email_proxy.mapToSource(index)
        row_data = self.email_model.get_row_data(source_idx.row())
        col_name = self.email_model._headers[index.column()]

        menu = QMenu(self)
        edit_action = menu.addAction("✏️ Edit Email Details")
        copy_action = menu.addAction("📋 Copy Email Address") if col_name == "Sender" else None

        action = menu.exec_(self.email_view.viewport().mapToGlobal(pos))

        if action == edit_action:
            dialog = EditEmailDialog(row_data, self)
            if dialog.exec_() == QDialog.Accepted:
                updated_data = dialog.get_data()
                self.db.update_email(row_data["id"], updated_data)
                self.lbl_status.setText("✅ Updated email details.")
                self._refresh_table()
        elif copy_action and action == copy_action:
            email = row_data.get("sender_email", "")
            if email:
                QApplication.clipboard().setText(email)
                self.lbl_status.setText(f"Copied {email} to clipboard.")
