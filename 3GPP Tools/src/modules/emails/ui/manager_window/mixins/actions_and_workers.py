"""Star/follow actions, Outlook file opening, and sync/move/rescan/stats workers."""

from __future__ import annotations

import logging
import os
import platform
import subprocess
import webbrowser
from pathlib import Path

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QMessageBox

from modules.emails.core.email_threads import (
    EmailMoveThread,
    EmailSyncThread,
    EmailTargetRescanThread,
)
from modules.emails.core.stats.statistics_threads import EmailStatsExporterThread


class EmailManagerActionsAndWorkersMixin:
    """User actions and background QThread orchestration for EmailManagerWindow."""

    def _toggle_current_star(self):
        if not getattr(self, "current_tdoc_id", ""):
            return
        is_starred = self.current_tdoc_id in self.email_model.starred_tdocs
        new_status = not is_starred

        logging.info(f"[EmailWindow] Toggling Star for {self.current_tdoc_id} -> {new_status}")
        self.db.toggle_tdoc_star(self.current_tdoc_id, new_status)

        if new_status:
            self.email_model.starred_tdocs.add(self.current_tdoc_id)
            self.tdoc_model.starred_tdocs.add(self.current_tdoc_id)
        else:
            self.email_model.starred_tdocs.discard(self.current_tdoc_id)
            self.tdoc_model.starred_tdocs.discard(self.current_tdoc_id)

        self.email_model.dataChanged.emit(
            self.email_model.index(0, 0),
            self.email_model.index(self.email_model.rowCount() - 1, self.email_model.columnCount() - 1),
        )
        self.tdoc_model.dataChanged.emit(
            self.tdoc_model.index(0, 0),
            self.tdoc_model.index(self.tdoc_model.rowCount() - 1, self.tdoc_model.columnCount() - 1),
        )
        self.tdoc_proxy.invalidateFilter()

        self.btn_toggle_star.setText(
            f"❌ Unstar {self.current_tdoc_id}" if new_status else f"⭐ Star {self.current_tdoc_id}"
        )

    def _toggle_current_ai_follow(self):
        if not getattr(self, "current_agenda_item", ""):
            return
        is_followed = self.current_agenda_item in self.email_model.followed_ais
        new_status = not is_followed

        logging.info(f"[EmailWindow] Toggling AI Follow for {self.current_agenda_item} -> {new_status}")
        self.db.toggle_ai_follow(self.current_agenda_item, new_status)

        if new_status:
            self.email_model.followed_ais.add(self.current_agenda_item)
            self.tdoc_model.followed_ais.add(self.current_agenda_item)
        else:
            self.email_model.followed_ais.discard(self.current_agenda_item)
            self.tdoc_model.followed_ais.discard(self.current_agenda_item)

        self.email_model.dataChanged.emit(
            self.email_model.index(0, 0),
            self.email_model.index(self.email_model.rowCount() - 1, self.email_model.columnCount() - 1),
        )
        self.tdoc_model.dataChanged.emit(
            self.tdoc_model.index(0, 0),
            self.tdoc_model.index(self.tdoc_model.rowCount() - 1, self.tdoc_model.columnCount() - 1),
        )
        self.tdoc_proxy.invalidateFilter()

        ai_display = self.current_agenda_item or "Unknown AI"
        self.btn_toggle_follow.setText(f"❌ Unfollow {ai_display}" if new_status else f"★ Follow {ai_display}")

    def _open_current_msg(self):
        if self.current_msg_path and Path(self.current_msg_path).exists():
            try:
                os.startfile(self.current_msg_path)
            except Exception as e:
                QMessageBox.warning(self, "Error", f"Could not open file: {e}")

    def _open_current_msg_folder(self):
        if getattr(self, "current_msg_path", "") and Path(self.current_msg_path).exists():
            try:
                if platform.system() == "Windows":
                    windows_path = os.path.normpath(self.current_msg_path)
                    subprocess.run(["explorer", "/select,", windows_path])
                else:
                    os.startfile(Path(self.current_msg_path).parent)
            except Exception as e:
                QMessageBox.warning(self, "Error", f"Could not open folder: {e}")

    def _run_sync(self):
        if not self.source_folder:
            QMessageBox.warning(self, "Setup Required", "Please configure the Outlook Source Folder first.")
            return
        self._set_buttons_enabled(False)
        self.btn_sync.setText("⏳ Syncing...")
        sd = self.dt_start.date().toString(Qt.ISODate)
        ed = self.dt_end.date().toString(Qt.ISODate)
        self.sync_thread = EmailSyncThread(self.source_folder, self.meeting_dir, self.ai_lookup, self.db, sd, ed)
        self.sync_thread.log_msg.connect(lambda m, _: self.lbl_status.setText(m))
        self.sync_thread.progress_update.connect(
            lambda c, t: self.lbl_status.setText(f"Scanning Source: {c} / {t} items...")
        )
        self.sync_thread.finished.connect(self._on_sync_finished)
        self.sync_thread.start()

    def _on_sync_finished(self, success: bool, msg: str):
        self._set_buttons_enabled(True)
        self.btn_sync.setText("🔄 Sync Source")
        self.lbl_status.setText(msg)
        self._refresh_table()

    def _run_move(self):
        if not self.target_folder:
            QMessageBox.warning(self, "Setup Required", "Please configure the Outlook Target Folder first.")
            return
        indexes = self.email_view.selectionModel().selectedRows()
        if not indexes:
            QMessageBox.information(self, "Selection Required", "Please select at least one email to move.")
            return
        items_to_move = []
        for idx in indexes:
            source_idx = self.email_proxy.mapToSource(idx)
            row_data = self.email_model.get_row_data(source_idx.row())
            if row_data.get("outlook_location") == "Source":
                items_to_move.append((row_data.get("id"), row_data.get("agenda_item")))
        if not items_to_move:
            QMessageBox.information(self, "Notice", "Selected emails have already been moved to the Target.")
            return
        self._set_buttons_enabled(False)
        self.btn_move.setText("⏳ Moving...")
        self.move_thread = EmailMoveThread(items_to_move, self.target_folder, self.db)
        self.move_thread.progress_update.connect(lambda c, t: self.lbl_status.setText(f"Moving {c}/{t}..."))
        self.move_thread.finished.connect(self._on_move_finished)
        self.move_thread.start()

    def _run_move_all(self):
        if not self.target_folder:
            QMessageBox.warning(self, "Setup Required", "Please configure the Outlook Target Folder first.")
            return
        items_to_move = []
        for row_data in self.email_model._data:
            if row_data.get("outlook_location") == "Source":
                items_to_move.append((row_data.get("id"), row_data.get("agenda_item")))
        if not items_to_move:
            QMessageBox.information(self, "Notice", "There are no emails in the Source folder to move.")
            return
        reply = QMessageBox.question(
            self,
            "Confirm Move All",
            f"Move ALL {len(items_to_move)} source emails to the target?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if reply == QMessageBox.No:
            return
        self._set_buttons_enabled(False)
        self.btn_move_all.setText("⏳ Moving All...")
        self.move_thread = EmailMoveThread(items_to_move, self.target_folder, self.db)
        self.move_thread.progress_update.connect(lambda c, t: self.lbl_status.setText(f"Moving {c}/{t}..."))
        self.move_thread.finished.connect(self._on_move_all_finished)
        self.move_thread.start()

    def _on_move_all_finished(self, success: bool, msg: str):
        self._set_buttons_enabled(True)
        self.btn_move_all.setText("⏭️ Move All")
        self.lbl_status.setText(msg)
        self._refresh_table()

    def _on_move_finished(self, success: bool, msg: str):
        self._set_buttons_enabled(True)
        self.btn_move.setText("➡️ Move Selected")
        self.lbl_status.setText(msg)
        self._refresh_table()

    def _run_target_rescan(self):
        if not self.target_folder:
            QMessageBox.warning(self, "Setup Required", "Please configure the Outlook Target Folder first.")
            return
        self._set_buttons_enabled(False)
        self.btn_rescan.setText("⏳ Scanning...")
        sd = self.dt_start.date().toString(Qt.ISODate)
        ed = self.dt_end.date().toString(Qt.ISODate)
        self.rescan_thread = EmailTargetRescanThread(
            self.target_folder, self.meeting_dir, self.ai_lookup, self.db, sd, ed
        )
        self.rescan_thread.log_msg.connect(lambda m, _: self.lbl_status.setText(m))
        self.rescan_thread.progress_update.connect(
            lambda c, t: self.lbl_status.setText(f"Scanning Target: {c} / {t} items...")
        )
        self.rescan_thread.finished.connect(self._on_rescan_finished)
        self.rescan_thread.start()

    def _on_rescan_finished(self, success: bool, msg: str):
        self._set_buttons_enabled(True)
        self.btn_rescan.setText("🔁 Scan Target")
        self.lbl_status.setText(msg)
        self._refresh_table()

    def _generate_statistics(self):
        all_emails = self.email_model._data
        if not all_emails:
            QMessageBox.warning(self, "No Data", "There are no emails loaded to analyze.")
            return

        self._set_buttons_enabled(False)
        self.btn_stats.setText("⏳ Generating...")

        current_stats = getattr(self, "stats_config", self._load_config().get("stats_config", {}))
        meeting_name = self.meeting_dir.name if self.meeting_dir else "Meeting"

        self.stats_thread = EmailStatsExporterThread(self.meeting_dir, all_emails, meeting_name, current_stats)
        self.stats_thread.finished.connect(self._on_stats_finished)
        self.stats_thread.start()

    def _on_stats_finished(self, success: bool, msg: str):
        self._set_buttons_enabled(True)
        self.btn_stats.setText("📊 Statistics")
        if success:
            self.lbl_status.setText("✅ Analytics Report Generated.")
            if hasattr(os, "startfile"):
                os.startfile(msg)
            else:
                webbrowser.open(f"file:///{msg}")
        else:
            self.lbl_status.setText("❌ Analytics Generation Failed.")
            QMessageBox.warning(self, "Error", f"Could not generate statistics:\n{msg}")

    def _set_buttons_enabled(self, enabled: bool):
        self.toolbar.set_buttons_enabled(enabled)
