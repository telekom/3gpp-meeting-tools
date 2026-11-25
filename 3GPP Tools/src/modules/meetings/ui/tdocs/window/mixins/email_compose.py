# --- File: src/modules/meetings/ui/tdocs/window/mixins/email_compose.py ---
"""
E-mail: compose/send TDocs via Outlook, email manager, per-TDoc emails, general email sync.

Mixin of TDocsWindow: relies on attributes set in TDocsWindow.__init__ / _setup_*.
"""
import logging
import os
import re
import urllib.parse
import webbrowser
from pathlib import Path
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QCursor
from PyQt5.QtWidgets import (
    QMessageBox,
    QApplication,
    QToolTip,
    QDialog,
)
from modules.meetings.core.tdocs.tdoc_file_handler import TDocFileHandler
from modules.emails.core.general_email_sync import GeneralEmailSyncThread
from modules.emails.core.outlook_client import OutlookClient
from modules.emails.core.tracked_threads_service import pin_thread_by_subject_or_cid
from modules.emails.ui.general_email_dialog import (
    GeneralEmailFoldersDialog,
    GeneralEmailSyncDialog,
    TDocEmailsDialog,
    load_wg_email_config,
)
from core.utils.folder_opener import open_folder as _open_folder


class EmailMixin:
    """E-mail: compose/send TDocs via Outlook, email manager, per-TDoc emails, general email sync."""

    def _build_email_subject(self, tdoc_id: str) -> str:
        clean_tdoc = str(tdoc_id or "").strip()
        row_data = next(
            (r for r in self.model._data if str(r.get("TDoc", "")).strip().upper() == clean_tdoc.upper()),
            None,
        )

        if not row_data:
            base_tdoc = self._extract_base_tdoc(clean_tdoc)
            row_data = next(
                (r for r in self.model._data if str(r.get("TDoc", "")).strip().upper() == base_tdoc),
                {},
            )

        wg = str(self.mtg_info.get("wg_name", "")).strip().upper()
        mtg_num = str(self.mtg_info.get("meeting_number", "")).strip()
        ai = str(row_data.get("Agenda Item", "")).strip()
        title = str(row_data.get("Title", "")).strip()

        tag_parts = []
        if wg and mtg_num:
            tag_parts.append(f"{wg}#{mtg_num}")
        elif wg or mtg_num:
            tag_parts.append(f"{wg}{mtg_num}")

        if ai and ai.upper() not in ["N/A", "NONE", "UNKNOWN", ""]:
            tag_parts.append(ai)

        tag_str = f"[{', '.join(tag_parts)}] " if tag_parts else ""
        return f"{tag_str}{clean_tdoc} {title}".strip()

    def _compose_email_draft(self, tdoc_id: str):
        subject = self._build_email_subject(tdoc_id)
        encoded_subject = urllib.parse.quote(subject, safe='')
        mailto_url = f"mailto:?subject={encoded_subject}"

        try:
            if hasattr(os, 'startfile'):
                os.startfile(mailto_url)
            else:
                webbrowser.open(mailto_url)
            logging.info(f"📧 [Email Draft] Opened draft with subject: '{subject}'")
        except Exception as e:
            logging.warning(f"Failed to launch email client automatically: {e}")
            QApplication.clipboard().setText(subject)
            QToolTip.showText(QCursor.pos(), "📋 Subject copied to clipboard!", self)

    def _send_tdoc_via_email(self, target_tdoc: str):
        """Drafts an email with the TDoc subject and attaches the TDoc document(s), downloading first if needed."""
        clean_target = str(target_tdoc or "").strip()
        if not clean_target:
            return

        local_files = TDocFileHandler.find_local_tdoc_files(self.meeting_dir, clean_target)
        if local_files:
            self._open_email_draft_with_attachments(clean_target, local_files)
            return

        base_tdoc = self._extract_base_tdoc(clean_target)
        if base_tdoc in self.model.loading_tdocs:
            return
        self._trigger_download_thread(
            base_tdoc,
            clean_target,
            send_email_on_finish=True,
        )

    def _open_email_draft_with_attachments(self, target_tdoc: str, attachments: list):
        subject = self._build_email_subject(target_tdoc)
        valid_paths = [Path(p) for p in (attachments or []) if p and Path(p).is_file()]
        if not valid_paths:
            QMessageBox.warning(
                self,
                "Attachment Missing",
                f"Could not locate a local document file to attach for {target_tdoc}.",
            )
            return

        ok, err = OutlookClient.create_email_draft(subject, valid_paths)
        if ok:
            logging.info(
                f"📧 [Send TDoc] Opened Outlook draft for '{target_tdoc}' with {len(valid_paths)} attachment(s)."
            )
            return

        # Graceful fallback when Classic Outlook COM is unavailable
        QApplication.clipboard().setText(subject)
        encoded_subject = urllib.parse.quote(subject, safe='')
        mailto_url = f"mailto:?subject={encoded_subject}"
        try:
            if hasattr(os, 'startfile'):
                os.startfile(mailto_url)
            else:
                webbrowser.open(mailto_url)
        except Exception:
            pass

        try:
            _open_folder(valid_paths[0].parent)
        except Exception:
            pass

        QMessageBox.information(
            self,
            "Attach TDoc Manually",
            f"Could not attach files automatically via Outlook COM ({err}).\n\n"
            f"The email subject has been copied to your clipboard and the local folder for {target_tdoc} was opened so you can drag-and-drop the file(s).",
        )

    def _open_email_manager(self):
        """Legacy entry point redirected to the unified Meeting Email Inbox cockpit."""
        self._open_meeting_email_inbox()

    def _on_generate_meeting_email_statistics(self):
        """Generates the generic HTML Email Analytics Dashboard from the unified email databases."""
        from modules.emails.core.stats.dataset_builder import build_unified_email_stats_dataset
        from modules.emails.core.stats.statistics_threads import EmailStatsExporterThread
        from modules.meetings.ui.tdocs_dialogs import StatisticsSettingsDialog

        tdoc_ai_lookup = {
            str(r.get("TDoc", "")).strip().upper(): str(r.get("Agenda Item", "")).strip()
            for r in (getattr(self.model, "_data", None) or [])
            if r.get("TDoc") and not r.get("_is_tracked_thread")
        }
        dataset = build_unified_email_stats_dataset(
            general_email_db=getattr(self, "general_email_db", None),
            tracked_db=getattr(self, "tracked_threads_db", None),
            meeting_id=self._get_meeting_id_label(),
            tdoc_ai_lookup=tdoc_ai_lookup,
            tdoc_meta_map=self._build_tdoc_meta_map(),
        )
        if not dataset:
            QMessageBox.information(
                self,
                "No Emails Indexed",
                "There are no synced emails or tracked threads for this meeting yet.\n\n"
                "Sync Outlook folders or pin threads first.",
            )
            return
        cfg = StatisticsSettingsDialog().load_config()
        meeting_name = self._get_meeting_id_label() or self.meeting_dir.name
        self._email_stats_thread = EmailStatsExporterThread(self.meeting_dir, dataset, meeting_name, cfg)
        self._email_stats_thread.finished.connect(self._on_email_stats_export_finished)
        self._email_stats_thread.start()

    def _on_email_stats_export_finished(self, success: bool, msg: str):
        if success:
            if hasattr(os, "startfile"):
                os.startfile(msg)
            else:
                webbrowser.open(f"file:///{msg}")
        else:
            QMessageBox.warning(self, "Email Statistics Failed", f"Could not generate email statistics:\n{msg}")

    def _refresh_email_counts(self):
        try:
            counts = self.general_email_db.get_email_counts_per_tdoc()
            self.model.set_email_counts(counts)
            self._sync_tracked_threads_to_table()
            if self.sender() is not getattr(self, "_inbox_dialog", None):
                self._refresh_inbox_dialog_if_open()
        except Exception as e:
            logging.error(f"Error refreshing email counts: {e}")

    def _open_tdoc_emails(self, tdoc_id: str):
        if not tdoc_id:
            return
        tdoc_clean = tdoc_id.strip().upper()
        match = re.search(r'^(.*?)-?(?:r|rev)\d{1,2}[a-zA-Z]?$', tdoc_clean, re.IGNORECASE)
        base_tdoc = match.group(1).upper() if match else tdoc_clean

        if base_tdoc in self._email_dialogs:
            dlg = self._email_dialogs[base_tdoc]
            dlg.setWindowState(dlg.windowState() & ~Qt.WindowMinimized | Qt.WindowActive)
            dlg.raise_()
            dlg.activateWindow()
            return

        family = self.model.get_family_tdocs(base_tdoc)
        wg = self.mtg_info.get("wg_name", "SA2")

        # Save callback supporting partial note appends without forcing a status overwrite
        def _save_notes_proxy(target_id: str, new_status: str, new_notes: str):
            existing_row = next((r for r in self.model._data if str(r.get("TDoc", "")).strip().upper() == target_id.upper()), {})
            status_to_use = new_status if new_status else existing_row.get("My Status", "⚪ Neutral")
            curr_notes = str(existing_row.get("My Notes", "")).strip()
            merged_notes = f"{curr_notes}\n\n{new_notes}" if curr_notes else new_notes
            self._save_user_data(target_id, status_to_use, merged_notes)

        def _pin_discussion_as_thread(info: dict):
            if not hasattr(self, "tracked_threads_db"):
                return
            agenda_map, valid_wis = self._get_pin_meeting_context()
            subj = str(info.get("subject") or f"Discussion on {base_tdoc}").strip()
            norm = str(info.get("normalized_subject") or subj).strip()
            tid = pin_thread_by_subject_or_cid(
                self.tracked_threads_db,
                title=norm or subj,
                normalized_subject=norm,
                meeting_id=self._get_meeting_id_label(),
                wg=wg,
                linked_tdocs=list(info.get("linked_tdocs") or family),
                agenda_map=agenda_map,
                valid_wis=valid_wis,
                tdoc_meta_map=self._build_tdoc_meta_map(),
            )
            self._on_tracked_thread_changed()
            self._open_tracked_thread_dialog(tid)

        base_row = next(
            (r for r in self.model._data if str(r.get("TDoc", "")).strip().upper() == base_tdoc),
            {},
        )
        tdoc_info = {
            "Title": base_row.get("Title", ""),
            "Source": base_row.get("Source", ""),
            "Agenda Item": base_row.get("Agenda Item", ""),
            "Sub-Item": base_row.get("Sub-Item", ""),
            "Related WIs": base_row.get("Related WIs", ""),
            "TDoc Status": base_row.get("TDoc Status", ""),
            "My Status": base_row.get("My Status", "⚪ Neutral"),
            "My Notes": base_row.get("My Notes", ""),
        }

        dialog = TDocEmailsDialog(
            base_tdoc,
            family,
            self.general_email_db.db_path,
            wg=wg,
            save_callback=_save_notes_proxy,
            parent=None,
            pin_thread_callback=_pin_discussion_as_thread,
            tdoc_info=tdoc_info,
            open_tdoc_callback=getattr(self, "_open_tdoc_document_directly", None),
            user_data_callback=lambda tid, st, notes: self._save_user_data(tid, st, notes),
        )
        dialog.data_changed.connect(self._refresh_email_counts)
        dialog.tdoc_selected.connect(self._open_tdoc_emails)

        self._email_dialogs[base_tdoc] = dialog
        dialog.finished.connect(lambda _, t=base_tdoc: self._email_dialogs.pop(t, None))
        dialog.show()

    def _on_configure_email_folders(self):
        wg = self.mtg_info.get("wg_name", "SA2")
        dialog = GeneralEmailFoldersDialog(wg, self)
        dialog.exec_()

    def _on_sync_general_emails(self):
        wg = self.mtg_info.get("wg_name", "SA2")
        folders = load_wg_email_config(wg)
        if not folders:
            if QMessageBox.question(self, "No Folders Configured",
                                    f"No Outlook folders configured for {wg}.\nConfigure folders now?",
                                    QMessageBox.Yes | QMessageBox.No) == QMessageBox.Yes:
                self._on_configure_email_folders()
            return

        dialog = GeneralEmailSyncDialog(wg, self.mtg_info.get("start_date", ""), self.mtg_info.get("end_date", ""), self)
        if dialog.exec_() != QDialog.Accepted:
            return

        params = dialog.get_params()
        if hasattr(self, "_start_unified_email_sync"):
            self._start_unified_email_sync(
                sync_mode="full",
                start_date=params["start_date"],
                end_date=params["end_date"],
                days_buffer=params["buffer"],
                silent=False,
            )
            return

        self.email_btn.setEnabled(False)
        self.email_btn.setText("⏳ Syncing...")

        self.general_email_sync_thread = GeneralEmailSyncThread(
            folders, self.general_email_db.db_path,
            start_date=params["start_date"], end_date=params["end_date"], days_buffer=params["buffer"]
        )
        self.general_email_sync_thread.progress_msg.connect(lambda m: self.email_btn.setText(f"⏳ {m[:20]}"))
        self.general_email_sync_thread.finished.connect(self._on_general_sync_finished)
        self.general_email_sync_thread.start()

    def _on_general_sync_finished(self, success: bool, msg: str, count: int):
        self.email_btn.setEnabled(True)
        self.email_btn.setText("📧 Emails ▾")
        self._refresh_email_counts()
        if success:
            QMessageBox.information(self, "Email Sync Complete", msg)
        else:
            QMessageBox.warning(self, "Email Sync Failed", msg)

    def _on_mark_all_emails_read(self):
        from modules.emails.core.tracked_threads_service import mark_thread_read_with_outlook
        synced_count = self.general_email_db.mark_all_read(sync_outlook=True)
        if hasattr(self, "tracked_threads_db"):
            for t in self.tracked_threads_db.get_threads_for_meeting(self._get_meeting_id_label()):
                synced_count += mark_thread_read_with_outlook(self.tracked_threads_db, t["thread_id"], sync_outlook=True)
        self._refresh_email_counts()
        self._refresh_comboboxes()
        QMessageBox.information(
            self,
            "Mark All Read",
            f"All meeting emails and tracked threads have been marked as read (synced {synced_count} Outlook item(s)).",
        )

    def _on_wipe_generic_emails(self):
        reply = QMessageBox.question(
            self,
            "Confirm Wipe Generic Emails",
            "Are you sure you want to completely wipe all generic email records for this meeting?\n\n"
            "• SA2 eMeeting tables will be preserved.\n"
            "• You can re-sync from Outlook anytime.",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        if reply == QMessageBox.Yes:
            self.general_email_db.wipe_generic_emails()
            self._refresh_email_counts()
            QMessageBox.information(self, "Database Wiped", "Generic email records have been wiped.")
