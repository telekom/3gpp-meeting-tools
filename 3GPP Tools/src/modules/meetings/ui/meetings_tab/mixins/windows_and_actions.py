"""TDocsWindow launcher, inline downloaders, global actions, and report workflows."""

import logging
import re
from pathlib import Path

from PyQt5.QtWidgets import QMessageBox

from core.ui.toast import notify_success
from modules.meetings.core.agenda_manager import AgendaDownloaderThread
from modules.meetings.core.tdocs.tdocs_downloader import TDocsDownloaderThread
from modules.meetings.core.tdocs.tdocs_parser import TDocsParser
from modules.meetings.ui.meetings_actions import (
    export_merged_tdocs,
    on_merger_finished,
    on_tdocs_caching_finished,
    open_agreement_tracker,
    open_contribution_report,
    open_last_meeting,
    open_meeting_from_report,
    open_topic_dossier_studio,
    start_tdocs_caching,
)
from modules.meetings.ui.tdocs_window import TDocsWindow


class MeetingsTabWindowsActionsMixin:
    """Opens TDocsWindow instances, manages inline FTP downloads, and dispatches reports."""

    def _get_tdoc_list_path(self, row_data: dict) -> Path:
        mtg_id = row_data.get("mtg_id")
        if not mtg_id:
            return None

        current_cache = self.dl_dir_input.text().strip() if hasattr(self, 'dl_dir_input') else self.settings.cache_dir
        folder_name = row_data.get("folder_name") or row_data.get("meeting_number", "")
        agenda_dir = Path(current_cache) / folder_name / "Agenda"

        if agenda_dir.exists() and agenda_dir.is_dir():
            for file_path in agenda_dir.iterdir():
                filename = file_path.name.lower()
                if (filename.startswith("tdoc_list_meeting_") or filename.startswith(
                        "tdocs_list_")) and filename.endswith(".xlsx"):
                    return file_path

        return agenda_dir / f"TDoc_List_Meeting_{mtg_id}.xlsx"

    def _open_tdocs_window(self, mtg_info: dict, filepath: str, quiet: bool = False):
        self.settings.save_last_meeting(mtg_info)
        self._update_last_meeting_btn()
        mtg_id = mtg_info.get("mtg_id")

        mtg_info["is_active_sync"] = self.db.is_active_sync_meeting(
            mtg_info.get("wg_name", ""),
            mtg_info.get("start_date", ""),
            mtg_info.get("end_date", ""),
            mtg_info.get("is_electronic", 0)
        )

        if mtg_id in self.tdoc_windows:
            existing_win = self.tdoc_windows[mtg_id]
            if existing_win and existing_win.isVisible():
                existing_win.raise_()
                existing_win.activateWindow()
                return
            else:
                existing_win.close()
                existing_win.deleteLater()
                del self.tdoc_windows[mtg_id]

        if not filepath or not Path(filepath).exists():
            logging.error(f"[Meetings] TDoc list for {mtg_id} not found at '{filepath}'.")
            if not quiet:
                QMessageBox.warning(self, "TDoc List Missing", f"The TDoc list file was not found:\n{filepath}")
            return

        tdocs_data = TDocsParser.parse_tdocs_excel(filepath)
        if not tdocs_data:
            logging.warning(f"[Meetings] No TDocs could be read from '{filepath}' (meeting {mtg_id}).")
            if quiet:
                return
            reply = QMessageBox.question(
                self, "No TDocs Available",
                "The cached TDoc list for this meeting contains no TDocs.\n\n"
                "This usually means it was downloaded before any TDocs were allocated, "
                "or the file is unreadable (see the console log for details).\n\n"
                "Re-download the latest TDoc list from the 3GPP portal now?",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
            if reply == QMessageBox.Yes:
                self._download_and_open_tdocs(mtg_info, -1)
            return

        window = TDocsWindow(mtg_info, tdocs_data, filepath)
        window.global_action_requested.connect(self._handle_global_action_from_window)
        self.tdoc_windows[mtg_id] = window
        window.show()

    def _on_inline_download_finished(self, success: bool, result: str, mtg_id: str, row_data: dict, row_idx: int):
        if mtg_id in self._active_dl_threads:
            del self._active_dl_threads[mtg_id]

        self.refresh_table()
        if success:
            self._open_tdocs_window(row_data, result)
        else:
            QMessageBox.critical(self, "Download Error", f"Failed to download TDocs List:\n{result}")

    def _update_last_meeting_btn(self):
        last_id, last_num, last_wg = self.settings.get_last_meeting()

        if last_id and last_num and not last_wg:
            results = self.db.search_meetings(search_term=last_num)
            matched = next((m for m in results if m.get("mtg_id") == last_id), None)
            if matched:
                last_wg = matched.get("wg_name")

        if last_num:
            display_tag = f"{last_wg}#{last_num}" if last_wg else f"{last_num}"
            self.btn_open_last.setText(f"Last Meeting ({display_tag})")
            self.btn_open_last.setToolTip(f"Instantly resume working on meeting {display_tag}")
            self.btn_open_last.setEnabled(True)
        else:
            self.btn_open_last.setText("Open Last Meeting")
            self.btn_open_last.setToolTip("No recent meeting history found. Please open a meeting first.")
            self.btn_open_last.setEnabled(False)

    def _download_agenda_csv(self, row_data: dict):
        current_cache = self.dl_dir_input.text().strip() if hasattr(self, 'dl_dir_input') else self.settings.cache_dir
        folder_name = row_data.get("folder_name") or row_data.get("meeting_number", "")
        agenda_dir = Path(current_cache) / folder_name / "Agenda"

        candidate_urls = []
        raw_url = row_data.get("url_key", "")
        if raw_url:
            full_ftp = raw_url if raw_url.startswith("http") else f"https://www.3gpp.org/ftp/{raw_url.lstrip('/')}"
            candidate_urls.append(full_ftp)

        docs_url = row_data.get("docs_folder_url", "")
        if docs_url:
            candidate_urls.append(re.sub(r'/Docs/?$', '', docs_url, flags=re.IGNORECASE))

        self.agenda_dl_thread = AgendaDownloaderThread(candidate_urls, agenda_dir, self)
        self.agenda_dl_thread.finished.connect(self._on_agenda_download_finished)
        self.agenda_dl_thread.start()

    def _on_agenda_download_finished(self, success: bool, msg: str):
        if success:
            notify_success(self, "Agenda Download", msg)
        else:
            QMessageBox.warning(self, "Agenda Download", msg)

    def _download_and_open_tdocs(self, row_data: dict, row_idx: int = -1):
        mtg_id = row_data.get("mtg_id")
        row_data['tdoc_btn_status'] = 'fetching'

        if not isinstance(row_idx, int) or row_idx < 0:
            row_idx = -1
            if hasattr(self.table_model, '_data'):
                for i, r in enumerate(self.table_model._data):
                    if r.get('mtg_id') == mtg_id:
                        row_idx = i
                        break

        if isinstance(row_idx, int) and row_idx >= 0:
            index = self.table_model.index(row_idx, 1)
            self.table_model.dataChanged.emit(index, index)

        current_cache = self.dl_dir_input.text().strip() if hasattr(self, 'dl_dir_input') else self.settings.cache_dir
        folder_name = row_data.get("folder_name") or row_data.get("meeting_number", "")
        local_path = Path(current_cache) / folder_name

        thread = TDocsDownloaderThread(mtg_id, local_path, self)
        self._active_dl_threads[mtg_id] = thread
        thread.finished.connect(
            lambda success, res, m_id: self._on_inline_download_finished(success, res, m_id, row_data, row_idx))
        thread.start()

    def _open_last_meeting(self):
        open_last_meeting(self)

    def _start_tdocs_caching(self, docs_url: str, local_path: Path):
        start_tdocs_caching(self, docs_url, local_path)

    def _on_tdocs_caching_finished(self, success: bool, msg: str):
        on_tdocs_caching_finished(self, success, msg)

    def _handle_global_action_from_window(self, tdoc_str: str, action: str):
        self.global_tdoc_input.setText(tdoc_str)
        self.search_controller.on_tdoc_input_changed(tdoc_str)

        if not self.search_controller.current_found_meeting:
            logging.warning(f"⚠️ MeetingsTab couldn't find meeting for '{tdoc_str}'")
            QMessageBox.warning(self, "Not Found",
                                f"Could not find '{tdoc_str}' in the global database. Ensure the meeting is synced.")
            return

        if action == 'open_meeting':
            self.search_controller.action_open_meeting_list()
        elif action == 'open_doc':
            self.search_controller.action_open_tdoc_only()
        elif action == 'add_to_cart':
            self.search_controller.action_add_to_cart()
        elif action == 'compare_active_word':
            self.search_controller.action_compare_with_active_word()

    def _export_merged_tdocs(self):
        export_merged_tdocs(self)

    def _on_merger_finished(self, success: bool, msg: str):
        on_merger_finished(self, success, msg)

    def _open_contribution_report(self):
        open_contribution_report(self)

    def _open_topic_dossier_studio(self):
        open_topic_dossier_studio(self)

    def _open_agreement_tracker(self):
        open_agreement_tracker(self)

    def _open_meeting_from_report(self, wg_name: str, meeting_number: str):
        open_meeting_from_report(self, wg_name, meeting_number)

    def open_cr_pcr_for_meeting(self, row_data: dict):
        from modules.word_tools.ui.cr_pcr_dialog import CrPcrDialog
        mtg_info = {
            "wg_name": row_data.get("wg_name", ""),
            "meeting_number": row_data.get("meeting_number", ""),
            "location": row_data.get("location", ""),
            "dates": row_data.get("dates", ""),
            "start_date": row_data.get("start_date", ""),
            "end_date": row_data.get("end_date", ""),
            "title": row_data.get("title", ""),
        }
        if not hasattr(self, "_active_cr_dialogs"):
            self._active_cr_dialogs = []
        dlg = CrPcrDialog(
            spec_num="",
            meeting_info=mtg_info,
            parent=self,
        )
        self._active_cr_dialogs.append(dlg)
        dlg.destroyed.connect(lambda: self._active_cr_dialogs.remove(dlg) if dlg in self._active_cr_dialogs else None)
        dlg.show()
        dlg.raise_()
        dlg.activateWindow()
