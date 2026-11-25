# --- File: src/modules/meetings/ui/tdocs/window/mixins/meeting_data.py ---
"""
Data refresh: Excel/revisions/agenda fetch, WI enrichment, agenda/chair-notes controller
delegates, Session Plan window integration.

Mixin of TDocsWindow: relies on attributes set in TDocsWindow.__init__ / _setup_*.
"""
import logging
from pathlib import Path
from PyQt5.QtCore import QTimer
from PyQt5.QtWidgets import QMessageBox
from core.network.network_state import NetworkState
from modules.meetings.core.tdocs.tdoc_file_handler import TDocFileHandler
from modules.meetings.core.tdocs.tdocs_downloader import TDocsDownloaderThread
from modules.meetings.core.tdocs.tdocs_parser import TDocsParser
from modules.meetings.core.meeting_wi_enricher import MeetingWIEnricherThread
from modules.meetings.core.tdocs.tdocs_threads import TDocsRevisionsFetcherThread, TdocsByAgendaThread
from modules.meetings.core.url_router import URLRouter
from modules.meetings.ui.tdocs_agenda_controller import (
    import_word_agenda_dialog,
    process_imported_agenda_files,
    process_next_in_import_queue,
    on_word_import_progress,
    on_word_import_finished,
    finalize_import_batch,
    download_agenda_csv,
    on_agenda_download_finished,
    download_chair_notes,
    scan_existing_chair_notes,
    on_chair_notes_finished,
    open_chair_notes_folder,
)


class MeetingDataMixin:
    """Data refresh: Excel/revisions/agenda fetch, WI enrichment, agenda/chair-notes controller"""

    def _refresh_both(self):
        self._refresh_excel()
        self._refresh_revisions(silent=True)

    def _refresh_excel(self):
        if not self.mtg_info.get("mtg_id"):
            return QMessageBox.warning(self, "Missing ID", "Cannot refresh: Missing 3GPP Portal ID.")
        self.refresh_btn.setText("⏳ Downloading...")
        self.refresh_btn.setEnabled(False)
        self.dl_thread = TDocsDownloaderThread(self.mtg_info.get("mtg_id"), self.meeting_dir, self)
        self.dl_thread.finished.connect(self._on_refresh_excel_finished)
        self.dl_thread.start()

    def _on_refresh_excel_finished(self, success: bool, result: str, mtg_id: str):
        self.refresh_btn.setText("🔄 Refresh")
        self.refresh_btn.setEnabled(True)
        if success:
            self.filepath = result
            if new_data := TDocsParser.parse_tdocs_excel(self.filepath):
                self.model.update_data(new_data)
                self._setup_cache()
                self._refresh_comboboxes()
                self._refresh_email_counts()
                self.last_mod_lbl.setText(self._get_mod_date_str())
                self._start_background_wi_enrichment()
            else:
                QMessageBox.warning(self, "Parse Error", "Downloaded, but could not parse the Excel file.")
        else:
            QMessageBox.critical(self, "Download Error", f"Failed to refresh TDocs:\n{result}")

    def _start_background_wi_enrichment(self):
        """Dispatches an unnoticeable, asynchronous background thread to enrich missing Work Items."""
        try:
            if not getattr(self, "filepath", None):
                return
            excel_path = Path(self.filepath)
            if not excel_path.exists():
                return
            if "wi_enrichment" in self.active_threads:
                existing = self.active_threads["wi_enrichment"]
                if existing.isRunning():
                    return
            thread = MeetingWIEnricherThread(excel_path, parent=self)
            thread.finished_enrichment.connect(self._on_wi_enrichment_finished)
            self.active_threads["wi_enrichment"] = thread
            thread.start()
        except Exception as e:
            logging.debug(f"[TDocsWindow] Failed to dispatch background WI enrichment: {e}")

    def _on_wi_enrichment_finished(self, success: bool, count: int, msg: str):
        if "wi_enrichment" in self.active_threads:
            del self.active_threads["wi_enrichment"]
        if success and count > 0:
            try:
                self._load_tracked_wi_tokens()
                self._update_filter_facets()
            except Exception as e:
                logging.debug(f"[TDocsWindow] Filter refresh after WI enrichment: {e}")

    def _refresh_revisions(self, silent=False):
        if not self.revisions_url:
            return
        self.rev_thread = TDocsRevisionsFetcherThread(self.revisions_url, self.meeting_dir)
        self.rev_thread.finished.connect(lambda s, d, m: self._on_revisions_fetched(s, d, m, silent))
        self.rev_thread.start()

    def _on_revisions_fetched(self, success: bool, data: dict, msg: str, silent: bool):
        if success:
            self.model.revisions = data
            self.model.dataChanged.emit(self.model.index(0, 0), self.model.index(self.model.rowCount() - 1, 0))
            if not silent:
                self.refresh_btn.setText(f"✅ {len(data)} Revs")
                QTimer.singleShot(4000, lambda: self.refresh_btn.setText("🔄 Refresh"))
        elif not silent:
            QMessageBox.warning(self, "Revisions Error", f"Failed to sync revisions:\n{msg}")

    def _fetch_tdocs_by_agenda(self):
        wg_name = self.mtg_info.get("wg_name", "").upper()
        main_url = self.mtg_info.get("url_key", "")
        if main_url and not main_url.startswith("http"):
            main_url = f"https://www.3gpp.org/ftp/{main_url.lstrip('/')}"

        is_active = self.mtg_info.get("is_active_sync", False)

        candidate_urls = []
        if NetworkState.get_instance().is_local_active():
            local_base = URLRouter._get_local_server_base(wg_name)
            candidate_urls.append(local_base)

        if is_active:
            sync_wg = "SA3LI" if wg_name == "SA3LI" else wg_name
            candidate_urls.append(f"https://www.3gpp.org/ftp/Meetings_3GPP_SYNC/{sync_wg}")

        if main_url:
            candidate_urls.append(main_url)

        if not candidate_urls:
            return QMessageBox.warning(self, "No URL Available", "Cannot sync agenda: No valid meeting URL found.")

        self.refresh_btn.setText("⏳ Parsing HTML...")
        self.agenda_thread = TdocsByAgendaThread(candidate_urls, self.meeting_dir)
        self.agenda_thread.finished.connect(self._on_agenda_fetched)
        self.agenda_thread.start()

    def _on_agenda_fetched(self, success: bool, agenda_data: dict):
        if success and agenda_data:
            self.model.merge_agenda_data(agenda_data)
            self._refresh_comboboxes()
            self.refresh_btn.setText(f"✅ {len(agenda_data)} Merged")
            QTimer.singleShot(4000, lambda: self.refresh_btn.setText("🔄 Refresh"))
        else:
            self.refresh_btn.setText("🔄 Refresh")
            QMessageBox.warning(self, "Error", "Failed to parse TdocsByAgenda.htm.")

    def _open_session_plan_window(self):
        """Launches the dedicated Meeting Session Plan viewer."""
        if not hasattr(self, "_session_plan_window") or self._session_plan_window is None:
            from modules.meetings.ui.session_plan.session_plan_window import SessionPlanWindow
            followed = getattr(self.proxy, "followed_wi_tokens", set()) if hasattr(self, "proxy") else set()
            starred = getattr(self.proxy, "starred_wi_tokens", set()) if hasattr(self, "proxy") else set()
            tracked_tokens = set(followed) | set(starred)
            self._session_plan_window = SessionPlanWindow(
                meeting_dir=self.meeting_dir,
                mtg_info=self.mtg_info,
                agenda_map=self.agenda_map,
                tracked_tokens=tracked_tokens,
            )
        self._session_plan_window.show()
        self._session_plan_window.raise_()
        self._session_plan_window.activateWindow()

    def _load_tdoc_as_session_plan(self, tdoc_id: str):
        """Fetches the target TDoc, stages it to Agenda/SessionPlans, and launches SessionPlanWindow."""
        from modules.meetings.core.session_plan.session_plan_parser import SessionPlanParser
        from modules.meetings.core.session_plan.session_plan_db import SessionPlanDatabase
        from modules.meetings.core.url_router import URLRouter
        from core.network.network_state import NetworkState

        wg_name = self.mtg_info.get("wg_name", "").upper()
        candidate_urls = []
        if NetworkState.get_instance().is_local_active():
            candidate_urls.append(URLRouter._get_local_server_base(wg_name) + "/Docs")
        if self.docs_ftp_url:
            candidate_urls.append(self.docs_ftp_url)

        tdoc_dir = self.meeting_dir / tdoc_id
        extracted_docs = []
        try:
            for url in candidate_urls:
                extracted_docs = TDocFileHandler.download_and_extract_tdoc(tdoc_id, url, tdoc_dir, timeout=10)
                if extracted_docs:
                    break
        except Exception as e:
            logging.warning(f"Failed to fetch {tdoc_id}: {e}")

        if not extracted_docs and tdoc_dir.exists():
            extracted_docs = list(tdoc_dir.glob("*.docx")) + list(tdoc_dir.glob("*.doc"))

        if not extracted_docs:
            QMessageBox.warning(self, "Download Failed", f"Could not acquire document files for {tdoc_id}.")
            return

        doc_file = extracted_docs[0]
        staged_file = SessionPlanParser.stage_session_plan_file(doc_file, self.meeting_dir, tdoc_id)
        db = SessionPlanDatabase(self.meeting_dir / "Agenda" / "session_plans.db")

        if staged_file.suffix.lower() == ".docx":
            snap = SessionPlanParser.parse_docx(staged_file, meeting_number=str(self.mtg_info.get("meeting_number", "")))
        else:
            snap = SessionPlanParser.parse_text_or_markdown(
                staged_file.read_text(encoding="utf-8", errors="replace"),
                meeting_number=str(self.mtg_info.get("meeting_number", "")),
                cached_file_path=str(staged_file),
            )

        db.save_snapshot(snap)
        self._open_session_plan_window()
        if hasattr(self, "_session_plan_window") and self._session_plan_window:
            self._session_plan_window._load_data()
        QMessageBox.information(self, "Session Plan Loaded", f"Successfully loaded {tdoc_id} into Session Plan viewer!")

    def _import_word_agenda_dialog(self):
        import_word_agenda_dialog(self)

    def _process_imported_agenda_files(self, source_paths: list):
        process_imported_agenda_files(self, source_paths)

    def _process_next_in_import_queue(self):
        process_next_in_import_queue(self)

    def _on_word_import_progress(self, msg: str):
        on_word_import_progress(self, msg)

    def _on_word_import_finished(self, success: bool, agenda_data: dict, filename: str, error_msg: str):
        on_word_import_finished(self, success, agenda_data, filename, error_msg)

    def _finalize_import_batch(self):
        finalize_import_batch(self)

    def _download_agenda_csv(self):
        download_agenda_csv(self)

    def _on_agenda_download_finished(self, success: bool, msg: str):
        on_agenda_download_finished(self, success, msg)

    def _download_chair_notes(self):
        download_chair_notes(self)

    def _scan_existing_chair_notes(self):
        scan_existing_chair_notes(self)

    def _on_chair_notes_finished(self, success: bool, count: int, files: list, msg: str):
        on_chair_notes_finished(self, success, count, files, msg)

    def _open_chair_notes_folder(self):
        open_chair_notes_folder(self)
