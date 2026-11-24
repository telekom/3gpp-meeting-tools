"""Row navigation, TDoc download/open actions, exports, and session state persistence."""

from __future__ import annotations

import re
import webbrowser
from pathlib import Path

from PyQt5.QtCore import QDate, QModelIndex, QPoint, Qt
from PyQt5.QtWidgets import QApplication, QMenu, QMessageBox

from core.utils.portal_utils import open_portal_tdoc
from modules.meetings.core.tdocs.tdocs_threads import TDocActionThread
from modules.meetings.core.url_router import URLRouter
from modules.meetings.ui.dialogs.contribution_export_controller import (
    export_to_excel,
    generate_statistics,
    on_export_finished,
    on_stats_finished,
)
from modules.meetings.ui.tdocs_dialogs import ReadOnlyViewerDialog


class ContributionDialogActionsMixin:
    """Mixin handling table double-clicks, context menus, TDoc actions, exports, and session state."""

    def _request_open_meeting(self, row_data: dict):
        """Requests opening the full meeting window in the main application."""
        wg = str(row_data.get("WG", "")).strip()
        mtg_raw = str(row_data.get("Meeting", "")).strip()
        mtg_num = mtg_raw.split("#")[-1].strip() if "#" in mtg_raw else mtg_raw
        if wg and mtg_num:
            self.open_meeting_requested.emit(wg, mtg_num)
        else:
            QMessageBox.warning(self, "Missing Info", "Cannot identify the Working Group or meeting number for this row.")

    def _open_docs_folder(self, row_data: dict):
        """Opens the Docs folder of the meeting for this TDoc in the default web browser."""
        docs_url = str(row_data.get("docs_folder_url", "")).strip()
        if not docs_url:
            wg = row_data.get("WG", "")
            mtg_num = row_data.get("meeting_number", "")
            mtgs = self.meetings_db.search_meetings(wg_name=[wg] if wg else None, search_term=mtg_num)
            if mtgs:
                docs_url = mtgs[0].get("docs_folder_url", "")

        if docs_url:
            full_url = docs_url if docs_url.startswith("http") else f"https://www.3gpp.org/ftp/{docs_url.lstrip('/')}"
            webbrowser.open(full_url)
        else:
            QMessageBox.warning(self, "Missing URL", "No documents folder URL available for this meeting.")

    def _get_tdoc_download_url(self, row_data: dict) -> str:
        """Constructs the fully qualified 3GPP FTP/HTTP download URL for the TDoc."""
        tdoc_id = str(row_data.get("TDoc", "")).strip()
        if not tdoc_id:
            return ""

        docs_url = str(row_data.get("docs_folder_url", "")).strip()
        if not docs_url:
            wg = row_data.get("WG", "")
            mtg_num = row_data.get("meeting_number", "")
            mtgs = self.meetings_db.search_meetings(wg_name=[wg] if wg else None, search_term=mtg_num)
            if mtgs:
                docs_url = mtgs[0].get("docs_folder_url", "")

        if docs_url:
            full_url = docs_url if docs_url.startswith("http") else f"https://www.3gpp.org/ftp/{docs_url.lstrip('/')}"
            return f"{full_url.rstrip('/')}/{tdoc_id}.zip"
        return ""

    def _open_tdoc(self, row_data: dict):
        """Downloads (if missing) and opens the actual document using TDocActionThread."""
        tdoc_id = str(row_data.get("TDoc", "")).strip()
        if not tdoc_id:
            return

        match = re.search(r"^(.*?)-?(?:r|rev)(\d{1,2}[a-zA-Z]?)$", tdoc_id, re.IGNORECASE)
        base_tdoc = match.group(1).upper() if match else tdoc_id.upper()
        target_filename = tdoc_id.upper()

        if target_filename in self.active_tdoc_threads:
            th = self.active_tdoc_threads[target_filename]
            if th and th.isRunning():
                return

        wg_name = row_data.get("WG") or row_data.get("wg_name", "")
        folder_name = row_data.get("folder_name") or row_data.get("meeting_number", "")
        raw_url = row_data.get("url_key", "")
        docs_url = row_data.get("docs_folder_url", "")

        main_ftp = raw_url if raw_url.startswith("http") else (f"https://www.3gpp.org/ftp/{raw_url.lstrip('/')}" if raw_url else "")

        is_active_sync = self.meetings_db.is_active_sync_meeting(
            wg_name,
            row_data.get("start_date", ""),
            row_data.get("end_date", ""),
            int(row_data.get("is_electronic", 0) or 0),
        )

        candidate_urls = URLRouter.build_priority_url_list(
            wg_name,
            folder_name,
            main_ftp,
            is_active_sync,
            target_filename=target_filename,
        )

        if docs_url:
            clean_docs = docs_url if docs_url.startswith("http") else f"https://www.3gpp.org/ftp/{docs_url.lstrip('/')}"
            clean_docs = clean_docs.rstrip("/")
            if clean_docs not in candidate_urls:
                candidate_urls.append(clean_docs)

        if not candidate_urls:
            QMessageBox.warning(self, "Missing URL", f"No candidate download routes found for {tdoc_id}.")
            return

        meeting_dir = Path(self.settings.cache_dir) / folder_name
        meeting_dir.mkdir(parents=True, exist_ok=True)

        thread = TDocActionThread(base_tdoc, target_filename, candidate_urls, meeting_dir, open_file=True)
        self.active_tdoc_threads[target_filename] = thread
        thread.finished_action.connect(lambda t, s, m, th=thread: self._on_tdoc_action_finished(t, s, m, th))
        thread.start()

    def _on_tdoc_action_finished(self, tdoc: str, success: bool, msg: str, thread: TDocActionThread):
        if thread.target_filename in self.active_tdoc_threads:
            del self.active_tdoc_threads[thread.target_filename]
        if not success:
            QMessageBox.warning(self, f"Action Failed: {tdoc}", msg)

    def _on_table_double_clicked(self, index: QModelIndex):
        """Handles double-click navigation based on the clicked column."""
        row_data = self.table_model.data(index, Qt.UserRole)
        if not row_data:
            return

        col_name = self.table_model.COLUMNS[index.column()]

        if col_name in ["WG", "Meeting"]:
            self._request_open_meeting(row_data)
            return

        if col_name == "Abstract":
            abstract_text = str(row_data.get("Abstract", "")).strip()
            tdoc_id = str(row_data.get("TDoc", "")).strip()
            if abstract_text:
                ReadOnlyViewerDialog(self, f"📄 Abstract: {tdoc_id}", abstract_text).exec_()
                return

        self._open_tdoc(row_data)

    def _on_table_context_menu(self, pos: QPoint):
        """Row-level right-click menu with explicit actions."""
        index = self.table_view.indexAt(pos)
        if not index.isValid():
            return

        row_data = self.table_model.data(index, Qt.UserRole)
        if not row_data:
            return

        menu = QMenu(self)
        mtg = row_data.get("Meeting", "")
        tdoc = row_data.get("TDoc", "")

        act_open_mtg = menu.addAction(f"🗓️ Open Meeting Table ({mtg})")
        act_open_mtg.triggered.connect(lambda: self._request_open_meeting(row_data))

        act_open_docs = menu.addAction("📂 Open Docs Folder")
        act_open_docs.triggered.connect(lambda: self._open_docs_folder(row_data))

        menu.addSeparator()

        if tdoc:
            act_open_tdoc = menu.addAction(f"📄 Open TDoc Document ({tdoc})")
            act_open_tdoc.triggered.connect(lambda: self._open_tdoc(row_data))

            act_portal = menu.addAction(f"🌐 View {tdoc} on 3GU Portal")
            act_portal.triggered.connect(lambda: open_portal_tdoc(tdoc))

            menu.addSeparator()

            act_copy_tdoc = menu.addAction(f"📋 Copy TDoc Number ({tdoc})")
            act_copy_tdoc.triggered.connect(lambda: QApplication.clipboard().setText(tdoc))

            tdoc_url = self._get_tdoc_download_url(row_data)
            if tdoc_url:
                act_copy_url = menu.addAction("🔗 Copy URL")
                act_copy_url.triggered.connect(lambda: QApplication.clipboard().setText(tdoc_url))

        title = str(row_data.get("Title", "")).strip()
        if title:
            act_copy_title = menu.addAction("📋 Copy Title")
            act_copy_title.triggered.connect(lambda: QApplication.clipboard().setText(title))

        abstract = str(row_data.get("Abstract", "")).strip()
        if abstract:
            menu.addSeparator()
            act_view_abs = menu.addAction("📝 View Abstract...")
            act_view_abs.triggered.connect(lambda: ReadOnlyViewerDialog(self, f"📄 Abstract: {tdoc}", abstract).exec_())

        menu.exec_(self.table_view.viewport().mapToGlobal(pos))

    def _export_to_excel(self):
        export_to_excel(self)

    def _on_export_finished(self, success: bool, msg: str):
        on_export_finished(self, success, msg)

    def _generate_statistics(self):
        generate_statistics(self)

    def _on_stats_finished(self, success: bool, msg: str):
        on_stats_finished(self, success, msg)

    def get_session_state(self) -> dict:
        """Return serializable ContributionReportDialog filter state."""
        return {
            "wgs": self.wg_filter.getCheckedItems() if hasattr(self, "wg_filter") else [],
            "date_from": self.date_from.date().toString("yyyy-MM-dd") if hasattr(self, "date_from") else "",
            "date_to": self.date_to.date().toString("yyyy-MM-dd") if hasattr(self, "date_to") else "",
            "wi_tokens": self.wi_token_widget.get_tokens() if hasattr(self, "wi_token_widget") else [],
            "ai_tokens": self.ai_token_widget.get_tokens() if hasattr(self, "ai_token_widget") else [],
            "wi_ai_or": bool(self.chk_wi_ai_or.isChecked()) if hasattr(self, "chk_wi_ai_or") else False,
            "companies": sorted(self.company_widget.get_selected_companies()) if hasattr(self, "company_widget") else [],
            "primary_only": bool(self.company_widget.is_primary_only()) if hasattr(self, "company_widget") else False,
            "include_tentative": bool(self.company_widget.is_include_tentative()) if hasattr(self, "company_widget") else False,
            "bypass_cache": bool(self.chk_bypass_cache.isChecked()) if hasattr(self, "chk_bypass_cache") else False,
        }

    def restore_session_state(self, state: dict) -> None:
        """Restore WG, date, WI/AI tokens, company filters, and flags."""
        if not isinstance(state, dict):
            return
        wgs = state.get("wgs")
        if isinstance(wgs, list) and hasattr(self, "set_active_wgs"):
            self.set_active_wgs([str(w) for w in wgs])

        d_from = QDate.fromString(str(state.get("date_from") or ""), "yyyy-MM-dd")
        if d_from.isValid() and hasattr(self, "date_from"):
            self.date_from.setDate(d_from)
        d_to = QDate.fromString(str(state.get("date_to") or ""), "yyyy-MM-dd")
        if d_to.isValid() and hasattr(self, "date_to"):
            self.date_to.setDate(d_to)

        if isinstance(state.get("wi_tokens"), list) and hasattr(self.wi_token_widget, "set_tokens"):
            self.wi_token_widget.set_tokens(state["wi_tokens"])
        if isinstance(state.get("ai_tokens"), list) and hasattr(self.ai_token_widget, "set_tokens"):
            self.ai_token_widget.set_tokens(state["ai_tokens"])
        if "wi_ai_or" in state and hasattr(self, "chk_wi_ai_or"):
            self.chk_wi_ai_or.setChecked(bool(state["wi_ai_or"]))

        if isinstance(state.get("companies"), list) and hasattr(self.company_widget, "set_selection"):
            self.company_widget.set_selection(
                companies=state["companies"],
                primary_only=bool(state.get("primary_only", False)),
                include_tentative=bool(state.get("include_tentative", False)),
            )

        if "bypass_cache" in state and hasattr(self, "chk_bypass_cache"):
            self.chk_bypass_cache.setChecked(bool(state["bypass_cache"]))
