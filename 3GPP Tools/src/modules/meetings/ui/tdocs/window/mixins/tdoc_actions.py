# --- File: src/modules/meetings/ui/tdocs/window/mixins/tdoc_actions.py ---
"""
Per-TDoc actions: action/related menus, cell popups, notes, open/locate/reveal TDocs,
agreements counter & dialogs, draft CR, triage.

Mixin of TDocsWindow: relies on attributes set in TDocsWindow.__init__ / _setup_*.
"""
import logging
import os
import re
import webbrowser
from pathlib import Path
from PyQt5.QtCore import Qt, QPoint
from PyQt5.QtGui import QCursor
from PyQt5.QtWidgets import QTableView, QMessageBox
from core.ui.ui_components import BUTTON_STYLE_TOOLBAR_SECONDARY
from modules.meetings.core.compare_manager import ComparisonManager
from modules.meetings.core.tdocs.tdocs_threads import TDocActionThread
from modules.meetings.ui.tdocs_dialogs import (
    ReadOnlyViewerDialog,
    InteractiveNotesDialog,
    TDocInfoDialog,
)
from modules.meetings.ui.tdocs_menus import build_action_menu, build_related_menu
from modules.meetings.ui.tdoc_triage_dialog import TDocTriageDialog
from core.utils.paths import get_project_root
from core.ui.theme import C


class TDocActionsMixin:
    """Per-TDoc actions: action/related menus, cell popups, notes, open/locate/reveal TDocs,"""

    def _update_agreements_counter(self):
        """Updates the ⚖️ Agreements toolbar button counter badge and active styling."""
        if not hasattr(self, "btn_meeting_agreements"):
            return
        if not getattr(self, "is_ran", False):
            self.btn_meeting_agreements.setVisible(False)
            return
        ag_list = getattr(self, "chair_agreements", [])
        non_recap = [a for a in ag_list if not a.get("is_previous_meeting_recap")]
        cnt = len(non_recap)
        if cnt == 0 and hasattr(self, "model") and hasattr(self.model, "_data"):
            cnt = sum(1 for r in self.model._data if r.get("_has_agreement"))

        if cnt > 0:
            self.btn_meeting_agreements.setText(f"⚖️ Agreements ({cnt})")
            self.btn_meeting_agreements.setStyleSheet(f"""
                QPushButton {{
                    background-color: {C.SUCCESS_SOFT};
                    color: {C.SUCCESS_DARK};
                    border: 1.5px solid {C.SUCCESS_BORDER_MEDIUM};
                    border-radius: 4px;
                    padding: 4px 10px;
                    font-size: 12px;
                    font-weight: bold;
                }}
                QPushButton:hover {{
                    background-color: {C.SUCCESS_SOFT_ALT};
                    border-color: #22C55E;
                }}
            """)
        else:
            self.btn_meeting_agreements.setText("⚖️ Agreements")
            self.btn_meeting_agreements.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)

    def _open_meeting_agreements_dialog(self, target_tdoc: str = None):
        """Opens the dedicated Meeting Agreements Viewer & Management Console."""
        if not getattr(self, "is_ran", False):
            QMessageBox.information(
                self,
                "RAN-Specific Feature",
                "Chairman's Notes agreements scanning and viewing are specific to RAN working groups (RAN1, RAN2, RAN3, etc.).\n\n"
                "In SA/CT working groups, contribution agreements are formally tracked via TDoc Status (agreed/approved) in the meeting table."
            )
            return
        from modules.meetings.ui.meeting_agreements_dialog import MeetingAgreementsDialog

        ag_list = getattr(self, "chair_agreements", [])
        if not ag_list:
            from modules.meetings.core.tdocs.chair_notes_agreements_parser import ChairNotesAgreementsParser
            _, ag_list = ChairNotesAgreementsParser.load_cached_chair_notes_data(self.meeting_dir)
            if ag_list:
                self.chair_agreements = ag_list

        if not ag_list:
            # Fallback: construct lightweight agreement entries from model rows if remarks contain agreement tags
            synthetic_ags = []
            if hasattr(self, "model") and hasattr(self.model, "_data"):
                for row in self.model._data:
                    if row.get("_has_agreement"):
                        synthetic_ags.append({
                            "kind": row.get("_agreement_kind", "Agreement"),
                            "tdoc": row.get("TDoc", ""),
                            "ai": row.get("Agenda Item", ""),
                            "wi": row.get("Related WIs", ""),
                            "title": row.get("Title", ""),
                            "company": row.get("Source", ""),
                            "text": row.get("Secretary Remarks", ""),
                        })
            if synthetic_ags:
                ag_list = synthetic_ags
                self.chair_agreements = ag_list

        if not ag_list:
            reply = QMessageBox.question(
                self,
                "No Agreements Loaded",
                "No Chairman's Notes agreements have been scanned or cached for this meeting yet.\n\n"
                "Would you like to scan existing Chairman's Notes in Agenda/ChairNotes now?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.Yes,
            )
            if reply == QMessageBox.Yes:
                self._scan_existing_chair_notes()
            return

        mtg_name = self.mtg_info.get("meeting_name", "") or f"{self.mtg_info.get('wg_name', '')} {self.mtg_info.get('meeting_number', '')}".strip()
        dlg = MeetingAgreementsDialog(
            agreements=ag_list,
            meeting_name=mtg_name,
            meeting_dir=self.meeting_dir,
            parent=self,
        )
        dlg.locate_tdoc_requested.connect(self.locate_tdoc)
        if target_tdoc:
            dlg.focus_tdoc(target_tdoc)
        dlg.exec_()

    def _handle_tdoc_action(self, base_tdoc: str):
        if str(base_tdoc or "").strip().upper().startswith("THR-"):
            self._open_tracked_thread_dialog(str(base_tdoc).strip().upper())
            return
        if base_tdoc in self.model.loading_tdocs or not self.docs_ftp_url:
            return
        revisions = self.model.revisions.get(base_tdoc, [])
        build_action_menu(
            self.table,
            base_tdoc,
            self.docs_ftp_url,
            self.revisions_url,
            revisions,
            self.meeting_dir,
            self._trigger_download_thread,
            self._export_llm_single,
            self._compose_email_draft,
            QCursor.pos(),
            compare_active_word_callback=self._compare_tdoc_with_active_word,
        )

    def _show_related_menu(self, target_tdoc: str, pos: QPoint):
        build_related_menu(
            self,
            target_tdoc,
            self.model.valid_tdocs,
            self.docs_ftp_url,
            self.revisions_url,
            self._scroll_to_tdoc,
            self._trigger_download_thread,
            self._export_llm_single,
            self.global_action_requested.emit,
            self._compose_email_draft,
            pos,
            triage_callback=self._open_tdoc_triage,
            compare_active_word_callback=self._compare_tdoc_with_active_word,
            launch_cli_callback=self._launch_cli_for_tdocs,
            send_tdoc_email_callback=self._send_tdoc_via_email,
        )

    @staticmethod
    def _extract_base_tdoc(tdoc_id: str) -> str:
        clean = str(tdoc_id or "").strip().upper()
        match = re.search(r'^(.*?)-?(?:r|rev)\d{1,2}[a-zA-Z]?$', clean, re.IGNORECASE)
        return match.group(1).upper() if match else clean

    def _open_interactive_notes_for_row(self, row_data: dict):
        if not row_data:
            return
        tdoc_id = str(row_data.get("TDoc", "") or "").strip()
        if not tdoc_id:
            return
        all_tags = self.db.get_all_unique_tags() if hasattr(self, "db") else []
        InteractiveNotesDialog(
            self,
            tdoc_id,
            row_data,
            self._save_user_data,
            all_existing_tags=all_tags,
        ).exec_()

    def _show_cell_popup(self, index):
        if not index.isValid():
            return
        row_data = self.model._data[self.proxy.mapToSource(index).row()]
        col_name = self.model._headers[index.column()]
        if row_data.get("_is_tracked_thread"):
            if col_name in ["My Notes", "My Status"]:
                self._open_interactive_notes_for_row(row_data)
            else:
                self._open_tracked_thread_dialog(row_data.get("_thread_id") or row_data.get("TDoc", ""))
            return

        if col_name == "Emails":
            row_data = self.model._data[self.proxy.mapToSource(index).row()]
            self._open_tdoc_emails(row_data.get("TDoc", ""))
            return

        if col_name == "TDoc":
            row_data = self.model._data[self.proxy.mapToSource(index).row()]
            tdoc_id = str(row_data.get("TDoc", "")).strip()
            revs = self.model.revisions.get(tdoc_id, [])
            TDocInfoDialog(
                row_data,
                docs_ftp_url=self.docs_ftp_url,
                revisions=revs,
                parent=self,
                triage_callback=self._open_tdoc_triage,  # <--- Added
            ).exec_()
            return

        row_data = self.model._data[self.proxy.mapToSource(index).row()]
        tdoc_id = row_data.get("TDoc", "")

        if col_name == "Secretary Remarks":
            if row_data.get("_has_agreement"):
                self._open_meeting_agreements_dialog(target_tdoc=tdoc_id)
                return
            remarks = str(row_data.get("Secretary Remarks", "") or "")
            if remarks.strip():
                ReadOnlyViewerDialog(self, f"📄 Viewing: Secretary Remarks ({tdoc_id})", remarks).exec_()
            else:
                self._open_interactive_notes_for_row(row_data)
            return

        if col_name == "TDoc Status" and row_data.get("_has_agreement"):
            self._open_meeting_agreements_dialog(target_tdoc=tdoc_id)
            return

        if col_name not in [
            "Title", "Source", "Abstract", "My Notes", "My Status"
        ]:
            return

        if col_name in ["Title", "Source", "Abstract"]:
            val = next(
                (
                    str(v)
                    for k, v in row_data.items()
                    if str(k).strip().lower() == col_name.lower() and v
                ),
                "",
            )
            ReadOnlyViewerDialog(self, f"📄 Viewing: {col_name} ({tdoc_id})", val).exec_()
        else:
            self._open_interactive_notes_for_row(row_data)

    def _find_proxy_row(self, base_tdoc: str):
        for row in range(self.proxy.rowCount()):
            if self.proxy.data(self.proxy.index(row, 1), Qt.UserRole) == base_tdoc:
                return row
        return None

    def _focus_and_highlight_proxy_row(self, row: int, base_tdoc: str):
        idx = self.proxy.index(row, 1)
        if not idx.isValid():
            return
        self.model.set_highlighted_tdoc(base_tdoc)
        if hasattr(self, "_highlight_clear_timer") and self._highlight_clear_timer is not None:
            self._highlight_clear_timer.start()
        self.table.selectRow(row)
        self.table.setCurrentIndex(idx)
        self.table.scrollTo(idx, QTableView.PositionAtCenter)

    def _reveal_tdoc_in_table(self, target_tdoc: str):
        """
        Syncs the table view after a successful Quick Open:
        - If the current filter already includes the TDoc, scroll to and highlight it.
        - If hidden by active filters, clear filters and filter to the TDoc, then scroll & highlight.
        """
        base_tdoc = self._extract_base_tdoc(target_tdoc)
        if not base_tdoc or base_tdoc not in self.model.valid_tdocs:
            return

        row = self._find_proxy_row(base_tdoc)
        if row is not None:
            self._focus_and_highlight_proxy_row(row, base_tdoc)
            return

        self._clear_all_filters()
        self.search_timer.stop()
        self.search_input.blockSignals(True)
        self.search_input.setText(base_tdoc)
        self.search_input.blockSignals(False)
        self.proxy.setGlobalFilter(base_tdoc)
        self._update_filter_facets()
        self._update_count_label()
        self._adjust_table_row_heights()

        row = self._find_proxy_row(base_tdoc)
        if row is not None:
            self._focus_and_highlight_proxy_row(row, base_tdoc)

    def _scroll_to_tdoc(self, target_tdoc: str):
        base_tdoc = self._extract_base_tdoc(target_tdoc)
        if base_tdoc in self.model.valid_tdocs:
            row = self._find_proxy_row(base_tdoc)
            if row is not None:
                self.table.scrollTo(self.proxy.index(row, 1), QTableView.PositionAtCenter)
                self.table.selectRow(row)
                return
            if QMessageBox.question(self, "Hidden TDoc",
                                    f"TDoc '{base_tdoc}' is hidden by active filters.\nClear filters to view?",
                                    QMessageBox.Yes | QMessageBox.No) == QMessageBox.Yes:
                self._clear_all_filters()
                self._scroll_to_tdoc(target_tdoc)
        else:
            if QMessageBox.question(self, "External TDoc",
                                    f"{base_tdoc} is not from this meeting.\nSearch global database?",
                                    QMessageBox.Yes | QMessageBox.No) == QMessageBox.Yes:
                self.global_action_requested.emit(base_tdoc, 'open_meeting')

    def _on_tdoc_action_finished(self, tdoc: str, success: bool, msg: str, thread: TDocActionThread):
        if tdoc in self.active_threads:
            del self.active_threads[tdoc]
        self.model.set_loading(tdoc, False)
        if not success:
            return QMessageBox.warning(self, f"Action Failed: {tdoc}", msg)
        if getattr(thread, "sync_table_on_success", False):
            self._reveal_tdoc_in_table(getattr(thread, "target_filename", "") or tdoc)
        if getattr(thread, "send_email_on_finish", False):
            target_name = getattr(thread, "target_filename", "") or tdoc
            attachments = list(getattr(thread, "extracted_doc_paths", []) or [])
            if not attachments:
                zip_candidate = self.meeting_dir / tdoc / f"{target_name}.zip"
                if zip_candidate.is_file():
                    attachments = [zip_candidate]
            self._open_email_draft_with_attachments(target_name, attachments)
            return
        if getattr(thread, "is_silent_compare", False):
            from modules.word_tools.core.document_opener import select_best_word_document

            files = getattr(thread, "extracted_doc_paths", [])
            best_word = select_best_word_document(files, thread.target_filename)
            if best_word:
                mgr = ComparisonManager.get_instance()
                if getattr(thread, "auto_compare_active_word", False):
                    mgr.set_slot_b(thread.target_filename, str(best_word))
                    if hasattr(self, "cart_frame") and hasattr(self.cart_frame, "run_comparison"):
                        self.cart_frame.run_comparison()
                else:
                    mgr.add_to_cart(thread.target_filename, str(best_word))
            else:
                found_exts = ", ".join(sorted({Path(f).suffix.lower() for f in files})) or "none"
                QMessageBox.warning(
                    self,
                    "Compare Failed",
                    f"No Word document (.docx / .doc) found in {thread.target_filename}.zip (found: {found_exts}).",
                )

    def _open_draft_cr_dialog(self, tdoc_row_data: dict = None):
        from modules.word_tools.ui.cr_pcr_dialog import CrPcrDialog

        mtg_info = {
            "wg_name": self.mtg_info.get("wg_name", ""),
            "meeting_number": self.mtg_info.get("meeting_number", ""),
            "location": self.mtg_info.get("location", ""),
            "dates": self.mtg_info.get("dates", ""),
            "start_date": self.mtg_info.get("start_date", ""),
            "end_date": self.mtg_info.get("end_date", ""),
            "title": self.mtg_info.get("title", ""),
        }

        tdoc_info = tdoc_row_data or {}
        spec_num = tdoc_info.get("Spec", "") or tdoc_info.get("Spec#", "")
        version = tdoc_info.get("Version", "")

        if not hasattr(self, "_active_cr_dialogs"):
            self._active_cr_dialogs = []
        dialog = CrPcrDialog(
            spec_num=spec_num,
            version=version,
            meeting_info=mtg_info,
            tdoc_info=tdoc_info,
            parent=self,
        )
        self._active_cr_dialogs.append(dialog)
        dialog.destroyed.connect(lambda: self._active_cr_dialogs.remove(dialog) if dialog in self._active_cr_dialogs else None)
        dialog.show()
        dialog.raise_()
        dialog.activateWindow()

    def _open_tdoc_from_signal(self, tdoc_id: str):
        main_app_window = self.window()
        main_app_window.setWindowState(main_app_window.windowState() & ~Qt.WindowMinimized | Qt.WindowActive)
        main_app_window.raise_()
        main_app_window.activateWindow()

        self._scroll_to_tdoc(tdoc_id)
        is_rev = re.search(r'(?:r|rev)\d{1,2}[a-zA-Z]?$', tdoc_id, re.IGNORECASE)
        target_url = self.revisions_url if is_rev and self.revisions_url else self.docs_ftp_url

        if not target_url:
            QMessageBox.warning(self, "Missing URL", "No FTP URL available to download this document.")
            return

        match = re.search(r'^(.*?)-?(?:r|rev)\d{1,2}[a-zA-Z]?$', tdoc_id, re.IGNORECASE)
        base_tdoc = match.group(1).upper() if match else tdoc_id.upper()
        self._trigger_download_thread(base_tdoc, tdoc_id, target_url, is_silent_compare=False)

    def _jump_to_tdoc_from_signal(self, tdoc_id: str):
        main_app_window = self.window()
        main_app_window.setWindowState(main_app_window.windowState() & ~Qt.WindowMinimized | Qt.WindowActive)
        main_app_window.raise_()
        main_app_window.activateWindow()
        self._scroll_to_tdoc(tdoc_id)

    def _open_doc_in_word(self, doc_path):
        import os
        try:
            if hasattr(os, "startfile"):
                os.startfile(str(doc_path))
            else:
                import webbrowser
                webbrowser.open(f"file:///{doc_path}")
        except Exception as e:
            QMessageBox.warning(self, "Open Failed", f"Could not open document:\n{e}")

    def _open_agreement_tracker(self):
        from modules.meetings.ui.agreement_tracker_dialog import AgreementTrackerDialog
        from modules.meetings.core.meetings_db import MeetingsDatabase
        from PyQt5.QtCore import QDate

        if not hasattr(self, "_agreement_tracker_dlg") or self._agreement_tracker_dlg is None:
            db_path = get_project_root() / "db" / "3gpp_data.db"
            self._agreement_tracker_dlg = AgreementTrackerDialog(MeetingsDatabase(db_path), self)
            self._agreement_tracker_dlg.open_meeting_requested.connect(
                lambda wg, num: self.global_action_requested.emit(f"{wg}#{num}", "open_meeting")
            )

        wg = str(self.mtg_info.get("wg_name", "")).strip()
        if wg:
            self._agreement_tracker_dlg.set_active_wgs([wg])
        s_date = QDate.fromString(str(self.mtg_info.get("start_date", "")), "yyyy-MM-dd")
        e_date = QDate.fromString(str(self.mtg_info.get("end_date", "")), "yyyy-MM-dd")
        if s_date.isValid() and e_date.isValid():
            self._agreement_tracker_dlg.set_date_range(s_date.addDays(-7), e_date.addDays(7))

        checked_wis = self.wi_combo.getCheckedItems() if hasattr(self, "wi_combo") else []
        all_wis = (
            self.wi_combo.model().rowCount() - 1
            if hasattr(self, "wi_combo") and self.wi_combo.model()
            else 0
        )
        if checked_wis and len(checked_wis) < all_wis:
            self._agreement_tracker_dlg._clear_wi_tokens()
            self._agreement_tracker_dlg.add_wi_tokens(checked_wis)

        self._agreement_tracker_dlg.show()
        self._agreement_tracker_dlg.raise_()
        self._agreement_tracker_dlg.activateWindow()

    def _open_tdoc_document_directly(self, tdoc_id: str):
        """
        Downloads/opens a TDoc contribution (or revision) directly without stealing focus from email dialogs.
        Automatically routes external / cross-WG TDocs (e.g. R3-263595, older SA2 TDocs) via
        global_action_requested -> MeetingsTab instead of failing against this meeting's FTP folder.
        """
        td = str(tdoc_id or "").strip().upper()
        if not td:
            return
        match = re.search(r"^(.*?)-?(?:r|rev)\d{1,2}[a-zA-Z]?$", td, re.IGNORECASE)
        base_tdoc = match.group(1).upper() if match else td
        valid_tdocs = getattr(self.model, "valid_tdocs", set()) if hasattr(self, "model") else set()
        in_current_meeting = (base_tdoc in valid_tdocs) or (td in valid_tdocs)

        if not in_current_meeting:
            try:
                from modules.meetings.core.meetings_db import MeetingsDatabase
                db_path = get_project_root() / "db" / "3gpp_data.db"
                if db_path.is_file():
                    ext_mtg = MeetingsDatabase(db_path).find_meeting_by_tdoc(base_tdoc)
                    if ext_mtg and ext_mtg.get("mtg_id") != self.mtg_info.get("mtg_id"):
                        self.global_action_requested.emit(td, "open_doc")
                        return
            except Exception as exc:
                logging.debug(f"Global TDoc lookup skipped for {td}: {exc}")

            cur_prefix = str(self._get_tdoc_prefix() or "").upper()
            if cur_prefix and not base_tdoc.startswith(cur_prefix):
                QMessageBox.information(
                    self,
                    "External TDoc",
                    f"{td} does not belong to {self._get_meeting_id_label()} and its meeting "
                    f"could not be resolved in the local Meetings Database.\n\n"
                    f"Sync that Working Group / meeting in the Meetings tab to open it directly.",
                )
                return

        is_rev = bool(re.search(r"(?:r|rev)\d{1,2}[a-zA-Z]?$", td, re.IGNORECASE))
        target_url = (
            getattr(self, "revisions_url", "")
            if is_rev and getattr(self, "revisions_url", "")
            else getattr(self, "docs_ftp_url", "")
        )
        if not target_url:
            self._scroll_to_tdoc(td)
            return
        self._trigger_download_thread(base_tdoc, td, target_url, is_silent_compare=False)


    def _open_tdoc_triage(self, tdoc_id: str):
        """Opens a modeless AI Triage dialog for the selected TDoc."""
        if not tdoc_id:
            return
        tdoc_clean = tdoc_id.strip().upper()

        if tdoc_clean in self._triage_dialogs:
            dlg = self._triage_dialogs[tdoc_clean]
            dlg.setWindowState(dlg.windowState() & ~Qt.WindowMinimized | Qt.WindowActive)
            dlg.raise_()
            dlg.activateWindow()
            return

        # Find the row data for this TDoc or its base
        row_data = next(
            (r for r in self.model._data if str(r.get("TDoc", "")).strip().upper() == tdoc_clean),
            None
        )
        if not row_data:
            match = re.search(r'^(.*?)-?(?:r|rev)\d{1,2}[a-zA-Z]?$', tdoc_clean, re.IGNORECASE)
            base_tdoc = match.group(1).upper() if match else tdoc_clean
            row_data = next(
                (r for r in self.model._data if str(r.get("TDoc", "")).strip().upper() == base_tdoc),
                {"TDoc": tdoc_clean}
            )

        dialog = TDocTriageDialog(
            tdoc_data=row_data,
            meeting_dir=self.meeting_dir,
            docs_ftp_url=self.docs_ftp_url,
            revisions_url=self.revisions_url,
            save_callback=self._save_user_data,
            parent=None  # Independent top-level modeless window
        )
        self._triage_dialogs[tdoc_clean] = dialog
        dialog.finished.connect(lambda _, t=tdoc_clean: self._triage_dialogs.pop(t, None))
        dialog.show()
