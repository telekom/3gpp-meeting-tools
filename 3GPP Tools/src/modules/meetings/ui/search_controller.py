# --- File: modules/meetings/ui/search_controller.py ---
import os
import re
import logging
from pathlib import Path
from PyQt5.QtCore import QPoint
from PyQt5.QtWidgets import QApplication, QMessageBox, QPushButton
from modules.meetings.core.tdocs.tdocs_threads import TDocActionThread

logger = logging.getLogger(__name__)


class GlobalSearchController:
    def __init__(self, main_tab):
        self.tab = main_tab  # Store a reference to the main MeetingsTab
        self.current_found_meeting = None
        self.current_incomplete_meetings = []

    def connect_signals(self):
        """Wires up the UI elements from the main tab."""
        self.tab.global_tdoc_input.textChanged.connect(self.on_tdoc_input_changed)
        self.tab.global_tdoc_input.returnPressed.connect(self.action_open_tdoc_only)
        self.tab.btn_open_tdoc.clicked.connect(self.action_open_tdoc_only)
        self.tab.btn_open_meeting.clicked.connect(self.action_open_meeting_list)
        if hasattr(self.tab, "btn_compare_active_word"):
            self.tab.btn_compare_active_word.clicked.connect(self.action_compare_with_active_word)
        if hasattr(self.tab, "btn_copy_tdoc"):
            self.tab.btn_copy_tdoc.clicked.connect(self.action_show_copy_menu)

    def on_tdoc_input_changed(self, text):
        text = text.strip()
        match = re.match(r'^([A-Za-z0-9]+-\d+)(r\d+[a-zA-Z]?)?$', text, re.IGNORECASE)

        self.current_found_meeting = None
        self.current_incomplete_meetings = []
        self.tab.btn_open_tdoc.setVisible(False)
        self.tab.btn_open_meeting.setVisible(False)
        if hasattr(self.tab, "btn_compare_active_word"):
            self.tab.btn_compare_active_word.setVisible(False)
        if hasattr(self.tab, "btn_copy_tdoc"):
            self.tab.btn_copy_tdoc.setVisible(False)
        self.tab._set_tdoc_metadata_warning([])

        if not match:
            return

        meeting = self.tab.db.find_meeting_by_tdoc(text)
        if meeting:
            self.current_found_meeting = meeting
            self.tab.btn_open_tdoc.setVisible(True)
            self.tab.btn_open_meeting.setVisible(True)
            if hasattr(self.tab, "btn_compare_active_word"):
                self.tab.btn_compare_active_word.setVisible(True)
            if hasattr(self.tab, "btn_copy_tdoc"):
                self.tab.btn_copy_tdoc.setVisible(True)

            mtg_name = f"{meeting.get('wg_name', '')} {meeting.get('meeting_number', '')}"
            self.tab.btn_open_meeting.setToolTip(f"Open the full TDocs table for {mtg_name}")
            self.tab.btn_open_tdoc.setToolTip(f"Instantly download {text.upper()} from {mtg_name}")
            if hasattr(self.tab, "btn_compare_active_word"):
                self.tab.btn_compare_active_word.setToolTip(
                    f"Download {text.upper()} ({mtg_name}) and immediately compare it with your currently active Microsoft Word document"
                )
            if hasattr(self.tab, "btn_copy_tdoc"):
                self.tab.btn_copy_tdoc.setToolTip(
                    f"Copy {text.upper()} ({mtg_name}) number, title, citation, or direct 3GPP download URLs to clipboard"
                )
            return

        # Strict lookup failed. Check same-WG/same-year meetings for missing
        # Docs-derived first/last TDoc metadata.
        self.current_incomplete_meetings = self.tab.db.find_incomplete_tdoc_meetings(text)
        self.tab._set_tdoc_metadata_warning(self.current_incomplete_meetings)

    def action_open_tdoc_only(self):
        # 🐛 FIX: Removed the .isVisible() check. If this is triggered programmatically
        # from the Work Items tab, the button is technically off-screen and would return False!
        if not self.current_found_meeting:
            return

        tdoc_str = self.tab.global_tdoc_input.text().strip()
        match = re.match(r'^([A-Za-z0-9]+-\d+)(r\d+[a-zA-Z]?)?$', tdoc_str, re.IGNORECASE)
        if not match: return

        base_tdoc = match.group(1).upper()
        target_filename = (base_tdoc + (match.group(2) or "")).upper()

        if target_filename in self.tab._active_dl_threads: return

        self.tab.btn_open_tdoc.setText("⏳ Get...")
        self.tab.btn_open_tdoc.setEnabled(False)
        self.tab.global_tdoc_input.setEnabled(False)
        QApplication.processEvents()

        self._download_global_tdoc(self.current_found_meeting, base_tdoc, target_filename, has_rev=bool(match.group(2)))

    def action_add_to_cart(self):
        """Silently downloads a document and adds it to the Word Comparison Cart."""
        if not self.current_found_meeting: return

        tdoc_str = self.tab.global_tdoc_input.text().strip()
        match = re.match(r'^([A-Za-z0-9]+-\d+)(r\d+[a-zA-Z]?)?$', tdoc_str, re.IGNORECASE)
        if not match: return

        base_tdoc = match.group(1).upper()
        target_filename = (base_tdoc + (match.group(2) or "")).upper()

        if target_filename in self.tab._active_dl_threads: return

        self._download_global_tdoc(self.current_found_meeting, base_tdoc, target_filename, has_rev=bool(match.group(2)),
                                   silent_cart=True)

    def action_compare_with_active_word(self):
        """
        1-click comparison of the currently active Microsoft Word document (Slot A)
        against the TDoc entered in the Quick TDoc Jump box (Slot B).
        """
        if not self.current_found_meeting:
            return

        from modules.meetings.core.compare_manager import ComparisonManager

        tdoc_str = self.tab.global_tdoc_input.text().strip()
        match = re.match(r'^([A-Za-z0-9]+-\d+)(r\d+[a-zA-Z]?)?$', tdoc_str, re.IGNORECASE)
        if not match:
            return

        base_tdoc = match.group(1).upper()
        target_filename = (base_tdoc + (match.group(2) or "")).upper()
        if target_filename in self.tab._active_dl_threads:
            return

        # 1. Capture active Word document immediately on click
        mgr = ComparisonManager.get_instance()
        active_doc = mgr.stage_active_word_document(slot="A")
        if not active_doc:
            QMessageBox.warning(
                self.tab,
                "No Active Word Document",
                "No active Microsoft Word document could be detected.\n\n"
                "Please ensure Microsoft Word is running and the document you wish to compare is open on your screen.",
            )
            return

        if hasattr(self.tab, "btn_compare_active_word"):
            self.tab.btn_compare_active_word.setText("⏳ Diff...")
            self.tab.btn_compare_active_word.setEnabled(False)
            QApplication.processEvents()

        self._download_global_tdoc(
            self.current_found_meeting,
            base_tdoc,
            target_filename,
            has_rev=bool(match.group(2)),
            silent_cart=True,
            auto_compare_active_word=True,
        )

    def action_open_meeting_list(self):
        if not self.current_found_meeting:
            return
        meeting = self.current_found_meeting

        self.tab.btn_open_meeting.setText("⏳ Load...")
        self.tab.btn_open_meeting.setEnabled(False)
        QApplication.processEvents()

        filepath = self.tab._get_tdoc_list_path(meeting)
        if filepath and filepath.exists():
            self.tab._open_tdocs_window(meeting, str(filepath))
        else:
            self.tab._download_and_open_tdocs(meeting)

        self.tab.btn_open_meeting.setText("🗓️ Mtg")
        self.tab.btn_open_meeting.setEnabled(True)

    def action_show_copy_menu(self):
        """Displays the 3GPP Quick Copy dropdown menu for the current TDoc and meeting."""
        if not self.current_found_meeting:
            return
        tdoc_str = self.tab.global_tdoc_input.text().strip()
        if not tdoc_str:
            return

        from modules.meetings.ui.tdocs_menus import build_quick_copy_menu

        meeting = self.current_found_meeting
        row_data = self._resolve_tdoc_row_data(meeting, tdoc_str)
        docs_url = meeting.get("docs_folder_url") or ""
        revisions_url = meeting.get("revisions_url") or ""
        main_url = meeting.get("url_key") or ""

        menu = build_quick_copy_menu(
            parent=self.tab.btn_copy_tdoc if hasattr(self.tab, "btn_copy_tdoc") else self.tab,
            tdoc_id=tdoc_str,
            row_data=row_data,
            docs_ftp_url=docs_url,
            revisions_url=revisions_url,
            main_url=main_url,
        )
        if hasattr(self.tab, "btn_copy_tdoc") and self.tab.btn_copy_tdoc.isVisible():
            menu.exec_(self.tab.btn_copy_tdoc.mapToGlobal(QPoint(0, self.tab.btn_copy_tdoc.height())))
        else:
            menu.exec_(self.tab.global_tdoc_input.mapToGlobal(QPoint(0, self.tab.global_tdoc_input.height())))

    def _resolve_tdoc_row_data(self, meeting: dict, tdoc_str: str) -> dict:
        """
        Looks up cached row metadata (Title, Source, etc.) for tdoc_str if available.
        Checks active TDocsWindow (in-memory) first, then local JSON cache and Excel.
        """
        clean_tdoc = tdoc_str.strip().upper()
        base_match = re.match(r"^([A-Za-z0-9]+-\d+)", clean_tdoc)
        base_tdoc = base_match.group(1).upper() if base_match else clean_tdoc

        # 1. Check open TDoc windows for instant in-memory hit
        mtg_id = meeting.get("mtg_id")
        if hasattr(self.tab, "tdoc_windows") and mtg_id in self.tab.tdoc_windows:
            win = self.tab.tdoc_windows.get(mtg_id)
            if win and hasattr(win, "model") and hasattr(win.model, "_data"):
                for row in win.model._data:
                    if str(row.get("TDoc", "")).strip().upper() == base_tdoc:
                        return dict(row)

        # 2. Check cached file on disk (via fast JSON cache if present)
        if hasattr(self.tab, "_get_tdoc_list_path"):
            filepath = self.tab._get_tdoc_list_path(meeting)
            if filepath and filepath.exists():
                json_cache = str(filepath) + ".json"
                if os.path.exists(json_cache):
                    try:
                        import json
                        with open(json_cache, "r", encoding="utf-8") as f:
                            cached = json.load(f)
                        if isinstance(cached, list):
                            for row in cached:
                                if str(row.get("TDoc", "")).strip().upper() == base_tdoc:
                                    return dict(row)
                    except Exception as e:
                        logger.debug(f"Could not read TDoc JSON cache: {e}")

                # If no JSON cache yet, attempt light parse
                try:
                    from modules.meetings.core.tdocs.tdocs_parser import TDocsParser
                    parsed = TDocsParser.parse_tdocs_excel(str(filepath))
                    if parsed:
                        for row in parsed:
                            if str(row.get("TDoc", "")).strip().upper() == base_tdoc:
                                return dict(row)
                except Exception as e:
                    logger.debug(f"Could not parse Excel for TDoc row data: {e}")

        return {"TDoc": clean_tdoc}

    def _download_global_tdoc(self, meeting: dict, base_tdoc: str, target_filename: str, has_rev: bool,
                              silent_cart: bool = False, auto_compare_active_word: bool = False):
        docs_url = meeting.get("docs_folder_url")
        if not docs_url:
            logger.warning(f"[GlobalSearch] No Docs URL available for {target_filename}.")
            if hasattr(self.tab, "btn_compare_active_word"):
                self.tab.btn_compare_active_word.setText("⚖️ vs Word")
                self.tab.btn_compare_active_word.setEnabled(True)
            return
        if not docs_url.startswith("http"): docs_url = "https://www.3gpp.org/ftp/" + docs_url.lstrip('/')

        current_cache = self.tab.dl_dir_input.text().strip() if hasattr(self.tab,
                                                                        'dl_dir_input') else self.tab.settings.cache_dir
        folder_name = meeting.get("folder_name") or meeting.get("meeting_number", "")
        meeting_dir = Path(current_cache) / folder_name

        dl_url = docs_url
        if has_rev:
            raw_url = meeting.get("url_key", "")
            main_ftp = raw_url if raw_url.startswith("http") else f"https://www.3gpp.org/ftp/{raw_url.lstrip('/')}"
            dl_url = main_ftp.rstrip('/') + '/INBOX/Revisions/'

        # Trigger background thread. Do NOT auto-open Word if it's a silent cart operation.
        thread = TDocActionThread(base_tdoc, target_filename, dl_url, meeting_dir, open_file=not silent_cart)

        # Tag the thread so the callback knows what to do
        thread.is_silent_compare = silent_cart
        thread.auto_compare_active_word = auto_compare_active_word
        thread.target_filename = target_filename

        self.tab._active_dl_threads[target_filename] = thread
        thread.finished_action.connect(
            lambda t, s, m, th=thread: self._on_global_tdoc_download_finished(target_filename, s, m, th)
        )
        thread.start()

    def _on_global_tdoc_download_finished(self, tdoc_name: str, success: bool, msg: str, thread: TDocActionThread):
        if tdoc_name in self.tab._active_dl_threads:
            th = self.tab._active_dl_threads.pop(tdoc_name, None)
            if th and th.isRunning():
                th.wait(50)

        self.tab.btn_open_tdoc.setText("📄 Doc")
        self.tab.btn_open_tdoc.setEnabled(True)
        self.tab.global_tdoc_input.setEnabled(True)
        if hasattr(self.tab, "btn_compare_active_word"):
            self.tab.btn_compare_active_word.setText("⚖️ vs Word")
            self.tab.btn_compare_active_word.setEnabled(True)

        if not success:
            logger.error(f"[GlobalSearch] TDoc action failed for {tdoc_name}: {msg}")
            QMessageBox.warning(self.tab, f"Action Failed: {tdoc_name}", msg)
            return

        # Automatically push to the comparison cart if it was a silent request
        if getattr(thread, "is_silent_compare", False):
            from modules.meetings.core.compare_manager import ComparisonManager
            from modules.word_tools.core.document_opener import select_best_word_document

            extracted_files = getattr(thread, "extracted_doc_paths", [])
            best_word = select_best_word_document(extracted_files, thread.target_filename)
            if best_word:
                mgr = ComparisonManager.get_instance()
                if getattr(thread, "auto_compare_active_word", False):
                    mgr.set_slot_b(thread.target_filename, str(best_word))
                    if hasattr(self.tab, "_run_comparison"):
                        self.tab._run_comparison()
                else:
                    mgr.add_to_cart(thread.target_filename, str(best_word))
            else:
                found_exts = ", ".join(sorted({Path(f).suffix.lower() for f in extracted_files})) or "none"
                QMessageBox.warning(
                    self.tab,
                    "Compare Failed",
                    f"No Word document (.docx / .doc) found inside {thread.target_filename}.zip (found: {found_exts}).",
                )