# --- File: modules/meetings/ui/search_controller.py ---
import re
from pathlib import Path
from PyQt5.QtWidgets import QApplication, QMessageBox, QPushButton
from modules.meetings.core.tdocs_threads import TDocActionThread


class GlobalSearchController:
    def __init__(self, main_tab):
        self.tab = main_tab  # Store a reference to the main MeetingsTab
        self.current_found_meeting = None

    def connect_signals(self):
        """Wires up the UI elements from the main tab."""
        self.tab.global_tdoc_input.textChanged.connect(self.on_tdoc_input_changed)
        self.tab.global_tdoc_input.returnPressed.connect(self.action_open_tdoc_only)
        self.tab.btn_open_tdoc.clicked.connect(self.action_open_tdoc_only)
        self.tab.btn_open_meeting.clicked.connect(self.action_open_meeting_list)

    def on_tdoc_input_changed(self, text):
        text = text.strip()
        match = re.match(r'^([A-Za-z0-9]+-\d+)(r\d+[a-zA-Z]?)?$', text, re.IGNORECASE)

        if not match:
            self.tab.btn_open_tdoc.setVisible(False)
            self.tab.btn_open_meeting.setVisible(False)
            self.current_found_meeting = None
            return

        meeting = self.tab.db.find_meeting_by_tdoc(text)
        if meeting:
            self.current_found_meeting = meeting
            self.tab.btn_open_tdoc.setVisible(True)
            self.tab.btn_open_meeting.setVisible(True)

            mtg_name = f"{meeting.get('wg_name', '')} {meeting.get('meeting_number', '')}"
            self.tab.btn_open_meeting.setToolTip(f"Open the full TDocs table for {mtg_name}")
            self.tab.btn_open_tdoc.setToolTip(f"Instantly download {text.upper()} from {mtg_name}")
        else:
            self.tab.btn_open_tdoc.setVisible(False)
            self.tab.btn_open_meeting.setVisible(False)
            self.current_found_meeting = None

    def action_open_tdoc_only(self):
        if not self.tab.btn_open_tdoc.isVisible() or not self.current_found_meeting:
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

    def action_open_meeting_list(self):
        if not self.current_found_meeting: return
        meeting = self.current_found_meeting

        self.tab.btn_open_meeting.setText("⏳ Load...")
        self.tab.btn_open_meeting.setEnabled(False)
        QApplication.processEvents()

        filepath = self.tab._get_tdoc_list_path(meeting)
        if filepath and filepath.exists():
            self.tab._open_tdocs_window(meeting, str(filepath))
        else:
            dummy_btn = QPushButton()
            self.tab._download_and_open_tdocs(meeting, dummy_btn)

        self.tab.btn_open_meeting.setText("🗓️ Mtg")
        self.tab.btn_open_meeting.setEnabled(True)

    def _download_global_tdoc(self, meeting: dict, base_tdoc: str, target_filename: str, has_rev: bool,
                              silent_cart: bool = False):
        docs_url = meeting.get("docs_folder_url")
        if not docs_url: return
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
        thread.target_filename = target_filename

        self.tab._active_dl_threads[target_filename] = thread
        thread.finished_action.connect(
            lambda t, s, m, th=thread: self._on_global_tdoc_download_finished(target_filename, s, m, th)
        )
        thread.start()

    def _on_global_tdoc_download_finished(self, tdoc_name: str, success: bool, msg: str, thread: TDocActionThread):
        if tdoc_name in self.tab._active_dl_threads:
            del self.tab._active_dl_threads[tdoc_name]

        self.tab.btn_open_tdoc.setText("📄 Doc")
        self.tab.btn_open_tdoc.setEnabled(True)
        self.tab.global_tdoc_input.setEnabled(True)

        if not success:
            QMessageBox.warning(self.tab, f"Action Failed: {tdoc_name}", msg)
            return

        # Automatically push to the comparison cart if it was a silent request
        if getattr(thread, "is_silent_compare", False):
            from modules.meetings.core.compare_manager import ComparisonManager
            extracted_files = getattr(thread, "extracted_doc_paths", [])
            if extracted_files:
                ComparisonManager.get_instance().add_to_cart(thread.target_filename, str(extracted_files[0]))
            else:
                QMessageBox.warning(self.tab, "Compare Failed", "No Word document found inside this TDoc ZIP.")