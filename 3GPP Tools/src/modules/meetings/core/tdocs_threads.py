# --- File: src/modules/meetings/core/tdocs_threads.py ---
import logging
import re
import json
from pathlib import Path

import requests
from PyQt5.QtCore import QThread, pyqtSignal

from core.network.session import NetworkSession
from modules.meetings.core.tdocs_parser import TDocsParser
from modules.meetings.core.tdoc_file_handler import TDocFileHandler


class TDocsRevisionsFetcherThread(QThread):
    finished = pyqtSignal(bool, dict, str)

    def __init__(self, url: str, meeting_dir: Path = None):
        super().__init__()
        self.url = url
        self.meeting_dir = meeting_dir

    def run(self):
        try:
            session = NetworkSession.get_instance()
            NetworkSession.apply_humanness(session)
            response = session.get(self.url, timeout=30)
            response.raise_for_status()

            html = response.text
            pattern = re.compile(r'href=["\']?(?:[^"\'>]*/)?(([A-Za-z0-9\-]+)(r\d+[a-zA-Z]?)\.zip)["\']?',
                                 re.IGNORECASE)
            matches = pattern.findall(html)

            revisions = {}
            for full_file, base_tdoc, rev_str in matches:
                base_tdoc = base_tdoc.upper()
                rev_str = rev_str.lower()
                if base_tdoc not in revisions:
                    revisions[base_tdoc] = []
                if rev_str not in revisions[base_tdoc]:
                    revisions[base_tdoc].append(rev_str)

            for k in revisions:
                revisions[k].sort()

            if self.meeting_dir:
                try:
                    agenda_dir = self.meeting_dir / "Agenda"
                    agenda_dir.mkdir(parents=True, exist_ok=True)
                    rev_file = agenda_dir / "revisions.json"
                    with open(rev_file, "w", encoding="utf-8") as f:
                        json.dump(revisions, f, indent=4)
                except Exception as e:
                    logging.warning(f"Failed to cache revisions locally: {e}")

            self.finished.emit(True, revisions, "Success")

        except requests.exceptions.HTTPError as e:
            if e.response.status_code == 404:
                self.finished.emit(True, {}, "No Revisions folder found.")
            else:
                self.finished.emit(False, {}, str(e))
        except Exception as e:
            self.finished.emit(False, {}, str(e))


# --- Inside: src/modules/meetings/core/tdocs_threads.py ---
class TDocActionThread(QThread):
    finished_action = pyqtSignal(str, bool, str)

    def __init__(self, base_tdoc: str, target_filename: str, base_urls, meeting_dir: Path, open_file: bool = True):
        super().__init__()
        self.base_tdoc = base_tdoc
        self.target_filename = target_filename
        # Accept a single URL (legacy) or a priority list of URLs
        self.base_urls = base_urls if isinstance(base_urls, list) else [base_urls]
        self.tdoc_dir = meeting_dir / base_tdoc
        self.open_file = open_file
        self.extracted_doc_paths = []

    def run(self):
        success = False
        last_err = "No valid URLs provided."

        for url in self.base_urls:
            try:
                self.extracted_doc_paths = TDocFileHandler.download_and_extract_tdoc(
                    self.target_filename, url, self.tdoc_dir
                )
                if self.extracted_doc_paths:
                    success = True
                    break  # Break out of the fallback loop on success!
            except requests.exceptions.HTTPError as e:
                if e.response.status_code == 404:
                    last_err = f"404 Not Found at {url}"
                    continue  # File isn't here, try the next fallback URL!
                last_err = str(e)
                continue
            except Exception as e:
                last_err = str(e)
                continue

        if not success:
            self.finished_action.emit(self.base_tdoc, False, f"Could not retrieve document.\nLast error: {last_err}")
            return

        if self.open_file:
            import os, webbrowser
            for doc in self.extracted_doc_paths:
                if hasattr(os, 'startfile'):
                    os.startfile(str(doc))
                else:
                    webbrowser.open(f"file:///{doc}")

        msg = "Opened successfully." if self.open_file else "Downloaded & Added successfully."
        self.finished_action.emit(self.base_tdoc, True, msg)


class TdocsByAgendaThread(QThread):
    ui_log_msg = pyqtSignal(str, int)
    finished = pyqtSignal(bool, dict)

    def __init__(self, meeting_ftp_url: str, local_folder: Path):
        super().__init__()
        self.meeting_ftp_url = meeting_ftp_url
        self.local_folder = local_folder

    def run(self):
        try:
            self.ui_log_msg.emit("⏳ Initiating TdocsByAgenda Sync...", logging.INFO)
            clean_base_url = self.meeting_ftp_url.rstrip('/')

            session = NetworkSession.get_instance()
            NetworkSession.apply_humanness(session)

            self.ui_log_msg.emit("🔍 Searching FTP for TdocsByAgenda file...", logging.INFO)
            response = session.get(clean_base_url, timeout=30)
            response.raise_for_status()

            pattern = re.compile(r'href=["\']?([^"\'>]*tdocsbyagenda[^"\'>]*\.html?)["\']?', re.IGNORECASE)
            matches = pattern.findall(response.text)

            if not matches:
                self.ui_log_msg.emit("❌ Could not find any TdocsByAgenda file on the FTP server.", logging.ERROR)
                self.finished.emit(False, {})
                return

            target_filename = matches[-1].split('/')[-1]
            agenda_url = f"{clean_base_url}/{target_filename}"

            agenda_dir = self.local_folder / "Agenda"
            agenda_dir.mkdir(parents=True, exist_ok=True)
            agenda_path = agenda_dir / "TdocsByAgenda.htm"

            self.ui_log_msg.emit(f"⬇️ Downloading: {agenda_url}", logging.INFO)
            NetworkSession.download_file(agenda_url, agenda_path)

            agenda_data = TDocsParser.parse_tdocs_by_agenda(str(agenda_path), self.ui_log_msg)
            self.finished.emit(True, agenda_data)

        except Exception as e:
            self.ui_log_msg.emit(f"❌ Failed to sync TdocsByAgenda: {str(e)}", logging.ERROR)
            self.finished.emit(False, {})