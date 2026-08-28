# --- File: src/modules/meetings/core/tdocs_threads.py ---
import json
import logging
import re
import os
import shutil
from pathlib import Path

import requests
from PyQt5.QtCore import QThread, pyqtSignal

from core.network.session import NetworkSession
from modules.meetings.core.tdoc_file_handler import TDocFileHandler
from modules.meetings.core.tdocs_parser import TDocsParser

from modules.word_tools.core.libreoffice_converter import (
    convert_doc_to_docx_libreoffice,
    is_libreoffice_available,
    get_libreoffice_missing_msg,
)


class TDocsRevisionsFetcherThread(QThread):
    finished = pyqtSignal(bool, dict, str)

    def __init__(self, url: str, meeting_dir: Path = None):
        super().__init__()
        self.url = url
        self.meeting_dir = meeting_dir

    def run(self):
        try:
            logging.info(f"🔍 [Revisions Sync] Fetching revision listings from: {self.url}")
            session = NetworkSession.get_instance()
            NetworkSession.apply_humanness(session)
            response = session.get(self.url, timeout=15)
            response.raise_for_status()

            html = response.text
            pattern = re.compile(
                r'href=["\']?(?:[^"\'>]*/)?(([A-Za-z0-9\-]+)(r\d+[a-zA-Z]?)\.zip)["\']?',
                re.IGNORECASE,
            )
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

            logging.info(
                f"✅ [Revisions Sync] Discovered revisions for {len(revisions)} base TDocs."
            )

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
                logging.warning(
                    f"⚠️ [Revisions Sync] Revisions directory 404 Not Found at {self.url}"
                )
                self.finished.emit(True, {}, "No Revisions folder found.")
            else:
                logging.error(f"❌ [Revisions Sync] HTTP Error: {e}")
                self.finished.emit(False, {}, str(e))
        except Exception as e:
            logging.error(f"❌ [Revisions Sync] Network Exception: {e}")
            self.finished.emit(False, {}, str(e))


class TDocActionThread(QThread):
    finished_action = pyqtSignal(str, bool, str)

    def __init__(
        self,
        base_tdoc: str,
        target_filename: str,
        base_urls,
        meeting_dir: Path,
        open_file: bool = True,
    ):
        super().__init__()
        self.base_tdoc = base_tdoc
        self.target_filename = target_filename
        self.base_urls = (
            base_urls if isinstance(base_urls, list) else [base_urls]
        )
        self.tdoc_dir = meeting_dir / base_tdoc
        self.open_file = open_file
        self.extracted_doc_paths = []

    def run(self):
        success = False
        last_err = "No valid URLs provided."

        logging.info("-" * 65)
        logging.info(
            f"🚀 [Action Thread] Initiating retrieval for '{self.target_filename}'"
        )
        logging.info(
            f"📋 [Priority Queue] Evaluating {len(self.base_urls)} candidate route(s):"
        )
        for idx, u in enumerate(self.base_urls, 1):
            logging.info(f"   [{idx}] {u}")
        logging.info("-" * 65)

        for idx, url in enumerate(self.base_urls, 1):
            try:
                logging.info(
                    f"➡️ [Route {idx}/{len(self.base_urls)}] Trying: {url}/{self.target_filename}.zip"
                )
                self.extracted_doc_paths = (
                    TDocFileHandler.download_and_extract_tdoc(
                        self.target_filename, url, self.tdoc_dir, timeout=6
                    )
                )
                if self.extracted_doc_paths:
                    success = True
                    logging.info(
                        f"🎯 [MATCH] Successfully acquired '{self.target_filename}' from Route [{idx}]."
                    )
                    break
            except requests.exceptions.HTTPError as e:
                if e.response.status_code == 404:
                    last_err = f"404 Not Found at {url}"
                    logging.info(f"   ↳ ❌ [404] Not present on route: {url}")
                    continue
                last_err = f"HTTP {e.response.status_code} Error: {e}"
                logging.warning(f"   ↳ ⚠️ [HTTP ERROR] {last_err}")
                continue
            except requests.exceptions.ConnectTimeout:
                last_err = f"Connection timed out at {url}"
                logging.warning(f"   ↳ ⏱️ [TIMEOUT] Route unreachable: {url}")
                continue
            except Exception as e:
                last_err = str(e)
                logging.warning(f"   ↳ ⚠️ [ERROR] Failed route {url}: {e}")
                continue

        if not success:
            logging.error(
                f"❌ [FETCH FAILED] All {len(self.base_urls)} candidate routes exhausted for '{self.target_filename}'."
            )
            self.finished_action.emit(
                self.base_tdoc,
                False,
                f"Could not retrieve document.\nLast error: {last_err}",
            )
            return

        if self.open_file:
            import os
            import webbrowser

            for doc in self.extracted_doc_paths:
                if hasattr(os, "startfile"):
                    os.startfile(str(doc))
                else:
                    webbrowser.open(f"file:///{doc}")

        msg = (
            "Opened successfully."
            if self.open_file
            else "Downloaded & Added successfully."
        )
        self.finished_action.emit(self.base_tdoc, True, msg)


class TdocsByAgendaThread(QThread):
    ui_log_msg = pyqtSignal(str, int)
    finished = pyqtSignal(bool, dict)

    def __init__(self, candidate_urls, local_folder: Path):
        super().__init__()
        # Accept either a single string URL or a list of candidate folder URLs
        self.candidate_urls = (
            candidate_urls
            if isinstance(candidate_urls, list)
            else [candidate_urls]
        )
        self.local_folder = local_folder

    def run(self):
        session = NetworkSession.get_instance()
        NetworkSession.apply_humanness(session)

        logging.info("=" * 65)
        logging.info("📄 [TdocsByAgenda Sync] Starting agenda retrieval...")
        logging.info(
            f"📋 [Priority Queue] Evaluating {len(self.candidate_urls)} candidate directory routes:"
        )
        for idx, u in enumerate(self.candidate_urls, 1):
            logging.info(f"   [{idx}] {u}")
        logging.info("=" * 65)

        pattern = re.compile(
            r'href=["\']?([^"\'>]*tdocsbyagenda[^"\'>]*\.html?)["\']?',
            re.IGNORECASE,
        )
        found_target_url = None
        last_error = "No candidate routes provided."

        for idx, folder_url in enumerate(self.candidate_urls, 1):
            clean_url = folder_url.rstrip("/")
            logging.info(
                f"🔍 [Route {idx}/{len(self.candidate_urls)}] Scanning directory for TdocsByAgenda: {clean_url}"
            )

            try:
                response = session.get(clean_url, timeout=6)
                if response.status_code != 200:
                    logging.info(
                        f"   ↳ ❌ HTTP {response.status_code} at {clean_url}"
                    )
                    continue

                matches = pattern.findall(response.text)
                if matches:
                    target_filename = matches[-1].split("/")[-1]
                    found_target_url = f"{clean_url}/{target_filename}"
                    logging.info(
                        f"🎯 [MATCH] Found agenda file '{target_filename}' at {clean_url}"
                    )
                    break
                else:
                    logging.info(
                        f"   ↳ ⚠️ Directory accessible but no TdocsByAgenda.htm link found."
                    )

            except Exception as e:
                last_error = str(e)
                logging.warning(
                    f"   ↳ ⏱️ Directory check failed for {clean_url}: {e}"
                )
                continue

        if not found_target_url:
            logging.error(
                f"❌ [TdocsByAgenda Sync Failed] Could not find TdocsByAgenda in any candidate location. Last error: {last_error}"
            )
            self.finished.emit(False, {})
            return

        try:
            agenda_dir = self.local_folder / "Agenda"
            agenda_dir.mkdir(parents=True, exist_ok=True)
            agenda_path = agenda_dir / "TdocsByAgenda.htm"

            logging.info(
                f"⬇️ [Downloading Agenda] Fetching: {found_target_url}"
            )
            NetworkSession.download_file(found_target_url, agenda_path)
            logging.info(
                f"💾 [Agenda Saved] Saved locally to: {agenda_path.resolve()}"
            )

            agenda_data = TDocsParser.parse_tdocs_by_agenda(
                str(agenda_path), self.ui_log_msg
            )
            logging.info(
                f"✅ [Parsing Complete] Successfully parsed {len(agenda_data)} items from TdocsByAgenda."
            )
            self.finished.emit(True, agenda_data)

        except Exception as e:
            logging.error(
                f"❌ [TdocsByAgenda Parse Failed] Error downloading/parsing: {e}"
            )
            self.finished.emit(False, {})

def _unblock_file(path: Path):
    """Removes NTFS Mark-of-the-Web alternate data stream on Windows."""
    if os.name == "nt":
        zone_stream = Path(f"{path.resolve()}:Zone.Identifier")
        try:
            if zone_stream.exists():
                zone_stream.unlink()
        except Exception:
            pass


class WordAgendaImporterThread(QThread):
    """
    Background worker that stages an imported Chairman's Notes/TDocs file,
    converts legacy .doc files to .docx via LibreOffice, and parses the table.
    """
    progress = pyqtSignal(str)
    finished = pyqtSignal(bool, dict, str, str)  # (success, agenda_data, filename, error_message)

    def __init__(self, source_path: Path, meeting_dir: Path):
        super().__init__()
        self.source_path = Path(source_path)
        self.meeting_dir = Path(meeting_dir)

    def run(self):
        try:
            if not self.source_path.exists():
                self.finished.emit(False, {}, self.source_path.name, f"File not found:\n{self.source_path}")
                return

            self.progress.emit("Staging file...")
            agenda_dir = self.meeting_dir / "Agenda" / "ChairNotes"
            agenda_dir.mkdir(parents=True, exist_ok=True)

            target_path = agenda_dir / self.source_path.name
            if self.source_path.resolve() != target_path.resolve():
                try:
                    shutil.copy2(str(self.source_path), str(target_path))
                except Exception as e:
                    self.finished.emit(False, {}, self.source_path.name, f"Failed to copy file to {agenda_dir}:\n{e}")
                    return

            _unblock_file(target_path)

            ext = target_path.suffix.lower()
            agenda_data = {}

            if ext == ".doc":
                if not is_libreoffice_available():
                    self.finished.emit(False, {}, target_path.name, get_libreoffice_missing_msg())
                    return

                self.progress.emit("Converting via LibreOffice...")
                try:
                    docx_path = convert_doc_to_docx_libreoffice(
                        doc_path=target_path,
                        output_path=target_path.with_suffix(".docx"),
                    )
                    if not docx_path or not docx_path.exists():
                        self.finished.emit(False, {}, target_path.name, "LibreOffice conversion produced no output file.")
                        return

                    self.progress.emit("Parsing converted document...")
                    agenda_data = TDocsParser.parse_tdocs_from_docx(str(docx_path))
                except Exception as e:
                    self.finished.emit(False, {}, target_path.name, f"LibreOffice conversion failed:\n{e}")
                    return

            elif ext == ".docx":
                self.progress.emit("Parsing .docx table...")
                agenda_data = TDocsParser.parse_tdocs_from_docx(str(target_path))

            elif ext in [".htm", ".html"]:
                self.progress.emit("Parsing HTML agenda...")
                agenda_data = TDocsParser.parse_tdocs_by_agenda(str(target_path))

            if not agenda_data:
                self.finished.emit(
                    False, {}, target_path.name, f"No valid TDocs table could be extracted from:\n{target_path.name}"
                )
                return

            self.finished.emit(True, agenda_data, target_path.name, "")

        except Exception as e:
            logging.error(f"[WordAgendaImporterThread] Import failed: {e}", exc_info=True)
            self.finished.emit(False, {}, self.source_path.name, str(e))