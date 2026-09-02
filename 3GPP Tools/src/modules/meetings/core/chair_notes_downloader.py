# --- File: src/modules/meetings/core/chair_notes_downloader.py ---
import os
import re
import logging
import urllib.parse
from pathlib import Path
from bs4 import BeautifulSoup
from PyQt5.QtCore import QThread, pyqtSignal

from core.network.session import NetworkSession


class ChairNotesDownloaderThread(QThread):
    """
    Asynchronously crawls /Inbox/Chair_Notes directory listings across candidate URLs,
    filters files containing 'ChairNotes', and downloads them to Agenda/ChairNotes/.
    """
    progress = pyqtSignal(str)
    # Emits: (success: bool, downloaded_count: int, downloaded_files: list, message: str)
    finished = pyqtSignal(bool, int, list, str)

    def __init__(self, candidate_base_urls: list, target_dir: Path, parent=None):
        super().__init__(parent)
        self.candidate_base_urls = candidate_base_urls
        self.target_dir = Path(target_dir)

    def run(self):
        try:
            self.target_dir.mkdir(parents=True, exist_ok=True)
            session = NetworkSession.get_instance()
            NetworkSession.apply_humanness(session)

            subpaths = [
                "Inbox/Chair_Notes",
                "INBOX/Chair_Notes",
                "Inbox/Chair_notes",
                "INBOX/Chair_notes",
                "Chair_Notes"
            ]

            target_dir_url = None
            html_content = ""

            # 1. Locate the active Chair_Notes folder across priority endpoints
            for base_url in self.candidate_base_urls:
                if self.isInterruptionRequested():
                    return

                if not base_url:
                    continue
                clean_base = base_url.rstrip("/")
                for sub in subpaths:
                    if self.isInterruptionRequested():
                        return

                    probe_url = f"{clean_base}/{sub}"
                    try:
                        self.progress.emit(f"Probing {probe_url[:35]}...")
                        resp = session.get(probe_url, timeout=15)
                        if resp.status_code == 200 and ("ChairNotes" in resp.text or "Directory Listing" in resp.text):
                            target_dir_url = probe_url
                            html_content = resp.text
                            logging.info(f"[ChairNotes] Found active directory listing at: {target_dir_url}")
                            break
                    except Exception as e:
                        logging.debug(f"[ChairNotes] Failed to probe {probe_url}: {e}")
                if target_dir_url:
                    break

            if not target_dir_url or not html_content:
                self.finished.emit(False, 0, [], "Could not locate an active /Inbox/Chair_Notes folder on available servers.")
                return

            # 2. Extract valid Chairman's Notes files
            files_to_download = self._parse_file_list(html_content, target_dir_url)

            if not files_to_download:
                self.finished.emit(True, 0, [], "Connected to Chair_Notes folder, but no matching ChairNotes files were found.")
                return

            # 3. Stream downloads to Agenda/ChairNotes/
            downloaded = []
            total_files = len(files_to_download)

            for idx, (filename, file_url) in enumerate(files_to_download, 1):
                if self.isInterruptionRequested():
                    return

                self.progress.emit(f"({idx}/{total_files}) Downloading {filename[:25]}...")
                dest_path = self.target_dir / filename

                try:
                    dl_resp = session.get(file_url, stream=True, timeout=60)
                    dl_resp.raise_for_status()

                    with open(dest_path, "wb") as f:
                        for chunk in dl_resp.iter_content(chunk_size=65536):
                            if self.isInterruptionRequested():
                                return
                            if chunk:
                                f.write(chunk)

                    downloaded.append(filename)
                    logging.info(f"[ChairNotes] Saved: {dest_path}")
                except Exception as dl_err:
                    logging.error(f"[ChairNotes] Error downloading {filename}: {dl_err}")

            success = len(downloaded) > 0
            msg = f"Downloaded {len(downloaded)} Chairman's Notes files to:\n{self.target_dir}"
            self.finished.emit(success, len(downloaded), downloaded, msg)

        except Exception as ex:
            logging.error(f"[ChairNotes] Unexpected failure: {ex}", exc_info=True)
            self.finished.emit(False, 0, [], f"Error downloading Chairman's Notes: {ex}")

    def _parse_file_list(self, html: str, base_url: str) -> list:
        """Extracts and validates ChairNotes filenames and URLs from HTML directory listings."""
        soup = BeautifulSoup(html, "html.parser")
        file_map = {}

        # 1. Check ASP.NET checkbox input controls: value="ChairNotes_Andy_...doc"
        for inp in soup.find_all("input", class_="downloadInput"):
            val = inp.get("value", "").strip()
            if val and "chairnotes" in val.lower() and not val.startswith("~$"):
                file_map[val] = f"{base_url.rstrip('/')}/{urllib.parse.quote(val)}"

        # 2. Check standard anchor tags: href=".../ChairNotes_...doc"
        for a in soup.find_all("a", href=True):
            href = a["href"].strip()
            parsed_name = urllib.parse.unquote(os.path.basename(href))
            link_text = a.get_text().strip()
            target_name = parsed_name if parsed_name else link_text

            if target_name and "chairnotes" in target_name.lower():
                # Filter out temporary Word lock files and navigation folders
                if target_name.startswith("~$") or target_name.upper() in ["OLDER", "PARENT DIRECTORY"]:
                    continue

                full_url = urllib.parse.urljoin(base_url.rstrip("/") + "/", href)
                if target_name not in file_map:
                    file_map[target_name] = full_url

        return list(file_map.items())