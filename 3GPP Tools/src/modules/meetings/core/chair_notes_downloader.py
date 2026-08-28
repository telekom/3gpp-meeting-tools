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
    filters files containing 'ChairNotes', and downloads them to the meeting Agenda folder.
    """
    progress = pyqtSignal(str)
    # finished emits: (success: bool, downloaded_count: int, downloaded_files: list, message: str)
    finished = pyqtSignal(bool, int, list, str)

    def __init__(self, candidate_base_urls: list, agenda_dir: Path, parent=None):
        super().__init__(parent)
        self.candidate_base_urls = candidate_base_urls
        self.agenda_dir = Path(agenda_dir)

    def run(self):
        try:
            self.agenda_dir.mkdir(parents=True, exist_ok=True)
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

            # 1. Probe candidate endpoints to locate the active Chair_Notes folder
            for base_url in self.candidate_base_urls:
                if not base_url:
                    continue
                clean_base = base_url.rstrip("/")
                for sub in subpaths:
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

            # 2. Parse HTML listing for files matching 'ChairNotes'
            files_to_download = self._parse_file_list(html_content, target_dir_url)

            if not files_to_download:
                self.finished.emit(True, 0, [], "Connected to Chair_Notes folder, but no matching ChairNotes files were found.")
                return

            # 3. Download files asynchronously
            downloaded = []
            total_files = len(files_to_download)

            for idx, (filename, file_url) in enumerate(files_to_download, 1):
                self.progress.emit(f"({idx}/{total_files}) Downloading {filename[:25]}...")
                dest_path = self.agenda_dir / filename

                try:
                    dl_resp = session.get(file_url, stream=True, timeout=60)
                    dl_resp.raise_for_status()

                    with open(dest_path, "wb") as f:
                        for chunk in dl_resp.iter_content(chunk_size=65536):
                            if chunk:
                                f.write(chunk)

                    downloaded.append(filename)
                    logging.info(f"[ChairNotes] Successfully saved: {dest_path}")
                except Exception as dl_err:
                    logging.error(f"[ChairNotes] Error downloading {filename}: {dl_err}")

            success = len(downloaded) > 0
            msg = f"Downloaded {len(downloaded)} Chairman's Notes files to:\n{self.agenda_dir}"
            self.finished.emit(success, len(downloaded), downloaded, msg)

        except Exception as ex:
            logging.error(f"[ChairNotes] Unexpected failure: {ex}", exc_info=True)
            self.finished.emit(False, 0, [], f"Error downloading Chairman's Notes: {ex}")

    def _parse_file_list(self, html: str, base_url: str) -> list:
        """Extracts and validates ChairNotes filenames and absolute URLs from HTML directory listings."""
        soup = BeautifulSoup(html, "html.parser")
        file_map = {}

        # Parse ASP.NET checkbox input tags: <input type="checkbox" class="downloadInput" value="...doc">
        for inp in soup.find_all("input", class_="downloadInput"):
            val = inp.get("value", "").strip()
            if val and "chairnotes" in val.lower() and not val.startswith("~$"):
                file_map[val] = f"{base_url.rstrip('/')}/{urllib.parse.quote(val)}"

        # Parse anchor tags: <a href="...">...</a>
        for a in soup.find_all("a", href=True):
            href = a["href"].strip()
            # Extract raw filename from URL or anchor text
            parsed_name = urllib.parse.unquote(os.path.basename(href))
            link_text = a.get_text().strip()
            target_name = parsed_name if parsed_name else link_text

            if target_name and "chairnotes" in target_name.lower():
                # Filter out temporary Word lock files and folders
                if target_name.startswith("~$") or target_name.upper() in ["OLDER", "PARENT DIRECTORY"]:
                    continue

                full_url = urllib.parse.urljoin(base_url.rstrip("/") + "/", href)
                if target_name not in file_map:
                    file_map[target_name] = full_url

        return list(file_map.items())