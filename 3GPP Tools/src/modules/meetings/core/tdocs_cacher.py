# --- File: modules/meetings/core/tdocs_cacher.py ---
import logging
import re
from pathlib import Path
from urllib.parse import urljoin, unquote
from concurrent.futures import ThreadPoolExecutor, as_completed

from PyQt5.QtCore import QThread, pyqtSignal
from core.network.session import NetworkSession


class TDocsCacherThread(QThread):
    # Emits (success_boolean, result_message)
    finished = pyqtSignal(bool, str)

    def __init__(self, docs_url: str, local_path: Path, parent=None):
        super().__init__(parent)
        self.docs_url = docs_url
        self.local_path = local_path
        self.is_cancelled = False

    def run(self):
        try:
            logging.info(f"🔍 Fetching TDoc directory listing from: {self.docs_url}")

            # 1. Fetch the HTML of the Docs folder
            html = NetworkSession.get_html(self.docs_url)

            # 2. Extract all .zip hrefs from the FTP page
            href_pattern = re.compile(r'href=["\']([^"\'>]+\.zip)["\']', re.IGNORECASE)
            hrefs = href_pattern.findall(html)

            # 3. Filter strictly for files named like TDocs (e.g., S2-2605693.zip or revisions like S2-2605693r1.zip)
            # The regex captures the file name without the .zip extension
            tdoc_pattern = re.compile(r'^([A-Za-z0-9]+-\d+.*)\.zip$', re.IGNORECASE)

            download_tasks = []
            for href in hrefs:
                filename = unquote(href.split('/')[-1])
                match = tdoc_pattern.match(filename)
                if match:
                    folder_name = match.group(1)  # Name without .zip (e.g., S2-2605693)
                    file_url = urljoin(self.docs_url, href)

                    target_dir = self.local_path / folder_name
                    target_file = target_dir / filename
                    download_tasks.append((file_url, target_dir, target_file, filename))

            # Deduplicate just in case 3GPP has wonky HTML
            download_tasks = list({t[2]: t for t in download_tasks}.values())
            total_files = len(download_tasks)

            if total_files == 0:
                msg = "No valid TDoc .zip files found in the Docs directory."
                logging.warning(f"⚠️ {msg}")
                self.finished.emit(True, msg)
                return

            logging.info(f"📥 Found {total_files} TDoc zip files. Starting cache process...")

            downloaded = 0
            skipped = 0
            processed = 0

            session = NetworkSession.get_instance()

            # 4. Use a moderate thread pool to speed up caching without overloading 3GPP's firewall
            with ThreadPoolExecutor(max_workers=5) as executor:
                future_to_task = {}
                for task in download_tasks:
                    future = executor.submit(self._download_file, session, *task)
                    future_to_task[future] = task

                for future in as_completed(future_to_task):
                    if self.is_cancelled:
                        break

                    filename = future_to_task[future][3]
                    try:
                        was_downloaded = future.result()
                        processed += 1
                        if was_downloaded:
                            downloaded += 1
                            logging.info(f"✅ [{processed}/{total_files}] Downloaded: {filename}")
                        else:
                            skipped += 1
                            # ---> OPTIMIZATION: Throttle skip logs so we don't crash the UI Event Loop!
                            if skipped % 50 == 0:
                                logging.info(f"⏭️ Skipped {skipped} existing files so far...")

                    except Exception as e:
                        processed += 1
                        logging.error(f"❌ [{processed}/{total_files}] Failed to download {filename}: {e}")

            # 5. Output Summary
            summary = f"Caching Complete! Downloaded: {downloaded}, Skipped: {skipped}, Total: {total_files}"
            logging.info(f"🏁 {summary}")
            self.finished.emit(True, summary)

        except Exception as e:
            error_msg = f"Failed to cache TDocs: {e}"
            logging.error(f"❌ {error_msg}")
            self.finished.emit(False, error_msg)

    def _download_file(self, session, file_url, target_dir, target_file, filename):
        """Worker function to download a single file."""
        # ONLY download if the file is not present
        if target_file.exists():
            return False

            # Create subfolders if necessary
        target_dir.mkdir(parents=True, exist_ok=True)

        # Inherit humanness and proxy rules
        NetworkSession.apply_humanness(session)
        response = session.get(file_url, stream=True, timeout=60)
        response.raise_for_status()

        with open(target_file, 'wb') as f:
            # ---> OPTIMIZATION: Increased chunk size from 16KB to 64KB
            for chunk in response.iter_content(chunk_size=65536):
                if chunk:
                    f.write(chunk)

        return True