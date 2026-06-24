import os
import re
import webbrowser
import zipfile
from pathlib import Path

import requests
from PyQt5.QtCore import QThread, pyqtSignal

from core.network.session import NetworkSession


class TDocsRevisionsFetcherThread(QThread):
    finished = pyqtSignal(bool, dict, str)

    def __init__(self, url: str):
        super().__init__()
        self.url = url

    def run(self):
        try:
            session = NetworkSession.get_instance()
            NetworkSession.apply_humanness(session)
            response = session.get(self.url, timeout=30)
            response.raise_for_status()

            html = response.text
            # Safely capture full filename, base TDoc, and revision string (e.g., S2-2605740r01)
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

            self.finished.emit(True, revisions, "Success")
        except requests.exceptions.HTTPError as e:
            if e.response.status_code == 404:
                self.finished.emit(True, {}, "No Revisions folder found.")
            else:
                self.finished.emit(False, {}, str(e))
        except Exception as e:
            self.finished.emit(False, {}, str(e))


class TDocActionThread(QThread):
    finished_action = pyqtSignal(str, bool, str)

    # ---> FIX 2: Added 'open_file' parameter
    def __init__(self, base_tdoc: str, target_filename: str, base_url: str, meeting_dir: Path, open_file: bool = True):
        super().__init__()
        self.base_tdoc = base_tdoc
        self.target_filename = target_filename
        self.base_url = base_url
        self.tdoc_dir = meeting_dir / base_tdoc
        self.zip_path = self.tdoc_dir / f"{target_filename}.zip"
        self.open_file = open_file

    def run(self):
        try:
            if not self.zip_path.exists():
                self.tdoc_dir.mkdir(parents=True, exist_ok=True)
                dl_url = self.base_url.rstrip('/') + f"/{self.target_filename}.zip"

                from core.network.session import NetworkSession
                session = NetworkSession.get_instance()
                NetworkSession.apply_humanness(session)
                response = session.get(dl_url, stream=True, timeout=30)
                response.raise_for_status()

                with open(self.zip_path, 'wb') as f:
                    for chunk in response.iter_content(chunk_size=16384):
                        if chunk: f.write(chunk)

            extracted_files = []
            import zipfile
            with zipfile.ZipFile(self.zip_path, 'r') as z:
                for info in z.infolist():
                    if '__MACOSX' in info.filename or info.filename.startswith('._'):
                        continue
                    if info.filename.lower().endswith(('.doc', '.docx', '.pdf', '.ppt', '.pptx')):
                        original_name = Path(info.filename).name

                        # ---> THE FIX: Smart Rename instead of Subfolders!
                        # If the inner file is missing the revision marker (e.g. S2-2603332r01),
                        # we prepend it so it doesn't collide with the base document in the folder.
                        if self.target_filename.lower() not in original_name.lower():
                            safe_name = f"{self.target_filename}_{original_name}"
                        else:
                            safe_name = original_name

                        # Extract directly into the root tdoc_dir (restoring your existing functionality)
                        out_path = self.tdoc_dir / safe_name

                        if not out_path.exists():
                            with open(out_path, 'wb') as f:
                                f.write(z.read(info.filename))

                        extracted_files.append(out_path)

            if not extracted_files:
                self.finished_action.emit(self.base_tdoc, False, "No viewable documents found inside the ZIP.")
                return

            # Keep the exact paths stored for the UI Comparison Cart
            self.extracted_doc_paths = extracted_files

            if self.open_file:
                import os, webbrowser
                for doc in extracted_files:
                    if hasattr(os, 'startfile'):
                        os.startfile(str(doc))
                    else:
                        webbrowser.open(f"file:///{doc}")

            msg = "Opened successfully." if self.open_file else "Downloaded & Added successfully."
            self.finished_action.emit(self.base_tdoc, True, msg)

        except Exception as e:
            self.finished_action.emit(self.base_tdoc, False, str(e))