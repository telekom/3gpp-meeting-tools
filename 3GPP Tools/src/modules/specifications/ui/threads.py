import logging
from pathlib import Path

from PyQt5.QtCore import QThread, pyqtSignal

from core.network.session import NetworkSession


class SpecDownloadThread(QThread):
    finished_success = pyqtSignal(Path)
    error = pyqtSignal(str)

    # ---> NEW: Add the log signal!
    ui_log_msg = pyqtSignal(str, int)

    def __init__(self, url: str, zip_path: Path):
        super().__init__()
        self.url = url
        self.zip_path = zip_path

    def run(self):
        try:
            # ---> NEW: Emit a starting message
            self.ui_log_msg.emit(f"⏳ Downloading specification archive: {self.zip_path.name}...", logging.INFO)

            NetworkSession.download_file(self.url, self.zip_path)

            # ---> NEW: Emit a success message
            self.ui_log_msg.emit(f"✅ Download complete: {self.zip_path.name}", logging.INFO)
            self.finished_success.emit(self.zip_path)
        except Exception as e:
            # ---> NEW: Emit an error message if something breaks
            self.ui_log_msg.emit(f"❌ Download Failed: {str(e)}", logging.ERROR)
            self.error.emit(str(e))