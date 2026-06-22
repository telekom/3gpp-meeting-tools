from pathlib import Path

from PyQt5.QtCore import QThread, pyqtSignal

from core.network.session import NetworkSession


class SpecDownloadThread(QThread):
    finished_success = pyqtSignal(Path)
    error = pyqtSignal(str)

    def __init__(self, url: str, zip_path: Path):
        super().__init__()
        self.url = url
        self.zip_path = zip_path

    def run(self):
        try:
            NetworkSession.download_file(self.url, self.zip_path)
            self.finished_success.emit(self.zip_path)
        except Exception as e:
            self.error.emit(str(e))
