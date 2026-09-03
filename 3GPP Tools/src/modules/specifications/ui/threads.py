# --- File: src/modules/specifications/ui/threads.py ---
import logging
from pathlib import Path
from typing import Union

from PyQt5.QtCore import QThread, pyqtSignal

from core.network.session import (
    DownloadCancelledError,
    HttpError,
    NetworkError,
    NetworkSession,
)


class SpecDownloadThread(QThread):
    """
    Background worker thread to download specification archives from 3GPP FTP
    using the centralized NetworkSession.
    """

    finished_success = pyqtSignal(Path)
    error = pyqtSignal(str)
    ui_log_msg = pyqtSignal(str, int)
    progress = pyqtSignal(int)

    def __init__(self, url: str, dest_path: Union[str, Path], timeout: int = 30):
        super().__init__()
        self.url = str(url).strip()
        self.dest_path = Path(dest_path)
        self.timeout = timeout

    def run(self):
        if not self.url:
            self.error.emit("Download URL is empty.")
            return

        try:
            self.ui_log_msg.emit(
                f"⬇️ Downloading: {self.dest_path.name}...", logging.INFO
            )

            # Delegate connection, delay, headers, streaming, and progress to NetworkSession
            downloaded_bytes = NetworkSession.download_file(
                url=self.url,
                dest_path=self.dest_path,
                timeout=self.timeout,
                progress_cb=self.progress.emit,
                cancel_cb=self.isInterruptionRequested,
                atomic=True,
            )

            size_kb = downloaded_bytes / 1024.0
            self.ui_log_msg.emit(
                f"✅ Successfully downloaded {self.dest_path.name} ({size_kb:.1f} KB)",
                logging.INFO,
            )
            self.finished_success.emit(self.dest_path)

        except DownloadCancelledError:
            self.ui_log_msg.emit(
                f"⚠️ Download cancelled: {self.dest_path.name}", logging.WARNING
            )

        except HttpError as e:
            err_msg = f"HTTP {e.status_code} downloading {self.dest_path.name}: {e}"
            logging.error(err_msg)
            self.ui_log_msg.emit(f"❌ {err_msg}", logging.ERROR)
            self.error.emit(str(e))

        except NetworkError as e:
            err_msg = f"Network error downloading {self.dest_path.name}: {e}"
            logging.error(err_msg)
            self.ui_log_msg.emit(f"❌ {err_msg}", logging.ERROR)
            self.error.emit(str(e))