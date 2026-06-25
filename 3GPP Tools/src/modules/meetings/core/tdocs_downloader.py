# --- File: modules/meetings/core/tdocs_downloader.py ---
import re
from pathlib import Path
from PyQt5.QtCore import QThread, pyqtSignal

# Import your global network session manager
from core.network.session import NetworkSession


class TDocsDownloaderThread(QThread):
    # Emits (success_boolean, file_path_or_error_message)
    finished = pyqtSignal(bool, str, str)

    def __init__(self, mtg_id: str, local_path: Path, parent=None):
        super().__init__(parent)
        self.mtg_id = mtg_id
        self.local_path = local_path

    def run(self):
        url = f"https://portal.3gpp.org/ngppapp/GenerateDocumentList.Aspx?meetingId={self.mtg_id}"
        agenda_dir = self.local_path / "Agenda"

        try:
            # 1. Create the Agenda subfolder safely
            agenda_dir.mkdir(parents=True, exist_ok=True)

            # 2. Utilize the common NetworkSession
            # This automatically inherits your proxies, retries, and User-Agent humanness configuration
            session = NetworkSession.get_instance()
            NetworkSession.apply_humanness(session)

            # 3. Make the HTTP request via the global session
            response = session.get(url, stream=True, timeout=45)
            response.raise_for_status()

            # 4. Attempt to extract the server's suggested filename from headers
            filename = f"TDocs_List_{self.mtg_id}.xlsx"  # Fallback name
            content_disposition = response.headers.get('content-disposition')

            if content_disposition:
                # Regex looks for: filename="Agenda_84089.xlsx" or filename=Agenda_84089.xlsx
                matches = re.findall(r'filename="?([^"]+)"?', content_disposition)
                if matches:
                    filename = matches[0]

            filepath = agenda_dir / filename

            # 5. Save the Excel file in chunks
            with open(filepath, 'wb') as f:
                # ---> OPTIMIZATION: Increased chunk size from 8KB to 64KB
                for chunk in response.iter_content(chunk_size=65536):
                    if chunk:
                        f.write(chunk)

            self.finished.emit(True, str(filepath), self.mtg_id)

        except Exception as e:
            self.finished.emit(False, str(e), self.mtg_id)