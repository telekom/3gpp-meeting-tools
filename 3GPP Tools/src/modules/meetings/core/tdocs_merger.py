# --- File: src/modules/meetings/core/tdocs_merger.py ---
import os
import re
from pathlib import Path
import pandas as pd
from PyQt5.QtCore import QThread, pyqtSignal

from core.network.session import NetworkSession
from modules.meetings.core.tdocs_parser import TDocsParser
import logging


class TDocsMergerThread(QThread):
    progress = pyqtSignal(str)
    finished = pyqtSignal(bool, str)

    def __init__(self, meetings_data: list, force_download: bool, save_path: str, cache_dir: str, parent=None):
        super().__init__(parent)
        self.meetings_data = meetings_data
        self.force_download = force_download
        self.save_path = save_path
        self.cache_dir = Path(cache_dir)

    def run(self):
        try:
            all_dfs = []

            # Utilize the global NetworkSession so we inherit proxies and Humanness Delays
            session = NetworkSession.get_instance()
            NetworkSession.apply_humanness(session)

            total = len(self.meetings_data)

            for idx, mtg in enumerate(self.meetings_data):
                mtg_id = mtg.get("mtg_id")
                if not mtg_id:
                    continue  # Skip if we don't have a 3GPP Portal ID

                wg = mtg.get("wg_name", "")
                mtg_num = mtg.get("meeting_number", "")
                folder_name = mtg.get("folder_name") or mtg_num
                start_date = mtg.get("start_date", "")
                end_date = mtg.get("end_date", "")

                self.progress.emit(f"Processing {wg} {mtg_num} ({idx + 1}/{total})...")

                agenda_dir = self.cache_dir / folder_name / "Agenda"
                agenda_dir.mkdir(parents=True, exist_ok=True)

                # Find cached file
                filepath = self._find_cached_file(agenda_dir, mtg_id)

                # Check if we need to force download or if it's missing
                if self.force_download or not filepath or not filepath.exists():
                    self.progress.emit(f"Downloading {wg} {mtg_num} ({idx + 1}/{total})...")
                    filepath = self._download_sync(session, mtg_id, agenda_dir)

                if not filepath or not filepath.exists():
                    logging.warning(f"Skipping {wg} {mtg_num}: Could not resolve Excel file.")
                    continue

                # Parse the Excel file using your existing parser
                parsed_data = TDocsParser.parse_tdocs_excel(str(filepath))
                if not parsed_data:
                    continue

                # Convert to a DataFrame
                df = pd.DataFrame(parsed_data)

                # Inject the 4 new columns at the very front of the table
                df.insert(0, "WG", wg)
                df.insert(1, "Meeting", mtg_num)
                df.insert(2, "Start Date", start_date)
                df.insert(3, "End Date", end_date)

                all_dfs.append(df)

            if not all_dfs:
                self.finished.emit(False, "No TDocs found to merge for the selected meetings.")
                return

            self.progress.emit("Concatenating and saving master Excel file...")

            # Concat perfectly handles varying column layouts across different WG files
            master_df = pd.concat(all_dfs, ignore_index=True)
            master_df.to_excel(self.save_path, index=False)

            self.finished.emit(True,
                               f"Successfully merged {len(master_df)} TDocs across {len(all_dfs)} meetings!\n\nSaved to:\n{self.save_path}")

        except Exception as e:
            logging.error(f"Error merging TDocs: {e}", exc_info=True)
            self.finished.emit(False, f"Error merging TDocs:\n{str(e)}")

    def _find_cached_file(self, agenda_dir: Path, mtg_id: str) -> Path:
        """Looks for the existing Excel file in the local cache."""
        if agenda_dir.exists():
            for f in agenda_dir.iterdir():
                if ("tdoc_list_meeting_" in f.name.lower() or "tdocs_list_" in f.name.lower()) and f.name.endswith(
                        ".xlsx"):
                    return f
        return agenda_dir / f"TDoc_List_Meeting_{mtg_id}.xlsx"

    def _download_sync(self, session, mtg_id: str, agenda_dir: Path):
        """Downloads the file synchronously within this background thread."""
        url = f"https://portal.3gpp.org/ngppapp/GenerateDocumentList.Aspx?meetingId={mtg_id}"
        try:
            response = session.get(url, stream=True, timeout=45)
            response.raise_for_status()

            filename = f"TDoc_List_Meeting_{mtg_id}.xlsx"
            content_disposition = response.headers.get('content-disposition')
            if content_disposition:
                matches = re.findall(r'filename="?([^"]+)"?', content_disposition)
                if matches:
                    filename = matches[0]

            filepath = agenda_dir / filename
            with open(filepath, 'wb') as f:
                for chunk in response.iter_content(chunk_size=65536):
                    if chunk:
                        f.write(chunk)
            return filepath
        except Exception as e:
            logging.error(f"Sync download failed for meeting {mtg_id}: {e}")
            return None