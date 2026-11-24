"""QThread workers for listing and downloading Chairman's Notes."""

import logging
import os
from pathlib import Path
from typing import List, Tuple

from PyQt5.QtCore import QThread, pyqtSignal

from core.network.session import NetworkSession
from core.utils.paths import to_extended_path

from .selector import (
    SUBPATHS,
    extract_timestamp_sort_key,
    group_and_pick_latest_by_track,
    is_valid_chair_notes_filename,
    looks_like_directory_listing,
    parse_file_list,
    pick_latest_file,
    prefer_clean_variants,
    select_best_chair_notes_files,
)
from .transfer import (
    download_files_sync,
    extract_chair_notes_zip,
    format_byte_size,
    locate_chair_notes_dir,
)


class ChairNotesListingThread(QThread):
    """Lightweight asynchronous probe thread checking candidate 3GPP URLs for Chair_Notes."""

    progress = pyqtSignal(str)
    listing_finished = pyqtSignal(bool, list, str, str)

    def __init__(self, candidate_base_urls: list, parent=None):
        super().__init__(parent)
        self.candidate_base_urls = candidate_base_urls

    def run(self):
        try:
            session = NetworkSession.get_instance()
            NetworkSession.apply_humanness(session)

            target_dir_url, html_content = ChairNotesDownloaderThread.locate_chair_notes_dir(
                session,
                self.candidate_base_urls,
                progress_callback=lambda msg: self.progress.emit(msg),
                cancel_check=self.isInterruptionRequested,
            )

            if self.isInterruptionRequested():
                self.listing_finished.emit(False, [], "", "Chairman's Notes check cancelled.")
                return

            if not target_dir_url or not html_content:
                self.listing_finished.emit(
                    False,
                    [],
                    "",
                    "Could not locate an active /Inbox/Chairs_Notes or /Inbox/Chair_Notes folder on available servers.",
                )
                return

            files = ChairNotesDownloaderThread.parse_file_list_static(html_content, target_dir_url)
            if not files:
                self.listing_finished.emit(
                    True,
                    [],
                    target_dir_url,
                    "Connected to Chairman's Notes folder, but no matching files were found.",
                )
                return

            self.listing_finished.emit(True, files, target_dir_url, "")
        except Exception as ex:
            logging.error(f"[ChairNotesListing] Unexpected failure: {ex}", exc_info=True)
            self.listing_finished.emit(False, [], "", f"Error scanning Chairman's Notes: {ex}")


class ChairNotesDownloaderThread(QThread):
    """Asynchronously crawls Chair_Notes directories, selects best notes, and extracts ZIPs."""

    progress = pyqtSignal(str)
    download_finished = pyqtSignal(bool, int, list, str)

    SUBPATHS = list(SUBPATHS)

    def __init__(
        self,
        candidate_base_urls: list,
        target_dir: Path,
        parent=None,
        only_final_or_latest: bool = True,
        selected_files: list = None,
        overwrite: bool = True,
    ):
        super().__init__(parent)
        self.candidate_base_urls = candidate_base_urls
        self.target_dir = Path(target_dir)
        self.only_final_or_latest = only_final_or_latest
        self.selected_files = selected_files
        self.overwrite = overwrite

    def run(self):
        try:
            os.makedirs(to_extended_path(self.target_dir), exist_ok=True)
            session = NetworkSession.get_instance()
            NetworkSession.apply_humanness(session)

            if self.selected_files is not None:
                files_to_download = list(self.selected_files)
            else:
                target_dir_url, html_content = self.locate_chair_notes_dir(
                    session,
                    self.candidate_base_urls,
                    progress_callback=lambda msg: self.progress.emit(msg),
                    cancel_check=self.isInterruptionRequested,
                )

                if self.isInterruptionRequested():
                    self.download_finished.emit(False, 0, [], "Chairman's Notes download cancelled.")
                    return

                if not target_dir_url or not html_content:
                    self.download_finished.emit(
                        False,
                        0,
                        [],
                        "Could not locate an active /Inbox/Chairs_Notes or /Inbox/Chair_Notes folder on available servers.",
                    )
                    return

                files_to_download = self._parse_file_list(html_content, target_dir_url)
                if self.only_final_or_latest and len(files_to_download) > 1:
                    files_to_download = self.select_best_chair_notes_files(files_to_download)

            if not files_to_download:
                self.download_finished.emit(
                    True,
                    0,
                    [],
                    "No matching Chairman's Notes files were selected or found.",
                )
                return

            downloaded = self.download_files_sync(
                session,
                files_to_download,
                self.target_dir,
                progress_callback=lambda msg: self.progress.emit(msg),
                cancel_check=self.isInterruptionRequested,
                overwrite=self.overwrite,
            )

            if self.isInterruptionRequested():
                self.download_finished.emit(
                    False, len(downloaded), downloaded, "Chairman's Notes download cancelled."
                )
                return

            success = len(downloaded) > 0
            msg = f"Downloaded {len(downloaded)} Chairman's Notes file(s) to:\n{self.target_dir}"
            self.download_finished.emit(success, len(downloaded), downloaded, msg)

        except Exception as ex:
            logging.error(f"[ChairNotes] Unexpected failure: {ex}", exc_info=True)
            self.download_finished.emit(False, 0, [], f"Error downloading Chairman's Notes: {ex}")

    @classmethod
    def locate_chair_notes_dir(
        cls,
        session,
        candidate_base_urls: list,
        progress_callback=None,
        cancel_check=None,
    ) -> Tuple[str, str]:
        """Probe candidate base URLs across `SUBPATHS`."""
        return locate_chair_notes_dir(
            session,
            candidate_base_urls,
            subpaths=cls.SUBPATHS,
            progress_callback=progress_callback,
            cancel_check=cancel_check,
        )

    @staticmethod
    def _looks_like_directory_listing(html_text: str) -> bool:
        return looks_like_directory_listing(html_text)

    @classmethod
    def is_valid_chair_notes_filename(cls, filename: str) -> bool:
        return is_valid_chair_notes_filename(filename)

    def _parse_file_list(self, html: str, base_url: str) -> List[Tuple[str, str]]:
        return self.parse_file_list_static(html, base_url)

    @classmethod
    def parse_file_list_static(cls, html: str, base_url: str) -> List[Tuple[str, str]]:
        return parse_file_list(html, base_url)

    @classmethod
    def select_best_chair_notes_files(
        cls,
        file_tuples: List[Tuple[str, str]],
        only_final_or_latest: bool = True,
    ) -> List[Tuple[str, str]]:
        return select_best_chair_notes_files(file_tuples, only_final_or_latest=only_final_or_latest)

    @classmethod
    def _prefer_clean_variants(cls, file_tuples: List[Tuple[str, str]]) -> List[Tuple[str, str]]:
        return prefer_clean_variants(file_tuples)

    @classmethod
    def _extract_timestamp_sort_key(cls, filename: str) -> Tuple:
        return extract_timestamp_sort_key(filename)

    @classmethod
    def _pick_latest_file(cls, items: List[Tuple[str, str]]) -> Tuple[str, str]:
        return pick_latest_file(items)

    @classmethod
    def _group_and_pick_latest_by_track(cls, items: List[Tuple[str, str]]) -> List[Tuple[str, str]]:
        return group_and_pick_latest_by_track(items)

    @staticmethod
    def _format_byte_size(num_bytes: int) -> str:
        return format_byte_size(num_bytes)

    @classmethod
    def download_files_sync(
        cls,
        session,
        files_to_download: List[Tuple[str, str]],
        target_dir: Path,
        progress_callback=None,
        cancel_check=None,
        overwrite: bool = True,
    ) -> List[str]:
        return download_files_sync(
            session,
            files_to_download,
            target_dir,
            progress_callback=progress_callback,
            cancel_check=cancel_check,
            overwrite=overwrite,
            zip_extractor=cls.extract_chair_notes_zip,
        )

    @classmethod
    def extract_chair_notes_zip(cls, zip_path: Path, target_dir: Path) -> List[str]:
        return extract_chair_notes_zip(zip_path, target_dir)
