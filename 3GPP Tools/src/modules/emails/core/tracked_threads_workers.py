# --- File: src/modules/emails/core/tracked_threads_workers.py ---
"""Background QThread workers for asynchronous Outlook COM email-thread syncing and scanning."""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Dict, List

from PyQt5.QtCore import QThread, pyqtSignal

from modules.emails.core.tracked_threads_db import TrackedThreadsDatabase
from modules.emails.core.tracked_threads_service import (
    list_recent_threads_in_folders,
    refresh_thread,
)

logger = logging.getLogger(__name__)


class _BaseComWorker(QThread):
    """Base worker providing cancellation and safe COM apartment initialization."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self._cancel_requested = False

    def request_cancel(self) -> None:
        self._cancel_requested = True

    def is_cancelled(self) -> bool:
        return bool(self._cancel_requested)

    def _run_with_com(self, fn):
        try:
            import pythoncom
            pythoncom.CoInitialize()
        except Exception:
            pythoncom = None  # Non-Windows or headless mock environment
        try:
            return fn()
        finally:
            if pythoncom is not None:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass


class TrackedThreadsRefreshWorker(_BaseComWorker):
    """Refreshes all pinned tracked threads for a meeting asynchronously."""

    progress = pyqtSignal(int, int, str, list)  # completed, total, current_title, remaining_titles
    thread_updated = pyqtSignal(str, int)  # thread_id, new_count
    finished_sync = pyqtSignal(bool, dict, str)  # ok, updates_dict, error_message

    def __init__(self, db_path: Path, meeting_id: str, parent=None):
        super().__init__(parent)
        self.db_path = Path(db_path)
        self.meeting_id = meeting_id

    def run(self):
        def _work():
            db = TrackedThreadsDatabase(self.db_path)
            threads = db.get_threads_for_meeting(self.meeting_id)
            total = len(threads)
            titles = [str(t.get("title") or t.get("thread_id") or "") for t in threads]
            updates: Dict[str, int] = {}

            if total == 0:
                self.finished_sync.emit(True, {}, "")
                return

            for idx, thread in enumerate(threads):
                if self.is_cancelled():
                    self.finished_sync.emit(False, updates, "Sync cancelled by user.")
                    return
                self.progress.emit(idx, total, titles[idx], titles[idx + 1 :])
                tid = thread["thread_id"]
                new_count = refresh_thread(db, tid, cancel_check=self.is_cancelled)
                if new_count > 0:
                    updates[tid] = new_count
                    self.thread_updated.emit(tid, new_count)
                self.progress.emit(idx + 1, total, titles[idx], titles[idx + 1 :])

            self.finished_sync.emit(True, updates, "")

        try:
            self._run_with_com(_work)
        except Exception as exc:
            logger.error("TrackedThreadsRefreshWorker failed: %s", exc, exc_info=True)
            self.finished_sync.emit(False, {}, str(exc))


class SingleThreadRefreshWorker(_BaseComWorker):
    """Refreshes a single tracked email thread across configured Outlook folders asynchronously."""

    progress = pyqtSignal(int, int, str, list)  # done_folders, total_folders, current_folder, remaining_folders
    finished_sync = pyqtSignal(bool, int, str)  # ok, new_count, error_message

    def __init__(self, db_path: Path, thread_id: str, parent=None):
        super().__init__(parent)
        self.db_path = Path(db_path)
        self.thread_id = thread_id

    def run(self):
        def _work():
            db = TrackedThreadsDatabase(self.db_path)
            new_count = refresh_thread(
                db,
                self.thread_id,
                progress_cb=lambda done, tot, cur, rem: self.progress.emit(done, tot, cur, rem),
                cancel_check=self.is_cancelled,
            )
            if self.is_cancelled():
                self.finished_sync.emit(False, new_count, "Sync cancelled by user.")
            else:
                self.finished_sync.emit(True, new_count, "")

        try:
            self._run_with_com(_work)
        except Exception as exc:
            logger.error("SingleThreadRefreshWorker failed: %s", exc, exc_info=True)
            self.finished_sync.emit(False, 0, str(exc))


class OutlookFolderScanWorker(_BaseComWorker):
    """Scans configured WG Outlook folders asynchronously for PinFromOutlookDialog."""

    progress = pyqtSignal(int, int, str, list)  # done_folders, total_folders, current_folder, remaining_folders
    finished_scan = pyqtSignal(bool, list, str)  # ok, threads_list, error_message

    def __init__(self, wg: str, max_items_per_folder: int = 300, parent=None):
        super().__init__(parent)
        self.wg = wg
        self.max_items_per_folder = max_items_per_folder

    def run(self):
        def _work():
            threads = list_recent_threads_in_folders(
                self.wg,
                max_items_per_folder=self.max_items_per_folder,
                progress_cb=lambda done, tot, cur, rem: self.progress.emit(done, tot, cur, rem),
                cancel_check=self.is_cancelled,
            )
            if self.is_cancelled():
                self.finished_scan.emit(False, threads, "Folder scan cancelled by user.")
            else:
                self.finished_scan.emit(True, threads, "")

        try:
            self._run_with_com(_work)
        except Exception as exc:
            logger.error("OutlookFolderScanWorker failed: %s", exc, exc_info=True)
            self.finished_scan.emit(False, [], str(exc))
