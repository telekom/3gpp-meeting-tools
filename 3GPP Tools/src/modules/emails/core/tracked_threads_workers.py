# --- File: src/modules/emails/core/tracked_threads_workers.py ---
"""Background QThread workers for asynchronous Outlook COM email-thread syncing and scanning."""

from __future__ import annotations

from datetime import datetime, timedelta
import logging
from pathlib import Path
from typing import Dict, List, Optional

from PyQt5.QtCore import QThread, pyqtSignal

from modules.emails.core.email_config import load_wg_email_config
from modules.emails.core.general_email_db import GeneralEmailDatabase
from modules.emails.core.general_email_sync import GeneralEmailSyncThread
from modules.emails.core.sync_watermark import (
    LAST_SYNC_META_KEY,
    format_utc_iso,
    resolve_sync_window,
    utc_now,
)
from modules.emails.core.tracked_threads_db import TrackedThreadsDatabase
from modules.emails.core.tracked_threads_service import (
    list_recent_threads_in_folders,
    refresh_meeting_threads,
    refresh_thread,
    reload_all_thread_messages_from_outlook,
    sync_scoped_inbox_items,
)

_ORIG_REFRESH_THREAD = refresh_thread

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

            if refresh_thread is _ORIG_REFRESH_THREAD:
                updates = refresh_meeting_threads(
                    db,
                    self.meeting_id,
                    progress_cb=lambda idx, tot, cur, rem: self.progress.emit(idx, tot, cur, rem),
                    cancel_check=self.is_cancelled,
                )
                if self.is_cancelled():
                    self.finished_sync.emit(False, updates, "Sync cancelled by user.")
                    return
                for tid, new_count in (updates or {}).items():
                    self.thread_updated.emit(tid, new_count)
                self.finished_sync.emit(True, updates, "")
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
    """Refreshes a single tracked email thread (or reloads all cached EntryIDs) asynchronously."""

    progress = pyqtSignal(int, int, str, list)  # done, total, current_item, remaining_items
    finished_sync = pyqtSignal(bool, int, str)  # ok, updated_or_new_count, error_message

    def __init__(self, db_path: Path, thread_id: str, reload_cached: bool = False, parent=None):
        super().__init__(parent)
        self.db_path = Path(db_path)
        self.thread_id = thread_id
        self.reload_cached = bool(reload_cached)

    def run(self):
        def _work():
            db = TrackedThreadsDatabase(self.db_path)
            cb = lambda done, tot, cur, rem: self.progress.emit(done, tot, cur, rem)
            if self.reload_cached:
                count = reload_all_thread_messages_from_outlook(
                    db,
                    self.thread_id,
                    progress_cb=cb,
                    cancel_check=self.is_cancelled,
                )
            else:
                count = refresh_thread(
                    db,
                    self.thread_id,
                    progress_cb=cb,
                    cancel_check=self.is_cancelled,
                )
            if self.is_cancelled():
                self.finished_sync.emit(False, count, "Sync cancelled by user.")
            else:
                self.finished_sync.emit(True, count, "")

        try:
            self._run_with_com(_work)
        except Exception as exc:
            logger.error("SingleThreadRefreshWorker failed: %s", exc, exc_info=True)
            self.finished_sync.emit(False, 0, str(exc))


class OutlookFolderScanWorker(_BaseComWorker):
    """Scans configured WG Outlook folders asynchronously for PinFromOutlookDialog."""

    progress = pyqtSignal(int, int, str, list)  # done_folders, total_folders, current_folder, remaining_folders
    finished_scan = pyqtSignal(bool, list, str)  # ok, threads_list, error_message

    def __init__(self, wg: str, max_items_per_folder: int = 1500, parent=None):
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


class InboxOutlookSyncWorker(_BaseComWorker):
    """Incrementally syncs both Tracked Threads and WG Related Emails from Outlook off the UI thread."""

    progress_text = pyqtSignal(str)
    finished_sync = pyqtSignal(bool, int, int, str)  # ok, new_thread_msgs, general_synced, error_msg

    def __init__(
        self,
        tracked_db_path: Optional[Path],
        general_db_path: Optional[Path],
        meeting_id: str,
        wg: str,
        start_date: str = "",
        end_date: str = "",
        parent=None,
        *,
        scoped_thread_ids: Optional[List[str]] = None,
        scoped_email_ids: Optional[List[str]] = None,
        sync_mode: str = "incremental",
        days_buffer: int = 5,
        silent: bool = False,
    ):
        super().__init__(parent)
        self.tracked_db_path = Path(tracked_db_path) if tracked_db_path else None
        self.general_db_path = Path(general_db_path) if general_db_path else None
        self.meeting_id = str(meeting_id or "").strip()
        self.wg = str(wg or "SA2").strip().upper()
        self.start_date = str(start_date or "").strip()
        self.end_date = str(end_date or "").strip()
        self.scoped_thread_ids = list(scoped_thread_ids) if scoped_thread_ids is not None else None
        self.scoped_email_ids = list(scoped_email_ids) if scoped_email_ids is not None else None
        is_scoped = self.scoped_thread_ids is not None or self.scoped_email_ids is not None
        self.sync_mode = "scoped" if is_scoped else (str(sync_mode or "incremental").strip().lower() or "incremental")
        self.days_buffer = max(0, int(days_buffer if days_buffer is not None else 5))
        self.silent = bool(silent)
        self.last_sync_window_info: Optional[dict] = None

    @staticmethod
    def _resolve_incremental_window(
        general_db_path: Optional[Path],
        start_date: str,
        end_date: str,
    ) -> tuple[str, str]:
        today = datetime.now().date()
        gdb = GeneralEmailDatabase(general_db_path) if (general_db_path and Path(general_db_path).exists()) else None
        win = resolve_sync_window(gdb, start_date, end_date, sync_mode="incremental")
        if win.get("start_date") and win.get("end_date"):
            return win["start_date"], win["end_date"]
        since_utc = win.get("since_dt_utc")
        if since_utc is not None:
            return since_utc.astimezone().strftime("%Y-%m-%d"), today.strftime("%Y-%m-%d")
        return start_date or (today - timedelta(days=10)).strftime("%Y-%m-%d"), end_date or today.strftime("%Y-%m-%d")

    def run(self):
        def _work():
            sync_started_utc = format_utc_iso(utc_now())
            new_thread_msgs = 0
            general_synced = 0
            is_scoped = self.scoped_thread_ids is not None or self.scoped_email_ids is not None

            if is_scoped:
                tdb = TrackedThreadsDatabase(self.tracked_db_path) if self.tracked_db_path else None
                gdb = GeneralEmailDatabase(self.general_db_path) if self.general_db_path else None
                self.progress_text.emit("Syncing filtered conversations...")
                new_thread_msgs, general_synced = sync_scoped_inbox_items(
                    tdb,
                    gdb,
                    tracked_thread_ids=self.scoped_thread_ids,
                    unpinned_email_ids=self.scoped_email_ids,
                    progress_cb=lambda msg: self.progress_text.emit(msg),
                    cancel_check=self.is_cancelled,
                )
            else:
                gdb_meta = GeneralEmailDatabase(self.general_db_path) if self.general_db_path else None
                win = resolve_sync_window(
                    gdb_meta,
                    self.start_date,
                    self.end_date,
                    sync_mode=self.sync_mode,
                    days_buffer=self.days_buffer,
                )
                self.last_sync_window_info = win
                gap_lbl = win.get("gap_description") or self.sync_mode

                if self.tracked_db_path and self.meeting_id and not self.is_cancelled():
                    self.progress_text.emit(f"Syncing pinned threads ({gap_lbl})...")
                    tdb = TrackedThreadsDatabase(self.tracked_db_path)
                    updates = refresh_meeting_threads(
                        tdb,
                        self.meeting_id,
                        progress_cb=lambda idx, tot, cur, _rem: self.progress_text.emit(
                            f"Folder ({idx + 1}/{tot}): {cur[:28]}"
                        ),
                        cancel_check=self.is_cancelled,
                        since_dt_utc=win.get("since_dt_utc"),
                    )
                    new_thread_msgs = sum((updates or {}).values())

                if self.general_db_path and not self.is_cancelled():
                    folders = load_wg_email_config(self.wg) if self.wg else []
                    self.progress_text.emit(f"Scanning Outlook folders ({gap_lbl})...")
                    syncer = GeneralEmailSyncThread(
                        folders,
                        self.general_db_path,
                        start_date=str(win.get("start_date") or ""),
                        end_date=str(win.get("end_date") or ""),
                        days_buffer=int(win.get("days_buffer", 0)),
                        since_dt_utc=win.get("since_dt_utc"),
                    )
                    general_synced = syncer.sync_folders(
                        progress_cb=lambda m: self.progress_text.emit(m),
                        cancel_check=self.is_cancelled,
                    )

                if not self.is_cancelled() and gdb_meta is not None:
                    try:
                        gdb_meta.set_sync_meta(LAST_SYNC_META_KEY, sync_started_utc)
                    except Exception as exc:
                        logger.debug("Could not persist %s watermark: %s", LAST_SYNC_META_KEY, exc)

            if self.is_cancelled():
                self.finished_sync.emit(False, new_thread_msgs, general_synced, "Outlook sync cancelled.")
            else:
                self.finished_sync.emit(True, new_thread_msgs, general_synced, "")

        try:
            self._run_with_com(_work)
        except Exception as exc:
            logger.error("InboxOutlookSyncWorker failed: %s", exc, exc_info=True)
            self.finished_sync.emit(False, 0, 0, str(exc))
