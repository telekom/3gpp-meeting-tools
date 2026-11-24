"""Background QThread workers for GitHub update checking and ZIP staging."""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Optional

from PyQt5.QtCore import QThread, pyqtSignal
from PyQt5.QtWidgets import QWidget

from core.updater.app_updater import (
    CommitInfo,
    LocalInstallInfo,
    UpdateCheckResult,
    check_for_updates,
    inspect_local_install,
    stage_zip_update,
)

logger = logging.getLogger(__name__)


class AppUpdateCheckThread(QThread):
    """Non-blocking background worker that checks GitHub `master` for updates."""

    result_ready = pyqtSignal(object)  # UpdateCheckResult

    def __init__(
        self,
        app_root: Optional[Path] = None,
        record_check_time: bool = True,
        parent: Optional[QWidget] = None,
    ) -> None:
        super().__init__(parent)
        self._app_root = app_root
        self._record_check_time = record_check_time

    def run(self) -> None:
        try:
            res = check_for_updates(
                app_root=self._app_root,
                timeout=15,
                record_check_time=self._record_check_time,
            )
        except Exception as exc:
            logger.warning("[Updater] Update check failed: %s", exc)
            res = UpdateCheckResult(
                status="error",
                local_info=inspect_local_install(self._app_root),
                error=str(exc),
            )
        self.result_ready.emit(res)


class AppUpdateDownloadThread(QThread):
    """Background worker that downloads and stages the commit-pinned ZIP archive."""

    progress_changed = pyqtSignal(int, str)
    staging_finished = pyqtSignal(object)  # StagedUpdateInfo
    staging_failed = pyqtSignal(str)
    staging_cancelled = pyqtSignal()

    def __init__(
        self,
        target_commit: CommitInfo,
        local_info: LocalInstallInfo,
        parent: Optional[QWidget] = None,
    ) -> None:
        super().__init__(parent)
        self._target_commit = target_commit
        self._local_info = local_info
        self._cancelled = False

    def request_cancel(self) -> None:
        self._cancelled = True

    def run(self) -> None:
        try:
            self.progress_changed.emit(0, f"Downloading {self._target_commit.short_sha} from GitHub...")
            staged = stage_zip_update(
                target_commit=self._target_commit,
                local_info=self._local_info,
                progress_cb=lambda pct: self.progress_changed.emit(
                    pct,
                    f"Downloading archive ({pct}%)...",
                ),
                cancel_cb=lambda: self._cancelled or self.isInterruptionRequested(),
                timeout=90,
            )
            self.progress_changed.emit(100, "Archive verified and staged.")
            self.staging_finished.emit(staged)
        except Exception as exc:
            if self._cancelled or "cancel" in str(exc).lower():
                self.staging_cancelled.emit()
            else:
                logger.exception("[Updater] Failed to stage update: %s", exc)
                self.staging_failed.emit(str(exc))
