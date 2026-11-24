"""AppUpdateDialog composing UI setup, commit HTML feed, and update/relaunch actions."""

from __future__ import annotations

from typing import Optional

from PyQt5.QtCore import QTimer, pyqtSignal
from PyQt5.QtWidgets import QDialog, QWidget

from core.ui.updater_ui.html_renderer import _format_local_baseline
from core.ui.updater_ui.mixins import AppUpdateActionsMixin, AppUpdateUiSetupMixin
from core.ui.updater_ui.workers import AppUpdateCheckThread, AppUpdateDownloadThread
from core.updater.app_updater import UpdateCheckResult, inspect_local_install


class AppUpdateDialog(AppUpdateUiSetupMixin, AppUpdateActionsMixin, QDialog):
    """Modal dialog displaying available `master` updates, commit details, and 1-click apply."""

    update_status_cleared = pyqtSignal()

    def __init__(
        self,
        initial_result: Optional[UpdateCheckResult] = None,
        parent: Optional[QWidget] = None,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle("3GPP Delegate Tools — Application Updates")
        self.setMinimumSize(680, 520)
        self.resize(740, 580)

        self._result: Optional[UpdateCheckResult] = initial_result
        self._check_thread: Optional[AppUpdateCheckThread] = None
        self._download_thread: Optional[AppUpdateDownloadThread] = None

        self._setup_ui()
        self._apply_styles()

        if self._result is not None:
            self._render_result(self._result)
        else:
            local_info = inspect_local_install()
            self.local_ver_lbl.setText(_format_local_baseline(local_info))
            self.remote_ver_lbl.setText("Checking GitHub...")
            QTimer.singleShot(50, self.start_check)

    def closeEvent(self, event) -> None:
        if self._download_thread and self._download_thread.isRunning():
            self._download_thread.request_cancel()
            self._download_thread.wait(500)
        if self._check_thread and self._check_thread.isRunning():
            self._check_thread.requestInterruption()
            self._check_thread.wait(300)
        super().closeEvent(event)
