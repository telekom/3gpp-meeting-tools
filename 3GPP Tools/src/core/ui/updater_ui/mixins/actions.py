"""Update check, state rendering, ZIP/Git application, and relaunch actions mixin."""

from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path

from PyQt5.QtCore import QTimer
from PyQt5.QtWidgets import QApplication, QMessageBox

from core.ui.updater_ui.html_renderer import (
    _format_commit_date,
    _format_local_baseline,
    _render_commits_html,
)
from core.ui.updater_ui.workers import AppUpdateCheckThread, AppUpdateDownloadThread
from core.updater.app_updater import (
    UpdateCheckResult,
    apply_git_update,
    launch_staged_update_and_relaunch,
    save_update_state,
)
from core.utils.paths import get_temp_root


class AppUpdateActionsMixin:
    """Interactive update workflow actions for AppUpdateDialog."""

    def start_check(self) -> None:
        if self._check_thread and self._check_thread.isRunning():
            return
        self.check_btn.setEnabled(False)
        self.apply_btn.setEnabled(False)
        self.status_title_lbl.setText("Checking GitHub master for updates...")
        self.notice_lbl.setVisible(False)

        self._check_thread = AppUpdateCheckThread(record_check_time=True, parent=self)
        self._check_thread.result_ready.connect(self._on_check_finished)
        self._check_thread.start()

    def _on_check_finished(self, result: UpdateCheckResult) -> None:
        self.check_btn.setEnabled(True)
        self._result = result
        self._render_result(result)

    def _render_result(self, result: UpdateCheckResult) -> None:
        local = result.local_info
        self.local_ver_lbl.setText(_format_local_baseline(local))

        if result.latest_commit is not None:
            latest_date = _format_commit_date(result.latest_commit)
            self.remote_ver_lbl.setText(f"{result.latest_commit.short_sha} · {latest_date}")
        else:
            self.remote_ver_lbl.setText("Unavailable")

        self.notice_lbl.setVisible(False)
        self.skip_btn.setVisible(False)
        self.force_zip_btn.setVisible(False)

        if result.status == "error":
            self.status_title_lbl.setText("Could not check for updates")
            self.notice_lbl.setText(
                f"{result.error}\nIf you are behind a corporate proxy, verify Settings -> Proxy Settings."
            )
            self.notice_lbl.setVisible(True)
            self.apply_btn.setEnabled(False)
            self.changes_browser.setHtml(_render_commits_html(result))
            return

        if result.status == "update_available" and result.latest_commit is not None:
            count = len(result.pending_commits)
            noun = "change" if count == 1 else "changes"
            self.status_title_lbl.setText(f"Update available ({count} new {noun} on master)")
            self.skip_btn.setVisible(True)

            if local.install_mode == "git":
                if local.is_git_dirty:
                    self.notice_lbl.setText(
                        "Developer Git workspace has uncommitted local code changes. "
                        "Commit or stash your edits and run 'git pull', or click 'Overwrite via ZIP Update...' "
                        "to replace application code while preserving your databases and settings."
                    )
                    self.notice_lbl.setVisible(True)
                    self.apply_btn.setText("Git Workspace Dirty")
                    self.apply_btn.setEnabled(False)
                    self.force_zip_btn.setVisible(True)
                    self.force_zip_btn.setEnabled(True)
                else:
                    self.apply_btn.setText("Git Pull && Restart")
                    self.apply_btn.setEnabled(True)
            else:
                if local.install_mode == "zip_manual":
                    self.notice_lbl.setText(
                        "Installed from a manual GitHub ZIP without a version marker "
                        "(detected from file timestamps). Updating once locks your exact commit ID."
                    )
                    self.notice_lbl.setVisible(True)
                self.apply_btn.setText("Update && Restart")
                self.apply_btn.setEnabled(True)
        else:
            self.status_title_lbl.setText("You are up to date")
            self.update_status_cleared.emit()
            if local.install_mode == "zip_manual" and result.latest_commit is not None:
                self.notice_lbl.setText(
                    "Your files match the latest commit timestamp. You can still click "
                    "'Re-sync Latest' to pin your exact commit ID for future incremental diffs."
                )
                self.notice_lbl.setVisible(True)
                self.apply_btn.setText("Re-sync Latest && Restart")
                self.apply_btn.setEnabled(True)
            else:
                self.apply_btn.setText("Up to Date")
                self.apply_btn.setEnabled(False)

        self.changes_browser.setHtml(_render_commits_html(result))

    def _skip_current_update(self) -> None:
        if self._result and self._result.latest_commit:
            save_update_state({"skipped_sha": self._result.latest_commit.sha})
            self.update_status_cleared.emit()
            self.close()

    def _confirm_and_force_zip_update(self) -> None:
        if not self._result or not self._result.latest_commit:
            return
        reply = QMessageBox.question(
            self,
            "Overwrite Local Git Changes?",
            "This will download the latest GitHub snapshot and overwrite modified application "
            "source files in '3GPP Tools'.\n\n"
            "Your local databases (3gpp_data.db), proxy/AI settings, user presets, PlantUML JAR, "
            "and output/ files will be preserved.\n\n"
            "Proceed with ZIP update?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if reply == QMessageBox.Yes:
            self._start_zip_download()

    def _start_apply_update(self) -> None:
        if not self._result or not self._result.latest_commit:
            return

        local = self._result.local_info
        if local.install_mode == "git":
            self._apply_git_and_restart()
            return

        self._start_zip_download()

    def _start_zip_download(self) -> None:
        if not self._result or not self._result.latest_commit:
            return
        if self._download_thread and self._download_thread.isRunning():
            return

        self.check_btn.setEnabled(False)
        self.apply_btn.setEnabled(False)
        self.force_zip_btn.setEnabled(False)
        self.skip_btn.setEnabled(False)
        self.close_btn.setEnabled(False)
        self.progress_widget.setVisible(True)
        self.progress_bar.setValue(0)
        self.progress_label.setText("Preparing download...")
        self.cancel_dl_btn.setVisible(True)

        self._download_thread = AppUpdateDownloadThread(
            target_commit=self._result.latest_commit,
            local_info=self._result.local_info,
            parent=self,
        )
        self._download_thread.progress_changed.connect(self._on_download_progress)
        self._download_thread.staging_finished.connect(self._on_staging_finished)
        self._download_thread.staging_failed.connect(self._on_staging_failed)
        self._download_thread.staging_cancelled.connect(self._on_staging_cancelled)
        self._download_thread.start()

    def _apply_git_and_restart(self) -> None:
        self.check_btn.setEnabled(False)
        self.apply_btn.setEnabled(False)
        self.progress_widget.setVisible(True)
        self.progress_bar.setRange(0, 0)
        self.progress_label.setText("Running git pull --ff-only...")
        QApplication.processEvents()

        ok, msg, _req_changed = apply_git_update(self._result.local_info.app_root)
        self.progress_bar.setRange(0, 100)
        if not ok:
            self.progress_widget.setVisible(False)
            self.check_btn.setEnabled(True)
            self.apply_btn.setEnabled(True)
            reply = QMessageBox.question(
                self,
                "Git Update Failed",
                f"Git pull could not complete automatically:\n\n{msg}\n\n"
                "Would you like to apply the update using the GitHub ZIP updater instead "
                "(preserving your databases, settings, user presets, and output/ files)?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.Yes,
            )
            if reply == QMessageBox.Yes:
                self._start_zip_download()
            return

        self.progress_bar.setValue(100)
        self.progress_label.setText("Git update complete. Restarting application...")
        self._relaunch_and_close_main_window(self._result.local_info.app_root)

    def _on_download_progress(self, pct: int, msg: str) -> None:
        self.progress_bar.setValue(max(0, min(100, int(pct))))
        self.progress_label.setText(msg)

    def _cancel_download(self) -> None:
        if self._download_thread and self._download_thread.isRunning():
            self.cancel_dl_btn.setEnabled(False)
            self.progress_label.setText("Cancelling download...")
            self._download_thread.request_cancel()

    def _reset_download_ui(self) -> None:
        self.progress_widget.setVisible(False)
        self.cancel_dl_btn.setVisible(False)
        self.cancel_dl_btn.setEnabled(True)
        self.check_btn.setEnabled(True)
        self.skip_btn.setEnabled(True)
        self.close_btn.setEnabled(True)
        if self._result:
            self._render_result(self._result)

    def _on_staging_cancelled(self) -> None:
        self._reset_download_ui()

    def _on_staging_failed(self, err: str) -> None:
        self._reset_download_ui()
        QMessageBox.warning(
            self,
            "Update Download Failed",
            f"Could not download or stage the update:\n\n{err}",
        )

    def _on_staging_finished(self, staged) -> None:
        self.cancel_dl_btn.setVisible(False)
        self.progress_label.setText("Launching updater and restarting...")
        try:
            from core.ui.window_session import mark_pending_restart_restore

            mark_pending_restart_restore()
        except Exception:
            pass
        try:
            launch_staged_update_and_relaunch(staged, parent_pid=os.getpid(), relaunch=True)
        except Exception as exc:
            self._reset_download_ui()
            QMessageBox.critical(
                self,
                "Update Launch Failed",
                f"Staged update could not start:\n\n{exc}",
            )
            return

        self.accept()
        main_win = self.parentWidget()
        if main_win is not None:
            QTimer.singleShot(0, main_win.close)
            QTimer.singleShot(150, QApplication.quit)
        else:
            QTimer.singleShot(0, QApplication.quit)

    def _relaunch_and_close_main_window(self, app_root: Path) -> None:
        """Spawn a detached delayed relauncher for Git pull restarts, then close main UI."""
        try:
            from core.ui.window_session import mark_pending_restart_restore

            mark_pending_restart_restore()
        except Exception:
            pass
        work_root = get_temp_root() / "app_update"
        work_root.mkdir(parents=True, exist_ok=True)
        helper_code = (
            "import ctypes, os, subprocess, sys, time\n"
            f"pid = {os.getpid()}\n"
            f"app_root = r'''{str(app_root)}'''\n"
            f"py_exe = r'''{sys.executable}'''\n"
            "if os.name == 'nt' and pid > 0:\n"
            "    k32 = ctypes.windll.kernel32\n"
            "    k32.OpenProcess.argtypes = [ctypes.c_uint32, ctypes.c_int, ctypes.c_uint32]\n"
            "    k32.OpenProcess.restype = ctypes.c_void_p\n"
            "    k32.WaitForSingleObject.argtypes = [ctypes.c_void_p, ctypes.c_uint32]\n"
            "    k32.CloseHandle.argtypes = [ctypes.c_void_p]\n"
            "    h = k32.OpenProcess(0x00100000, 0, pid)\n"
            "    if h:\n"
            "        k32.WaitForSingleObject(h, 20000)\n"
            "        k32.CloseHandle(h)\n"
            "    time.sleep(0.4)\n"
            "else:\n"
            "    time.sleep(1.2)\n"
            "launch_cmd = os.path.join(app_root, 'src', 'Launch.cmd')\n"
            "if os.name == 'nt' and os.path.isfile(launch_cmd):\n"
            "    subprocess.Popen(['cmd.exe', '/c', launch_cmd], cwd=app_root, creationflags=subprocess.CREATE_NEW_CONSOLE)\n"
            "else:\n"
            "    env = os.environ.copy()\n"
            "    env['PYTHONPATH'] = app_root + os.pathsep + os.path.join(app_root, 'src')\n"
            "    subprocess.Popen([py_exe, '-m', 'main_tools'], cwd=app_root, env=env)\n"
        )
        helper_path = work_root / "relaunch_after_git.py"
        helper_path.write_text(helper_code, encoding="utf-8")
        kwargs = {"cwd": str(work_root)}
        if os.name == "nt":
            kwargs["creationflags"] = 0x08000000
        subprocess.Popen([sys.executable, str(helper_path)], **kwargs)

        self.accept()
        main_win = self.parentWidget()
        if main_win is not None:
            QTimer.singleShot(0, main_win.close)
            QTimer.singleShot(150, QApplication.quit)
        else:
            QTimer.singleShot(0, QApplication.quit)
