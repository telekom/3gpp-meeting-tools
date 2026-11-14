# --- File: src/core/ui/update_dialog.py ---
"""
Application Self-Update Dialog & Background Workers for 3GPP Delegate Tools.

Displays continuous commit changes from GitHub `master`, allows checking on
demand, skipping a commit, or downloading and applying the latest update while
preserving user databases and configuration files.
"""

from __future__ import annotations

import datetime
import html
import logging
import os
import subprocess
import sys
from pathlib import Path
from typing import List, Optional

from PyQt5.QtCore import QSize, Qt, QThread, QTimer, pyqtSignal
from PyQt5.QtWidgets import (
    QApplication,
    QDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QTextBrowser,
    QVBoxLayout,
    QWidget,
)

from core.config.config import GITHUB_BRANCH, HELP_URL
from core.ui.icons import lucide_icon
from core.ui.theme import C, R, S, T
from core.updater.app_updater import (
    CommitInfo,
    LocalInstallInfo,
    UpdateCheckResult,
    apply_git_update,
    check_for_updates,
    inspect_local_install,
    launch_staged_update_and_relaunch,
    save_update_state,
    stage_zip_update,
)
from core.utils.paths import get_temp_root

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


def _format_commit_date(commit: CommitInfo) -> str:
    if commit.updated_ts > 0:
        try:
            dt = datetime.datetime.fromtimestamp(commit.updated_ts)
            return dt.strftime("%Y-%m-%d %H:%M")
        except Exception:
            pass
    return commit.updated_iso or ""


def _format_local_baseline(local: LocalInstallInfo) -> str:
    if local.install_mode == "git":
        dirty_tag = " (uncommitted local edits)" if local.is_git_dirty else ""
        sha_part = local.short_sha or "unknown"
        return f"Git checkout · {sha_part}{dirty_tag}"
    if local.install_mode == "zip_tracked":
        sha_part = local.short_sha or "unknown"
        return f"Tracked ZIP · {sha_part}"
    if local.estimated_ts > 0:
        try:
            dt_str = datetime.datetime.fromtimestamp(local.estimated_ts).strftime("%Y-%m-%d %H:%M")
            return f"Manual ZIP · ~{dt_str}"
        except Exception:
            pass
    return "Manual ZIP · unknown baseline"


def _render_commits_html(result: UpdateCheckResult) -> str:
    """Render pending and recent `master` commits into styled HTML."""
    pending_shas = {c.sha for c in result.pending_commits}
    commits = result.pending_commits or result.all_commits

    if not commits:
        return (
            f"<div style='color:{C.TEXT_MUTED}; font-family:{T.FONT_FAMILY_UI}; padding:12px;'>"
            "No commit history loaded yet. Click <b>Check Now</b> to fetch the latest changes from GitHub."
            "</div>"
        )

    parts: List[str] = [
        f"<div style='font-family:{T.FONT_FAMILY_UI}; font-size:12px; color:{C.TEXT_BODY};'>"
    ]

    if result.status == "update_available" and result.pending_commits:
        count = len(result.pending_commits)
        noun = "change" if count == 1 else "changes"
        parts.append(
            f"<div style='margin-bottom:10px; font-weight:600; color:{C.ACCENT_PRIMARY};'>"
            f"New on <code>{html.escape(GITHUB_BRANCH)}</code> ({count} {noun} since your version):"
            f"</div>"
        )
        if result.history_truncated:
            parts.append(
                f"<div style='margin-bottom:10px; padding:6px 10px; background:{C.INFO_SOFT}; "
                f"border:1px solid {C.INFO_BORDER}; border-radius:6px; color:{C.TEXT_BODY}; font-size:11px;'>"
                "Showing the latest 20 commits from GitHub. Updating applies the complete latest snapshot."
                "</div>"
            )
    else:
        parts.append(
            f"<div style='margin-bottom:10px; font-weight:600; color:{C.TEXT_SECONDARY};'>"
            f"Recent commits on <code>{html.escape(GITHUB_BRANCH)}</code>:"
            f"</div>"
        )

    for c in commits:
        is_new = c.sha in pending_shas
        card_bg = C.BG_CARD if is_new else C.BG_SURFACE
        border_col = C.ACCENT_RING if is_new else C.BORDER_SUBTLE
        badge_html = (
            f"<span style='background:{C.ACCENT_SOFT}; color:{C.ACCENT_PRIMARY}; "
            f"padding:1px 6px; border-radius:4px; font-size:10px; font-weight:700;'>NEW</span> "
            if is_new
            else ""
        )
        msg_lines = [ln for ln in (c.message or c.title).splitlines()]
        title_line = html.escape(msg_lines[0] if msg_lines else c.title)
        body_lines = "\n".join(msg_lines[1:]).strip()
        body_html = ""
        if body_lines:
            body_html = (
                f"<div style='margin-top:6px; white-space:pre-wrap; color:{C.TEXT_SECONDARY}; "
                f"font-size:11px; line-height:1.4;'>{html.escape(body_lines)}</div>"
            )

        date_str = html.escape(_format_commit_date(c))
        author_str = html.escape(c.author)
        sha_link = (
            f"<a href='{html.escape(c.url)}' style='color:{C.ACCENT_PRIMARY}; text-decoration:none; "
            f"font-family:Consolas,monospace; font-size:11px;'>{html.escape(c.short_sha)}</a>"
        )

        parts.append(
            f"<div style='background:{card_bg}; border:1px solid {border_col}; border-radius:8px; "
            f"padding:9px 11px; margin-bottom:8px;'>"
            f"<div>{badge_html}<span style='font-weight:600; color:{C.TEXT_PRIMARY};'>{title_line}</span></div>"
            f"{body_html}"
            f"<div style='margin-top:5px; font-size:11px; color:{C.TEXT_MUTED};'>"
            f"{sha_link} &nbsp;&middot;&nbsp; {date_str} &nbsp;&middot;&nbsp; {author_str}"
            f"</div>"
            f"</div>"
        )

    parts.append("</div>")
    return "".join(parts)


class AppUpdateDialog(QDialog):
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

    def _setup_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(S.LG, S.LG, S.LG, S.LG)
        layout.setSpacing(S.MD)

        # Header summary card
        self.summary_card = QFrame()
        self.summary_card.setObjectName("updateSummaryCard")
        card_layout = QVBoxLayout(self.summary_card)
        card_layout.setContentsMargins(S.MD, S.MD, S.MD, S.MD)
        card_layout.setSpacing(6)

        top_row = QHBoxLayout()
        self.status_title_lbl = QLabel("Checking for updates...")
        self.status_title_lbl.setObjectName("updateStatusTitle")
        top_row.addWidget(self.status_title_lbl, 1)

        self.open_repo_btn = QPushButton("Open GitHub Repo")
        self.open_repo_btn.setCursor(Qt.PointingHandCursor)
        self.open_repo_btn.setIcon(lucide_icon("external-link", color=C.TEXT_SECONDARY, size=14))
        self.open_repo_btn.setIconSize(QSize(14, 14))
        self.open_repo_btn.clicked.connect(self._open_github_commits)
        top_row.addWidget(self.open_repo_btn)
        card_layout.addLayout(top_row)

        meta_row = QHBoxLayout()
        meta_row.setSpacing(S.LG)

        col_local = QVBoxLayout()
        col_local.setSpacing(2)
        lbl_l = QLabel("Installed Version")
        lbl_l.setObjectName("updateMetaCaption")
        self.local_ver_lbl = QLabel("Inspecting...")
        self.local_ver_lbl.setObjectName("updateMetaValue")
        col_local.addWidget(lbl_l)
        col_local.addWidget(self.local_ver_lbl)
        meta_row.addLayout(col_local, 1)

        col_remote = QVBoxLayout()
        col_remote.setSpacing(2)
        lbl_r = QLabel("Latest on GitHub (master)")
        lbl_r.setObjectName("updateMetaCaption")
        self.remote_ver_lbl = QLabel("—")
        self.remote_ver_lbl.setObjectName("updateMetaValue")
        col_remote.addWidget(lbl_r)
        col_remote.addWidget(self.remote_ver_lbl)
        meta_row.addLayout(col_remote, 1)

        card_layout.addLayout(meta_row)

        self.notice_lbl = QLabel("")
        self.notice_lbl.setObjectName("updateNoticeLabel")
        self.notice_lbl.setWordWrap(True)
        self.notice_lbl.setVisible(False)
        card_layout.addWidget(self.notice_lbl)

        layout.addWidget(self.summary_card)

        # Changes viewer
        feed_header = QLabel("What's Changed on master")
        feed_header.setObjectName("updateSectionHeader")
        layout.addWidget(feed_header)

        self.changes_browser = QTextBrowser()
        self.changes_browser.setObjectName("updateChangesBrowser")
        self.changes_browser.setOpenExternalLinks(True)
        layout.addWidget(self.changes_browser, 1)

        # Preservation info
        pres_lbl = QLabel(
            "Local databases (3gpp_data.db), proxy/AI settings, user presets, PlantUML JAR, "
            "and output/ files are always preserved."
        )
        pres_lbl.setObjectName("updatePreservationHint")
        pres_lbl.setWordWrap(True)
        layout.addWidget(pres_lbl)

        # Progress row (visible during download/staging)
        self.progress_widget = QWidget()
        prog_layout = QVBoxLayout(self.progress_widget)
        prog_layout.setContentsMargins(0, 0, 0, 0)
        prog_layout.setSpacing(4)
        self.progress_label = QLabel("")
        self.progress_label.setObjectName("updateMetaCaption")
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setTextVisible(True)
        prog_layout.addWidget(self.progress_label)
        prog_layout.addWidget(self.progress_bar)
        self.progress_widget.setVisible(False)
        layout.addWidget(self.progress_widget)

        # Action bar
        btn_row = QHBoxLayout()
        btn_row.setSpacing(S.SM)

        self.check_btn = QPushButton("Check Now")
        self.check_btn.setCursor(Qt.PointingHandCursor)
        self.check_btn.setIcon(lucide_icon("refresh-cw", color=C.TEXT_SECONDARY, size=14))
        self.check_btn.setIconSize(QSize(14, 14))
        self.check_btn.clicked.connect(self.start_check)
        btn_row.addWidget(self.check_btn)

        self.skip_btn = QPushButton("Skip This Version")
        self.skip_btn.setCursor(Qt.PointingHandCursor)
        self.skip_btn.clicked.connect(self._skip_current_update)
        self.skip_btn.setVisible(False)
        btn_row.addWidget(self.skip_btn)

        btn_row.addStretch(1)

        self.cancel_dl_btn = QPushButton("Cancel Download")
        self.cancel_dl_btn.setCursor(Qt.PointingHandCursor)
        self.cancel_dl_btn.clicked.connect(self._cancel_download)
        self.cancel_dl_btn.setVisible(False)
        btn_row.addWidget(self.cancel_dl_btn)

        self.apply_btn = QPushButton("Update && Restart")
        self.apply_btn.setObjectName("updatePrimaryBtn")
        self.apply_btn.setCursor(Qt.PointingHandCursor)
        self.apply_btn.setIcon(lucide_icon("download", color=C.TEXT_ON_ACCENT, size=14))
        self.apply_btn.setIconSize(QSize(14, 14))
        self.apply_btn.clicked.connect(self._start_apply_update)
        self.apply_btn.setEnabled(False)
        btn_row.addWidget(self.apply_btn)

        self.close_btn = QPushButton("Close")
        self.close_btn.setCursor(Qt.PointingHandCursor)
        self.close_btn.clicked.connect(self.close)
        btn_row.addWidget(self.close_btn)

        layout.addLayout(btn_row)

    def _apply_styles(self) -> None:
        self.setStyleSheet(
            f"""
            QDialog {{
                background-color: {C.BG_CANVAS};
            }}
            QFrame#updateSummaryCard {{
                background-color: {C.BG_CARD};
                border: 1px solid {C.BORDER_SUBTLE};
                border-radius: {R.LG}px;
            }}
            QLabel#updateStatusTitle {{
                font-family: {T.FONT_FAMILY_UI};
                font-size: {T.SIZE_SECTION_TITLE}px;
                font-weight: 700;
                color: {C.TEXT_PRIMARY};
            }}
            QLabel#updateSectionHeader {{
                font-family: {T.FONT_FAMILY_UI};
                font-size: {T.SIZE_BODY}px;
                font-weight: 700;
                color: {C.TEXT_PRIMARY};
            }}
            QLabel#updateMetaCaption {{
                font-family: {T.FONT_FAMILY_UI};
                font-size: {T.SIZE_CAPTION}px;
                color: {C.TEXT_MUTED};
            }}
            QLabel#updateMetaValue {{
                font-family: {T.FONT_FAMILY_UI};
                font-size: {T.SIZE_BODY}px;
                font-weight: 600;
                color: {C.TEXT_PRIMARY};
            }}
            QLabel#updateNoticeLabel {{
                font-family: {T.FONT_FAMILY_UI};
                font-size: {T.SIZE_CAPTION}px;
                color: {C.WARNING};
                background-color: {C.WARNING_SOFT};
                border: 1px solid {C.WARNING_BORDER};
                border-radius: {R.SM}px;
                padding: 6px 8px;
            }}
            QLabel#updatePreservationHint {{
                font-family: {T.FONT_FAMILY_UI};
                font-size: {T.SIZE_CAPTION}px;
                color: {C.TEXT_MUTED};
            }}
            QTextBrowser#updateChangesBrowser {{
                background-color: {C.BG_CANVAS};
                border: 1px solid {C.BORDER_SUBTLE};
                border-radius: {R.MD}px;
                padding: 4px;
            }}
            QPushButton#updatePrimaryBtn {{
                background-color: {C.ACCENT_PRIMARY};
                color: {C.TEXT_ON_ACCENT};
                border: 1px solid {C.ACCENT_PRIMARY};
                border-radius: {R.MD}px;
                padding: 6px 14px;
                font-weight: 600;
            }}
            QPushButton#updatePrimaryBtn:hover {{
                background-color: {C.ACCENT_HOVER};
            }}
            QPushButton#updatePrimaryBtn:disabled {{
                background-color: {C.BORDER_STRONG};
                border-color: {C.BORDER_STRONG};
                color: {C.TEXT_MUTED};
            }}
            """
        )

    def _open_github_commits(self) -> None:
        import webbrowser

        url = "https://github.com/telekom/3gpp-meeting-tools/commits/master/3GPP%20Tools"
        try:
            webbrowser.open(url)
        except Exception:
            webbrowser.open(HELP_URL)

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
                        "Developer Git workspace has uncommitted local changes. "
                        "Commit or stash your edits and run 'git pull' manually to avoid losing work."
                    )
                    self.notice_lbl.setVisible(True)
                    self.apply_btn.setText("Git Workspace Dirty")
                    self.apply_btn.setEnabled(False)
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

    def _start_apply_update(self) -> None:
        if not self._result or not self._result.latest_commit:
            return

        local = self._result.local_info
        if local.install_mode == "git":
            self._apply_git_and_restart()
            return

        if self._download_thread and self._download_thread.isRunning():
            return

        self.check_btn.setEnabled(False)
        self.apply_btn.setEnabled(False)
        self.skip_btn.setEnabled(False)
        self.close_btn.setEnabled(False)
        self.progress_widget.setVisible(True)
        self.progress_bar.setValue(0)
        self.progress_label.setText("Preparing download...")
        self.cancel_dl_btn.setVisible(True)

        self._download_thread = AppUpdateDownloadThread(
            target_commit=self._result.latest_commit,
            local_info=local,
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
            QMessageBox.warning(self, "Git Update Failed", msg)
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
        else:
            QTimer.singleShot(0, QApplication.quit)

    def _relaunch_and_close_main_window(self, app_root: Path) -> None:
        """Spawn a detached delayed relauncher for Git pull restarts, then close main UI."""
        work_root = get_temp_root() / "app_update"
        work_root.mkdir(parents=True, exist_ok=True)
        helper_code = (
            "import os, subprocess, sys, time\n"
            f"pid = {os.getpid()}\n"
            f"app_root = r'''{str(app_root)}'''\n"
            f"py_exe = r'''{sys.executable}'''\n"
            "time.sleep(1.2)\n"
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
        else:
            QTimer.singleShot(0, QApplication.quit)

    def closeEvent(self, event) -> None:
        if self._download_thread and self._download_thread.isRunning():
            self._download_thread.request_cancel()
            self._download_thread.wait(500)
        if self._check_thread and self._check_thread.isRunning():
            self._check_thread.requestInterruption()
            self._check_thread.wait(300)
        super().closeEvent(event)
