"""Modular Application Self-Update UI dialog, HTML commit renderer, and worker threads."""

from core.ui.updater_ui.dialog import AppUpdateDialog
from core.ui.updater_ui.html_renderer import (
    _format_commit_date,
    _format_local_baseline,
    _render_commits_html,
)
from core.ui.updater_ui.workers import AppUpdateCheckThread, AppUpdateDownloadThread, logger

__all__ = [
    "AppUpdateCheckThread",
    "AppUpdateDialog",
    "AppUpdateDownloadThread",
    "_format_commit_date",
    "_format_local_baseline",
    "_render_commits_html",
    "logger",
]
