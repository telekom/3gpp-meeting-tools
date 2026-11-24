"""Backward-compatible re-export shim for Application Self-Update dialog and workers.

Canonical implementation lives in ``core.ui.updater_ui``.
"""

from core.ui.updater_ui import (
    AppUpdateCheckThread,
    AppUpdateDialog,
    AppUpdateDownloadThread,
    _format_commit_date,
    _format_local_baseline,
    _render_commits_html,
    logger,
)

__all__ = [
    "AppUpdateCheckThread",
    "AppUpdateDialog",
    "AppUpdateDownloadThread",
    "_format_commit_date",
    "_format_local_baseline",
    "_render_commits_html",
    "logger",
]
