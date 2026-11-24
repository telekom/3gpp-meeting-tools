"""Commit history HTML and baseline formatting helpers for AppUpdateDialog."""

from __future__ import annotations

import datetime
import html
from typing import List

from core.config.config import GITHUB_BRANCH
from core.ui.theme import C, T
from core.updater.app_updater import CommitInfo, LocalInstallInfo, UpdateCheckResult


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
