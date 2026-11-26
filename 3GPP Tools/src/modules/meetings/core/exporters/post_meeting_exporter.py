# --- File: src/modules/meetings/core/exporters/post_meeting_exporter.py ---
"""
Backward-compatible import shim for the Post-Meeting Markdown Dossier Exporter.

The implementation was split into ``src/modules/meetings/core/exporters/post_meeting/``
(see ``post_meeting/README.md`` for a 'where to edit what' map).
"""
from modules.meetings.core.exporters.post_meeting import (
    PostMeetingExporterThread,
    logger,
    quote_markdown_path,
    safe_markdown_rel_link,
    safe_relpath,
    sanitize_folder_name,
    slugify_anchor,
)

__all__ = [
    "PostMeetingExporterThread",
    "logger",
    "sanitize_folder_name",
    "safe_relpath",
    "quote_markdown_path",
    "safe_markdown_rel_link",
    "slugify_anchor",
]
