# --- File: src/modules/meetings/core/exporters/pre_meeting_exporter.py ---
"""
Backward-compatible import shim for the Pre-Meeting Markdown Dossier Exporter.

The implementation was split into ``src/modules/meetings/core/exporters/pre_meeting/``
(see ``pre_meeting/README.md`` for a 'where to edit what' map).
"""

from modules.meetings.core.exporters.pre_meeting import (
    PreMeetingExporterThread,
    logger,
)

__all__ = [
    "PreMeetingExporterThread",
    "logger",
]
