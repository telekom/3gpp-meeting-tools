# --- File: modules/meetings/ui/ui_tabs.py ---
"""Backward-compatible re-export shim for MeetingsTab.

The implementation was split into ``modules.meetings.ui.meetings_tab``
(see ``meetings_tab/README.md`` for a submodule map).
"""

from __future__ import annotations

from .meetings_tab import MeetingsTab  # noqa: F401

__all__ = ["MeetingsTab"]
