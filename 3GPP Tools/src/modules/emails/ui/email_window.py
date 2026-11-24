"""Backward-compatible re-export shim for EmailManagerWindow.

Canonical implementation lives in ``modules.emails.ui.manager_window``.
"""

from modules.emails.ui.manager_window import EditEmailDialog, EmailManagerWindow

__all__ = ["EditEmailDialog", "EmailManagerWindow"]
