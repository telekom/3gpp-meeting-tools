# --- File: src/modules/meetings/ui/tdocs_window.py ---
"""
Backward-compatible import shim for the TDocs window.

The implementation was split into the ``tdocs/window/`` package (see
tdocs/window/README.md for a 'where to edit what' map).

Note for tests: patch module globals where they are *used* (e.g.
``modules.meetings.ui.tdocs.window.mixins.exports._open_folder``); patching them here
has no effect. Class-attribute patches (``QMessageBox.warning``,
``OutlookClient.create_email_draft``) still work through this module.
"""
from PyQt5.QtWidgets import QMessageBox  # noqa: F401  (kept for patch("...tdocs_window.QMessageBox.x"))

from modules.emails.core.outlook_client import OutlookClient  # noqa: F401
from modules.meetings.ui.tdocs.window import TDocsWindow  # noqa: F401

__all__ = ["TDocsWindow", "QMessageBox", "OutlookClient"]
