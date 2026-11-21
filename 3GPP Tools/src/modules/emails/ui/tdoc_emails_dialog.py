"""Backward-compatible re-export shim for TDocEmailsDialog and StandaloneEmailReaderWindow."""

from core.ai.cli_launcher import launch_cli_in_terminal
from modules.emails.ui.tdoc_emails import (
    StandaloneEmailReaderWindow,
    TDocEmailsDialog,
    logger as _dialog_logger,
)

logger = _dialog_logger

__all__ = [
    "StandaloneEmailReaderWindow",
    "TDocEmailsDialog",
    "launch_cli_in_terminal",
]
