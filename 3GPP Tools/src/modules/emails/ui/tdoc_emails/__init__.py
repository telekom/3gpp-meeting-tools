"""TDoc-related email viewer package."""

from modules.emails.ui.tdoc_emails.dialog import TDocEmailsDialog, logger
from modules.emails.ui.tdoc_emails.standalone_reader import StandaloneEmailReaderWindow

__all__ = [
    "StandaloneEmailReaderWindow",
    "TDocEmailsDialog",
    "logger",
]
