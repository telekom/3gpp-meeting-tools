"""Feature mixins composing TDocEmailsDialog."""

from modules.emails.ui.tdoc_emails.mixins.actions import TDocEmailsActionsMixin
from modules.emails.ui.tdoc_emails.mixins.rendering import TDocEmailsRenderingMixin
from modules.emails.ui.tdoc_emails.mixins.ui_setup import TDocEmailsUiSetupMixin

__all__ = [
    "TDocEmailsActionsMixin",
    "TDocEmailsRenderingMixin",
    "TDocEmailsUiSetupMixin",
]
