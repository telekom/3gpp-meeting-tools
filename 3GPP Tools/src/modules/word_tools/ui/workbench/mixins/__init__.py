"""Feature mixins composing WordExtractorTab."""

from __future__ import annotations

from .compat_and_callbacks import WordTabCompatAndCallbacksMixin
from .status_and_poll import WordTabStatusAndPollMixin
from .ui_setup import WordTabUiSetupMixin

__all__ = [
    "WordTabUiSetupMixin",
    "WordTabStatusAndPollMixin",
    "WordTabCompatAndCallbacksMixin",
]
