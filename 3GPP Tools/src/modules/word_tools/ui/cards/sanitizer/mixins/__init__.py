"""Feature mixins composing WordSanitizerCard."""

from __future__ import annotations

from .audit_and_actions import SanitizerActionsMixin
from .ui_setup import SanitizerUiSetupMixin

__all__ = [
    "SanitizerUiSetupMixin",
    "SanitizerActionsMixin",
]
