"""Feature mixins composing `CallFlowDialog`."""

from __future__ import annotations

from modules.puml2visio.ui.callflow.mixins.generation_and_validation import (
    CallFlowGenerationMixin,
)
from modules.puml2visio.ui.callflow.mixins.ui_setup import CallFlowUiSetupMixin

__all__ = [
    "CallFlowGenerationMixin",
    "CallFlowUiSetupMixin",
]
