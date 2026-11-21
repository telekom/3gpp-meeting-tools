"""Backward-compatible re-export shim for TopicInputsViewerDialog."""

from modules.meetings.ui.dialogs.topic_inputs import (
    BADGE_COLORS,
    STYLE_PRIMARY_BTN,
    STYLE_SECONDARY_BTN,
    STYLE_TABLE,
    InputsCollectorThread,
    TopicInputsViewerDialog,
    logger as _dialog_logger,
)

logger = _dialog_logger

__all__ = [
    "TopicInputsViewerDialog",
    "InputsCollectorThread",
    "STYLE_PRIMARY_BTN",
    "STYLE_SECONDARY_BTN",
    "STYLE_TABLE",
    "BADGE_COLORS",
]
