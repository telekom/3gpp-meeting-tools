"""Modular Fluent-style drag & drop target package for 3GPP Tools."""

from .metadata import (
    COLOR_BUSY_BG,
    COLOR_BUSY_BORDER,
    COLOR_DRAG_ACCENT,
    COLOR_DRAG_BG,
    COLOR_DRAG_BORDER,
    COLOR_ERROR_BG,
    COLOR_ERROR_BORDER,
    COLOR_IDLE_BG,
    COLOR_IDLE_BORDER,
    COLOR_IDLE_HOVER_BORDER,
    COLOR_LOADED_BG,
    COLOR_LOADED_BORDER,
    COLOR_TEXT_MUTED,
    COLOR_TEXT_PRIMARY,
    COLOR_TEXT_SECONDARY,
    DEFAULT_FILE_TYPE,
    FILE_TYPE_METADATA,
    format_file_size,
)
from .widget import (
    ClickableDropLabel,
    InteractiveDropLabel,
    ModernDropZone,
    logger as _widget_logger,
)

logger = _widget_logger

__all__ = [
    "COLOR_BUSY_BG",
    "COLOR_BUSY_BORDER",
    "COLOR_DRAG_ACCENT",
    "COLOR_DRAG_BG",
    "COLOR_DRAG_BORDER",
    "COLOR_ERROR_BG",
    "COLOR_ERROR_BORDER",
    "COLOR_IDLE_BG",
    "COLOR_IDLE_BORDER",
    "COLOR_IDLE_HOVER_BORDER",
    "COLOR_LOADED_BG",
    "COLOR_LOADED_BORDER",
    "COLOR_TEXT_MUTED",
    "COLOR_TEXT_PRIMARY",
    "COLOR_TEXT_SECONDARY",
    "ClickableDropLabel",
    "DEFAULT_FILE_TYPE",
    "FILE_TYPE_METADATA",
    "InteractiveDropLabel",
    "ModernDropZone",
    "format_file_size",
]
