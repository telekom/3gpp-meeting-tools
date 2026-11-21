"""Backward-compatible re-export shim for CrPcrDialog."""

from modules.word_tools.ui.cr_pcr import (
    BUTTON_STYLE_PRIMARY,
    CrPcrDialog,
    logger as _dialog_logger,
)

logger = _dialog_logger

__all__ = [
    "CrPcrDialog",
    "BUTTON_STYLE_PRIMARY",
]
