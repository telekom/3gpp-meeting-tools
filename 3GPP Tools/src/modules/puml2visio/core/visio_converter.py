"""Backward-compatible facade re-exporting `modules.puml2visio.core.visio`."""

import logging

from modules.puml2visio.core.visio import (
    ConverterThread,
    SvgConverterThread,
    VisioReaderThread,
)

logger = logging.getLogger(__name__)

__all__ = [
    "ConverterThread",
    "SvgConverterThread",
    "VisioReaderThread",
    "logger",
]
