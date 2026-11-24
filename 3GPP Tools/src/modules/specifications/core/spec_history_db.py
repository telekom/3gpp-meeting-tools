"""Backward-compatible re-export shim for `modules.specifications.core.history_db`."""

from __future__ import annotations

from modules.specifications.core.history_db import (
    SpecHistoryDatabase,
    logger,
)

__all__ = [
    "SpecHistoryDatabase",
    "logger",
]
