"""Public facade for the Specification Change History SQLite database package."""

from __future__ import annotations

from modules.specifications.core.history_db.database import (
    SpecHistoryDatabase,
    logger,
)

__all__ = [
    "SpecHistoryDatabase",
    "logger",
]
