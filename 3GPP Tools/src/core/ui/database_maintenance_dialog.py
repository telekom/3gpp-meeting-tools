# --- File: src/core/ui/database_maintenance_dialog.py ---
"""Backward-compatible re-export shim for :mod:`core.ui.database_maintenance`."""

from __future__ import annotations

from core.ui.database_maintenance import (
    DatabaseMaintenanceDialog,
    KpiCard,
    format_file_size,
    logger as _db_maint_logger,
)

logger = _db_maint_logger

__all__ = [
    "DatabaseMaintenanceDialog",
    "KpiCard",
    "format_file_size",
]
