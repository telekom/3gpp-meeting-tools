"""Modular SQLite Database Maintenance & Compaction Hub package."""

from .dialog import DatabaseMaintenanceDialog, logger as _dialog_logger
from .storage_ops import (
    format_file_size,
    get_tracked_databases,
    query_db_metrics,
    vacuum_database_file,
)
from .widgets import KpiCard

logger = _dialog_logger

__all__ = [
    "DatabaseMaintenanceDialog",
    "KpiCard",
    "format_file_size",
    "get_tracked_databases",
    "query_db_metrics",
    "vacuum_database_file",
]
