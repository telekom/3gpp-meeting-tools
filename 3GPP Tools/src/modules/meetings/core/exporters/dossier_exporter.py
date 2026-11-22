# --- File: src/modules/meetings/core/exporters/dossier_exporter.py ---
"""Backward-compatible re-export shim for :mod:`modules.meetings.core.exporters.dossier`."""

from __future__ import annotations

from modules.meetings.core.exporters.dossier import (
    DOSSIER_CSS,
    DossierExporterThread,
    logger as _dossier_logger,
)

logger = _dossier_logger

__all__ = [
    "DOSSIER_CSS",
    "DossierExporterThread",
]
