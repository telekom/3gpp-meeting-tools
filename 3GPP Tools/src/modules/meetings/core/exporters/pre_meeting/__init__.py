# --- File: src/modules/meetings/core/exporters/pre_meeting/__init__.py ---
"""Modular package for the Pre-Meeting Report & Markdown Dossier Exporter."""

from .exporter_thread import PreMeetingExporterThread, logger

__all__ = [
    "PreMeetingExporterThread",
    "logger",
]
