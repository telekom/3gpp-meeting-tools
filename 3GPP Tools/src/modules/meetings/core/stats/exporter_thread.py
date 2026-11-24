"""Backward-compatible re-export shim for meeting statistics exporter.

Canonical implementation lives in ``modules.meetings.core.stats.exporter``.
"""

from modules.meetings.core.stats.exporter import StatisticsExporterThread
from modules.meetings.core.stats.html_dashboard_compiler import compile_view_block

__all__ = ["StatisticsExporterThread", "compile_view_block"]
