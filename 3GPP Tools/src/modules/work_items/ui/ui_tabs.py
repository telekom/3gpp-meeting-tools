"""Backward-compatible re-export shim for Work Items tab and table widgets."""

from modules.work_items.ui.work_items_delegates import (
    ReleaseBadgeDelegate,
    RemarksDelegate,
    SpecsDelegate,
    StatusBadgeDelegate,
    TimelineDelegate,
    TrackingDelegate,
    WgBadgeDelegate,
    WidDelegate,
)
from modules.work_items.ui.work_items_model import WorkItemsTableModel
from modules.work_items.ui.work_items_tab import WorkItemsTab, logger as _tab_logger

logger = _tab_logger

__all__ = [
    "TrackingDelegate",
    "WgBadgeDelegate",
    "WidDelegate",
    "ReleaseBadgeDelegate",
    "StatusBadgeDelegate",
    "SpecsDelegate",
    "TimelineDelegate",
    "RemarksDelegate",
    "WorkItemsTableModel",
    "WorkItemsTab",
]
