# --- File: src/modules/work_items/ui/work_items_dialogs.py ---
"""Backward-compatible re-export shim for :mod:`modules.work_items.ui.dialogs`."""

from __future__ import annotations

from modules.work_items.ui.dialogs import (
    AddWorkItemDialog,
    LinkedSpecsDialog,
    ManualWIFetcherThread,
    WorkItemInfoDialog,
    build_wi_copy_data,
    copy_wis_default,
    logger as _dialogs_logger,
    populate_wi_batch_copy_menu,
    populate_wi_copy_menu,
)

logger = _dialogs_logger

__all__ = [
    "AddWorkItemDialog",
    "LinkedSpecsDialog",
    "ManualWIFetcherThread",
    "WorkItemInfoDialog",
    "build_wi_copy_data",
    "copy_wis_default",
    "populate_wi_batch_copy_menu",
    "populate_wi_copy_menu",
]
