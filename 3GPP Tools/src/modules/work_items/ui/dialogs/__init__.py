"""Modular Work Items UI dialogs and Word-compatible Quick Copy helpers."""

import logging

from .add_dialog import AddWorkItemDialog, ManualWIFetcherThread
from .copy_menus import (
    build_wi_copy_data,
    copy_wis_default,
    populate_wi_batch_copy_menu,
    populate_wi_copy_menu,
)
from .info_dialog import WorkItemInfoDialog
from .linked_specs_dialog import LinkedSpecsDialog

logger = logging.getLogger(__name__)

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
