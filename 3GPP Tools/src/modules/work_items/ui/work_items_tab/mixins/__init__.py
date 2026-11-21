"""Feature mixins composing WorkItemsTab."""

from modules.work_items.ui.work_items_tab.mixins.data_and_sync import (
    WorkItemsDataSyncMixin,
)
from modules.work_items.ui.work_items_tab.mixins.dialogs_and_menus import (
    WorkItemsDialogsMenusMixin,
)
from modules.work_items.ui.work_items_tab.mixins.filters_and_tracking import (
    WorkItemsFiltersTrackingMixin,
)
from modules.work_items.ui.work_items_tab.mixins.ui_setup import (
    WorkItemsUiSetupMixin,
)

__all__ = [
    "WorkItemsDataSyncMixin",
    "WorkItemsDialogsMenusMixin",
    "WorkItemsFiltersTrackingMixin",
    "WorkItemsUiSetupMixin",
]
