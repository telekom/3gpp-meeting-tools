"""Composite dataset loading & table filtering mixin for `MeetingEmailInboxDialog`."""

from __future__ import annotations

from .data_loading import InboxDataLoadingMixin
from .table_filtering import InboxTableFilteringMixin


class InboxDataAndFilteringMixin(
    InboxDataLoadingMixin,
    InboxTableFilteringMixin,
):
    """Composite mixin combining dataset loading/grouping and table filtering."""


__all__ = [
    "InboxDataAndFilteringMixin",
    "InboxDataLoadingMixin",
    "InboxTableFilteringMixin",
]
