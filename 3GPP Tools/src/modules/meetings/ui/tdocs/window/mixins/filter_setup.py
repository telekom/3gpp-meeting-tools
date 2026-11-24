# --- File: src/modules/meetings/ui/tdocs/window/mixins/filter_setup.py ---
"""Composite filter setup mixin combining persistence and reactive filter controls."""

from modules.meetings.ui.tdocs.window.mixins.filter_controls import FilterControlsMixin
from modules.meetings.ui.tdocs.window.mixins.filter_persistence import (
    _TDOC_COMBO_KEYS,
    FilterPersistenceMixin,
)


class FilterSetupMixin(FilterPersistenceMixin, FilterControlsMixin):
    """Filter bar construction, persistence, and simple filter-change/toggle handlers."""


__all__ = ["FilterSetupMixin", "_TDOC_COMBO_KEYS"]
