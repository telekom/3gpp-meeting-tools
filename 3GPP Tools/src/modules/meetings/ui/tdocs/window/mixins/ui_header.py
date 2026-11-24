# --- File: src/modules/meetings/ui/tdocs/window/mixins/ui_header.py ---
"""Composite header mixin combining toolbar setup and menus/clock/instant fetch helpers."""

from modules.meetings.ui.tdocs.window.mixins.header_menus_and_clock import HeaderMenusAndClockMixin
from modules.meetings.ui.tdocs.window.mixins.header_widgets import HeaderWidgetsMixin


class HeaderMixin(HeaderWidgetsMixin, HeaderMenusAndClockMixin):
    """Header/toolbar construction, resources menu, instant fetch/compare, routing indicator, venue clock."""
