"""Feature mixins composing TDocsWindow (see ../README.md)."""
from modules.meetings.ui.tdocs.window.mixins.ui_header import HeaderMixin
from modules.meetings.ui.tdocs.window.mixins.filter_setup import FilterSetupMixin
from modules.meetings.ui.tdocs.window.mixins.table_layout import TableLayoutMixin
from modules.meetings.ui.tdocs.window.mixins.filter_tracked_wis import TrackedWIFilterMixin
from modules.meetings.ui.tdocs.window.mixins.filter_facets import FilterFacetsMixin
from modules.meetings.ui.tdocs.window.mixins.tdoc_actions import TDocActionsMixin
from modules.meetings.ui.tdocs.window.mixins.email_compose import EmailMixin
from modules.meetings.ui.tdocs.window.mixins.tracked_threads import TrackedThreadsMixin
from modules.meetings.ui.tdocs.window.mixins.exports import ExportsMixin
from modules.meetings.ui.tdocs.window.mixins.meeting_data import MeetingDataMixin
from modules.meetings.ui.tdocs.window.mixins.drag_drop import DragDropMixin

__all__ = ["HeaderMixin", "FilterSetupMixin", "TableLayoutMixin", "TrackedWIFilterMixin", "FilterFacetsMixin", "TDocActionsMixin", "EmailMixin", "TrackedThreadsMixin", "ExportsMixin", "MeetingDataMixin", "DragDropMixin"]
