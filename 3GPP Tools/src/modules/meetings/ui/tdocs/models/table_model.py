"""TDocsTableModel combining precomputation, state, agreements, roles, and agenda merge mixins."""

import logging
from pathlib import Path

from PyQt5.QtCore import QAbstractTableModel, QModelIndex, Qt

from modules.meetings.ui.tdocs.models.mixins import (
    TDocsAgendaMergeMixin,
    TDocsAgreementsHtmlMixin,
    TDocsDataRolesMixin,
    TDocsIndexingMixin,
    TDocsStateThreadsMixin,
)

logger = logging.getLogger("modules.meetings.ui.tdocs_models")


class TDocsTableModel(
    TDocsIndexingMixin,
    TDocsStateThreadsMixin,
    TDocsAgreementsHtmlMixin,
    TDocsDataRolesMixin,
    TDocsAgendaMergeMixin,
    QAbstractTableModel,
):
    """High-performance Qt table model for 3GPP TDocs, tracked email threads, and agreements."""

    def __init__(self, meeting_dir: Path = None, data=None, user_data=None, agenda_map=None):
        super().__init__()
        self.meeting_dir = meeting_dir
        self.user_data = user_data or {}
        self._data = data or []
        self.agenda_map = agenda_map or {}

        # Appending "Emails" at the end prevents breaking delegates for columns 0, 7, 12!
        self._headers = [
            "", "TDoc", "Title", "Source", "Type", "For",
            "Abstract", "Secretary Remarks", "My Status", "My Notes", "Agenda Item", "TDoc Status", "Related TDocs", "Emails"
        ]
        self.valid_tdocs = {str(r.get("TDoc", "")) for r in self._data if r.get("TDoc")}
        self.loading_tdocs = set()
        self._highlighted_tdoc = None
        self.revisions = {}
        self.email_counts = {}  # {tdoc_id: {'total': int, 'unread': int}}
        self._email_counts = self.email_counts
        self._tdoc_map = {}
        self._family_cache = {}

        self._apply_company_sanitization(self._data)
        self._precompute_row_wi_tokens(self._data)
        self._precompute_agenda_sort_tuples(self._data)
        self._apply_user_data_logic()
        self._rebuild_family_cache()
        self._recompute_cached_email_stats()
        self.chair_agreements = []
        self._detect_and_tag_agreements(self._data)

    def rowCount(self, index=QModelIndex()):
        return len(self._data)

    def columnCount(self, index=QModelIndex()):
        return len(self._headers)

    def headerData(self, section, orientation, role):
        if orientation == Qt.Horizontal:
            col_name = self._headers[section]
            if role == Qt.DisplayRole:
                if col_name == "Abstract":
                    return "📝"
                if col_name == "My Notes":
                    return "📓"
                return col_name
            elif role == Qt.ToolTipRole:
                if col_name == "Abstract":
                    return "Abstract"
                if col_name == "My Notes":
                    return "My Notes"
                return col_name
        return None
