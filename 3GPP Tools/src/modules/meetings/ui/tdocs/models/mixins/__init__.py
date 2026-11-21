"""Feature mixins composing TDocsTableModel."""

from modules.meetings.ui.tdocs.models.mixins.agenda_merge import TDocsAgendaMergeMixin
from modules.meetings.ui.tdocs.models.mixins.agreements_and_html import (
    TDocsAgreementsHtmlMixin,
)
from modules.meetings.ui.tdocs.models.mixins.data_roles import TDocsDataRolesMixin
from modules.meetings.ui.tdocs.models.mixins.indexing import TDocsIndexingMixin
from modules.meetings.ui.tdocs.models.mixins.state_and_threads import (
    TDocsStateThreadsMixin,
)

__all__ = [
    "TDocsAgendaMergeMixin",
    "TDocsAgreementsHtmlMixin",
    "TDocsDataRolesMixin",
    "TDocsIndexingMixin",
    "TDocsStateThreadsMixin",
]
