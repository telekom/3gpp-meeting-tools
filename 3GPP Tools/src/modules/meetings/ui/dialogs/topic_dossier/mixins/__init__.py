"""Feature mixins composing TopicDossierDialog."""

from modules.meetings.ui.dialogs.topic_dossier.mixins.export_and_cli import (
    TopicDossierExportCliMixin,
)
from modules.meetings.ui.dialogs.topic_dossier.mixins.topics_and_status import (
    TopicDossierTopicsStatusMixin,
)
from modules.meetings.ui.dialogs.topic_dossier.mixins.ui_setup import (
    TopicDossierUiSetupMixin,
)

__all__ = [
    "TopicDossierExportCliMixin",
    "TopicDossierTopicsStatusMixin",
    "TopicDossierUiSetupMixin",
]
