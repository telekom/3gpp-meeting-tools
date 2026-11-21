"""Topic Dossier & Briefing Studio dialog package."""

from modules.meetings.ui.dialogs.topic_dossier.dialog import TopicDossierDialog
from modules.meetings.ui.dialogs.topic_dossier.editor_dialog import TopicEditorDialog
from modules.meetings.ui.dialogs.topic_dossier.styles import (
    ALL_COMMON_WGS,
    STYLE_CARD,
    STYLE_PRIMARY_BTN,
    STYLE_SECONDARY_BTN,
)

__all__ = [
    "TopicDossierDialog",
    "TopicEditorDialog",
    "ALL_COMMON_WGS",
    "STYLE_CARD",
    "STYLE_PRIMARY_BTN",
    "STYLE_SECONDARY_BTN",
]
