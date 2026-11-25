"""Backward-compatible re-export shim for TopicDossierDialog and TopicEditorDialog."""

from core.ai.cli_launcher import launch_cli_in_terminal
from core.utils.folder_opener import open_folder as _open_folder
from modules.meetings.core.exporters.topic_dossier_exporter import TopicDossierExporterThread
from modules.meetings.ui.dialogs.topic_dossier import (
    ALL_COMMON_WGS,
    STYLE_CARD,
    STYLE_PRIMARY_BTN,
    STYLE_SECONDARY_BTN,
    CorpusStatusReportDialog,
    TopicDossierDialog,
    TopicEditorDialog,
)

__all__ = [
    "TopicDossierDialog",
    "TopicEditorDialog",
    "CorpusStatusReportDialog",
    "ALL_COMMON_WGS",
    "STYLE_CARD",
    "STYLE_PRIMARY_BTN",
    "STYLE_SECONDARY_BTN",
    "launch_cli_in_terminal",
    "_open_folder",
    "TopicDossierExporterThread",
]
