"""Backward-compatible re-export shim for TDocs export & CLI session controllers.

Canonical implementation lives in ``modules.meetings.ui.dialogs.export_controller``.
"""

from __future__ import annotations

import logging

from core.ai.cli_launcher import launch_cli_in_terminal
from modules.meetings.core.exporters.llm_exporter import TDocMarkdownExporterThread
from modules.meetings.ui.dialogs.export_controller import (
    _expand_revision_rows,
    _show_export_complete_with_cli_prompt,
    export_markdown_single,
    export_markdown_visible,
    export_tdocs_package,
    get_package_export_rows,
    launch_cli_for_tdocs,
    on_dossier_export_finished,
    on_excel_export_finished,
    on_package_export_finished,
    on_package_files_finished,
    on_post_meeting_export_finished,
    on_pre_meeting_export_finished,
    open_dossier_export_dialog,
    open_excel_export_dialog,
    open_post_meeting_export_dialog,
    open_pre_meeting_export_dialog,
    prepare_and_launch_cli_session,
)

logger = logging.getLogger(__name__)

__all__ = [
    "TDocMarkdownExporterThread",
    "_expand_revision_rows",
    "_show_export_complete_with_cli_prompt",
    "export_markdown_single",
    "export_markdown_visible",
    "export_tdocs_package",
    "get_package_export_rows",
    "launch_cli_for_tdocs",
    "launch_cli_in_terminal",
    "logger",
    "on_dossier_export_finished",
    "on_excel_export_finished",
    "on_package_export_finished",
    "on_package_files_finished",
    "on_post_meeting_export_finished",
    "on_pre_meeting_export_finished",
    "open_dossier_export_dialog",
    "open_excel_export_dialog",
    "open_post_meeting_export_dialog",
    "open_pre_meeting_export_dialog",
    "prepare_and_launch_cli_session",
]
