"""Modular Cross-WG Topic Dossier Exporter package."""

from __future__ import annotations

from .checkup import (
    _rebuild_meeting_state_for_checkup,
    _resolve_topic_for_checkup,
    check_and_refresh_dossier,
)
from .exporter_thread import TopicDossierExporterThread
from .path_utils import (
    _copy_markdown_with_media,
    _decision_checkpoint_key,
    _normalize_tdoc_id,
    _resolve_doc_server_url,
    get_default_topic_export_base,
    safe_markdown_rel_link,
    safe_relpath,
    sanitize_folder_name,
)
from .stage2_discovery import (
    _evaluate_stage2_synthesis_state,
    _extract_report_header_meta,
    _parse_stage2_subfolder,
    _read_json_dict,
    discover_stage2_packages,
)

__all__ = [
    "TopicDossierExporterThread",
    "_copy_markdown_with_media",
    "_decision_checkpoint_key",
    "_evaluate_stage2_synthesis_state",
    "_extract_report_header_meta",
    "_normalize_tdoc_id",
    "_parse_stage2_subfolder",
    "_read_json_dict",
    "_rebuild_meeting_state_for_checkup",
    "_resolve_doc_server_url",
    "_resolve_topic_for_checkup",
    "check_and_refresh_dossier",
    "discover_stage2_packages",
    "get_default_topic_export_base",
    "safe_markdown_rel_link",
    "safe_relpath",
    "sanitize_folder_name",
]
