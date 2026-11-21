"""Backward-compatible re-export shim for Topic Dossier exporter.

The implementation is organized under ``modules.meetings.core.exporters.topic_dossier``.
All public and internal symbols continue to be re-exported here so existing imports
and monkeypatches targeting ``topic_dossier_exporter`` remain unaffected.
"""

from __future__ import annotations

from modules.meetings.core.tdocs.tdoc_markdown_converter import convert_document_to_markdown

from .topic_dossier import (
    TopicDossierExporterThread,
    _copy_markdown_with_media,
    _decision_checkpoint_key,
    _evaluate_stage2_synthesis_state,
    _extract_report_header_meta,
    _normalize_tdoc_id,
    _parse_stage2_subfolder,
    _read_json_dict,
    _rebuild_meeting_state_for_checkup,
    _resolve_doc_server_url,
    _resolve_topic_for_checkup,
    check_and_refresh_dossier,
    discover_stage2_packages,
    get_default_topic_export_base,
    safe_markdown_rel_link,
    safe_relpath,
    sanitize_folder_name,
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
    "convert_document_to_markdown",
    "discover_stage2_packages",
    "get_default_topic_export_base",
    "safe_markdown_rel_link",
    "safe_relpath",
    "sanitize_folder_name",
]
