"""Backward-compatible re-export shim for external agent tools.

The implementation is organized under ``modules.assistant.core.tools.external_agent``.
All public and internal symbols continue to be re-exported here so existing imports
and monkeypatches targeting ``external_agent_tools`` remain unaffected.
"""

from __future__ import annotations

from .external_agent.common import (
    ToolArgumentSpec,
    _COM_LOCK,
    _disable_faulthandler_for_com,
    _get_meetings_cache_roots,
    _parse_csv_list,
    _resolve_db_paths,
    _wrap_handler,
)
from .external_agent.diagram_and_wi_tools import (
    _handle_plantuml_office_studio,
    _handle_work_items_and_emails,
    _normalize_release_filters,
)
from .external_agent.history_and_protocol_tools import (
    _handle_protocol_inspector,
    _handle_spec_history_and_attribution,
)
from .external_agent.meeting_helpers import (
    _collect_user_tdoc_annotations,
    _load_tdocs_for_meeting_or_db,
    _matches_company_filter,
    _normalize_tdoc_record,
    _resolve_meeting_cache_dir,
    _resolve_meeting_record,
    _search_meetings_compatible,
)
from .external_agent.meetings_tools import (
    _handle_chair_notes_and_session_plan,
    _handle_meetings_and_tdocs,
)
from .external_agent.onsite_tools import (
    _handle_onsite_meeting_cockpit,
    _parse_10_10_10_10_directory_html,
)
from .external_agent.registry import (
    PROFILE_TOOLS,
    build_external_agent_tool_registry,
)
from .external_agent.spec_tools import (
    _handle_spec_catalogue_and_files,
    _handle_spec_text_and_clauses,
)
from .external_agent.word_studio_tools import (
    _handle_cr_pcr_generator,
    _handle_word_markdown_studio,
    _handle_word_sanitizer_and_compliance,
)

__all__ = [
    "PROFILE_TOOLS",
    "ToolArgumentSpec",
    "_COM_LOCK",
    "_collect_user_tdoc_annotations",
    "_disable_faulthandler_for_com",
    "_get_meetings_cache_roots",
    "_handle_chair_notes_and_session_plan",
    "_handle_cr_pcr_generator",
    "_handle_meetings_and_tdocs",
    "_handle_onsite_meeting_cockpit",
    "_handle_plantuml_office_studio",
    "_handle_protocol_inspector",
    "_handle_spec_catalogue_and_files",
    "_handle_spec_history_and_attribution",
    "_handle_spec_text_and_clauses",
    "_handle_word_markdown_studio",
    "_handle_word_sanitizer_and_compliance",
    "_handle_work_items_and_emails",
    "_load_tdocs_for_meeting_or_db",
    "_matches_company_filter",
    "_normalize_release_filters",
    "_normalize_tdoc_record",
    "_parse_10_10_10_10_directory_html",
    "_parse_csv_list",
    "_resolve_db_paths",
    "_resolve_meeting_cache_dir",
    "_resolve_meeting_record",
    "_search_meetings_compatible",
    "_wrap_handler",
    "build_external_agent_tool_registry",
]
