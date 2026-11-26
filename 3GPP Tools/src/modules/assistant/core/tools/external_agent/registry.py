"""Compact 12-tool external agent ToolRegistry builder and profile definitions."""

from __future__ import annotations

from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple

from modules.assistant.core.agent_models import ToolResult, ToolStatus
from modules.assistant.core.knowledge import (
    ProtocolKnowledgeService,
    SpecificationKnowledgeService,
    SpecSearchKnowledgeService,
)
from modules.assistant.core.tool_registry import (
    ToolDefinition,
    ToolPermission,
    ToolRegistry,
)
from .common import ToolArgumentSpec, _resolve_db_paths, _wrap_handler
from .diagram_and_wi_tools import (
    _handle_plantuml_office_studio,
    _handle_work_items_and_emails,
)
from .history_and_protocol_tools import (
    _handle_protocol_inspector,
    _handle_spec_history_and_attribution,
)
from .meetings_tools import (
    _handle_chair_notes_and_session_plan,
    _handle_meetings_and_tdocs,
)
from .onsite_tools import _handle_onsite_meeting_cockpit
from .spec_tools import (
    _handle_spec_catalogue_and_files,
    _handle_spec_text_and_clauses,
)
from .word_studio_tools import (
    _handle_cr_pcr_generator,
    _handle_word_markdown_studio,
    _handle_word_sanitizer_and_compliance,
)


# =========================================================================
# REGISTRY BUILDER (Compact 12-Tool Schema + Profile Filtering)
# =========================================================================

PROFILE_TOOLS: Dict[str, Tuple[str, ...]] = {
    "all": (
        "spec_catalogue_and_files",
        "spec_text_and_clauses",
        "spec_history_and_attribution",
        "protocol_inspector",
        "onsite_meeting_cockpit",
        "meetings_and_tdocs",
        "chair_notes_and_session_plan",
        "word_markdown_studio",
        "cr_pcr_generator",
        "word_sanitizer_and_compliance",
        "plantuml_office_studio",
        "work_items_and_emails",
    ),
    "knowledge": (
        "spec_catalogue_and_files",
        "spec_text_and_clauses",
        "spec_history_and_attribution",
        "protocol_inspector",
        "work_items_and_emails",
    ),
    "onsite": (
        "onsite_meeting_cockpit",
        "meetings_and_tdocs",
        "chair_notes_and_session_plan",
        "spec_text_and_clauses",
        "cr_pcr_generator",
    ),
    "authoring": (
        "word_markdown_studio",
        "cr_pcr_generator",
        "word_sanitizer_and_compliance",
        "plantuml_office_studio",
        "spec_text_and_clauses",
    ),
}


def build_external_agent_tool_registry(
    *,
    profile: str = "all",
    read_only: bool = False,
    specs_db_path: Optional[Path] = None,
    spec_search_db_path: Optional[Path] = None,
    protocol_db_path: Optional[Path] = None,
    spec_history_db_path: Optional[Path] = None,
    user_tdocs_db_path: Optional[Path] = None,
    session_plan_db_path: Optional[Path] = None,
) -> ToolRegistry:
    """Builds the token-efficient 12-tool external agent registry for Claude CLI / Codex CLI."""
    paths = _resolve_db_paths(
        specs_db_path=specs_db_path,
        spec_search_db_path=spec_search_db_path,
        protocol_db_path=protocol_db_path,
        spec_history_db_path=spec_history_db_path,
        user_tdocs_db_path=user_tdocs_db_path,
        session_plan_db_path=session_plan_db_path,
    )
    spec_service = SpecificationKnowledgeService(paths["specs_db"])
    search_service = SpecSearchKnowledgeService(paths["spec_search_db"])
    protocol_service = ProtocolKnowledgeService(paths["protocol_db"])

    all_defs: List[ToolDefinition] = [
        ToolDefinition(
            name="spec_catalogue_and_files",
            description="Search 3GPP specs/versions, download/resolve FTP archives, index for FTS5, or convert specs to Markdown.",
            permission=ToolPermission.KNOWLEDGE_READ,
            arguments=(
                ToolArgumentSpec("action", "string", "", required=True, enum=("search", "get", "resolve_files", "download", "index_for_search", "convert_to_md")),
                ToolArgumentSpec("spec_number", "string", "e.g. '23.501'"),
                ToolArgumentSpec("query", "string", ""),
                ToolArgumentSpec("version", "string", "e.g. '19.2.0'"),
                ToolArgumentSpec("release", "string", "e.g. 'Rel-19'"),
                ToolArgumentSpec("wg", "string", ""),
                ToolArgumentSpec("mode", "string", "'single'|'split'"),
                ToolArgumentSpec("limit", "integer", ""),
            ),
            handler=_wrap_handler(lambda a: _handle_spec_catalogue_and_files(a, spec_service=spec_service, spec_search_db_path=paths["spec_search_db"])),
        ),
        ToolDefinition(
            name="spec_text_and_clauses",
            description="FTS5 search across 3GPP specs, get clause text, trace evolution across releases, diff versions, or extract clauses to .docx/.md.",
            permission=ToolPermission.KNOWLEDGE_READ,
            arguments=(
                ToolArgumentSpec("action", "string", "", required=True, enum=("search_text", "get_clause", "evolution_matrix", "diff_versions", "extract_clause")),
                ToolArgumentSpec("query", "string", ""),
                ToolArgumentSpec("spec_number", "string", "e.g. '23.501'"),
                ToolArgumentSpec("clause_id", "string", "e.g. '5.5.1'"),
                ToolArgumentSpec("result_id", "string", "Handle from search_text"),
                ToolArgumentSpec("version", "string", ""),
                ToolArgumentSpec("old_version", "string", ""),
                ToolArgumentSpec("new_version", "string", ""),
                ToolArgumentSpec("release", "string", ""),
                ToolArgumentSpec("latest_only", "boolean", ""),
                ToolArgumentSpec("limit", "integer", ""),
            ),
            handler=_wrap_handler(lambda a: _handle_spec_text_and_clauses(a, search_service=search_service, spec_search_db_path=paths["spec_search_db"])),
        ),
        ToolDefinition(
            name="spec_history_and_attribution",
            description="Query Change History tables, attribute clauses to origin CR/pCR/TDoc, rank origin CRs, or generate TDoc dossiers.",
            permission=ToolPermission.KNOWLEDGE_READ,
            arguments=(
                ToolArgumentSpec("action", "string", "", required=True, enum=("query_history", "ingest_history", "attribute_clause", "rank_origin_crs", "tdoc_dossier")),
                ToolArgumentSpec("spec_number", "string", "e.g. '23.501'"),
                ToolArgumentSpec("clause_id", "string", "e.g. '5.34.2'"),
                ToolArgumentSpec("query", "string", ""),
                ToolArgumentSpec("version", "string", ""),
                ToolArgumentSpec("meeting", "string", ""),
                ToolArgumentSpec("category", "string", ""),
                ToolArgumentSpec("tdoc_id", "string", ""),
                ToolArgumentSpec("deep_verify_docx", "boolean", "Verify CR .docx cover"),
                ToolArgumentSpec("output_dir", "string", ""),
                ToolArgumentSpec("limit", "integer", ""),
            ),
            handler=_wrap_handler(lambda a: _handle_spec_history_and_attribution(
                a,
                specs_db_path=paths["specs_db"],
                spec_search_db_path=paths["spec_search_db"],
                spec_history_db_path=paths["spec_history_db"],
            )),
        ),
        ToolDefinition(
            name="protocol_inspector",
            description="Inspect 3GPP protocols (5GS/EPS NAS, UE Policies, NR/LTE RRC, NGAP, GTPv1-U, UP, PFCP): messages, IEs, and evolution matrices.",
            permission=ToolPermission.KNOWLEDGE_READ,
            arguments=(
                ToolArgumentSpec("action", "string", "", required=True, enum=("list_protocols", "list_messages", "find_message", "find_ie", "evolution_matrix")),
                ToolArgumentSpec("protocol", "string", "e.g. '5GS_NAS', 'RRC_NR', 'NGAP'"),
                ToolArgumentSpec("query", "string", "Message or IE name"),
                ToolArgumentSpec("version", "string", ""),
                ToolArgumentSpec("target_type", "string", "'message'|'ie'"),
                ToolArgumentSpec("format", "string", "'markdown'|'tsv'"),
                ToolArgumentSpec("limit", "integer", ""),
            ),
            handler=_wrap_handler(lambda a: _handle_protocol_inspector(a, protocol_service=protocol_service, protocol_db_path=paths["protocol_db"])),
        ),
        ToolDefinition(
            name="onsite_meeting_cockpit",
            description="On-site 10.10.10.10 LAN (3GPPWIFI) & live meeting cockpit: probe LAN, browse 10.10.10.10/ftp/, fetch/diff live r01/r02 revisions, sync Chair Notes, check venue schedule & free delegates, or batch-cache TDocs.",
            permission=ToolPermission.NETWORK_FETCH,
            arguments=(
                ToolArgumentSpec("action", "string", "", required=True, enum=("network_status", "browse_lan_server", "fetch_live_revisions", "sync_chair_notes", "where_should_we_go_now", "batch_cache_tdocs")),
                ToolArgumentSpec("wg", "string", "Default 'SA2'"),
                ToolArgumentSpec("meeting_number", "string", "e.g. '174'"),
                ToolArgumentSpec("tdoc_id", "string", "TDoc ID(s)"),
                ToolArgumentSpec("subfolder", "string", "e.g. 'Inbox/Revisions', 'Docs'"),
                ToolArgumentSpec("query", "string", ""),
                ToolArgumentSpec("day_name", "string", ""),
                ToolArgumentSpec("time_hhmm", "string", ""),
                ToolArgumentSpec("force_probe", "boolean", ""),
                ToolArgumentSpec("limit", "integer", ""),
            ),
            handler=_wrap_handler(lambda a: _handle_onsite_meeting_cockpit(a, specs_db_path=paths["specs_db"], session_plan_db_path=paths["session_plan_db"])),
        ),
        ToolDefinition(
            name="meetings_and_tdocs",
            description="Search meetings & TDocs (disambiguating confirmed vs '?'/'[...]' tentative companies), read TDocs as Markdown, manage stances/notes, or get meeting stats.",
            permission=ToolPermission.KNOWLEDGE_READ,
            arguments=(
                ToolArgumentSpec("action", "string", "", required=True, enum=("search_meetings", "search_tdocs", "read_tdoc_markdown", "get_or_set_stance", "meeting_stats")),
                ToolArgumentSpec("wg", "string", ""),
                ToolArgumentSpec("meeting_number", "string", ""),
                ToolArgumentSpec("year", "string", ""),
                ToolArgumentSpec("location", "string", ""),
                ToolArgumentSpec("tdoc_id", "string", ""),
                ToolArgumentSpec("query", "string", ""),
                ToolArgumentSpec("source_company", "string", ""),
                ToolArgumentSpec("confirmed_company_only", "boolean", "Exclude '?'/'[...]'"),
                ToolArgumentSpec("agenda_item", "string", ""),
                ToolArgumentSpec("work_item", "string", ""),
                ToolArgumentSpec("status", "string", ""),
                ToolArgumentSpec("revisions", "string", ""),
                ToolArgumentSpec("max_chars", "integer", ""),
                ToolArgumentSpec("stance", "string", ""),
                ToolArgumentSpec("notes", "string", ""),
                ToolArgumentSpec("limit", "integer", ""),
            ),
            handler=_wrap_handler(lambda a: _handle_meetings_and_tdocs(
                a,
                specs_db_path=paths["specs_db"],
                user_tdocs_db_path=paths["user_tdocs_db"],
                spec_history_db_path=paths["spec_history_db"],
            )),
        ),
        ToolDefinition(
            name="chair_notes_and_session_plan",
            description="Parse Chairman's Notes .docx for TDoc outcomes, or ingest, query, and diff SA2 Session Plan snapshots (TU 1/TU 2 sub-slots & rooms).",
            permission=ToolPermission.KNOWLEDGE_READ,
            arguments=(
                ToolArgumentSpec("action", "string", "", required=True, enum=("parse_chair_notes_file", "ingest_session_plan", "query_session_plan", "diff_session_plans")),
                ToolArgumentSpec("meeting_number", "string", ""),
                ToolArgumentSpec("file_path", "string", ""),
                ToolArgumentSpec("query", "string", ""),
                ToolArgumentSpec("day_name", "string", ""),
                ToolArgumentSpec("limit", "integer", ""),
            ),
            handler=_wrap_handler(lambda a: _handle_chair_notes_and_session_plan(a, session_plan_db_path=paths["session_plan_db"])),
        ),
        ToolDefinition(
            name="word_markdown_studio",
            description="Bidirectional Word (.docx/.doc) ⇄ Markdown (.md) conversion (with Track Changes, tables, PlantUML & figures), clause slicing, or Visio (.vsdx) extraction.",
            permission=ToolPermission.DOCUMENT_WRITE,
            arguments=(
                ToolArgumentSpec("action", "string", "", required=True, enum=("word_to_md", "md_to_word", "slice_clauses", "extract_visio")),
                ToolArgumentSpec("input_path", "string", "", required=True),
                ToolArgumentSpec("output_path", "string", ""),
                ToolArgumentSpec("revisions", "string", "'markup'|'accept'|'reject'"),
                ToolArgumentSpec("clauses", "string", "Comma-separated"),
            ),
            handler=_wrap_handler(_handle_word_markdown_studio),
        ),
        ToolDefinition(
            name="cr_pcr_generator",
            description="Generate official MCC-styled 3GPP CR or pCR .docx contributions with cover sheet, blue change separators, and differentiated naming.",
            permission=ToolPermission.DOCUMENT_WRITE,
            arguments=(
                ToolArgumentSpec("action", "string", "", required=True, enum=("generate_pcr", "generate_cr", "preview_filename")),
                ToolArgumentSpec("spec_number", "string", "e.g. '23.501'", required=True),
                ToolArgumentSpec("version", "string", ""),
                ToolArgumentSpec("clauses", "string", "Clauses to excerpt"),
                ToolArgumentSpec("tdoc_id", "string", ""),
                ToolArgumentSpec("title", "string", ""),
                ToolArgumentSpec("source_company", "string", ""),
                ToolArgumentSpec("work_item", "string", ""),
                ToolArgumentSpec("reason_for_change", "string", ""),
                ToolArgumentSpec("summary_of_change", "string", ""),
                ToolArgumentSpec("consequences_if_not_approved", "string", ""),
                ToolArgumentSpec("source_docx_path", "string", ""),
                ToolArgumentSpec("output_path", "string", ""),
            ),
            handler=_wrap_handler(_handle_cr_pcr_generator),
        ),
        ToolDefinition(
            name="word_sanitizer_and_compliance",
            description="Audit/sanitize Word (.docx) hygiene (F-07: language unification, purge language revisions, author consolidation), Word COM compare, or MIP sensitivity labels.",
            permission=ToolPermission.DOCUMENT_WRITE,
            arguments=(
                ToolArgumentSpec("action", "string", "", required=True, enum=("audit_hygiene", "sanitize_document", "compare_documents", "inspect_label", "apply_label")),
                ToolArgumentSpec("input_path", "string", "", required=True),
                ToolArgumentSpec("revised_path", "string", ""),
                ToolArgumentSpec("output_path", "string", ""),
                ToolArgumentSpec("target_language", "string", "Default 'en-GB'"),
                ToolArgumentSpec("target_author", "string", ""),
                ToolArgumentSpec("label_preset", "string", "Sensitivity label name, display name, or GUID from user policy (defaults to active configured label)"),
            ),
            handler=_wrap_handler(_handle_word_sanitizer_and_compliance),
        ),
        ToolDefinition(
            name="plantuml_office_studio",
            description="Validate/render PlantUML to SVG, export to editable Visio (.vsdx) or PowerPoint (.pptx/append), bridge .pptx⇄.vsdx, list 3GPP NF presets, or extract PlantUML from 3GPP-Tools-generated .vsdx.",
            permission=ToolPermission.DOCUMENT_WRITE,
            arguments=(
                ToolArgumentSpec("action", "string", "", required=True, enum=("validate_or_render_svg", "export_vsdx", "export_pptx", "bridge_pptx_vsdx", "extract_puml_from_vsdx", "list_presets")),
                ToolArgumentSpec("puml_code", "string", ""),
                ToolArgumentSpec("input_path", "string", ""),
                ToolArgumentSpec("output_path", "string", ""),
                ToolArgumentSpec("append_to_open", "boolean", ""),
                ToolArgumentSpec("query", "string", ""),
            ),
            handler=_wrap_handler(_handle_plantuml_office_studio),
        ),
        ToolDefinition(
            name="work_items_and_emails",
            description="Search 3GPP Work Items, fetch WI details from Portal, toggle tracking, search WG emails, or read/save incremental email summaries.",
            permission=ToolPermission.KNOWLEDGE_READ,
            arguments=(
                ToolArgumentSpec("action", "string", "", required=True, enum=("search_work_items", "fetch_wi_from_portal", "toggle_wi_tracking", "search_emails", "list_email_summaries", "get_email_summary", "save_email_summary")),
                ToolArgumentSpec("query", "string", ""),
                ToolArgumentSpec("wi_code", "string", ""),
                ToolArgumentSpec("wg", "string", ""),
                ToolArgumentSpec("release", "string", ""),
                ToolArgumentSpec("meeting_number", "string", ""),
                ToolArgumentSpec("scope_type", "string", "'thread'|'tdoc'|'ai'|'selection'"),
                ToolArgumentSpec("scope_key", "string", ""),
                ToolArgumentSpec("summary_md", "string", ""),
                ToolArgumentSpec("covered_ids", "string", "Comma-separated message IDs"),
                ToolArgumentSpec("cutoff_utc", "string", ""),
                ToolArgumentSpec("include_delta", "boolean", ""),
                ToolArgumentSpec("tracked_only", "boolean", ""),
                ToolArgumentSpec("starred", "boolean", ""),
                ToolArgumentSpec("following", "boolean", ""),
                ToolArgumentSpec("limit", "integer", ""),
            ),
            handler=_wrap_handler(lambda a: _handle_work_items_and_emails(a, specs_db_path=paths["specs_db"])),
        ),
    ]

    allowed_names: Optional[Set[str]] = None
    prof_key = (profile or "all").strip().lower()
    if prof_key in PROFILE_TOOLS:
        allowed_names = set(PROFILE_TOOLS[prof_key])

    allowed_perms = (
        {ToolPermission.KNOWLEDGE_READ}
        if read_only
        else {ToolPermission.KNOWLEDGE_READ, ToolPermission.DOCUMENT_WRITE, ToolPermission.NETWORK_FETCH}
    )
    registry = ToolRegistry(allowed_permissions=allowed_perms)
    for tool in all_defs:
        if allowed_names is not None and tool.name not in allowed_names:
            continue
        if read_only and tool.permission != ToolPermission.KNOWLEDGE_READ:
            continue
        if read_only:
            orig_handler = tool.handler
            mutating_actions = {
                "download",
                "index_for_search",
                "convert_to_md",
                "extract_clause",
                "ingest_history",
                "tdoc_dossier",
                "ingest_session_plan",
                "toggle_wi_tracking",
                "save_email_summary",
            }
            def _ro_guard(orig=orig_handler, tname=tool.name, **kw: Any) -> ToolResult:
                act = str(kw.get("action") or "").strip().lower()
                if act in mutating_actions or (
                    act == "get_or_set_stance"
                    and (kw.get("stance") is not None or kw.get("notes") is not None)
                ):
                    return ToolResult(
                        status=ToolStatus.FOUND,
                        data={"ok": False, "error": f"Read-only mode enabled: action '{act}' on '{tname}' is disabled."},
                        message="Read-only mode enabled",
                    )
                return orig(**kw)
            tool = ToolDefinition(
                name=tool.name,
                description=tool.description,
                arguments=tool.arguments,
                handler=_ro_guard,
                permission=tool.permission,
            )
        registry.register(tool)
    return registry
