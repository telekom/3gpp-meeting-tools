# External Agent Domain Tools (`external_agent/`)

Modular decomposition of the 12 multi-action MCP and CLI domain tools previously housed in `external_agent_tools.py`.

## Submodule Navigation Map

| Module | Ownership / Domains | Key Symbols |
|---|---|---|
| [`common.py`](common.py) | Shared argument builders, handler wrapper, Win32 COM serialization, DB paths, CSV parsing, DocResolver/Heading/Markdown helpers & shim forwarding | `ToolArgumentSpec`, `_wrap_handler`, `_COM_LOCK`, `_disable_faulthandler_for_com`, `_resolve_db_paths`, `_parse_csv_list`, `_extract_resolved_word_docs`, `_match_heading_for_clause`, `_export_doc_to_markdown`, `_get_meetings_cache_roots`, `resolve_meetings_cache_roots` |
| [`spec_tools.py`](spec_tools.py) | Domain 1 (`spec_catalogue_and_files`), Domain 2 (`spec_text_and_clauses`) | `_handle_spec_catalogue_and_files`, `_handle_spec_text_and_clauses` |
| [`history_and_protocol_tools.py`](history_and_protocol_tools.py) | Domain 3 (`spec_history_and_attribution`), Domain 4 (`protocol_inspector`) | `_handle_spec_history_and_attribution`, `_handle_protocol_inspector` |
| [`onsite_tools.py`](onsite_tools.py) | Domain 5 (`onsite_meeting_cockpit`: `10.10.10.10` LAN directory & live sync) | `_parse_10_10_10_10_directory_html`, `_handle_onsite_meeting_cockpit` |
| [`meeting_helpers.py`](meeting_helpers.py) | Shared meeting lookup, local cache resolution, TDoc normalization, Excel/SQLite fallback, user stances/notes, company filtering | `_search_meetings_compatible`, `_resolve_meeting_record`, `_resolve_meeting_cache_dir`, `_normalize_tdoc_record`, `_load_tdocs_for_meeting_or_db`, `_collect_user_tdoc_annotations`, `_matches_company_filter` |
| [`meetings_tools.py`](meetings_tools.py) | Domain 6 (`meetings_and_tdocs`), Domain 7 (`chair_notes_and_session_plan`) | `_handle_meetings_and_tdocs`, `_handle_chair_notes_and_session_plan` |
| [`word_studio_tools.py`](word_studio_tools.py) | Domain 8 (`word_markdown_studio`), Domain 9 (`cr_pcr_generator`), Domain 10 (`word_sanitizer_and_compliance`) | `_handle_word_markdown_studio`, `_handle_cr_pcr_generator`, `_handle_word_sanitizer_and_compliance` |
| [`diagram_and_wi_tools.py`](diagram_and_wi_tools.py) | Domain 11 (`plantuml_office_studio`), Domain 12 (`work_items_and_emails`) | `_handle_plantuml_office_studio`, `_normalize_release_filters`, `_handle_work_items_and_emails` |
| [`email_summary_tools.py`](email_summary_tools.py) | Persistent AI email summary & delta message helpers for Domain 12 | `handle_email_summary_mcp_action` |
| [`registry.py`](registry.py) | 12-tool compact MCP schema registration with MCP `annotations`, profile filtering (`all`, `knowledge`, `onsite`, `authoring`), `--read-only` action guard | `PROFILE_TOOLS`, `build_external_agent_tool_registry` |
| [`__init__.py`](__init__.py) | Package exports | `PROFILE_TOOLS`, `ToolArgumentSpec`, `_resolve_db_paths`, `build_external_agent_tool_registry` |

## Backward Compatibility & Shim Monkeypatching

- [`../external_agent_tools.py`](../external_agent_tools.py) remains as a thin re-export shim exposing all public and private module-level symbols.
- Submodules invoke `resolve_meetings_cache_roots()` (aliased as `_get_meetings_cache_roots`), which inspects `sys.modules` for `modules.assistant.core.tools.external_agent_tools._get_meetings_cache_roots` before calling the default implementation. Existing test patches targeting `external_agent_tools._get_meetings_cache_roots` therefore apply automatically across all domain handlers.
