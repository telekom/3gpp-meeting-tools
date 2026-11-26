"""Default example arguments for the MCP Monitor Tool Playground."""

from typing import Any, Dict

# Example arguments for interactive tool playground
DEFAULT_PLAYGROUND_ARGUMENTS: Dict[str, Dict[str, Dict[str, Any]]] = {
    "spec_text_and_clauses": {
        "search_text": {"spec_number": "23.501", "query": "URSP", "limit": 3},
        "get_clause": {"spec_number": "23.501", "clause_id": "5.15.1"},
        "diff_versions": {"spec_number": "23.501", "clause_id": "5.15.1", "old_version": "19.0.0", "new_version": "20.2.0"},
        "evolution_matrix": {"spec_number": "23.501", "clause_id": "5.15.1"},
        "extract_clause": {"spec_number": "23.501", "clause_id": "5.15.1"},
    },
    "protocol_inspector": {
        "list_protocols": {},
        "find_message": {"protocol": "5GS_NAS", "query": "REGISTRATION REQUEST"},
        "find_ie": {"protocol": "5GS_NAS", "query": "5GS tracking area identity"},
        "evolution_matrix": {"protocol": "5GS_NAS", "query": "REGISTRATION REQUEST", "target_type": "message"},
    },
    "meetings_and_tdocs": {
        "search_meetings": {"wg": "SA2", "limit": 3},
        "search_tdocs": {"wg": "SA2", "meeting_number": "174", "limit": 3},
        "read_tdoc_markdown": {"tdoc_id": "S2-2601720"},
        "meeting_stats": {"wg": "SA2", "meeting_number": "174"},
        "get_or_set_stance": {"tdoc_id": "S2-2601720"},
        "list_topics": {"limit": 10},
    },
    "onsite_meeting_cockpit": {
        "network_status": {"wg": "SA2"},
        "where_should_we_go_now": {"wg": "SA2"},
        "browse_lan_server": {"subfolder": "Docs"},
    },
    "spec_history_and_attribution": {
        "query_history": {"spec_number": "23.501", "limit": 3},
        "attribute_clause": {"spec_number": "23.501", "clause_id": "5.15.1", "limit": 3},
        "rank_origin_crs": {"spec_number": "23.501", "query": "URSP", "limit": 3},
        "tdoc_dossier": {"spec_number": "23.501", "tdoc_id": "S2-2601720"},
    },
    "work_items_and_emails": {
        "search_work_items": {"query": "6G", "limit": 3},
        "search_emails": {"query": "URSP", "limit": 5},
        "search_tracked_threads": {"limit": 5},
    },
    "cr_pcr_generator": {
        "preview_filename": {
            "mode": "pcr",
            "spec_number": "23.501",
            "clauses": "5.15.1",
            "tdoc_id": "S2-2601720",
        },
    },
}
