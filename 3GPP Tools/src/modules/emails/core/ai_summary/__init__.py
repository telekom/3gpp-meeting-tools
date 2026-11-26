"""Persistent incremental AI email summary store, Map-Reduce orchestrator, and workers."""

from .formatting import (
    extract_clean_direct_body,
    extract_deterministic_facts,
    format_messages_for_llm,
    get_message_timestamp_str,
    parse_summary_output,
    sort_messages_chronologically,
)
from .m365_handoff import (
    M365_COPILOT_APP_URI,
    M365_COPILOT_WEB_URL,
    M365HandoffPacket,
    build_ai_handoff_packet,
    build_leaf_handoff_packet,
    launch_m365_copilot,
    save_external_summary_from_handoff,
    stage_handoff_markdown,
)
from .models import EmailSummaryRecord, SummaryFreshness, SummaryScope, VALID_SCOPE_TYPES
from .orchestrator import EmailSummaryOrchestrator, load_summary_prompt
from .store import EmailSummaryStore, utc_now_iso
from .workers import AiBriefingWorker, EmailSummaryTaskWorker

__all__ = [
    "AiBriefingWorker",
    "EmailSummaryOrchestrator",
    "EmailSummaryRecord",
    "EmailSummaryStore",
    "EmailSummaryTaskWorker",
    "M365HandoffPacket",
    "M365_COPILOT_APP_URI",
    "M365_COPILOT_WEB_URL",
    "SummaryFreshness",
    "SummaryScope",
    "VALID_SCOPE_TYPES",
    "build_ai_handoff_packet",
    "build_leaf_handoff_packet",
    "extract_clean_direct_body",
    "extract_deterministic_facts",
    "format_messages_for_llm",
    "get_message_timestamp_str",
    "launch_m365_copilot",
    "load_summary_prompt",
    "parse_summary_output",
    "save_external_summary_from_handoff",
    "sort_messages_chronologically",
    "stage_handoff_markdown",
    "utc_now_iso",
]
