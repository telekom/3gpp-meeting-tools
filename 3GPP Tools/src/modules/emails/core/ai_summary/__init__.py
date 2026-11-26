"""Persistent incremental AI email summary store, Map-Reduce orchestrator, and workers."""

from .formatting import (
    extract_clean_direct_body,
    extract_deterministic_facts,
    format_messages_for_llm,
    get_message_timestamp_str,
    parse_summary_output,
    sort_messages_chronologically,
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
    "SummaryFreshness",
    "SummaryScope",
    "VALID_SCOPE_TYPES",
    "extract_clean_direct_body",
    "extract_deterministic_facts",
    "format_messages_for_llm",
    "get_message_timestamp_str",
    "load_summary_prompt",
    "parse_summary_output",
    "sort_messages_chronologically",
    "utc_now_iso",
]
