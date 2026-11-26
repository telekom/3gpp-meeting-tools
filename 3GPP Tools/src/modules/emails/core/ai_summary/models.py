"""Data models for persistent multi-scope AI email summaries."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Dict, List

VALID_SCOPE_TYPES = ("thread", "tdoc", "ai", "selection")


@dataclass(frozen=True)
class SummaryScope:
    """Identifies the target slice of emails being summarized."""

    scope_type: str  # "thread" | "tdoc" | "ai" | "selection"
    scope_key: str
    title: str = ""
    focus_prompt: str = ""

    def normalized(self) -> "SummaryScope":
        stype = (self.scope_type or "thread").strip().lower()
        if stype not in VALID_SCOPE_TYPES:
            stype = "thread"
        skey = str(self.scope_key or "").strip()
        if stype == "tdoc":
            skey = skey.upper()
        return SummaryScope(
            scope_type=stype,
            scope_key=skey,
            title=str(self.title or skey).strip(),
            focus_prompt=str(self.focus_prompt or "").strip(),
        )


@dataclass
class EmailSummaryRecord:
    """Persisted AI summary + watermark provenance for a thread, TDoc, selection, or AI."""

    scope_type: str
    scope_key: str
    title: str
    summary_md: str
    delta_md: str = ""
    structured_meta: Dict[str, Any] = field(default_factory=dict)
    deterministic_facts: Dict[str, Any] = field(default_factory=dict)
    cutoff_utc: str = ""
    covered_message_ids: List[str] = field(default_factory=list)
    message_count: int = 0
    incremental_hops: int = 0
    provider: str = ""
    model: str = ""
    reasoning_effort: str = ""
    focus_prompt: str = ""
    created_at: str = ""
    updated_at: str = ""


@dataclass
class SummaryFreshness:
    """Comparison of a stored summary against the current messages in scope."""

    has_summary: bool
    total_messages: int
    covered_count: int
    new_count: int
    new_message_ids: List[str] = field(default_factory=list)
    cutoff_utc: str = ""
    incremental_hops: int = 0
    recommend_full_rebuild: bool = False
    provider: str = ""
    model: str = ""
    reasoning_effort: str = ""
    updated_at: str = ""

    @property
    def is_stale(self) -> bool:
        return self.has_summary and self.new_count > 0

    @property
    def is_current(self) -> bool:
        return self.has_summary and self.new_count == 0
