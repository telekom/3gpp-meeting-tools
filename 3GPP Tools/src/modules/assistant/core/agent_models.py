"""Application-owned domain models for the 3GPP Assistant.

These models deliberately contain no Qt, Ollama, or SQLite dependencies.
"""

from __future__ import annotations

import threading
import uuid
from dataclasses import dataclass, field
from enum import Enum
from typing import Any, Dict, List, Optional


class ToolStatus(str, Enum):
    """Machine-readable outcome of a knowledge/tool operation."""

    FOUND = "found"
    NOT_FOUND = "not_found"
    SOURCE_UNAVAILABLE = "source_unavailable"
    SOURCE_INCOMPLETE = "source_incomplete"
    AMBIGUOUS = "ambiguous"
    INVALID_REQUEST = "invalid_request"
    ERROR = "error"


class EvidenceSourceType(str, Enum):
    """Origin/type of evidence made available to an agent research run."""

    SPECIFICATION_CATALOGUE = "specification_catalogue"
    SPECIFICATION_TEXT = "specification_text"
    PROTOCOL_KNOWLEDGE = "protocol_knowledge"


class ResearchOutcome(str, Enum):
    """Top-level lifecycle state of a research run."""

    COMPLETED = "completed"
    CANCELLED = "cancelled"
    FAILED = "failed"


@dataclass(frozen=True)
class Evidence:
    """Application-owned evidence with explicit provenance."""

    id: str
    source_type: EvidenceSourceType
    source_name: str
    specification: Optional[str] = None
    version: Optional[str] = None
    release_date: Optional[str] = None
    clause: Optional[str] = None
    title: Optional[str] = None
    content: Any = None
    content_complete: bool = True
    metadata: Dict[str, Any] = field(default_factory=dict)


@dataclass
class ToolResult:
    """Normalized result returned by a knowledge service/tool."""

    status: ToolStatus
    data: Any = None
    message: str = ""
    evidence_ids: List[str] = field(default_factory=list)
    metadata: Dict[str, Any] = field(default_factory=dict)

    @property
    def ok(self) -> bool:
        return self.status == ToolStatus.FOUND


@dataclass(frozen=True)
class ResearchActivity:
    """Objective action suitable for UI display and diagnostic tracing."""

    kind: str
    message: str
    metadata: Dict[str, Any] = field(default_factory=dict)


@dataclass
class DiagnosticTrace:
    """In-memory structured trace for one research run."""

    run_id: str
    question: str
    model: str = ""
    events: List[Dict[str, Any]] = field(default_factory=list)

    def add(self, event_type: str, **metadata: Any) -> None:
        self.events.append({"type": event_type, **metadata})


class CancellationToken:
    """Thread-safe cooperative cancellation state."""

    def __init__(self) -> None:
        self._event = threading.Event()

    def cancel(self) -> None:
        self._event.set()

    @property
    def is_cancelled(self) -> bool:
        return self._event.is_set()

    def raise_if_cancelled(self) -> None:
        if self.is_cancelled:
            raise ResearchCancelledError("Research run was cancelled.")


class ResearchCancelledError(RuntimeError):
    """Raised internally when a cooperative research cancellation is observed."""


@dataclass
class ResearchRun:
    """Application-owned state for a single research question."""

    question: str
    conversation_context: List[Dict[str, str]] = field(default_factory=list)
    id: str = field(default_factory=lambda: uuid.uuid4().hex)
    evidence: Dict[str, Evidence] = field(default_factory=dict)
    trace: Optional[DiagnosticTrace] = None
    cancellation_token: CancellationToken = field(default_factory=CancellationToken)
    _next_evidence_number: int = 1

    def __post_init__(self) -> None:
        if self.trace is None:
            self.trace = DiagnosticTrace(run_id=self.id, question=self.question)

    def add_evidence(
        self,
        source_type: EvidenceSourceType,
        source_name: str,
        *,
        specification: Optional[str] = None,
        version: Optional[str] = None,
        release_date: Optional[str] = None,
        clause: Optional[str] = None,
        title: Optional[str] = None,
        content: Any = None,
        content_complete: bool = True,
        metadata: Optional[Dict[str, Any]] = None,
    ) -> Evidence:
        evidence_id = f"E{self._next_evidence_number}"
        self._next_evidence_number += 1
        item = Evidence(
            id=evidence_id,
            source_type=source_type,
            source_name=source_name,
            specification=specification,
            version=version,
            release_date=release_date,
            clause=clause,
            title=title,
            content=content,
            content_complete=content_complete,
            metadata=dict(metadata or {}),
        )
        self.evidence[evidence_id] = item
        return item


@dataclass
class AgentResult:
    """Final result produced by a research run.

    Evidence and trace are returned to the presentation layer so citations and
    diagnostics remain application-owned rather than being reconstructed from
    model text.
    """

    outcome: ResearchOutcome
    answer: str = ""
    cited_evidence_ids: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)
    error: str = ""
    run_id: str = ""
    evidence: List[Evidence] = field(default_factory=list)
    trace: Optional[DiagnosticTrace] = None
