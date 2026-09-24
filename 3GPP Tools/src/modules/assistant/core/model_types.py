"""Provider-independent model response types used by the Assistant."""
from dataclasses import dataclass, field
from typing import Any, Dict, Tuple

@dataclass(frozen=True)
class ModelToolCall:
    name: str
    arguments: Dict[str, Any]
    call_id: str = ""

@dataclass(frozen=True)
class ModelResponse:
    content: str
    tool_calls: Tuple[ModelToolCall, ...] = ()
    raw_message: Dict[str, Any] = field(default_factory=dict)
    metrics: Dict[str, Any] = field(default_factory=dict)
    model: str = ""
    @property
    def requests_tools(self) -> bool:
        return bool(self.tool_calls)
