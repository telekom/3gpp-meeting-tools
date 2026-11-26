"""Data models for unified Ollama + external CLI LLM execution."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Dict, Tuple

SUPPORTED_PROVIDERS: Tuple[str, ...] = ("ollama", "codex", "claude", "gemini")

PROVIDER_META: Dict[str, Tuple[str, str]] = {
    "ollama": ("Ollama (Local)", "🦙"),
    "codex": ("Codex CLI", "🤖"),
    "claude": ("Claude CLI", "🧠"),
    "gemini": ("Gemini CLI", "✨"),
}


@dataclass(frozen=True)
class ModelSpec:
    """Unified provider + model + optional reasoning effort identifier."""

    provider: str
    model: str = ""
    reasoning_effort: str = ""

    def normalized(self) -> "ModelSpec":
        prov = (self.provider or "ollama").strip().lower()
        if prov not in PROVIDER_META:
            prov = "ollama"
        return ModelSpec(
            provider=prov,
            model=(self.model or "").strip(),
            reasoning_effort=(self.reasoning_effort or "").strip().lower(),
        )

    def display_label(self) -> str:
        norm = self.normalized()
        label, icon = PROVIDER_META.get(norm.provider, (norm.provider, "🤖"))
        model_txt = norm.model or "CLI Default"
        if norm.reasoning_effort:
            return f"{icon} {norm.provider} · {model_txt} ({norm.reasoning_effort})"
        return f"{icon} {norm.provider} · {model_txt}"

    def short_badge(self) -> str:
        norm = self.normalized()
        model_txt = norm.model or "default"
        if norm.reasoning_effort:
            return f"{norm.provider} · {model_txt} · {norm.reasoning_effort}"
        return f"{norm.provider} · {model_txt}"

    def to_dict(self) -> Dict[str, str]:
        norm = self.normalized()
        return {
            "provider": norm.provider,
            "model": norm.model,
            "reasoning_effort": norm.reasoning_effort,
        }

    @classmethod
    def from_dict(cls, data: Any, *, default_provider: str = "ollama") -> "ModelSpec":
        if not isinstance(data, dict):
            return cls(provider=default_provider)
        return cls(
            provider=str(data.get("provider") or default_provider).strip().lower(),
            model=str(data.get("model") or "").strip(),
            reasoning_effort=str(data.get("reasoning_effort") or "").strip().lower(),
        ).normalized()


@dataclass(frozen=True)
class ProviderModelEntry:
    """Selectable model option exposed in unified model menus & defaults dialog."""

    provider: str
    provider_label: str
    provider_icon: str
    model_slug: str
    display_name: str
    reasoning_levels: Tuple[str, ...] = ()
    default_effort: str = ""
    available: bool = True


@dataclass
class RunnerResult:
    """Result returned by `run_llm_text()` regardless of execution backend."""

    ok: bool
    text: str = ""
    error: str = ""
    provider: str = ""
    model: str = ""
    reasoning_effort: str = ""
    elapsed_s: float = 0.0
    usage: Dict[str, Any] = field(default_factory=dict)
    cancelled: bool = False
