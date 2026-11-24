"""Data models, curated model presets, and CLI specs for external LLM CLIs."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Callable, Dict, List, Optional, Tuple

DEFAULT_CONFIG: Dict[str, Any] = {
    "preferred_agent": "",
    "custom_paths": {
        "codex": "",
        "claude": "",
        "gemini": "",
    },
    "agent_models": {
        "codex": {"model": "", "reasoning_effort": ""},
        "claude": {"model": "", "reasoning_effort": ""},
        "gemini": {"model": "", "reasoning_effort": ""},
    },
}

_FALLBACK_CODEX_MODELS: Tuple[Tuple[str, str, Tuple[str, ...], str], ...] = (
    ("gpt-6.1-sol", "GPT-6.1-Sol", ("low", "medium", "high", "xhigh", "max", "ultra"), "medium"),
    ("gpt-6-astra", "GPT-6-Astra", ("low", "medium", "high", "xhigh", "max", "ultra"), "low"),
    ("gpt-6.1-terra", "GPT-6.1 Terra", ("low", "medium", "high", "xhigh", "max", "ultra"), "low"),
    ("gpt-6-sol", "GPT-6-Sol", ("low", "medium", "high", "xhigh", "max", "ultra"), "medium"),
    ("gpt-5.6-sol", "GPT-5.6-Sol", ("low", "medium", "high", "xhigh", "max", "ultra"), "medium"),
)

_CURATED_CLAUDE_MODELS: Tuple[Tuple[str, str, Tuple[str, ...], str], ...] = (
    ("sonnet", "Claude Sonnet (alias)", ("low", "medium", "high"), "medium"),
    ("opus", "Claude Opus (alias)", ("low", "medium", "high"), "medium"),
    ("haiku", "Claude Haiku (alias)", ("low", "medium"), "low"),
    ("claude-sonnet-4-5", "Claude Sonnet 4.5", ("low", "medium", "high"), "medium"),
    ("claude-opus-4-5", "Claude Opus 4.5", ("low", "medium", "high"), "medium"),
)

_CURATED_GEMINI_MODELS: Tuple[Tuple[str, str, Tuple[str, ...], str], ...] = (
    ("gemini-2.5-pro", "Gemini 2.5 Pro", (), ""),
    ("gemini-2.5-flash", "Gemini 2.5 Flash", (), ""),
    ("gemini-2.0-flash", "Gemini 2.0 Flash", (), ""),
)


@dataclass(frozen=True)
class CliModelOption:
    """Selectable model and supported reasoning effort levels for an LLM CLI."""

    slug: str
    display_name: str
    reasoning_levels: Tuple[str, ...] = ()
    default_reasoning_level: str = ""
    description: str = ""


@dataclass(frozen=True)
class CliAgentSpec:
    """Static descriptor for a supported external LLM CLI agent."""

    agent_id: str
    label: str
    vendor: str
    icon: str
    binary_names: Tuple[str, ...]
    install_hint: str
    args_builder: Callable[..., List[str]]


@dataclass
class CliAgentStatus:
    """Resolved installation and MCP registration status for a CLI agent."""

    agent_id: str
    label: str
    vendor: str
    icon: str
    installed: bool
    executable_path: Optional[str] = None
    source: str = "not_found"  # "config" | "path" | "user_dir" | "not_found"
    mcp_configured: bool = False
    install_hint: str = ""


def _build_codex_args(
    executable: str,
    prompt: str,
    model: str = "",
    reasoning_effort: str = "",
) -> List[str]:
    """Build interactive command arguments for OpenAI Codex CLI."""
    args = [executable]
    if model:
        args.extend(["-m", model])
    if reasoning_effort:
        args.extend(["-c", f"model_reasoning_effort={reasoning_effort}"])
    if prompt:
        args.append(prompt)
    return args


def _build_claude_args(
    executable: str,
    prompt: str,
    model: str = "",
    reasoning_effort: str = "",
) -> List[str]:
    """Build interactive command arguments for Anthropic Claude CLI."""
    args = [executable]
    if model:
        args.extend(["--model", model])
    if reasoning_effort:
        args.extend(["--effort", reasoning_effort])
    if prompt:
        args.append(prompt)
    return args


def _build_gemini_args(
    executable: str,
    prompt: str,
    model: str = "",
    reasoning_effort: str = "",  # Gemini CLI encodes thinking tier in model slug
) -> List[str]:
    """Build interactive command arguments for Google Gemini CLI."""
    args = [executable]
    if model:
        args.extend(["-m", model])
    if prompt:
        args.extend(["-i", prompt])
    return args


CLI_AGENT_SPECS: Dict[str, CliAgentSpec] = {
    "codex": CliAgentSpec(
        agent_id="codex",
        label="Codex CLI",
        vendor="OpenAI",
        icon="🤖",
        binary_names=("codex.cmd", "codex.exe", "codex"),
        install_hint="npm install -g @openai/codex",
        args_builder=_build_codex_args,
    ),
    "claude": CliAgentSpec(
        agent_id="claude",
        label="Claude CLI",
        vendor="Anthropic",
        icon="🧠",
        binary_names=("claude.cmd", "claude.exe", "claude"),
        install_hint="npm install -g @anthropic-ai/claude-code",
        args_builder=_build_claude_args,
    ),
    "gemini": CliAgentSpec(
        agent_id="gemini",
        label="Gemini CLI",
        vendor="Google",
        icon="✨",
        binary_names=("gemini.cmd", "gemini.exe", "gemini"),
        install_hint="npm install -g @google/gemini-cli",
        args_builder=_build_gemini_args,
    ),
}
