# --- File: src/core/ai/cli_detector.py ---
"""
External LLM CLI Agent Discovery & MCP Readiness Detector.

Detects installed AI coding/standards CLIs (`codex`, `claude`, `gemini`) with
zero-lag caching (<2 ms on the fast path), inspects MCP server configuration
files without spawning slow subprocesses, and manages optional overrides in
`src/config/cli_agents_config.json`.
"""

from __future__ import annotations

import json
import logging
import os
import shutil
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Tuple

from core.utils.paths import get_project_root

logger = logging.getLogger(__name__)

CONFIG_PATH = get_project_root() / "config" / "cli_agents_config.json"
_CACHE_TTL_SECONDS = 30.0

DEFAULT_CONFIG: Dict[str, Any] = {
    "preferred_agent": "",
    "custom_paths": {
        "codex": "",
        "claude": "",
        "gemini": "",
    },
}


@dataclass(frozen=True)
class CliAgentSpec:
    """Static descriptor for a supported external LLM CLI agent."""

    agent_id: str
    label: str
    vendor: str
    icon: str
    binary_names: Tuple[str, ...]
    install_hint: str
    args_builder: Callable[[str, str], List[str]]


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


def _build_codex_args(executable: str, prompt: str) -> List[str]:
    return [executable, prompt] if prompt else [executable]


def _build_claude_args(executable: str, prompt: str) -> List[str]:
    return [executable, prompt] if prompt else [executable]


def _build_gemini_args(executable: str, prompt: str) -> List[str]:
    return [executable, "-i", prompt] if prompt else [executable]


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

_STATUS_CACHE: Optional[Dict[str, CliAgentStatus]] = None
_STATUS_CACHE_TS: float = 0.0


def invalidate_cli_detection_cache() -> None:
    """Clear the cached CLI agent detection snapshot."""
    global _STATUS_CACHE, _STATUS_CACHE_TS
    _STATUS_CACHE = None
    _STATUS_CACHE_TS = 0.0


def load_cli_agents_config() -> Dict[str, Any]:
    """Load `cli_agents_config.json` from disk with safe defaults."""
    cfg: Dict[str, Any] = {
        "preferred_agent": DEFAULT_CONFIG["preferred_agent"],
        "custom_paths": dict(DEFAULT_CONFIG["custom_paths"]),
    }
    if CONFIG_PATH.exists():
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                loaded = json.load(f)
            if isinstance(loaded, dict):
                if isinstance(loaded.get("preferred_agent"), str):
                    cfg["preferred_agent"] = loaded["preferred_agent"].strip().lower()
                if isinstance(loaded.get("custom_paths"), dict):
                    for k, v in loaded["custom_paths"].items():
                        if isinstance(k, str) and isinstance(v, str):
                            cfg["custom_paths"][k.strip().lower()] = v.strip()
        except Exception as exc:
            logger.debug("Could not load cli_agents_config.json: %s", exc)
    return cfg


def save_cli_agents_config(cfg: Dict[str, Any]) -> None:
    """Persist `cli_agents_config.json` and invalidate detection cache."""
    try:
        CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(cfg, f, indent=2)
        invalidate_cli_detection_cache()
    except Exception as exc:
        logger.warning("Could not save cli_agents_config.json: %s", exc)


def get_preferred_cli_agent() -> Optional[str]:
    """Return the user's preferred CLI agent ID if set and recognized."""
    pref = load_cli_agents_config().get("preferred_agent", "").strip().lower()
    return pref if pref in CLI_AGENT_SPECS else None


def set_preferred_cli_agent(agent_id: str) -> None:
    """Persist the preferred CLI agent ID in `cli_agents_config.json`."""
    cfg = load_cli_agents_config()
    clean_id = (agent_id or "").strip().lower()
    cfg["preferred_agent"] = clean_id if clean_id in CLI_AGENT_SPECS else ""
    save_cli_agents_config(cfg)


def _prefer_windows_executable(candidate: Path) -> Path:
    """
    On Windows, npm/pnpm create both an extensionless POSIX shell script (`codex`)
    and a `.cmd`/`.exe` wrapper (`codex.cmd`). Prefer `.exe`/`.cmd`/`.bat` so
    PowerShell / CreateProcess launches cleanly without file-association prompts.
    """
    if os.name != "nt":
        return candidate
    if candidate.suffix.lower() in {".exe", ".cmd", ".bat", ".ps1"}:
        return candidate
    for ext in (".exe", ".cmd", ".bat"):
        sibling = candidate.with_name(candidate.name + ext)
        if sibling.is_file():
            return sibling
    return candidate


def _candidate_user_dirs() -> List[Path]:
    """Return common zero-admin / user-space binary directories on Windows & POSIX."""
    dirs: List[Path] = []
    home = Path.home()

    appdata = os.environ.get("APPDATA", "").strip()
    localappdata = os.environ.get("LOCALAPPDATA", "").strip()

    if appdata:
        dirs.append(Path(appdata) / "npm")
    dirs.append(home / ".local" / "bin")
    dirs.append(home / ".bun" / "bin")
    dirs.append(home / "scoop" / "shims")
    if localappdata:
        dirs.append(Path(localappdata) / "Programs" / "claude-code")
        dirs.append(Path(localappdata) / "Microsoft" / "WinGet" / "Links")
        dirs.append(Path(localappdata) / "pnpm")

    # Deduplicate preserving order
    seen: set = set()
    unique: List[Path] = []
    for d in dirs:
        norm = os.path.normcase(str(d))
        if norm not in seen:
            seen.add(norm)
            unique.append(d)
    return unique


def _resolve_executable(
    spec: CliAgentSpec,
    custom_path: str = "",
    which_fn: Optional[Callable[[str], Optional[str]]] = None,
    user_dirs: Optional[List[Path]] = None,
) -> tuple[Optional[str], str]:
    """
    Resolve an executable path for `spec`.
    Returns `(executable_path, source)`.
    """
    if custom_path:
        p = Path(custom_path).expanduser()
        if p.is_file():
            return str(_prefer_windows_executable(p)), "config"

    which_lookup = which_fn or shutil.which
    for name in spec.binary_names:
        found = which_lookup(name)
        if found:
            resolved = _prefer_windows_executable(Path(found))
            if resolved.is_file():
                return str(resolved), "path"
            return str(found), "path"

    for directory in (user_dirs if user_dirs is not None else _candidate_user_dirs()):
        if not directory.is_dir():
            continue
        for name in spec.binary_names:
            cand = directory / name
            if cand.is_file():
                return str(_prefer_windows_executable(cand)), "user_dir"

    return None, "not_found"


def check_mcp_configured(agent_id: str, home_dir: Optional[Path] = None) -> bool:
    """
    Check local CLI configuration files to determine whether `3gpp-tools` MCP
    is already registered. Pure file inspection (no subprocess calls).
    """
    home = home_dir or Path.home()
    cid = (agent_id or "").strip().lower()

    if cid == "codex":
        candidates = [
            home / ".codex" / "config.toml",
            home / ".codex" / "config.json",
        ]
    elif cid == "claude":
        candidates = [
            home / ".claude.json",
            home / ".claude" / "settings.json",
            home / ".claude" / "settings.local.json",
        ]
    elif cid == "gemini":
        candidates = [
            home / ".gemini" / "settings.json",
        ]
    else:
        return False

    for cfg_path in candidates:
        if not cfg_path.is_file():
            continue
        try:
            text = cfg_path.read_text(encoding="utf-8", errors="ignore")
            lower_text = text.lower()
            if "3gpp-tools" in lower_text or "mcp_server.py" in lower_text:
                return True
        except Exception:
            continue
    return False


def build_interactive_cli_args(agent_id: str, executable: str, prompt: str) -> List[str]:
    """Build interactive-mode CLI arguments for the given agent ID."""
    spec = CLI_AGENT_SPECS.get((agent_id or "").strip().lower())
    if spec is None:
        raise ValueError(f"Unsupported CLI agent: {agent_id}")
    return spec.args_builder(executable, prompt)


def detect_cli_agents(
    force_refresh: bool = False,
    *,
    which_fn: Optional[Callable[[str], Optional[str]]] = None,
    user_dirs: Optional[List[Path]] = None,
    home_dir: Optional[Path] = None,
) -> Dict[str, CliAgentStatus]:
    """
    Detect installed external LLM CLIs (`codex`, `claude`, `gemini`).

    Results are cached for 30 seconds on standard calls so UI context menus
    and toolbars render instantaneously without disk/PATH overhead.
    """
    global _STATUS_CACHE, _STATUS_CACHE_TS

    use_cache = (
        not force_refresh
        and which_fn is None
        and user_dirs is None
        and home_dir is None
        and _STATUS_CACHE is not None
        and (time.monotonic() - _STATUS_CACHE_TS) < _CACHE_TTL_SECONDS
    )
    if use_cache and _STATUS_CACHE is not None:
        return _STATUS_CACHE

    cfg = load_cli_agents_config()
    custom_paths = cfg.get("custom_paths", {}) if isinstance(cfg.get("custom_paths"), dict) else {}

    statuses: Dict[str, CliAgentStatus] = {}
    for agent_id, spec in CLI_AGENT_SPECS.items():
        custom = str(custom_paths.get(agent_id, "") or "").strip()
        exe_path, source = _resolve_executable(
            spec,
            custom_path=custom,
            which_fn=which_fn,
            user_dirs=user_dirs,
        )
        statuses[agent_id] = CliAgentStatus(
            agent_id=agent_id,
            label=spec.label,
            vendor=spec.vendor,
            icon=spec.icon,
            installed=bool(exe_path),
            executable_path=exe_path,
            source=source,
            mcp_configured=check_mcp_configured(agent_id, home_dir=home_dir),
            install_hint=spec.install_hint,
        )

    if which_fn is None and user_dirs is None and home_dir is None:
        _STATUS_CACHE = statuses
        _STATUS_CACHE_TS = time.monotonic()

    return statuses


def get_installed_cli_agents(force_refresh: bool = False) -> List[CliAgentStatus]:
    """Return a list of currently installed CLI agents."""
    return [st for st in detect_cli_agents(force_refresh=force_refresh).values() if st.installed]
