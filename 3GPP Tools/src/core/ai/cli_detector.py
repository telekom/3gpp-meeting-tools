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
import re
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


def _build_codex_args(
    executable: str,
    prompt: str,
    model: str = "",
    reasoning_effort: str = "",
) -> List[str]:
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
        "agent_models": {aid: dict(vals) for aid, vals in DEFAULT_CONFIG["agent_models"].items()},
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
                if isinstance(loaded.get("agent_models"), dict):
                    for k, v in loaded["agent_models"].items():
                        if isinstance(k, str) and isinstance(v, dict):
                            aid = k.strip().lower()
                            cfg["agent_models"][aid] = {
                                "model": str(v.get("model", "") or "").strip(),
                                "reasoning_effort": str(v.get("reasoning_effort", "") or "").strip().lower(),
                            }
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


def get_cli_agent_model_config(agent_id: str) -> Tuple[str, str]:
    """Return configured ``(model, reasoning_effort)`` override for ``agent_id``."""
    clean_id = (agent_id or "").strip().lower()
    models_map = load_cli_agents_config().get("agent_models", {})
    entry = models_map.get(clean_id, {}) if isinstance(models_map, dict) else {}
    return (
        str(entry.get("model", "") or "").strip(),
        str(entry.get("reasoning_effort", "") or "").strip().lower(),
    )


def set_cli_agent_model_config(agent_id: str, model: str = "", reasoning_effort: str = "") -> None:
    """Persist per-CLI ``model`` and ``reasoning_effort`` override in `cli_agents_config.json`."""
    clean_id = (agent_id or "").strip().lower()
    if clean_id not in CLI_AGENT_SPECS:
        return
    cfg = load_cli_agents_config()
    models_map = cfg.setdefault("agent_models", {})
    models_map[clean_id] = {
        "model": (model or "").strip(),
        "reasoning_effort": (reasoning_effort or "").strip().lower(),
    }
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


_CODEX_MODELS_CACHE: Optional[Tuple[str, float, List[CliModelOption]]] = None


def _parse_codex_config_defaults(home: Path) -> Tuple[str, str]:
    """Read top-level ``model`` and ``model_reasoning_effort`` from ``~/.codex/config.toml``."""
    cfg_toml = home / ".codex" / "config.toml"
    if not cfg_toml.is_file():
        return "", ""
    default_model = ""
    default_effort = ""
    try:
        for raw_line in cfg_toml.read_text(encoding="utf-8", errors="ignore").splitlines():
            line = raw_line.strip()
            if not line or line.startswith("#"):
                continue
            if line.startswith("["):
                break
            m_mod = re.match(r'^model\s*=\s*["\']([^"\']+)["\']', line)
            if m_mod:
                default_model = m_mod.group(1).strip()
                continue
            m_eff = re.match(r'^model_reasoning_effort\s*=\s*["\']([^"\']+)["\']', line)
            if m_eff:
                default_effort = m_eff.group(1).strip().lower()
    except Exception as exc:
        logger.debug("Could not parse .codex/config.toml defaults: %s", exc)
    return default_model, default_effort


def _load_codex_cached_models(home: Path) -> List[CliModelOption]:
    """Read selectable models and reasoning efforts from ``~/.codex/models_cache.json`` (<1 ms)."""
    global _CODEX_MODELS_CACHE
    cache_file = home / ".codex" / "models_cache.json"
    if not cache_file.is_file():
        return []
    try:
        mtime = cache_file.stat().st_mtime
        key = str(cache_file)
        if _CODEX_MODELS_CACHE and _CODEX_MODELS_CACHE[0] == key and _CODEX_MODELS_CACHE[1] == mtime:
            return list(_CODEX_MODELS_CACHE[2])
        raw = json.loads(cache_file.read_text(encoding="utf-8", errors="ignore"))
        raw_models = raw.get("models", raw) if isinstance(raw, dict) else raw
        parsed: List[CliModelOption] = []
        if isinstance(raw_models, list):
            for entry in raw_models:
                if not isinstance(entry, dict):
                    continue
                if str(entry.get("visibility", "list")).strip().lower() == "hide":
                    continue
                slug = str(entry.get("slug") or entry.get("id") or "").strip()
                if not slug:
                    continue
                display_name = str(entry.get("display_name") or entry.get("name") or slug).strip()
                levels: List[str] = []
                for lvl in entry.get("supported_reasoning_levels", []) or []:
                    if isinstance(lvl, dict) and lvl.get("effort"):
                        eff = str(lvl["effort"]).strip().lower()
                        if eff and eff not in levels:
                            levels.append(eff)
                    elif isinstance(lvl, str) and lvl.strip():
                        eff = lvl.strip().lower()
                        if eff not in levels:
                            levels.append(eff)
                parsed.append(
                    CliModelOption(
                        slug=slug,
                        display_name=display_name,
                        reasoning_levels=tuple(levels),
                        default_reasoning_level=str(entry.get("default_reasoning_level") or "").strip().lower(),
                        description=str(entry.get("description") or "").strip(),
                    )
                )
        if parsed:
            _CODEX_MODELS_CACHE = (key, mtime, parsed)
            return list(parsed)
    except Exception as exc:
        logger.debug("Could not parse .codex/models_cache.json: %s", exc)
    return []


def discover_cli_models(
    agent_id: str,
    *,
    home_dir: Optional[Path] = None,
) -> Tuple[List[CliModelOption], str, str]:
    """
    Return ``(models, cli_default_model, cli_default_effort)`` for ``agent_id``.

    Uses `~/.codex/models_cache.json` + `~/.codex/config.toml` for Codex (with zero
    subprocess lag) and inspects local settings / curated presets for Claude & Gemini.
    """
    home = home_dir or Path.home()
    cid = (agent_id or "").strip().lower()

    if cid == "codex":
        default_model, default_effort = _parse_codex_config_defaults(home)
        models = _load_codex_cached_models(home)
        if not models:
            models = [CliModelOption(s, d, lvls, deff) for s, d, lvls, deff in _FALLBACK_CODEX_MODELS]
        if default_model and not any(m.slug == default_model for m in models):
            models.insert(
                0,
                CliModelOption(
                    default_model,
                    default_model,
                    ("low", "medium", "high", "xhigh", "max", "ultra"),
                    default_effort or "medium",
                ),
            )
        return models, default_model, default_effort

    if cid == "claude":
        default_model = ""
        default_effort = ""
        cfg_file = home / ".claude" / "settings.json"
        if cfg_file.is_file():
            try:
                data = json.loads(cfg_file.read_text(encoding="utf-8", errors="ignore"))
                if isinstance(data, dict):
                    default_model = str(data.get("model") or "").strip()
                    default_effort = str(data.get("effortLevel") or data.get("effort") or "").strip().lower()
            except Exception:
                pass
        models = [CliModelOption(s, d, lvls, deff) for s, d, lvls, deff in _CURATED_CLAUDE_MODELS]
        if default_model and not any(m.slug == default_model for m in models):
            models.insert(0, CliModelOption(default_model, default_model, ("low", "medium", "high"), default_effort))
        return models, default_model, default_effort

    if cid == "gemini":
        default_model = ""
        cfg_file = home / ".gemini" / "settings.json"
        if cfg_file.is_file():
            try:
                data = json.loads(cfg_file.read_text(encoding="utf-8", errors="ignore"))
                if isinstance(data, dict):
                    m_val = data.get("model")
                    if isinstance(m_val, dict):
                        default_model = str(m_val.get("name") or "").strip()
                    elif isinstance(m_val, str):
                        default_model = m_val.strip()
            except Exception:
                pass
        models = [CliModelOption(s, d, lvls, deff) for s, d, lvls, deff in _CURATED_GEMINI_MODELS]
        if default_model and not any(m.slug == default_model for m in models):
            models.insert(0, CliModelOption(default_model, default_model, (), ""))
        return models, default_model, ""

    return [], "", ""


def format_active_model_badge(agent_id: str, *, home_dir: Optional[Path] = None) -> str:
    """Format a compact UI label describing the active or CLI-default model and effort."""
    cid = (agent_id or "").strip().lower()
    cfg_model, cfg_effort = get_cli_agent_model_config(cid)
    models, cli_def_model, cli_def_effort = discover_cli_models(cid, home_dir=home_dir)
    display_lookup = {m.slug: m.display_name for m in models}

    if cfg_model or cfg_effort:
        model_disp = display_lookup.get(cfg_model, cfg_model) if cfg_model else (
            display_lookup.get(cli_def_model, cli_def_model) or "CLI Default"
        )
        eff = cfg_effort or cli_def_effort
        return f"{model_disp} · {eff}" if eff else model_disp

    if cli_def_model:
        model_disp = display_lookup.get(cli_def_model, cli_def_model)
        eff = cli_def_effort
        return f"default: {model_disp}" + (f" · {eff}" if eff else "")
    return "CLI Default"


def build_interactive_cli_args(
    agent_id: str,
    executable: str,
    prompt: str,
    *,
    model: Optional[str] = None,
    reasoning_effort: Optional[str] = None,
) -> List[str]:
    """Build interactive-mode CLI arguments for the given agent ID with optional model/effort overrides."""
    clean_id = (agent_id or "").strip().lower()
    spec = CLI_AGENT_SPECS.get(clean_id)
    if spec is None:
        raise ValueError(f"Unsupported CLI agent: {agent_id}")
    if model is None or reasoning_effort is None:
        cfg_model, cfg_effort = get_cli_agent_model_config(clean_id)
        eff_model = cfg_model if model is None else model
        eff_level = cfg_effort if reasoning_effort is None else reasoning_effort
    else:
        eff_model, eff_level = model, reasoning_effort
    return spec.args_builder(executable, prompt, (eff_model or "").strip(), (eff_level or "").strip().lower())


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
