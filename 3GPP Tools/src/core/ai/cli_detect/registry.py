"""High-level CLI agent detection, argument assembly, and UI formatting."""

from __future__ import annotations

import time
from pathlib import Path
from typing import Callable, Dict, List, Optional

from .config_store import (
    _CACHE_TTL_SECONDS,
    get_cli_agent_model_config,
    get_status_cache,
    load_cli_agents_config,
    set_status_cache,
)
from .discovery import _resolve_executable, check_mcp_configured, discover_cli_models
from .models import CLI_AGENT_SPECS, CliAgentStatus


def format_active_model_badge(agent_id: str, *, home_dir: Optional[Path] = None) -> str:
    """Format a compact UI label describing the active model/effort override or ``CLI Default``."""
    cid = (agent_id or "").strip().lower()
    cfg_model, cfg_effort = get_cli_agent_model_config(cid)
    if not cfg_model and not cfg_effort:
        return "CLI Default"

    models, _, _ = discover_cli_models(cid, home_dir=home_dir)
    display_lookup = {m.slug: m.display_name for m in models}
    model_disp = display_lookup.get(cfg_model, cfg_model) if cfg_model else "CLI Default"
    return f"{model_disp} · {cfg_effort}" if cfg_effort else model_disp


def build_interactive_cli_args(
    agent_id: str,
    executable: str,
    prompt: str,
    *,
    model: Optional[str] = None,
    reasoning_effort: Optional[str] = None,
) -> List[str]:
    """Build interactive-mode CLI arguments for the given agent ID with optional overrides."""
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
    """Detect installed external LLM CLIs (`codex`, `claude`, `gemini`) with 30s caching."""
    cached_statuses, cached_ts = get_status_cache()
    use_cache = (
        not force_refresh
        and which_fn is None
        and user_dirs is None
        and home_dir is None
        and cached_statuses is not None
        and (time.monotonic() - cached_ts) < _CACHE_TTL_SECONDS
    )
    if use_cache and cached_statuses is not None:
        return cached_statuses

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
        set_status_cache(statuses, time.monotonic())

    return statuses


def get_installed_cli_agents(force_refresh: bool = False) -> List[CliAgentStatus]:
    """Return a list of currently installed CLI agents."""
    return [st for st in detect_cli_agents(force_refresh=force_refresh).values() if st.installed]
