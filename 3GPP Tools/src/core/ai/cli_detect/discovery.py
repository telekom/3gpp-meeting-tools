"""Executable resolution, MCP registration checks, and CLI model discovery."""

from __future__ import annotations

import json
import logging
import os
import re
import shutil
import sys
from pathlib import Path
from typing import Callable, List, Optional, Tuple

from .models import (
    _CURATED_CLAUDE_MODELS,
    _CURATED_GEMINI_MODELS,
    _FALLBACK_CODEX_MODELS,
    CliAgentSpec,
    CliModelOption,
)

logger = logging.getLogger(__name__)

_CODEX_MODELS_CACHE: Optional[Tuple[str, float, List[CliModelOption]]] = None


def _get_codex_models_cache() -> Optional[Tuple[str, float, List[CliModelOption]]]:
    """Return Codex models cache while honoring patches on `cli_detector`."""
    facade = sys.modules.get("core.ai.cli_detector")
    if facade is not None and hasattr(facade, "_CODEX_MODELS_CACHE"):
        return facade._CODEX_MODELS_CACHE
    return _CODEX_MODELS_CACHE


def _set_codex_models_cache(value: Optional[Tuple[str, float, List[CliModelOption]]]) -> None:
    """Update Codex models cache on both local module and `cli_detector` facade."""
    global _CODEX_MODELS_CACHE
    _CODEX_MODELS_CACHE = value
    facade = sys.modules.get("core.ai.cli_detector")
    if facade is not None:
        facade._CODEX_MODELS_CACHE = value


def _prefer_windows_executable(candidate: Path) -> Path:
    """Prefer `.exe`/`.cmd`/`.bat` wrappers over extensionless POSIX scripts on Windows."""
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
) -> Tuple[Optional[str], str]:
    """Resolve an executable path for ``spec`` and return ``(executable_path, source)``."""
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
    """Check local CLI config files to determine whether `3gpp-tools` MCP is registered."""
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
            lower_text = cfg_path.read_text(encoding="utf-8", errors="ignore").lower()
            if "3gpp-tools" in lower_text or "mcp_server.py" in lower_text:
                return True
        except Exception:
            continue
    return False


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
    """Read selectable models and reasoning efforts from ``~/.codex/models_cache.json``."""
    cache_file = home / ".codex" / "models_cache.json"
    if not cache_file.is_file():
        return []
    try:
        mtime = cache_file.stat().st_mtime
        key = str(cache_file)
        cached = _get_codex_models_cache()
        if cached and cached[0] == key and cached[1] == mtime:
            return list(cached[2])
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
            _set_codex_models_cache((key, mtime, parsed))
            return list(parsed)
    except Exception as exc:
        logger.debug("Could not parse .codex/models_cache.json: %s", exc)
    return []


def discover_cli_models(
    agent_id: str,
    *,
    home_dir: Optional[Path] = None,
) -> Tuple[List[CliModelOption], str, str]:
    """Return ``(models, cli_default_model, cli_default_effort)`` for ``agent_id``."""
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
