"""Persistent CLI configuration store and detection cache state."""

from __future__ import annotations

import json
import logging
import sys
from pathlib import Path
from typing import Any, Dict, Optional, Tuple

from core.utils.paths import get_project_root

from .models import CLI_AGENT_SPECS, DEFAULT_CONFIG

logger = logging.getLogger(__name__)

CONFIG_PATH = get_project_root() / "config" / "cli_agents_config.json"
_CACHE_TTL_SECONDS = 30.0

_STATUS_CACHE: Optional[Dict[str, Any]] = None
_STATUS_CACHE_TS: float = 0.0


def _resolve_config_path() -> Path:
    """Resolve config path dynamically via facade module to preserve test patching."""
    facade = sys.modules.get("core.ai.cli_detector")
    if facade is not None and hasattr(facade, "CONFIG_PATH"):
        return Path(facade.CONFIG_PATH)
    return CONFIG_PATH


def get_status_cache() -> Tuple[Optional[Dict[str, Any]], float]:
    """Return cached CLI status dictionary and timestamp."""
    return _STATUS_CACHE, _STATUS_CACHE_TS


def set_status_cache(statuses: Optional[Dict[str, Any]], timestamp: float) -> None:
    """Update cached CLI status dictionary and timestamp."""
    global _STATUS_CACHE, _STATUS_CACHE_TS
    _STATUS_CACHE = statuses
    _STATUS_CACHE_TS = timestamp
    facade = sys.modules.get("core.ai.cli_detector")
    if facade is not None:
        facade._STATUS_CACHE = statuses
        facade._STATUS_CACHE_TS = timestamp


def invalidate_cli_detection_cache() -> None:
    """Clear the cached CLI agent detection snapshot."""
    set_status_cache(None, 0.0)


def load_cli_agents_config() -> Dict[str, Any]:
    """Load `cli_agents_config.json` from disk with safe defaults."""
    cfg: Dict[str, Any] = {
        "preferred_agent": DEFAULT_CONFIG["preferred_agent"],
        "custom_paths": dict(DEFAULT_CONFIG["custom_paths"]),
        "agent_models": {aid: dict(vals) for aid, vals in DEFAULT_CONFIG["agent_models"].items()},
    }
    config_path = _resolve_config_path()
    if config_path.exists():
        try:
            with open(config_path, "r", encoding="utf-8") as f:
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
        config_path = _resolve_config_path()
        config_path.parent.mkdir(parents=True, exist_ok=True)
        with open(config_path, "w", encoding="utf-8") as f:
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
