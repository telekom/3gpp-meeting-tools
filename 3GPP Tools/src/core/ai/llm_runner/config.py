"""Persistent role defaults (`leaf` and `synthesis`) for unified email summarization."""

from __future__ import annotations

import json
import logging
from pathlib import Path
from typing import Any, Dict, Optional

from core.ai.cli_detect import detect_cli_agents, discover_cli_models, get_cli_agent_model_config
from core.ai.ollama import OllamaClient, load_ollama_config
from core.utils.paths import get_project_root

from .models import ModelSpec, SUPPORTED_PROVIDERS

logger = logging.getLogger(__name__)

CONFIG_PATH: Path = get_project_root() / "config" / "email_summary_config.json"

VALID_ROLES = ("leaf", "synthesis")

DEFAULT_SUMMARY_RUNNER_CONFIG: Dict[str, Any] = {
    "leaf": {"provider": "", "model": "", "reasoning_effort": ""},
    "synthesis": {"provider": "", "model": "", "reasoning_effort": ""},
}


def load_summary_runner_config(config_path: Optional[Path] = None) -> Dict[str, Any]:
    """Load `email_summary_config.json` with safe defaults."""
    path = Path(config_path) if config_path else CONFIG_PATH
    cfg: Dict[str, Any] = {
        role: dict(DEFAULT_SUMMARY_RUNNER_CONFIG[role]) for role in VALID_ROLES
    }
    if path.is_file():
        try:
            loaded = json.loads(path.read_text(encoding="utf-8", errors="ignore"))
            if isinstance(loaded, dict):
                for role in VALID_ROLES:
                    raw_role = loaded.get(role)
                    if isinstance(raw_role, dict):
                        prov = str(raw_role.get("provider") or "").strip().lower()
                        if prov and prov not in SUPPORTED_PROVIDERS:
                            prov = ""
                        cfg[role] = {
                            "provider": prov,
                            "model": str(raw_role.get("model") or "").strip(),
                            "reasoning_effort": str(raw_role.get("reasoning_effort") or "").strip().lower(),
                        }
        except Exception as exc:
            logger.debug("Could not load %s: %s", path, exc)
    return cfg


def save_summary_runner_config(cfg: Dict[str, Any], config_path: Optional[Path] = None) -> None:
    """Persist `email_summary_config.json` safely."""
    path = Path(config_path) if config_path else CONFIG_PATH
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps(cfg, indent=2), encoding="utf-8")
    except Exception as exc:
        logger.warning("Could not save %s: %s", path, exc)


def get_default_model_spec(role: str = "leaf", config_path: Optional[Path] = None) -> Optional[ModelSpec]:
    """Return the explicitly configured ModelSpec for `role` ('leaf' or 'synthesis'), or None."""
    clean_role = role if role in VALID_ROLES else "leaf"
    cfg = load_summary_runner_config(config_path)
    entry = cfg.get(clean_role, {})
    prov = str(entry.get("provider") or "").strip().lower()
    if not prov or prov not in SUPPORTED_PROVIDERS:
        return None
    return ModelSpec(
        provider=prov,
        model=str(entry.get("model") or "").strip(),
        reasoning_effort=str(entry.get("reasoning_effort") or "").strip().lower(),
    ).normalized()


def set_default_model_spec(
    role: str,
    spec: Optional[ModelSpec],
    config_path: Optional[Path] = None,
) -> None:
    """Persist a default ModelSpec for `role` ('leaf' or 'synthesis')."""
    clean_role = role if role in VALID_ROLES else "leaf"
    cfg = load_summary_runner_config(config_path)
    if spec is None:
        cfg[clean_role] = {"provider": "", "model": "", "reasoning_effort": ""}
    else:
        cfg[clean_role] = spec.normalized().to_dict()
    save_summary_runner_config(cfg, config_path)


def _auto_detect_fallback_spec(role: str = "leaf") -> ModelSpec:
    """Choose a sensible default across Ollama and installed CLIs when none is configured."""
    # Prefer connected Ollama for fast leaf triage if online; prefer installed CLI for deep synthesis
    ollama_models = []
    try:
        online, models, _ = OllamaClient().ping_and_get_models()
        if online and models:
            ollama_models = list(models)
    except Exception:
        ollama_models = []

    ollama_cfg = load_ollama_config()
    ollama_sel = str(ollama_cfg.get("selected_model") or "").strip()
    if ollama_sel not in ollama_models and ollama_models:
        ollama_sel = ollama_models[0]

    try:
        cli_statuses = detect_cli_agents()
    except Exception:
        cli_statuses = {}

    def _cli_spec(agent_id: str, prefer_fast: bool) -> Optional[ModelSpec]:
        st = cli_statuses.get(agent_id)
        if not st or not st.installed:
            return None
        cfg_m, cfg_e = get_cli_agent_model_config(agent_id)
        if cfg_m:
            eff = cfg_e or ("low" if prefer_fast else "high")
            return ModelSpec(provider=agent_id, model=cfg_m, reasoning_effort=eff).normalized()
        models, def_m, def_e = discover_cli_models(agent_id)
        slug = def_m or (models[0].slug if models else "")
        eff = ("low" if prefer_fast else "high") if agent_id in ("codex", "claude") else def_e
        return ModelSpec(provider=agent_id, model=slug, reasoning_effort=eff).normalized()

    if role == "leaf":
        if ollama_models and ollama_sel:
            return ModelSpec(provider="ollama", model=ollama_sel)
        for aid in ("codex", "claude", "gemini"):
            cand = _cli_spec(aid, prefer_fast=True)
            if cand is not None:
                return cand
    else:
        for aid in ("codex", "claude", "gemini"):
            cand = _cli_spec(aid, prefer_fast=False)
            if cand is not None:
                return cand
        if ollama_models and ollama_sel:
            return ModelSpec(provider="ollama", model=ollama_sel)

    return ModelSpec(provider="ollama", model=ollama_sel or "")


def resolve_effective_model_spec(
    role: str = "leaf",
    override: Optional[ModelSpec] = None,
    config_path: Optional[Path] = None,
) -> ModelSpec:
    """Resolve an effective ModelSpec from explicit override -> saved role default -> live auto-fallback."""
    if override is not None and (override.provider or "").strip().lower() in SUPPORTED_PROVIDERS:
        return override.normalized()
    saved = get_default_model_spec(role, config_path=config_path)
    if saved is not None:
        return saved
    return _auto_detect_fallback_spec(role=role)
