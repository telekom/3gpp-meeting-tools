"""Unified text-generation runner package for Ollama and external LLM CLIs."""

from .config import (
    CONFIG_PATH,
    DEFAULT_SUMMARY_RUNNER_CONFIG,
    VALID_ROLES,
    get_default_model_spec,
    load_summary_runner_config,
    resolve_effective_model_spec,
    save_summary_runner_config,
    set_default_model_spec,
)
from .models import (
    PROVIDER_META,
    SUPPORTED_PROVIDERS,
    ModelSpec,
    ProviderModelEntry,
    RunnerResult,
)
from .runners import (
    CancellationFlag,
    ChunkCallback,
    CliSubprocessRunner,
    OllamaTextRunner,
    build_headless_cli_command,
    discover_all_runner_models,
    run_llm_text,
)

__all__ = [
    "CONFIG_PATH",
    "DEFAULT_SUMMARY_RUNNER_CONFIG",
    "PROVIDER_META",
    "SUPPORTED_PROVIDERS",
    "VALID_ROLES",
    "CancellationFlag",
    "ChunkCallback",
    "CliSubprocessRunner",
    "ModelSpec",
    "OllamaTextRunner",
    "ProviderModelEntry",
    "RunnerResult",
    "build_headless_cli_command",
    "discover_all_runner_models",
    "get_default_model_spec",
    "load_summary_runner_config",
    "resolve_effective_model_spec",
    "run_llm_text",
    "save_summary_runner_config",
    "set_default_model_spec",
]
