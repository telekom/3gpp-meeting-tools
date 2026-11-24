# --- File: src/core/ai/ollama_client.py ---
"""Backward-compatible re-export shim for ``core.ai.ollama``."""

from __future__ import annotations

from .ollama import (
    CONFIG_PATH,
    DEFAULT_CONFIG,
    OllamaClient,
    OllamaDiagnosticsWorker,
    OllamaMonitorThread,
    OllamaServiceWorker,
    _normalize_ollama_host,
    build_ollama_service_environment,
    describe_ollama_hardware_overrides,
    find_ollama_binary,
    is_local_host,
    is_ollama_port_open,
    load_ollama_config,
    restart_ollama_service,
    save_ollama_config,
    start_ollama_service,
    stop_ollama_service,
)

__all__ = [
    "CONFIG_PATH",
    "DEFAULT_CONFIG",
    "_normalize_ollama_host",
    "load_ollama_config",
    "save_ollama_config",
    "is_local_host",
    "is_ollama_port_open",
    "find_ollama_binary",
    "build_ollama_service_environment",
    "describe_ollama_hardware_overrides",
    "start_ollama_service",
    "stop_ollama_service",
    "restart_ollama_service",
    "OllamaServiceWorker",
    "OllamaDiagnosticsWorker",
    "OllamaClient",
    "OllamaMonitorThread",
]
