"""Modular local Ollama client, daemon control, and background workers package."""

from __future__ import annotations

from .client import OllamaClient
from .config import (
    CONFIG_PATH,
    DEFAULT_CONFIG,
    _normalize_ollama_host,
    is_local_host,
    is_ollama_port_open,
    load_ollama_config,
    save_ollama_config,
)
from .service import (
    build_ollama_service_environment,
    describe_ollama_hardware_overrides,
    find_ollama_binary,
    restart_ollama_service,
    start_ollama_service,
    stop_ollama_service,
)
from .workers import (
    OllamaDiagnosticsWorker,
    OllamaMonitorThread,
    OllamaServiceWorker,
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
