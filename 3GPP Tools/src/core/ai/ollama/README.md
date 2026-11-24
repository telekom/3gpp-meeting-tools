# `core.ai.ollama`

Acyclic package backing the legacy `core.ai.ollama_client` shim.

## Architecture

- `config.py` — Persistent `ollama_config.json` loading/saving, loopback normalization (`_normalize_ollama_host`), `is_local_host`, and fast TCP probe (`is_ollama_port_open`).
- `service.py` — Local Ollama binary discovery (`find_ollama_binary`), Vulkan/iGPU environment builders, and daemon lifecycle (`start_ollama_service`, `stop_ollama_service`, `restart_ollama_service`).
- `client.py` — Offline REST client `OllamaClient` (`chat`, `stream_chat`, `ping_and_get_models`, `get_diagnostics`) using `get_ai_session()`.
- `workers.py` — Background `QThread` workers (`OllamaServiceWorker`, `OllamaDiagnosticsWorker`, `OllamaMonitorThread`).
- `__init__.py` — Public package exports.

## Dependency Flow

`config` -> `service` / `client` -> `workers` -> `__init__`.
