# External CLI Detection Package (`core/ai/cli_detect`)

Acyclic modular breakdown behind the backward-compatible `core.ai.cli_detector` shim:

- **`models.py`**: Immutable descriptors (`CliModelOption`, `CliAgentSpec`, `CliAgentStatus`), default configuration schema (`DEFAULT_CONFIG`), curated fallback model catalogs, CLI argument builders, and `CLI_AGENT_SPECS`.
- **`config_store.py`**: Loads/saves `src/config/cli_agents_config.json`, manages preferred agent & model/effort overrides, and coordinates the 30-second detection status cache (respecting test patches on `core.ai.cli_detector.CONFIG_PATH`).
- **`discovery.py`**: Resolves Windows/POSIX binaries, scans CLI config files for `3gpp-tools` MCP registration, parses `~/.codex/config.toml` & `models_cache.json` (respecting patches on `core.ai.cli_detector._CODEX_MODELS_CACHE`), and discovers per-CLI models.
- **`registry.py`**: High-level orchestrators (`detect_cli_agents`, `get_installed_cli_agents`, `build_interactive_cli_args`, `format_active_model_badge`).
