# Unified LLM Runner (`core/ai/llm_runner/`)

Provider-agnostic text-generation runner and role-default configuration shared across local Ollama and headless external LLM CLIs (`codex`, `claude`, `gemini`).

## Modules

- `models.py`: `ModelSpec`, `ProviderModelEntry`, `RunnerResult`, and provider metadata.
- `config.py`: Persists `leaf` (fast per-thread/TDoc) and `synthesis` (deep per-AI/selection) defaults in `src/config/email_summary_config.json`.
- `runners.py`: `OllamaTextRunner`, `CliSubprocessRunner` (stdin prompt delivery, `CREATE_NO_WINDOW`, cooperative cancellation, JSONL/`-o` output extraction), `discover_all_runner_models()`, and `run_llm_text()`.
