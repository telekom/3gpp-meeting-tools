# Ollama Configuration Dialog Package (`core/ai/ollama_ui`)

Modular breakdown behind the backward-compatible `core.ai.ollama_dialog` shim:

- **`dialog.py`**: `OllamaConfigDialog` shell (`QDialog`) composing the mixins and initializing client/worker state.
- **`mixins/ui_setup.py`**: `OllamaUiSetupMixin` building the service control, server/proxy, model preference, runtime diagnostics table, and GPU/advanced settings groups.
- **`mixins/actions_and_diagnostics.py`**: `OllamaActionsDiagnosticsMixin` handling daemon start/stop/restart workers, connectivity tests, async runtime diagnostics rendering, and save/restart prompts.
