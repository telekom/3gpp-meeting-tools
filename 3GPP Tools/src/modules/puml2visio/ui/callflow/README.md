# `modules.puml2visio.ui.callflow`

Modeless AI-powered PlantUML sequence diagram generator from 3GPP specification call flows (`CallFlowDialog`).

## Package Architecture (Acyclic)

- `mixins/ui_setup.py`: `CallFlowUiSetupMixin` (model/type bar, samples menu, left call-flow editor + style toggles, right code/SVG preview tabs, syntax error banner, refinement & bottom action bars).
- `mixins/generation_and_validation.py`: `CallFlowGenerationMixin` (Ollama model refresh, directive assembly, streaming generation, PlantUML SVG validation, iterative refinement, and LLM syntax auto-fix).
- `dialog.py`: `CallFlowDialog` orchestrator composing the mixins with `super().__init__(None)` + `Qt.Window` modeless lifecycle and clean worker/validator shutdown.
- `__init__.py`: Public facade exporting `CallFlowDialog`.

## Compatibility

`src/modules/puml2visio/ui/callflow_dialog.py` remains a thin backward-compatible re-export shim.
