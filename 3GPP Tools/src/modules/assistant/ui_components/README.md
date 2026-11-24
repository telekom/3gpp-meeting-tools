# Assistant UI Components Package (`modules.assistant.ui_components`)

Acyclic modular package extracted from `ui/assistant_components.py` (placed alongside `ui/` to avoid eager circular imports via `ui/__init__.py`):

- `styles.py`: Design-token QSS presets, status pills, card badges, `STARTER_PROMPTS`, and `TOOL_ICONS`.
- `prompt_and_markdown.py`: `MultilinePromptEdit`, `StarterChipsWidget`, `render_markdown_to_html()`, and message card HTML renderers.
- `cockpit_widgets.py`: `EvidenceInspectorWidget`, `ActivityTimelineWidget`, and `DiagnosticsTraceWidget`.
