# `modules.assistant.ui.tab`

Modular workspace package for [`AssistantTab`](tab.py), decomposed from the former `assistant_tab.py` monolith:

- [`tab.py`](tab.py): `AssistantTab` QWidget composition, signals, and initialization.
- [`mixins/ui_setup.py`](mixins/ui_setup.py): Header ribbon, docked status strip, chat/input cards, and cockpit tabs.
- [`mixins/docking_and_layout.py`](mixins/docking_and_layout.py): Dock reparenting, splitter toggles, prompt/chip handlers, and evidence selection.
- [`mixins/session_and_export.py`](mixins/session_and_export.py): Ollama status pill/settings, new session reset, clipboard copy, and Markdown export.
- [`mixins/research_runner.py`](mixins/research_runner.py): `AgentResearchWorker` execution, activity/evidence ingestion, KPI metrics, and cooperative shutdown.
