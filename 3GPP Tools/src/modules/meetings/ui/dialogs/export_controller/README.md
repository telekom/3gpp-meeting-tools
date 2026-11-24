# `modules.meetings.ui.dialogs.export_controller`

Modular TDoc export orchestration package decomposed from the former `tdocs_export_controller.py` monolith:

- [`package_and_excel.py`](package_and_excel.py): Selected/visible row resolution, TDoc package bundles + Excel manifests, and configurable Excel workbook exports.
- [`markdown_and_dossiers.py`](markdown_and_dossiers.py): Single/bulk Markdown export, Pre-/Post-Meeting briefing workflows, navigable HTML/PDF dossiers, and post-export CLI prompts.
- [`cli_sessions.py`](cli_sessions.py): Revision-row expansion, canonical vs. batch CLI workspace staging (`_CLI_CONTEXT.md`), and interactive LLM CLI terminal launch.
