# `topic_dossier` Package

Modular implementation of `TopicDossierDialog` and `TopicEditorDialog`, decomposed cleanly from `src/modules/meetings/ui/dialogs/topic_dossier_dialog.py`.

## Structure & Mixin Composition (MRO)

- `styles.py`: `ALL_COMMON_WGS`, `STYLE_CARD`, `STYLE_PRIMARY_BTN`, `STYLE_SECONDARY_BTN`.
- `editor_dialog.py` (`TopicEditorDialog`): Sub-dialog for creating/editing `MetaTopicDefinition` rules with `AiSuggestionDialog` discovery.
- `dialog.py` (`TopicDossierDialog`): Modeless top-level studio shell (`super().__init__(None)`, `Qt.Window`), session persistence, and worker shutdown inheriting from three mixins ahead of `QDialog`:
  1. `TopicDossierUiSetupMixin` (`mixins/ui_setup.py`): Left topics navigation pane, timeframe/cycle picker, destination card, export options, and `FlowLayout` action ribbon.
  2. `TopicDossierTopicsStatusMixin` (`mixins/topics_and_status.py`): Topic CRUD, cohort resolution, incremental `.export_manifest.json` inspection, and offline status refresh (`_on_check_status`).
  3. `TopicDossierExportCliMixin` (`mixins/export_and_cli.py`): `TopicInputsViewerDialog` launch, `TopicDossierExporterThread` execution, and LLM CLI launcher (honoring monkeypatches on `topic_dossier_dialog`).

## Backward Compatibility

`src/modules/meetings/ui/dialogs/topic_dossier_dialog.py` acts as a thin re-export shim exposing `TopicDossierDialog`, `TopicEditorDialog`, shared style constants, and patchable helpers (`launch_cli_in_terminal`, `_open_folder`, `TopicDossierExporterThread`).
