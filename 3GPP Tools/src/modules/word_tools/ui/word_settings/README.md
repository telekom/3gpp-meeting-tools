# `modules.word_tools.ui.word_settings`

Modular configuration console for Word Tools preferences, corporate MIP/Purview sensitivity label presets, 3GPP template ribbon installation, and LibreOffice runtime fallback.

## Package Architecture (Acyclic)

- `sensitivity_edit_dialog.py`: `SensitivityLabelEditDialog` modal sub-dialog for creating/editing label presets.
- `mixins/ui_setup.py`: `WordSettingsUiSetupMixin` (3GPP Templates, Delegate Defaults, Enterprise Sensitivity Labels, and LibreOffice groups).
- `mixins/actions_and_persistence.py`: `WordSettingsActionsMixin` (`WordConfig` load/save, Windows CLP policy sync, COM/file label inspection, template/LibreOffice toggles).
- `dialog.py`: `WordSettingsDialog` orchestrator combining both mixins.
- `__init__.py`: Public facade exporting `WordSettingsDialog` and `SensitivityLabelEditDialog`.

## Compatibility

`src/modules/word_tools/ui/word_settings_dialog.py` remains a thin backward-compatible re-export shim.
