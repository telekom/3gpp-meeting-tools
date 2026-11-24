# Specifications Dialogs Package (`ui/spec_dialogs`)

Modular breakdown behind the backward-compatible `modules.specifications.ui.dialogs` shim:

- **`config_and_filter_dialogs.py`**: `SpecsConfigDialog`, `AdvancedSyncDialog`, `TableFilterDialog`, `TargetedSyncDialog` with centralized `core.ui.theme.C` styling.
- **`add_spec_dialog.py`**: `ManualSpecFetcherThread` and `AddSpecDialog` for DynaReport metadata lookup and manual spec registration.
