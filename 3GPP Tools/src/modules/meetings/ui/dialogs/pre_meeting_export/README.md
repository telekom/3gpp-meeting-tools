# `pre_meeting_export` Package

Modular implementation of `PreMeetingExportDialog`, decomposed cleanly from the legacy `src/modules/meetings/ui/dialogs/pre_meeting_export_dialog.py` monolith.

## Structure & Mixin Composition (MRO)

`PreMeetingExportDialog` inherits from four focused mixins ahead of `QDialog`:

1. `PreMeetingUiSetupMixin` (`mixins/ui_setup.py`): Builds the teal briefing header, scope radio grid, dual tree selector container, contributor filter dropdown, revision/reservation checkboxes, destination options, and action bar.
2. `PreMeetingAnalysisMixin` (`mixins/analysis_and_estimates.py`): Precomputes revision supersession chains, extracts Work Item / Agenda Item / Company distributions, and computes live TDoc scope estimates (`_update_estimates`).
3. `PreMeetingTreeSelectionMixin` (`mixins/tree_selection.py`): Populates and synchronizes bidirectional check states, tri-state parents, and search filters across the `WI -> AI` and `AI -> WI` trees.
4. `PreMeetingOutputConfigMixin` (`mixins/output_and_config.py`): Handles directory browsing, incremental export detection, recent export menu, validation, and `get_export_config()`.

Shared QSS styles and `natural_ai_sort_key()` live in `styles.py`.

## Backward Compatibility

`src/modules/meetings/ui/dialogs/pre_meeting_export_dialog.py` acts as a thin re-export shim exposing `PreMeetingExportDialog`, `natural_ai_sort_key`, and all style constants.
