# `post_meeting_export` Package

Modular implementation of `PostMeetingExportDialog`, decomposed cleanly from the legacy `src/modules/meetings/ui/dialogs/post_meeting_export_dialog.py` monolith.

## Structure & Mixin Composition (MRO)

`PostMeetingExportDialog` inherits from four focused mixins ahead of `QDialog`:

1. `PostMeetingUiSetupMixin` (`mixins/ui_setup.py`): Builds the header banner, scope radio grid, dual tree selector container, document category checkboxes, conversion/destination options, and action buttons.
2. `PostMeetingAnalysisMixin` (`mixins/analysis_and_estimates.py`): Indexes WI/AI/Chair-Agreement distributions (`_analyze_meeting_contributions`) and recomputes live TDoc count estimates (`_update_estimates`).
3. `PostMeetingTreeSelectionMixin` (`mixins/tree_selection.py`): Populates and synchronizes bidirectional check states, tri-state parents, and search filters across the `WI -> AI` and `AI -> WI` trees.
4. `PostMeetingOutputConfigMixin` (`mixins/output_and_config.py`): Handles directory browsing, incremental export detection, recent export menu, in-dialog Chairman's Notes rescanning, validation, and `get_export_config()`.

Shared QSS styles and `natural_ai_sort_key()` live in `styles.py`.

## Backward Compatibility

`src/modules/meetings/ui/dialogs/post_meeting_export_dialog.py` acts as a thin re-export shim exposing `PostMeetingExportDialog`, `natural_ai_sort_key`, all style constants, and `logger`.
