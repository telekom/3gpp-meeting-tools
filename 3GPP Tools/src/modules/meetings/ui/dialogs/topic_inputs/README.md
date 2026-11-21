# `topic_inputs` Package

Modular implementation of `TopicInputsViewerDialog` and `InputsCollectorThread`, decomposed cleanly from the legacy `src/modules/meetings/ui/dialogs/topic_inputs_dialog.py` monolith.

## Structure & Mixin Composition (MRO)

`TopicInputsViewerDialog` inherits from four focused mixins ahead of `QDialog`:

1. `TopicInputsUiSetupMixin` (`mixins/ui_setup.py`): Builds the topic header card, KPI counters, progress strip, WG/Category filter pills, 4-tab view (`Decisions`, `All Contributions`, `Liaisons`, `WG Matrix`), collapsible detail reader drawer, and bottom action ribbon.
2. `TopicInputsCollectionFilteringMixin` (`mixins/collection_and_filtering.py`): Orchestrates `InputsCollectorThread` (`collector_thread.py`), updates KPI counters and pill counts, and runs interactive filter logic.
3. `TopicInputsTablesMixin` (`mixins/tables.py`): Populates all four `QTableWidget` tabs with provenance tooltips, color badges, and summary totals.
4. `TopicInputsDrawerActionsMixin` (`mixins/drawer_and_actions.py`): Renders rich-HTML detail provenance in the reader drawer and handles local document opening, 3GPP Portal links, Markdown clipboard export, and CSV export.

Shared stylesheets and `BADGE_COLORS` live in `styles.py`.

## Backward Compatibility

`src/modules/meetings/ui/dialogs/topic_inputs_dialog.py` acts as a thin re-export shim exposing `TopicInputsViewerDialog`, `InputsCollectorThread`, `STYLE_PRIMARY_BTN`, `STYLE_SECONDARY_BTN`, `STYLE_TABLE`, `BADGE_COLORS`, and `logger`.
