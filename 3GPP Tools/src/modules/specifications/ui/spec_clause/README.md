# `modules.specifications.ui.spec_clause`

Modular mixin decomposition of `SpecClauseExcerptDialog` (aliased `SpecClauseDialog`), backing the legacy `modules.specifications.ui.spec_clause_dialog` shim with zero raw hex literals (`core.ui.theme.C`).

## Architecture

- `mixins/ui_setup.py` — `SpecClauseUiSetupMixin`: header card, progress bar, clause search controls, hierarchical `QTreeWidget`, and 2-tier export footer.
- `mixins/tree_and_filter.py` — `SpecClauseTreeMixin`: `SpecExcerptPrepThread` callbacks, tree population, check cascades, search filtering, and button counters.
- `mixins/extraction.py` — `SpecClauseExtractionMixin`: `WordExcerptThread` invocation, Word/Markdown notifications, and modeless `CrPcrDialog` launch.
- `dialog.py` — `SpecClauseExcerptDialog` + `SpecClauseDialog` alias preserving top-level modeless `Qt.Window`/`Qt.NonModal` flags, clean shutdown, and session persistence.
- `__init__.py` — Public package exports.
