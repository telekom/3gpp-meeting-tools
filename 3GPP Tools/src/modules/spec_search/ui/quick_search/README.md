# `modules.spec_search.ui.quick_search`

Modular decomposition of `QuickSpecSearchPanel` (`Ctrl+Shift+F`), backing the legacy `modules.spec_search.ui.quick_search_panel` shim.

## Architecture

- `delegate.py` — `_QuickSearchResultDelegate` HTML result card renderer and `_format_snippet_html`.
- `styles.py` — Centralized `C`/`R`/`T` token QSS (`apply_quick_search_styles`).
- `mixins/ui_setup.py` — `QuickSearchUiSetupMixin`: filter bar, result list, collapsible preview pane, and action buttons.
- `mixins/search_and_preview.py` — `QuickSearchLogicMixin`: version resolution, debounced `SpecSearchQueryWorker` execution, highlighted preview, and deep links.
- `panel.py` — `QuickSpecSearchPanel` QWidget composition and debounce timer setup.
- `__init__.py` — Public package exports.
