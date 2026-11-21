# Spec Search UI Components (`spec_search/ui/components/`)

Modular decomposition of the Spec Search Evolution Matrix delegate, Concept B release pill
cards, tree compatibility adapters, and Clause Content Inspector previously housed in the
1,458-line `spec_search_components.py`. `spec_search_components.py` remains a backward-compatible
re-export shim.

## Submodule Navigation Map

| Module | Classes / Constants | Responsibility |
|---|---|---|
| [`styles.py`](styles.py) | `INSPECTOR_BTN_BASE`, `INSPECTOR_BTN_SUCCESS`, `INSPECTOR_BTN_AI`, `INSPECTOR_BTN_WARNING`, `INSPECTOR_BTN_INFO` | Inspector toolbar button QSS (`core.ui.theme.C` tokens, 0 raw hexes) |
| [`delegates.py`](delegates.py) | `SpecEvolutionStatusDelegate` | Evolution Matrix pill badges & highlight rendering (`scaled_font` / `scale_px` guarded by `test_font_scale.py`) |
| [`spec_card.py`](spec_card.py) | `ReleasePillWidget`, `SpecCardWidget` | Checkable major-release pills, tri-state card header, and card context menus |
| [`version_cards.py`](version_cards.py) | `_TreeCompatChildItem`, `_TreeCompatSpecItem`, `_TreeCompatRootItem`, `SpecSearchVersionCardsWidget`, `SpecSearchVersionTreeWidget` | Scrollable flow cards container & legacy `QTreeWidget` compatibility adapters |
| [`clause_inspector_render.py`](clause_inspector_render.py) | `SpecClauseInspectorRenderMixin` | HTML clause highlighting, match excerpting, next/prev search, clipboard formatting |
| [`clause_inspector.py`](clause_inspector.py) | `SpecClauseInspector` | Two-tier header, action toolbar menus, layout/maximize toggles, and Qt signal emitters |
| [`__init__.py`](__init__.py) | All 14 public/compat exports | Package exports |

## Invariants
- **Font Scaling Source Guard (`tests/test_font_scale.py`):** `delegates.py` is monitored directly — always use `scaled_font()` / `scale_px()`, never `setPointSize` or multi-arg `QFont(..., N)`.
- **Zero Raw Hexes (`tests/test_ui_hex_lint.py`):** All stylesheets and HTML spans in `components/` use `core.ui.theme.C` tokens.
- Keep every file under **400 lines**.
