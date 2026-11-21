# Drop Zone Component (`core/ui/components/drop_zone/`)

Modular decomposition of the Fluent-style drag-and-drop widget previously housed in the
732-line `core/ui/drop_zone.py`. `core/ui/drop_zone.py` remains a thin, 100%-compatible
re-export shim preserving every public class, constant, and helper.

## Submodule Navigation Map

| Module | Exports / Responsibility | Key Symbols |
|---|---|---|
| [`metadata.py`](metadata.py) | `C.*`-backed state color tokens, file-type archetypes, byte formatter | `COLOR_*`, `FILE_TYPE_METADATA`, `DEFAULT_FILE_TYPE`, `format_file_size` |
| [`styles.py`](styles.py) | Scoped QSS strings & icon badge builder (`C.*` tokens only) | `DROP_ZONE_ACTION_BTN_STYLE`, `CLEAR_BTN_STYLE`, `build_file_icon_badge_style`, `MULTI_FILE_ICON_BADGE_STYLE` |
| [`pages.py`](pages.py) | Stacked empty prompt & active document card layout builders | `_DropZonePagesMixin` |
| [`widget.py`](widget.py) | Anti-aliased dashed-border `QPainter`, DnD/state machine, legacy labels | `ModernDropZone`, `InteractiveDropLabel`, `ClickableDropLabel`, `logger` |
| [`__init__.py`](__init__.py) | Package exports (all 21 public symbols in `__all__` + `logger`) | `ModernDropZone`, `InteractiveDropLabel`, `ClickableDropLabel`, `format_file_size` |

## Invariants
- **Zero Raw Hex Literals:** Every color and QSS string uses `core.ui.theme.C` tokens, enforced by `tests/test_theme.py` (`test_ui_components_uses_zero_raw_hex_literals`) and `tests/test_ui_hex_lint.py`.
- **No Package Shadowing / Circular Imports:** Housed under `core/ui/components/drop_zone/` rather than `core/ui/drop_zone/` so the legacy `core.ui.drop_zone` module path remains unambiguous, and none of its submodules import from `core.ui.components` or `core.ui.ui_components`.
- Keep every file under **450 lines**.
