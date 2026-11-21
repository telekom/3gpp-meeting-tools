# Core UI Components (`core/ui/components/`)

Modular decomposition of the global Qt stylesheet, reusable widget styles, 3GPP WG badge
palettes, mnemonic escaping, Windows taskbar/icon manager, and DPAPI-encrypted `ProxyDialog`
previously housed in the 1,434-line `ui_components.py`. `ui_components.py` remains a
100%-compatible re-export shim preserving all 45 symbols (including `drop_zone` re-exports).

## Submodule Navigation Map

| Module | Exports / Responsibility | Key Symbols |
|---|---|---|
| [`global_style.py`](global_style.py) | Application-wide All-Blue Qt stylesheet & combo chevron | `GLOBAL_STYLE`, `_combo_chevron_url`, `_GLOBAL_COMBO_CHEVRON`, `_GLOBAL_COMBO_ARROW_QSS` |
| [`styles.py`](styles.py) | Shared card, table, tree, input, badge, pill & button QSS constants | `CARD_FRAME_STYLE`, `TABLE_STYLE_CLEAN`, `TREE_STYLE_CLEAN`, `BADGE_STYLE_*`, `BUTTON_STYLE_TOOLBAR_*`, `TOGGLE_PILL_STYLE`, etc. |
| [`wg_and_text.py`](wg_and_text.py) | Working Group palette/descriptions & Qt mnemonic ampersand helper | `WG_DESCRIPTIONS`, `get_wg_colors`, `get_wg_badge_style_and_desc`, `escape_qt_mnemonic` |
| [`window_icon.py`](window_icon.py) | Multi-DPI `QIcon`, OneDrive hydration, Win32 `WM_SETICON`/`IPropertyStore` stamping & global filter | `APP_USER_MODEL_ID`, `APP_DISPLAY_NAME`, `create_app_icon`, `get_app_icon_path`, `apply_native_window_icon`, `apply_console_window_icon`, `install_window_icon_manager` |
| [`proxy_dialog.py`](proxy_dialog.py) | Network proxy connection worker & Windows DPAPI profile dialog | `ProxyTestWorker`, `ProxyDialog` |
| [`__init__.py`](__init__.py) | Package exports (re-exports all 45 symbols incl. `drop_zone`) | All symbols in `__all__` |

## Invariants
- **Zero Raw Hex Literals:** Every stylesheet and painter color in `components/` uses `core.ui.theme` (`C`, `R`, `T`) tokens, guarded by `tests/test_theme.py` (`test_ui_components_uses_zero_raw_hex_literals`) and `tests/test_ui_hex_lint.py`.
- **Single-Responsibility & Compatibility:** External code may import either from `core.ui.ui_components` or `core.ui.components`.
- Keep every file under **440 lines**.
