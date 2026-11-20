# Main Application Shell (`core/ui/main_window/`)

`DragDropUI` is composed of 4 feature **mixins** plus the thin shell class.
Each mixin holds verbatim methods from the former 1,771-line `src/main_window.py`; they share
state via `self` (attributes created in `shell_window.py → DragDropUI.__init__` and `shell_setup.py → _setup_ui`).
`src/main_window.py` remains a backward-compatible re-export shim.

> **Agents:** open only the file for your concern (table below).

## Where to edit what

| Concern | File |
|---|---|
| Constructor, shared state, `resizeEvent` / `showEvent` / `closeEvent`, `log_message` | `shell_window.py` |
| UI layout, `NavRail`, `self.tabs`, splitters, keyboard shortcuts, focus chain | `mixins/shell_setup.py` |
| `BottomDrawer`, `RightSidePanel` dock/promote, tab switching, Command Palette, cross-tab navigation, Specs indexing | `mixins/panels_navigation.py` |
| PlantUML editor cache/templates/snippets/drag-drop, Visio read, conversion queue, Word ↔ Markdown, toasts | `mixins/puml_conversion.py` |
| Settings/Help menus, DB/Task/MCP/Proxy/Ollama/Display/WorldClock/Java dialogs, monitors, font scaling, updater, restart | `mixins/system_services.py` |

## Invariants
- Keep `DragDropUI.setMinimumSize(780, 560)` and the `self.tabs` (`.tabs.`) API intact.
- Use Smart Toasts (`show_toast` / `notify_success` / `notify_info`) instead of modal `QMessageBox.information` popups.
- Keep `super()` Qt lifecycle overrides (`__init__`, `resizeEvent`, `showEvent`, `closeEvent`) in `shell_window.py` so MRO stays trivial.
- Zero raw `#RRGGBB` hex literals (enforced by `tests/test_ui_hex_lint.py` for `src/core/ui/`).
- Keep every file under **500 lines**.
