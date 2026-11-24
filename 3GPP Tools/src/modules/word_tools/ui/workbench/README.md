# `modules.word_tools.ui.workbench`

Modular mixin decomposition of `WordExtractorTab`, backing the legacy `modules.word_tools.ui.word_tabs` shim (named `workbench` so it never shadows `word_tabs.py`).

## Architecture

- `mixins/ui_setup.py` — `WordTabUiSetupMixin`: 2-tier ribbon, `FlowLayout` navigation pills, 9 stacked workbench cards, and backward-compatible widget aliases.
- `mixins/status_and_poll.py` — `WordTabStatusAndPollMixin`: workbench switching, visibility-gated Active Word COM polling (resolving `word_tabs.get_active_word_document` dynamically), LibreOffice/DOTX status, and settings menu.
- `mixins/compat_and_callbacks.py` — `WordTabCompatAndCallbacksMixin`: drop handlers, legacy delegating properties/presets, and `MainWindow` conversion callbacks.
- `tab.py` — `WordExtractorTab` widget and timers.
- `__init__.py` — Public package exports.
