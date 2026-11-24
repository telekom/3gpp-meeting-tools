# `modules.word_tools.ui.cards.sanitizer`

Modular mixin decomposition of `WordSanitizerCard`, backing the legacy `modules.word_tools.ui.cards.sanitizer_card` shim with zero raw hex literals (`core.ui.theme.C`).

## Architecture

- `mixins/ui_setup.py` — `SanitizerUiSetupMixin`: header, `DocumentSelectorPane`, KPI flow strip, presets, language/author configuration panels, reviewer table, and green result card.
- `mixins/audit_and_actions.py` — `SanitizerActionsMixin`: preset application, `inspect_document` audit population, open-in-Word prompt handling, and `sanitize_document` dispatch.
- `card.py` — `WordSanitizerCard` widget composition.
- `__init__.py` — Public package exports.
