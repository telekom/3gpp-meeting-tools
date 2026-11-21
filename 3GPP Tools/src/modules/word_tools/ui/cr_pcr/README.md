# `cr_pcr` Package

Modular implementation of `CrPcrDialog`, decomposed cleanly from `src/modules/word_tools/ui/cr_pcr_dialog.py`.

## Structure & Mixin Composition (MRO)

- `styles.py`: `BUTTON_STYLE_PRIMARY` styled with `core.ui.theme.C` tokens.
- `dialog.py` (`CrPcrDialog`): Non-modal top-level window shell (`super().__init__(None)`, `Qt.Window`), session state persistence, and clean thread shutdown inheriting from three mixins ahead of `QDialog`:
  1. `CrPcrUiSetupMixin` (`mixins/ui_setup.py`): Builds the mode selector, `3gpp_70.dotx` template banner, meeting/contribution forms, clause selector tree, naming preview, and action bar.
  2. `CrPcrTreeNamingMixin` (`mixins/tree_and_naming.py`): Orchestrates `SpecExcerptPrepThread`, clause tree population/filtering, change-block summary, and differentiated export naming.
  3. `CrPcrGenerationTemplateMixin` (`mixins/generation_and_template.py`): Runs `CrPcrGeneratorThread` and handles `3gpp_70.dotx` install/uninstall toggles.

## Backward Compatibility

`src/modules/word_tools/ui/cr_pcr_dialog.py` acts as a thin re-export shim exposing `CrPcrDialog`, `BUTTON_STYLE_PRIMARY`, and `logger`.
