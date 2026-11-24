# `modules.word_tools.ui.cards.markdown`

Modular mixin decomposition of `WordMarkdownCard`, backing the legacy `modules.word_tools.ui.cards.markdown_card` shim.

## Architecture

- `mixins/ui_w2m.py` — `MarkdownW2mUiMixin`: Page 0 (Word $\to$ Markdown) selector, revision mode, media options, and preview banner.
- `mixins/ui_m2w.py` — `MarkdownM2wUiMixin`: Page 1 (Markdown $\to$ Word) selector, PlantUML/3GPP style options, and compilation banner.
- `mixins/actions_and_routing.py` — `MarkdownActionsMixin`: bidirectional drag-drop smart routing, conversion signal emission, file/folder opening, and worker result callbacks.
- `card.py` — `WordMarkdownCard` widget and navigation strip.
- `__init__.py` — Public package exports.
