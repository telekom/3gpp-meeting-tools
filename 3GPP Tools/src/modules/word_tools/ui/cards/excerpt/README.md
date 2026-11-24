# Word Excerpt Card Package (`ui/cards/excerpt`)

Modular breakdown behind the backward-compatible `modules.word_tools.ui.cards.excerpt_card` shim:

- **`card.py`**: `WordExcerptCard` widget state & mixin composition.
- **`mixins/ui_setup.py`**: `ExcerptCardUiSetupMixin` building the source drop zone, tree view, naming preview bar, and action row using `core.ui.theme.C`.
- **`mixins/scanning_and_tree.py`**: `ExcerptCardScanningTreeMixin` handling source management, `WordHeadingScanThread` lifecycle, filtering, and checkbox selection.
- **`mixins/extraction.py`**: `ExcerptCardExtractionMixin` handling differentiated filename previews and `WordExcerptThread` execution (preserving patch hooks on `excerpt_card.WordExcerptThread`).
