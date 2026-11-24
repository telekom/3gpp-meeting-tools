# `modules.word_tools.core.word_excerpt`

Acyclic package backing the legacy `modules.word_tools.core.word_excerpt_extractor` shim.

## Architecture

- `scanner.py` — `HeadingInfo` dataclass, natural filename ordering (`natural_key`, `order_sources`), and `WordHeadingScanner`.
- `com_helpers.py` — ZIP formatting-part restore (`copy_formatting_parts`), lock polling (`wait_for_file_unlocked`), headless Word COM configuration, and isolated label application (`apply_label_in_fresh_word`).
- `engine.py` — `WordExcerptEngine` invoking sibling `core/word_excerpt_fragment_worker.py` in an isolated subprocess and merging multi-source fragments via Word COM or finalizing single fragments inside temp storage.
- `threads.py` — `WordExcerptThread` (DOCX + optional Markdown/media export) and `WordHeadingScanThread`.
- `__init__.py` — Public package exports.

## Dependency Flow

`scanner` + `com_helpers` -> `engine` -> `threads` -> `__init__`.
