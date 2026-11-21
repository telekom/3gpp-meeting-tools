# Chair's Notes & Agreements Parser Package (`chair_notes/`)

Modular decomposition of `ChairNotesAgreementsParser` and result/text utilities previously housed in `chair_notes_agreements_parser.py`.

## Submodule Navigation Map

| Module | Class / Responsibility | Key Methods / Symbols |
|---|---|---|
| [`parser.py`](parser.py) | Unified facade class & DOCX entry point | `ChairNotesAgreementsParser.parse_docx_with_agreements` |
| [`text_and_results.py`](text_and_results.py) | Regex constants, PUA/whitespace cleanup, OpenXML color detection, Chair Result parser | `CHAIR_NOTES_CACHE_VERSION`, `clean_chair_notes_text`, `iter_block_items`, `_get_paragraph_dominant_color`, `parse_chair_result` |
| [`block_parser.py`](block_parser.py) | Document-order block parser (RAN1/RAN2 tab-delimited & boxed notes) | `ChairNotesBlockParserMixin._parse_block_based_docx` |
| [`table_parser.py`](table_parser.py) | Table-based parser (RAN3 3-column color-coded & SA2 multi-column) | `ChairNotesTableParserMixin` (`_find_main_tdoc_tables`, `_parse_table_based_docx`) |
| [`paragraph_agreements.py`](paragraph_agreements.py) | OpenXML green/blue & explicit agreement extraction from paragraphs | `ChairNotesParagraphAgreementsMixin._extract_colored_and_explicit_agreements_from_paragraphs` |
| [`cache_and_scan.py`](cache_and_scan.py) | `Agenda/chair_notes_cache.json` persistence & `.docx`/`.doc`/`.zip`/`.html` scanning | `ChairNotesCacheAndScanMixin` (`save_cached_chair_notes_data`, `load_cached_chair_notes_data`, `parse_chair_notes_file`, `scan_and_cache_chair_notes`) |
| [`__init__.py`](__init__.py) | Package exports | `ChairNotesAgreementsParser` and text/result utilities |

## Backward Compatibility

- [`../chair_notes_agreements_parser.py`](../chair_notes_agreements_parser.py) remains as a thin re-export shim exposing `ChairNotesAgreementsParser`, `CHAIR_NOTES_CACHE_VERSION`, `clean_chair_notes_text`, `iter_block_items`, and `parse_chair_result`.
