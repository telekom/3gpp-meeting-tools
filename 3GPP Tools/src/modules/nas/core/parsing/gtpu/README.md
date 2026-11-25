# GTPv1-U Parser Package (`modules.nas.core.parsing.gtpu`)

Modular package extracted from `parsing/gtpu_parser.py` (kept as a thin backward-compatible re-export shim):

- `catalogs.py`: Clause/figure regexes, interface applicability maps, extension header descriptions, common GTP-U header field specs (`Clause 5.1`), TS 38.415 cross-spec container links, signalling extension header attachments, and HTML diagram fallbacks for Clause 8 WMF/EMF figures (`8.2`, `8.4`, `8.5`, `8.6`).
- `builders.py`: `build_clause_7_messages()`, `build_common_header_message()`, `build_gpdu_message()`, and `build_ext_header_rows()`.
- `parser.py`: `GTPUDocxParser` for TS 29.281 Clauses 5.1, 5.2.1 (`Figure 5.1-1`, `Figure 5.2.1-3` live registry parsing), 5.2.2.x, 7.2/7.3, and 8.x.
- `__init__.py`: Public package exports (`GTPUDocxParser` and catalog constants).
