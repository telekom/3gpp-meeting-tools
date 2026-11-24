# ASN.1 Base Parser Package (`core/parsing/asn1_base`)

Acyclic modular breakdown behind the backward-compatible `modules.nas.core.parsing.asn1_base_parser` shim:

- **`scanner.py`**: OpenXML ASN.1 block & field-description table scanning, prose/direction extraction, and QTextBrowser HTML assembly (`RE_ASN1_BOX_CAPTION`, `scan_xml_documents`, `build_inspector_html`).
- **`grammar.py`**: ASN.1 comment-aware brace matching, depth-0 comma token splitting, inline `SEQUENCE`/`CHOICE` field parsing, and type definition extraction.
- **`base_parser.py`**: `BaseAsn1DocxParser` preserving constructor state and all delegating protected/static methods used by `asn1_parser`, `rrc_asn1_parser`, and `ran3_asn1_parser`.
