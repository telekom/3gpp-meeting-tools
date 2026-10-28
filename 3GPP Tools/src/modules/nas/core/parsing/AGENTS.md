# Parsing module guide

This package reads 3GPP specification DOCX files and converts their protocol
message and information-element definitions into structured data. The common
parser interface returns `(messages, ie_definitions)` and accepts an optional
progress callback.

## Parsing flow

- `protocol_parser_common.py` provides `ProtocolDocxDispatcher`, the unified
  entry point. It normalizes input paths, orders multipart documents by part
  index, infers the specification number/version from the filename, and routes
  the document to a protocol parser.
- Protocol parsers extract WordprocessingML from DOCX files and interpret the
  clauses, tables, and ASN.1 content relevant to that protocol.
- `asn1_parser.py` is a compatibility/export module for the ASN.1 parser
  classes and their older aliases.

## Files

- `nas_parser.py` — NAS & UE Policy message definitions and IE coding tables
  (TS 24.501, TS 24.301, TS 24.526, TS 24.008); the dispatcher's fallback parser.
  Supports Clause 8 / Annex D.5 message content tables, Clause 9 / Annex D.6 IE
  coding definitions, container unrolling (`Payload container`, `UE policy section
  management list`, `Service-level-AA container`, `Type 6 IE container`, `ePCO`),
  and TS 24.526 Clause 5.2 (URSP rules) / Clause 5.3 (ANDSP / WLANSP / N3AN)
  policy rule hierarchies.
- `pfcp_parser.py` — PFCP message tables, type registry, grouped IEs, and field
  expansion (TS 29.244).
- `gtpu_parser.py` — GTP-U message tables and type/extension-header registries
  (TS 29.281).
- `pdu_session_up_parser.py` — PDU Session User Plane frame tables and field
  definitions (TS 38.415).
- `asn1_base_parser.py` — shared ASN.1 DOCX scanning, extraction of type
  definitions and field descriptions, ASN.1 token/field parsing, trailing comma
  sanitization, and inspector HTML construction.
- `rrc_asn1_parser.py` — RRC ASN.1 message selection and field unrolling, built
  on `BaseAsn1DocxParser` (TS 36.331/38.331).
- `ran3_asn1_parser.py` — RAN3 ASN.1 object-set, `SEQUENCE OF` list unrolling,
  `ProtocolExtensionContainer` resolution, `MessageType` synthesis, and protocol
  IE extraction with Assigned Criticality tracking, built on `BaseAsn1DocxParser`
  (supported specs are listed in the dispatcher).
- `protocol_parser_utils.py` — shared DOCX archive/XML loading, Word paragraph
  and table-cell text extraction with Unicode superscript/subscript preservation
  (e.g., preserving `2⁴⁰ - 1`), multi-line `<w:br/>` and paragraph break
  preservation inside coding tables, half-octet `<w:vMerge>` split-row collapsing
  (`_collapse_word_half_octet_layout_rows`), and conversion of Word tables to HTML.
- `protocol_parser_constants.py` — WordprocessingML namespace/tag constants
  and shared regular expressions for filenames, clauses, captions, and ASN.1.
- `__init__.py` — package marker.

## Maintenance notes

- The parsers account for 3GPP-specific document conventions, such as multipart
  specifications, clause and table layouts, grouped information elements, and
  ASN.1 structures. Check the relevant parser's context before changing or
  generalizing logic; similar-looking structures can have protocol-specific
  meaning.
- Keep specification routing in `ProtocolDocxDispatcher`; it also determines
  which parser handles otherwise-unmatched specifications.
- DOCX parsers share the `(messages, ie_definitions)` result shape and optional
  progress callback. Preserve this when adding or changing parser routes.
- Multipart ordering, filename patterns, XML tags, and shared clause patterns
  are centralized in `protocol_parser_common.py` and
  `protocol_parser_constants.py`; check these before adding duplicates.
