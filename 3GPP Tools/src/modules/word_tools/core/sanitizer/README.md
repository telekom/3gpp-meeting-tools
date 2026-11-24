# `modules.word_tools.core.sanitizer`

Acyclic OpenXML package backing the legacy `modules.word_tools.core.document_sanitizer` shim.

## Architecture

- `models.py` — XML namespaces (`W_NS`, `A_NS`, `CP_NS`, `DC_NS`), story/style regexes, `LANGUAGE_PRESETS`, `LanguageStats`, `DocumentHygieneAudit`, and `SanitizeConfig`.
- `xml_helpers.py` — Run-property signature comparison (`_extract_rpr_signature`, `_is_pure_language_change`, `_get_tag_name`).
- `inspector.py` — High-speed single-pass auditor (`inspect_document`).
- `engine.py` — ZIP-level document sanitizer (`sanitize_document`) for styles, stories, language revision marks, comments, authors, and core metadata.
- `__init__.py` — Public package exports.

## Dependency Flow

`models` -> `xml_helpers` -> `inspector` / `engine` -> `__init__`.
