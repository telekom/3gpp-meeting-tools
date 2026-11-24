# `modules.word_tools.core.styles`

Modular 3GPP Word styles registry and OpenXML packaging package decomposed from the former `styles_3gpp.py` monolith:

- [`registry.py`](registry.py): Official template URLs, `Style3GPP`, `StyleInfo`, `STYLES_3GPP_REGISTRY`, `ALL_3GPP_STYLE_NAMES_LOWER`, `LINKED_CHAR_STYLES`, and `TABLE_GRID_XML`.
- [`xml_promoter.py`](xml_promoter.py): `promote_3gpp_styles_in_xml`, non-destructive `merge_and_promote_3gpp_styles`, and Word COM `promote_3gpp_styles_in_word`.
- [`templates.py`](templates.py): Caching/auto-download for `Template_3GPP_CR.docx`, `Pseudo-CR-Form-v0-1.docx`, and `3gpp_70.dotx` installation into `%APPDATA%\Microsoft\Templates`.
- [`packager.py`](packager.py): `get_3gpp_package_part` and lock-resilient `ensure_3gpp_package_styles`.
