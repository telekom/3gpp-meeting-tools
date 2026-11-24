# `modules.word_tools.core.presentation`

Modular PowerPoint/PDF presentation-to-Markdown package decomposed from the former `presentation_to_markdown.py` monolith:

- [`models.py`](models.py): `ExtractedSlide`, `CONVERTER_VERSION`, `TDOC_RE`, `OPENXML_NS`, filename sanitization, and GFM table formatting.
- [`com_extractor.py`](com_extractor.py): Tier 1 Windows PowerPoint COM extractor (1080p slide screenshots, titles, indented paragraphs, tables, speaker notes).
- [`openxml_and_pdf.py`](openxml_and_pdf.py): Tier 2 pure-Python `.pptx` OpenXML parser and Tier 3 PDF slide extractor (`pypdfium2` + `pypdf`).
- [`converter.py`](converter.py): 3GPP TDoc/generic header synthesis and multi-tier dispatch orchestrator (`convert_presentation_to_markdown`).
