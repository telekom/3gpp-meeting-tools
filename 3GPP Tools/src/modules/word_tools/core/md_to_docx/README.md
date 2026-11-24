# `modules.word_tools.core.md_to_docx`

Modular Markdown-to-Word (`.docx`) synthesis package decomposed from the former `markdown_to_docx.py` monolith:

- [`media_renderer.py`](media_renderer.py): Callout config, image path resolution, QtSvg/metafile PNG rasterization, and PlantUML diagram compilation.
- [`inline_and_lists.py`](inline_and_lists.py): OpenXML run styling, hyperlinks, image embedding, and recursive nested list rendering.
- [`converter.py`](converter.py): Headless OpenXML document synthesis (`convert_markdown_to_docx`), 3GPP styles injection, and MIP labeling.
- [`worker.py`](worker.py): Background `QThread` runner (`MarkdownToWordThread`) supporting local paths, URL downloads, and optional Word auto-open.
