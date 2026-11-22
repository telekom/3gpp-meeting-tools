# Navigable Dossier Exporter (`meetings/core/exporters/dossier/`)

Modular decomposition of the Unified Navigable Dossier Exporter (F-04) previously housed in
the 1,380-line `dossier_exporter.py`. `dossier_exporter.py` remains a thin, 100%-compatible
re-export shim preserving `DossierExporterThread`, `DOSSIER_CSS`, and `logger`.

## Submodule Navigation Map

| Module | Exports / Responsibility | Key Symbols |
|---|---|---|
| [`css.py`](css.py) | Standalone HTML & print-PDF stylesheet (`@media print` rules) | `DOSSIER_CSS` |
| [`html_builder.py`](html_builder.py) | CommonMark rendering, `[INSERTED]`/`[DELETED]` diff styling, Base64 images, TOC & HTML dossier | `convert_markdown_to_html`, `build_html_dossier` |
| [`pdf_renderer.py`](pdf_renderer.py) | Vector `QTextDocument` + `QPdfWriter(A4)` page compilation & `pypdf` hierarchical outline bookmarks | `render_bookmarked_pdf` |
| [`prompt_builder.py`](prompt_builder.py) | Companion AI citation map with 1-indexed page spans & Prompts A/B/C | `generate_ai_prompt` |
| [`exporter_thread.py`](exporter_thread.py) | Background `QThread` (parallel download + sequential COM-safe conversion + HTML/PDF/Prompt) | `DossierExporterThread`, `logger` |
| [`__init__.py`](__init__.py) | Package exports | `DossierExporterThread`, `DOSSIER_CSS`, `build_html_dossier`, `convert_markdown_to_html`, `render_bookmarked_pdf`, `generate_ai_prompt` |

## Invariants
- **Parallel Download / Sequential Conversion:** Phase 1 pre-downloads missing TDocs via `ThreadPoolExecutor` while Phase 2 converts sequentially on the worker thread with `isInterruptionRequested()` checks.
- **Signature Compatibility:** `render_bookmarked_pdf` / `DossierExporterThread._render_bookmarked_pdf` supports both the 4-argument modular signature and 3-argument legacy signature.
- Keep every file under **450 lines**.
