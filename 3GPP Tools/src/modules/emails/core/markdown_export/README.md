# Email Thread Markdown Exporter Package (`emails/core/markdown_export`)

Modular subsystem for Safelinks/Proofpoint URL unwrapping, GFM table extraction from Outlook HTML, localized quote stripping, e-Meeting callouts, QTextBrowser styling/navigation, and Option A thread card export.

## Module Inventory

| File | Responsibility |
|---|---|
| `formatting_helpers.py` | TDoc ID normalization, sender/company splitting, compact timestamps, attachment partitioning, multilingual trailing-quote detection |
| `url_and_links.py` | Safelinks/Proofpoint/FireEye unwrapping, `10.10.10.10` / Drafts URL extraction, Markdown URL & TDoc linkification |
| `html_and_tables.py` | HTML angle-bracket escaping, BeautifulSoup HTML-to-GFM table injection, Outlook list-spacing cleanup |
| `reader_ui.py` | `<<START>>..<<END>>` callout rendering, QTextBrowser stylesheet, hover tooltips, deterministic block scrolling |
| `exporter.py` | `export_thread_to_markdown()` (Option A interactive HTML cards + clean LLM Markdown) |

`thread_markdown_exporter.py` remains as a thin backward-compatible re-export shim.
