# Post-Meeting Markdown Dossier Exporter (`exporters/post_meeting/`)

`PostMeetingExporterThread` is composed of 4 feature **mixins** plus the `QThread` shell class and
pure path/link helpers. Each mixin holds verbatim methods of the former 1,763-line
`post_meeting_exporter.py`; they share state via `self` (initialized in `exporter_thread.py → __init__`).
`post_meeting_exporter.py` remains a backward-compatible re-export shim.

> **Agents:** open only the file for your concern (table below).

## Where to edit what

| Concern | File |
|---|---|
| `QThread` signals (`progress`, `finished`), `__init__`, `run()` pipeline, consolidated dossier | `exporter_thread.py` |
| Folder sanitization, relative paths, Markdown link encoding, anchor slugs (`sanitize_folder_name`, `safe_relpath`, etc.) | `path_utils.py` |
| LS reply tracking, Work Items DB load, scope filtering, Rel-XX/WI grouping, WI/Release resolution | `mixins/grouping.py` |
| Phase 1 parallel pre-download (`ThreadPoolExecutor` → `.part`) & Phase 2 sequential Word COM conversion | `mixins/conversion.py` |
| Local WI `index.md`, Master cover `index.md`, Markdown summary table rendering | `mixins/indexing.py` |
| RAN Chairman's Notes agreement attachment (`agreements.md`, AI → WI inheritance) | `mixins/agreements.py` |

## Invariants
- **`PreMeetingExporterThread` subclasses `PostMeetingExporterThread`:** keep `QThread` last in `PostMeetingExporterThread`'s base list and `pyqtSignal` definitions directly on `PostMeetingExporterThread`.
- **Two-Phase Parallel Pipeline:** pre-download concurrently to `.part` files (`ThreadPoolExecutor`), then convert sequentially on a single thread to isolate Word COM.
- **LLM-Safe Two-Line Headers:** `- **Source:**` lists confirmed contributors only; tentative carry-over co-signers go on the secondary bullet/note.
- **Strict RAN WG Scoping (`is_ran = 'RAN' in wg_name`):** only attach Chairman's Notes agreements for RAN working groups.
- Keep every file under **560 lines**.
