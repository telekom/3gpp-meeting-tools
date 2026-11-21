# Pre-Meeting Markdown Dossier Exporter (`exporters/pre_meeting/`)

`PreMeetingExporterThread` subclasses `PostMeetingExporterThread` and overrides/extends its
pipeline via 3 feature **mixins** plus the worker shell class. Each mixin holds verbatim methods
of the former 1,560-line `pre_meeting_exporter.py`; `pre_meeting_exporter.py` remains a
backward-compatible re-export shim.

> **Agents:** open only the file for your concern (table below).

## Where to edit what

| Concern | File |
|---|---|
| `__init__`, `run()` pipeline, consolidated navigable HTML/PDF dossier | `exporter_thread.py` |
| Revision lineage deduplication, multi-candidate LS reply map, Release/WI grouping & filtering | `mixins/grouping.py` |
| Phase 0 parallel pre-download (`ThreadPoolExecutor`), Phase 1 Markdown conversion, Phase 2 callout/media assembly | `mixins/export_dossiers.py` |
| Contributor activity matrix, local WI briefing `index.md`, master cover `index.md`, summary tables | `mixins/indexing.py` |

## Invariants
- **MRO Precedence over `PostMeetingExporterThread`:** `PreMeetingGroupingMixin`, `PreMeetingExportMixin`, and `PreMeetingIndexingMixin` **must** appear before `PostMeetingExporterThread` in `PreMeetingExporterThread`'s base list so Pre-Meeting overrides (`_build_ls_reply_map`, `_filter_and_group_tdocs`, `_export_all_dossiers`, `_generate_local_wi_indices`, `_generate_master_index`, `_render_markdown_table`) resolve before `PostMeetingExporterThread`'s mixins while `super().__init__(...)` still initializes `PostMeetingExporterThread` and its `QThread` signals.
- Keep every file under **520 lines**.
