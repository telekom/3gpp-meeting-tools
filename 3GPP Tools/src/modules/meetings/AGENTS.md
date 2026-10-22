# Meetings Module Developer & AI Agent Guide

This document describes the architecture, subcomponents, and rules for the **Meetings & TDocs** module (`src/modules/meetings`).

---

## 1. Module Overview

The Meetings module manages 3GPP meeting synchronization, TDoc exploration, agenda reservations, technical triage, personal stance/notes tracking, and document exports (Excel packages, Markdown reports, and TDoc text conversion).

---

## 2. Directory Structure

```text
src/modules/meetings/
├── AGENTS.md                          # This guide
├── core/
│   ├── excel_exporter.py              # Formatted Excel (.xlsx) export
│   ├── llm_exporter.py                # Batch/single document extraction and export to Markdown
│   ├── markdown_exporter.py           # Meeting summary reports (agreed CRs, LSs, company contributions)
│   ├── meetings_data_loader.py        # 3gpp_data.db queries and meeting list loaders
│   ├── meetings_network_sync.py       # FTP sync for meetings and TDoc lists
│   ├── tdoc_file_handler.py           # Downloading and caching TDoc ZIP/Word documents
│   ├── tdoc_markdown_converter.py     # Application adapter & Office wrappers for tdoc2md
│   ├── tdoc2md.py                     # [PROTECTED] Standalone 3GPP TDoc to Markdown conversion core
│   ├── tdoc_triage_worker.py          # Local LLM technical triage generation
│   ├── tdocs_db.py                    # SQLite interface for user_tdocs.db
│   └── stats/                         # Meeting statistics calculation and HTML report generation
└── ui/
    ├── meeting_selection_dialog.py    # Meeting picker dialog
    ├── meetings_tab.py                # Meetings browser UI tab
    ├── tdoc_details_dialog.py         # Full TDoc metadata and history dialog
    ├── tdoc_info_dialog.py            # Quick info card modal
    ├── tdoc_triage_dialog.py          # AI TDoc Triage execution & preview dialog
    ├── tdocs_dialogs.py               # Statistics parameters, export configuration dialogs
    ├── tdocs_menus.py                 # Right-click context menus for TDoc tables
    ├── tdocs_models.py                # High-performance TDocsTableModel & TDocsFilterProxyModel
    └── tdocs_window.py                # Primary TDoc list viewer window
```

---

## 3. Critical Architectural Rules

### 3.1. STRICT PROTECTION: `tdoc2md.py`
> [!CAUTION]
> **DO NOT MODIFY `src/modules/meetings/core/tdoc2md.py`.**
> `tdoc2md.py` is a drop-in standalone conversion engine maintained in an external upstream repository (`3gpp-tdoc-markdown-exporter`).
> - It relies strictly on the Python standard library.
> - Any application adaptations, LibreOffice/Word COM bridges, custom callbacks, or GUI integration MUST be implemented in `tdoc_markdown_converter.py` or calling modules, **NEVER** in `tdoc2md.py`.

### 3.2. Background Threading
- All FTP downloads, file extractions, Word COM automations, and LLM triage must run in `QThread` workers (e.g. `LLMExporterThread`, `TDocTriageWorker`, `MeetingsNetworkSync`).
- COM automation must initialize COM with `pythoncom.CoInitialize()` and clean up with `pythoncom.CoUninitialize()` in `finally` blocks.

### 3.3. TDoc Markdown Conversion Flow
- Primary entry point for application code: `tdoc_markdown_converter.convert_document_to_markdown()`.
- Supports `.docx`, `.docm`, and legacy `.doc` (via sensitivity-aware Word COM / LibreOffice pre-conversion).
- Extracts track changes (`[INSERTED: ...]`, `[DELETED: ...]`), tables, figures, Visio diagrams, and metadata headers.

### 3.4. Tracked Work Items Quick Filters (Integration with Work Items Module)
- **Token Resolution & Precomputation:** When loading meeting TDocs, `TDocsTableModel._precompute_row_wi_tokens` parses and caches normalized WI tokens on each row dictionary (`row["_wi_tokens"]`). This covers explicit `Related WIs`, `agenda_map` acronyms, and parenthesized acronyms in `Agenda item description`.
- **Acronym Normalization:** `_extract_base_tokens` strips Working Group suffixes (e.g. `AmbientIoT_Ph2-CT` $\rightarrow$ `ambientiot_ph2`) and generates base tokens with and without `FS_` prefixes.
- **Proxy Model (`TDocsFilterProxyModel`):**
  - Manages `filter_followed_wis` and `filter_starred_wis` boolean flags.
  - `setTrackedWITokens(followed_tokens, starred_tokens)` updates target token sets.
  - `rowMatchesFilters(source_row, exclude_filter=...)` evaluates token intersection under `exclude_filter="tracked_wi"`. When both filters are active, it evaluates the union (`followed_tokens | starred_tokens`).
- **UI Quick-Chips (`TDocsWindow`):**
  - Filter bar contains `btn_followed_wis` (`🎯 Followed WIs`) and `btn_starred_wis` (`⭐ Starred WIs`) checkable chips.
  - Dynamically calculates facet badge counts (`(N)`) in `_update_filter_facets` matching active criteria.
  - Seamlessly resets in `_clear_all_filters`.
  - Re-evaluates on window focus (`showEvent`) to instantly pick up tracking changes from the Work Items tab without window reload.
