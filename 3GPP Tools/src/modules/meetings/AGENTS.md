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
│   ├── post_meeting_exporter.py       # Post-meeting report generator (per-release & per-WI dossier)
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
- **Document Type Revision Mode Resolution (`resolve_revision_mode_for_type`):**
  - **CR & pCR:** Normative specification change proposals retain specification revision markers (`revisions="markup"` with `[INSERTED]`/`[DELETED]` tags and markup legend).
  - **LS IN & LS OUT (Liaisons):** Inter-group correspondence statements do not indicate specification changes and are exported as plain documents in `revisions="accept"` mode (all insertions rendered as normal text, deletions dropped, and markup legend omitted).
  - **Stale Detection Upgrade:** `is_conversion_stale` incorporates `expected_mode` matching, automatically re-converting older cached `markup` dossiers of LS documents into clean `accept` mode.

### 3.4. Tracked Work Items Quick Filters (Integration with Work Items Module)
- **Token Resolution & Precomputation:** When loading meeting TDocs, `TDocsTableModel._precompute_row_wi_tokens` parses and caches normalized WI tokens on each row dictionary (`row["_wi_tokens"]`). This covers explicit `Related WIs`, `agenda_map` acronyms, and parenthesized acronyms in `Agenda item description`.
- **Acronym Normalization:** `_extract_base_tokens` strips Working Group suffixes (e.g. `AmbientIoT_Ph2-CT` $\rightarrow$ `ambientiot_ph2`) and generates base tokens with and without `FS_` prefixes.
- **Proxy Model (`TDocsFilterProxyModel`):**
  - Manages `filter_followed_wis` and `filter_starred_wis` boolean flags.
  - Supports `include_unassigned_ls_in` (default `True`), ensuring incoming Liaison Statements without an assigned Work Item remain visible and exportable alongside followed topics.
  - `setTrackedWITokens(followed_tokens, starred_tokens)` updates target token sets.
  - `rowMatchesFilters(source_row, exclude_filter=...)` evaluates token intersection under `exclude_filter="tracked_wi"`. When both filters are active, it evaluates the union (`followed_tokens | starred_tokens`).
- **UI Quick-Chips & Filter Bar Ergonomics (`TDocsWindow`):**
  - Filter bar contains compact icon-only `btn_followed_wis` (`🎯`) and `btn_starred_wis` (`⭐`) checkable chips with dynamic match counts in hover tooltips, freeing substantial horizontal space.
  - Right-click context menu on `btn_followed_wis`/`btn_starred_wis` allows toggling `Include unassigned Incoming LS (LS in)`, with preference persisted in `meetings_config.json`.
  - Search input enforces minimum width, stretch factor, and native clear button (`setClearButtonEnabled(True)`).
  - Checkable comboboxes elide long individual selections with full name in tooltip and dynamically reflect active filter status (`activeFilter="true"` blue highlight).
  - Dedicated `✕ Clear` reset button (`btn_reset_filters`) instantly restores all search criteria.
  - Re-evaluates on window focus (`showEvent`) to instantly pick up tracking changes from the Work Items tab without window reload.

### 3.5. Post-Meeting Report & Dossier Markdown Generator
- **UI Dialog (`PostMeetingExportDialog`):** Modeless configuration dialog accessible from `TDocsWindow`'s Export menu (`📑 Export Post-Meeting Report (.md)...`). Allows selecting export scopes (`🎯 Followed Work Items`, `⭐ Starred Work Items`, `🎯+⭐ All Tracked`, `🌐 Whole Meeting`, `👁️ Currently Visible`, or `📋 Custom Work Items Selection...`), contribution categories (`LS in`, `pCR`, `CR`, `LS out`, `General LS`), and conversion options (auto-open, force reconversion, local index generation).
- **Exporter Worker (`PostMeetingExporterThread`):**
  - **Structure:** Hierarchical dossier tree under `Export/PostMeeting_<WG>_<Meeting>_<timestamp>/` arranged chronologically by Release (`Rel-XX/`) and then by Work Item (`<WI>/`), containing type subfolders (`LS_IN/`, `pCR/`, `CR/`, `LS_OUT/`).
  - **Contribution Filtering:** Strictly captures received incoming liaison statements (`LS in`) and approved/agreed normative outputs (`pCR`, `CR`, `LS out`). Filters out discussion papers, working papers, and unapproved/withdrawn contributions.
  - **Cross-WI Contributions:** Replicates documents associated with multiple Work Items into every relevant WI folder while embedding a bidirectional callout notice (`> [!NOTE] Cross-WI Contribution: ...`).
  - **Multi-Agenda Item Tracking:** Aggregates and lists all Agenda Items associated with a Work Item (e.g., SA2's `FS_6G_ARC` spanning AI 5.2.1, 5.2.2, etc.).
  - **Media Isolation:** Preserves collision-free media by copying `{tdoc_id}_*` assets into local `media/` subdirectories.
  - **Executive Indexing:** Generates comprehensive master `index.md` overview tables and per-WI local `index.md` summaries with relative Markdown hyperlinks.

### 3.6. Meeting Analytics & Statistics Dashboard
- **Single-Meeting Dashboard (`StatisticsExporterThread`, `html_dashboard_compiler.py`):**
  - Generates `Export/Statistics_Report.html` and companion structured Excel workbook `Export/Statistics_Data.xlsx`.
  - **Dynamic Multi-Scope Navigation:** Seamlessly scopes data across Overall Meeting (`global`), per-Work Item (🧩 `WI_Name`), and per-Agenda Item (📌 `AI_Topic`).
  - **Executive KPI Strip:** 6 responsive summary cards calculating Total Contributions (Originals vs Revisions), CR Agreement Rate % (`Agreed CRs / Handled CRs`), Meeting Treatment Rate % (`Handled vs Not Treated`), Document Composition ratio (`CR / Disc / LS`), Active Entities (with Joint % share), and Scoped Topics/WIs.
  - **Unified Semantic Palette (`core/config/plot_styles.py`):** Guarantees uniform colors across all pie/bar charts (`Agreed` = `#10B981` emerald, `Approved` = `#059669` green, `Not Treated` = `#94A3B8` slate, `Revised` = `#6366F1` indigo, `Noted` = `#F59E0B` amber, `Withdrawn` = `#EF4444` red).
  - **Document Types by Outcome:** Explicitly separates Change Requests (CR/pCR) from Discussion Papers and Liaison Statements, resolving outcome distortion.
  - **Dual Chart / Data Ergonomics:** Every visualization card provides a one-click toggle to a sortable, searchable data table with instant clipboard copy and CSV download.
  - **Cross-Filtering:** Plotly click events dynamically filter matching TDocs in the collapsible interactive explorer.
  - **Excel Workbook Export:** Generates `Executive Summary`, `Document Types`, `AI Volume`, `Outcomes`, `WI-SI Volume`, `Company Volume`, `Revision Activity`, `Specialization`, `Company x AI`, and `WI-SI x AI` sheets.
- **Cross-Meeting Contribution Audit (`ContributionStatsExporterThread`, `contribution_stats.py`):**
  - Multi-meeting audit visualizing longitudinal contribution trends, WG distribution, top co-signers, and Work Item allocation.
