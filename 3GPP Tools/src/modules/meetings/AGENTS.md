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
│   ├── meeting_wi_enricher.py         # Asynchronous background Work Item discovery & enrichment
│   ├── meetings_data_loader.py        # 3gpp_data.db queries and meeting list loaders
│   ├── meetings_network_sync.py       # FTP sync for meetings and TDoc lists
│   ├── tdoc_file_handler.py           # Downloading and caching TDoc ZIP/Word documents
│   ├── tdoc_url_resolver.py           # Official 3GPP server HTTPS download URL resolution
│   ├── tdoc_markdown_converter.py     # Application adapter & Office wrappers for tdoc2md
│   ├── tdoc2md.py                     # [PROTECTED] Standalone 3GPP TDoc to Markdown conversion core
│   ├── tdoc_triage_worker.py          # Local LLM technical triage generation
│   ├── tdocs_db.py                    # SQLite interface for user_tdocs.db
│   └── stats/                         # Meeting statistics calculation and HTML report generation
└── ui/
    ├── contribution_dialog.py         # Multi-meeting company contribution audit dialog
    ├── dialogs.py                     # MeetingInfoDialog, AddMeetingDialog, MeetingsConfigDialog
    ├── meetings_actions.py            # Async download, caching, and export actions for MeetingsTab
    ├── meetings_delegates.py          # DynamicHoverMenuDelegate, TDocsButtonDelegate, WGTagDelegate
    ├── meetings_menu.py               # Right-click context menu builder for meetings table
    ├── models.py                      # MeetingsTableModel (7 consolidated columns, English date ranges)
    ├── post_meeting_export_dialog.py  # Post-meeting report generator dialog
    ├── search_controller.py           # GlobalSearchController for inline Quick TDoc Jump
    ├── tdocs_components.py            # CheckableComboBox and custom styled UI widgets
    ├── tdocs_dialogs.py               # Statistics parameters, export configuration dialogs
    ├── tdocs_menus.py                 # Right-click context menus for TDoc tables
    ├── tdocs_models.py                # High-performance TDocsTableModel & TDocsFilterProxyModel
    ├── tdocs_window.py                # Primary TDoc list viewer window
    └── ui_tabs.py                     # MeetingsTab (full-width browser, 2-row top ribbon, floating cart)
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

### 3.3. TDoc Markdown Conversion & Provenance Flow
- Primary entry point for application code: `tdoc_markdown_converter.convert_document_to_markdown()`.
- Supports `.docx`, `.docm`, and legacy `.doc` (via sensitivity-aware Word COM / LibreOffice pre-conversion).
- Extracts track changes (`[INSERTED: ...]`, `[DELETED: ...]`), tables, figures, Visio diagrams, and metadata headers.
- **Document Type Revision Mode Resolution (`resolve_revision_mode_for_type`):**
  - **CR & pCR:** Normative specification change proposals retain specification revision markers (`revisions="markup"` with `[INSERTED]`/`[DELETED]` tags and markup legend).
  - **LS IN & LS OUT (Liaisons):** Inter-group correspondence statements do not indicate specification changes and are exported as plain documents in `revisions="accept"` mode (all insertions rendered as normal text, deletions dropped, and markup legend omitted).
  - **Stale Detection Upgrade:** `is_conversion_stale` incorporates `expected_mode` matching, automatically re-converting older cached `markup` dossiers of LS documents into clean `accept` mode.
- **3GPP Server URL & Provenance Preservation (`tdoc_url_resolver.py`):**
  - Canonical resolution (`resolve_tdoc_download_url`) normalizes legacy FTP paths to official HTTPS (`https://www.3gpp.org/ftp/...`) and routes revisions (e.g. `r01`, `rev2`) to `/INBOX/Revisions/` and base contributions to `/Docs/`.
  - Embeds `- **3GPP Server URL:** [<url>](<url>)` directly into standalone `.md` headers, ensuring permanent provenance when summaries are mixed or exported downstream.
  - Manifest index (`write_index`) includes a dedicated `3GPP Server` download link column (`[📥 Download ZIP](<url>)`).
  - Post-Meeting reports display dual links (`[Dossier](<rel_path>) · [📥 3GPP](<url>)`) in executive and WI index tables.
  - AI TDoc Triage dialog provides 1-click header links and includes the download URL in copied summaries and saved user notes.

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
  - **Liaison Statement Reply Tracking & Dossier Linking:**
    - Scans meeting metadata for `Reply in` and `Reply to` attributes across incoming and outgoing LS documents, tracing revision chains (`Revised to`) until the final approved TDoc.
    - Master and local index tables link directly to reply documents: `LS in` rows display `replied to in [S2-XXXX](<relative_path>)`, and `LS out` rows display `approved (reply to [S2-YYYY](<relative_path>))`.
    - Automatically injects reciprocal Markdown callouts (`[!TIP] Outgoing Reply` and `[!NOTE] Reply Liaison Statement`) into dossier headers for seamless 1-click back-and-forth navigation.
    - **Guaranteed Reply Document Inclusion:** If an included `LS in` was replied to by an approved `LS out` whose Work Item was not in the export scope (or unassigned), the reply document is automatically exported under `General/LS_OUT/<tdoc>.md` (in `revisions="accept"` mode) so that all reply links resolve to valid dossiers on disk without creating unwanted non-followed WI sections in the master index.

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

### 3.7. Opportunistic Work Item Auto-Enrichment (`meeting_wi_enricher.py`)
- **Decoupled Asynchronous Architecture:** `MeetingWIEnricherThread` is fired fire-and-forget in the background when `TDocsWindow` is initialized or when a meeting list is refreshed. It never blocks or slows down initial TDocs table loading.
- **Bounded Hyperlink Extraction:** Scans populated rows of the "Related WIs" column for 3GPP Portal hyperlinks matching `workitemId=(\d+)`. Breaks boundedly when consecutive blank TDoc rows are encountered.
- **Database Cross-Checking:** Queries SQLite `work_items` via `WorkItemsDatabase.get_existing_wi_codes()` to isolate only unindexed WIs.
- **Background Scraping:** Dispatches polite concurrent requests (`fetch_wi_details_from_portal`) and atomically persists newly discovered WIs with their metadata, responsible groups, and remarks (`WorkItemsDatabase.upsert_single_work_item()`).
- **Audit Marker Caching:** Writes `<excel_path>.wi_checked.json` (recording timestamp, file mtime, discovered count, and scraped codes). On subsequent opens, `is_enrichment_needed` checks `mtime(marker) >= mtime(excel)` and exits in ~0 ms.
- **Dynamic Refresh & Config:** Quietly updates tracked tokens and filter badges if any newly scraped WI matches user watchlist tokens. Can be toggled on/off in `MeetingsConfigDialog`.

### 3.8. Meetings Browser Modern Architecture (`MeetingsTab`)
- **Full-Width Table & 2-Row Top Ribbon:** Eliminates the legacy fixed-width right sidebar, allocating 100% of horizontal space to the meetings table:
  - **Row 1 (Primary Navigation & Toolbar):** Houses `🚀 Last Meeting` quick launcher with active theme styling, inline `⚡ TDoc:` jump box with dynamic `📄 Doc` and `🗓️ Mtg` action buttons, right-aligned `🔄 Sync ▾` menu, `📊 Reports ▾` menu, and `⚙️ Settings`.
  - **Row 2 (Inline Filter Engine):** `🔍 Search` (by number, city, or title), TSG segmented buttons (`[All] [SA] [RAN] [CT]`), `Working Group` checkable multi-select combo, Nature (`Regular` / `Ad-Hoc`), Format (`In-Person` / `Electronic`), and Date Presets (`All Dates`, `Upcoming`, `This Year`, `Past 12 Months`, `Custom Range...`). Selecting *Custom Range...* dynamically reveals the calendar pickers.
- **7-Column Table Design & Delegates:**
  - `Action` (col 1): `TDocsButtonDelegate` painting interactive `✔ Open`, `⬇ Get`, or `⏳ Fetching` status pills with hover feedback.
  - `WG` (col 2): `WGTagDelegate` rendering color-coded TSG domain badges (Soft Indigo for SA, Soft Emerald for RAN, Soft Amber for CT) using `get_wg_colors()`.
  - `Location` (col 4): Icon-prefixed labels (`📍 City, Country` vs `🌐 Online (eMeeting)`).
  - `Dates` (col 5): English consolidated ranges (`12 – 16 Oct 2026`) via `format_meeting_dates()` with temporal awareness tooltips (`🟢 Ongoing`, `🔵 Upcoming`, `⚪ Past`).
  - `Allocated TDocs` (col 6): Preserves full WG prefixes (`C4-264009 → C4-264023`, `—` when unallocated).
  - Double-clicking any row directly launches the meeting TDocs window (or downloads it if uncached).
- **Floating Comparison Cart:** Compact lower-right status bar pill that **auto-hides completely when empty**, sliding into view only when items are staged into Slot A or Slot B (`⚖️ Cart: <Doc A> vs <Doc B> [Compare in Word] [Clear]`).
- **Safe Maintenance Grouping:** Relocated destructive `Clear All Meetings...` from the main UI into `MeetingsConfigDialog` (`⚙️ Settings`) under a dedicated *⚠️ Database Maintenance (Danger Zone)* section.
