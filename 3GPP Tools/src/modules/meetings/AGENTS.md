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
    ├── search_controller.py           # GlobalSearchController for inline Quick TDoc Jump & quick actions
    ├── tdocs_components.py            # CheckableComboBox and custom styled UI widgets
    ├── tdocs_dialogs.py               # Statistics parameters, export configuration dialogs
    ├── tdocs_menus.py                 # Right-click context menus & centralized Quick Copy menu (table & quick bar)
    ├── tdocs_models.py                # High-performance TDocsTableModel & TDocsFilterProxyModel
    ├── tdocs_window.py                # Primary TDoc list viewer window
    └── ui_tabs.py                     # MeetingsTab (full-width browser, 2-row top ribbon, floating cart)
```

---

## 3. Critical Architectural Rules

### 3.1. Standalone Core: `tdoc2md.py` & Upstream Backporting
> [!NOTE]
> `tdoc2md.py` is a drop-in standalone conversion engine maintained in an external upstream repository (`3gpp-tdoc-markdown-exporter`).
> - It relies strictly on the Python standard library (no third-party dependencies).
> - Any application adaptations, LibreOffice/Word COM bridges, custom callbacks, or GUI integration MUST be implemented in `tdoc_markdown_converter.py` or calling modules.
> - **Incremental All-New pCR Support (v2.2.0):** Features native detection of "ALL TEXT is NEW" change blocks that contain inline Word revisions relative to an earlier contribution (e.g. `* * * First Change ALL TEXT is NEW * * * *` with `<w:ins>`/`<w:del>` deltas wrt a postponed predecessor). Automatically suppresses `[ADDED BLOCK]:` on unmarked paragraphs (which represent the baseline starting point, not new blocks), parses unparenthesized boundary qualifiers, and annotates the dual-layer change scope with predecessor TDoc provenance. Maintained strictly backward-compatible for direct upstream backporting.
> - **Subscript, Superscript & Inline Formatting Preservation:** Preserves run-level OpenXML font styles (`<w:vertAlign>`, `<w:b>`, `<w:i>`, `<w:strike>`, `<w:u>`, `<w:rFonts>`, and Office Math `<m:oMath>`). Normative telecommunication keys (e.g. $K_{NF}$, $K_{SEAF}$, $K_{AMF}$) render as `K<sub>NF</sub>` in Markdown (`text`) and `K_NF` in clean plain text (`plain`), ensuring 100% round-trip fidelity with `markdown_to_docx.py` and seamless FTS5 / LLM triage consumption. Redundant heading bolding (`# **Heading**`) is automatically stripped while preserving inline subscripts. Maintained strictly stdlib-only and 100% backward-compatible.

### 3.2. Background Threading
- All FTP downloads, file extractions, Word COM automations, and LLM triage must run in `QThread` workers (e.g. `LLMExporterThread`, `TDocTriageWorker`, `MeetingsNetworkSync`).
- COM automation must initialize COM with `pythoncom.CoInitialize()` and clean up with `pythoncom.CoUninitialize()` in `finally` blocks.

### 3.3. TDoc Markdown Conversion & Provenance Flow
- Primary entry point for application code: `tdoc_markdown_converter.convert_document_to_markdown()`.
- Supports `.docx`, `.docm`, and legacy `.doc` (via sensitivity-aware Word COM / LibreOffice pre-conversion).
- Extracts track changes (`[INSERTED: ...]`, `[DELETED: ...]`), tables, figures, Visio diagrams, and metadata headers.
- **Document Type Revision Mode Resolution (`resolve_revision_mode_for_type`):**
  - **CR & pCR:** Normative specification change proposals retain specification revision markers (`revisions="markup"` with `[INSERTED]`/`[DELETED]` tags and markup legend).
  - **Incremental All-New pCRs:** Documents where a proposed clause is new to the target specification (`ALL TEXT is NEW`) but contains revisions relative to an earlier postponed/revised draft (e.g. `(revision of S2-2611386)`) suppress `[ADDED BLOCK]:` prefixes on unmarked paragraphs. The unmarked text represents the starting-point baseline, while `<w:ins>` / `<w:del>` render cleanly as the delta wrt the predecessor TDoc.
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
- **UI Dialog (`PostMeetingExportDialog`):** Modeless configuration dialog accessible from `TDocsWindow`'s Export menu (`📑 Export Post-Meeting Report (.md)...`). Designed with a responsive 2-column layout (Left: Scope selection and full-height WI/AI tree; Right: Document types and conversion options) to guarantee zero clipping on laptop screens. Allows selecting export scopes (`🎯 Followed Work Items`, `⭐ Starred Work Items`, `🎯+⭐ All Tracked`, `🌐 Whole Meeting`, `👁️ Currently Visible`, or `📋 Custom Selection (Work Items & Agenda Items)...`), contribution categories (`LS in`, `pCR`, `CR`, `LS out`, `Discussion / WID`, `General LS`), and conversion options (auto-open, force reconversion, local index generation).
  - **Dual-Mode Custom Selection Tree:** Replaces flat list selection with a synchronized dual-view tree switcher:
    - `🗂️ Work Items (WI ➔ AIs)`: Work Items as parents, expanding to show child Agenda Items with document count badges.
    - `📋 Agenda Items (AI ➔ WIs)`: Agenda Items as parents with natural numerical sorting (e.g. `5.1` before `5.10`), expanding to show child Work Items.
    - **Automatic Pre-Selection:** Automatically pre-checks followed and starred Work Items (and their child Agenda Items) upon opening, requiring 0 manual clicks to start fine-tuning.
    - **Synchronized Tri-State Selection:** Both views share a single atomic selection model (`(wi, ai)` pairs); checking or unchecking in either view updates the other view seamlessly.
    - **Inline Search Filtering:** Instant multi-token search filtering across Work Item names, AI numbers, and topic descriptions.
    - **Quick Actions:** Toolbar buttons for `[⊞ Expand]`, `[⊟ Collapse]`, `[☑ All]`, and `[☐ None]`.
  - **Live "Estimated Scope" Banner & Dynamic Badges:** Real-time KPI banner (`📊 Estimated Scope: X TDocs across Y Work Items and Z Agenda Items (A CRs, B pCRs, C LSs, D Disc)`) and live document count badges on all Section 2 checkboxes updated on every toggle.
- **Exporter Worker (`PostMeetingExporterThread`):**
  - **Structure:** Hierarchical dossier tree under `Export/PostMeeting_<WG>_<Meeting>_<timestamp>/` arranged chronologically by Release (`Rel-XX/`) and then by Work Item (`<WI>/`), containing type subfolders (`LS_IN/`, `pCR/`, `CR/`, `LS_OUT/`, and `Discussion/`).
  - **Contribution Filtering:** Strictly captures received incoming liaison statements (`LS in`), approved/agreed normative outputs (`pCR`, `CR`, `LS out`), and optionally noted/agreed discussion papers & work item status reports (`Disc`, `WID`) when `include_disc` is enabled. Filters out unapproved, postponed, or withdrawn contributions.
  - **Fine-Grained WI + AI Scoping:** When custom WI/AI scoping is active, filters contributions to only export TDocs matching the selected Work Item and allowed Agenda Items (e.g. exporting only AI 5.1 architecture for `TEI18` while excluding AI 6.2 protocols).
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
  - **Row 1 (Primary Navigation & Toolbar):** Houses `🚀 Last Meeting` quick launcher with active theme styling, inline `⚡ TDoc:` jump box with dynamic `📄 Doc`, `🗓️ Mtg`, and `⚖️ vs Word` action buttons, right-aligned `🔄 Sync ▾` menu, `📊 Reports ▾` menu, and `⚙️ Settings`.
  - **Row 2 (Inline Filter Engine):** `🔍 Search` (by number, city, or title), TSG segmented buttons (`[All] [SA] [RAN] [CT]`), `Working Group` checkable multi-select combo, Nature (`Regular` / `Ad-Hoc`), Format (`In-Person` / `Electronic`), and Date Presets (`All Dates`, `Upcoming`, `This Year`, `Past 12 Months`, `Custom Range...`). Selecting *Custom Range...* dynamically reveals the calendar pickers.
- **7-Column Table Design & Delegates:**
  - `Action` (col 1): `TDocsButtonDelegate` painting interactive `✔ Open`, `⬇ Get`, or `⏳ Fetching` status pills with hover feedback.
  - `WG` (col 2): `WGTagDelegate` rendering color-coded TSG domain badges (Soft Indigo for SA, Soft Emerald for RAN, Soft Amber for CT) using `get_wg_colors()`.
  - `Location` (col 4): Icon-prefixed labels (`📍 City, Country` vs `🌐 Online (eMeeting)`).
  - `Dates` (col 5): English consolidated ranges (`12 – 16 Oct 2026`) via `format_meeting_dates()` with temporal awareness tooltips (`🟢 Ongoing`, `🔵 Upcoming`, `⚪ Past`).
  - `Allocated TDocs` (col 6): Preserves full WG prefixes (`C4-264009 → C4-264023`, `—` when unallocated).
  - Double-clicking any row directly launches the meeting TDocs window (or downloads it if uncached).
- **Floating Comparison Cart & Active Word Comparison (`ComparisonCartWidget`, `ComparisonManager`):** Shared lower-right status bar widget across `MeetingsTab` and `TDocsWindow` that collapses to a compact `⚖️ Stage Active Word ▾` launcher when empty and expands when Slot A or Slot B is populated (`⚖️ Cart: [Slot A ▾] ⇄ [Slot B ▾] [🪟 Active Word] [Compare in Word] [Clear]`). Supports 1-click staging of the foreground Active Word document, any open Word document, TDocs, or local files into either slot, 1-click slot direction swapping (`⇄`), automatic snapshot refresh (`resolve_comparison_paths()`) for unsaved/edited open Word drafts, and direct 1-click **Active Word vs. TDoc** comparison from the Quick Jump bar, `TDocsWindow` Instant Fetch bar, row context menu, action button menu, and Related TDoc hyperlinks.
- **Safe Maintenance Grouping:** Relocated destructive `Clear All Meetings...` from the main UI into `MeetingsConfigDialog` (`⚙️ Settings`) under a dedicated *⚠️ Database Maintenance (Danger Zone)* section.

### 3.9. High-Speed Parallel Pre-Download Pipeline for Markdown Exports (`llm_exporter.py`, `post_meeting_exporter.py`, `tdoc_file_handler.py`)
- **Decoupled Two-Phase Processing:** Both visible-table Markdown exports (`TDocMarkdownExporterThread` in `llm_exporter.py`) and post-meeting report dossiers (`PostMeetingExporterThread` in `post_meeting_exporter.py`) employ a two-phase architecture to eliminate serial network latency when exporting unindexed or uncached TDocs:
  - **Phase 1 (Concurrent Pre-Download):** Sub-millisecond pre-scan identifies unindexed or missing source documents (`.docx`/`.doc`) across the export scope, skipping already-cached and fresh dossiers. Dispatches missing archives to a polite worker pool (`ThreadPoolExecutor(max_workers=5..6)`). Downloads stream concurrently, cutting total elapsed network time by 5x–6x (~60s down to ~10s for 40 contributions) while staying within 3GPP server rate-limits.
  - **Phase 2 (In-Memory Markdown Conversion):** Sequentially converts the now locally-cached documents via direct OpenXML parsing without intermittent network pauses. Guarantees 100% single-threaded COM isolation for Word fallbacks (`win32com.client`) and prevents system-freezing multi-process churn from concurrent PowerShell vector metafile conversions.
- **Atomic File Integrity & Self-Healing Cache (`TDocFileHandler`):**
  - Downloads stream to temporary `.zip.part` files prior to atomic `os.replace` renaming, guaranteeing that interrupted network transfers or user cancellations never leave corrupt, 0-byte ZIP files on disk.
  - Automatically verifies local ZIP integrity with `zipfile.is_zipfile`; corrupted or truncated archives are automatically purged and re-fetched fresh.
  - Extracted `.docx`/`.doc` documents are similarly written via atomic `.part` rename.
  - Bidirectional route fallback: Automatically tries primary and secondary routes (`revisions_url` and `docs_ftp_url`) with a generous 15s timeout and non-blocking `cancel_check()` responsiveness.

### 3.10. Chairman's Notes & Session Document Ingestion (F-19)
- **Multi-Surface Input:**
  - **Drag & Drop (`TDocsWindow.dropEvent`):** Dragging `.docx`, `.doc`, `.htm`, or `.html` session documents onto `TDocsWindow` renders an interactive frosted blue drop overlay (`TDocsDropOverlay`) with dashed drop borders.
  - **Menu Action:** **🔄 Refresh ▾ $\rightarrow$ 📝 Import Word Document (.docx / .doc)...** provides a manual file chooser.
  - **Downloader Bridge (`ChairNotesDownloaderThread`):** Clicking **📥 Notes** downloads all active notes from `/Inbox/Chair_Notes` into `Agenda/ChairNotes/` and prompts the delegate with 1-click parsing and merging.
- **Background Conversion & Staging Worker (`WordAgendaImporterThread`):**
  - Copies documents safely to `{meeting_dir}/Agenda/ChairNotes/`.
  - Strips NTFS `Zone.Identifier` stream metadata via `_unblock_file`.
  - Headless LibreOffice sanitizes legacy Word 97–2003 `.doc` files (which often embed VBA macros) to `.docx` without executing macros or freezing the GUI via Word COM.
- **Dual-Mode Word Document Parser (`TDocsParser.parse_tdocs_from_docx`):**
  - **Mode 1: Table-Based Parser (SA2, CT):**
    - **Multi-Row Header Scanning (Rows 0–4):** Robustly locates table headers even when session chairs place meeting banners, track titles, or date/time headers on rows 0–1.
    - **Comprehensive Column Synonyms:** Resolves heterogeneous Working Group secretary labels across SA2, SA1/3, and CT (`comments`/`discussion`/`remarks`, `result`/`status`/`decision`/`outcome`, `td#`/`tdoc`/`doc#`, `source`/`company`/`proposer`, `ai`/`agenda item`).
    - **Section Banners & TDoc Tokens:** Handles merged full-row Agenda Item section banners lacking TDoc IDs and tolerates decorated TDoc numbers (e.g. `[C1-2409999]`, `R2-2405678r01 (rev of ...)`).
    - **Sequential Deliberation Merging:** Combines sequential discussion notes (`\n---\n`) for documents revisited across morning/afternoon sessions.
  - **Mode 2: Tab-Delimited Paragraph Parser (RAN1, RAN2):**
    - Automatically falls back when documents contain no tables (RAN1 and RAN2 chairmen write notes as tab-separated contribution lines `<TDoc>\t<Title>\t<Source>\t<Type>` followed by agreement/decision and discussion paragraphs).
    - Maps headings (`Heading 1`, `Heading 2`) to `Agenda Item`, extracts decisions/statuses (`Approved`, `Agreed`, `Noted`, `Withdrawn`), and captures multi-paragraph discussion notes per TDoc.
  - **FTP Presence Reality:** Verified across live 3GPP FTP servers that `/INBOX/Chair_Notes/` is primarily maintained by SA2, RAN1, and RAN2; other groups (SA1, SA3, CT1, CT4, RAN3) rely instead on email reflector rulings (Feature F-15: `<<START>>...<<END>>`), server-side `agenda.csv` files, or formal end-of-meeting summary TDocs in `/Docs/`.
- **Live Model Synchronization (`TDocsTableModel.merge_agenda_data`):**
  - Populates in-room discussions directly into the `Secretary Remarks` column (with clickable TDoc hyperlinks).
  - Updates real-time `TDoc Status` (`Agreed`, `Revised to...`, `Postponed`) hours before official agenda publication.
  - Dynamically injects newly announced late revisions with predecessor linking (`Is revision of`, base TDoc resolution).

### 3.11. SA2 Meeting Session Plan Management, Version Diffing & Mobile Export
- **Problem Solved:** SA2 meeting session plans ("Draft time allocation") are published as living TDocs with frequent revisions (often reusing the same TDoc number `S2-2610195` across intra-day updates), 3 parallel streams, 45-minute half-quarter Time Units (TUs), and ad-hoc drafting sessions (08:00 morning, lunch, and 18:00 evening).
- **Timezone & Corporate Laptop Synchronization (`SessionPlanTimezone`):**
  - Solves the Austrian corporate laptop clock lock (`Europe/Vienna`, CET `UTC+1` / CEST `UTC+2`) by strictly anchoring internal timestamps in true UTC.
  - 3-tier meeting timezone resolution: (1) Auto-parsing from Agenda Deadlines table (`0900 Local time | 0700 UTC` $\rightarrow$ `UTC+02:00`); (2) City/Country lookup from meeting location; (3) Interactive `SessionPlanTimezoneDialog` confirmation prompt.
  - Dual-time display in UI and tooltips (`09:00 Venue (02:00 Vienna)`).
  - Live clock strip in toolbar updating Venue time and laptop time every second.
- **Interactive Session Plan Studio & Ergonomics (`SessionPlanWindow`):**
  - **10-Column Graphical Half-Quarter Layout:** Organizes the 5-day week into 10 sub-columns (2 sub-columns per day: `TU 1` and `TU 2`) using `QHeaderView.Stretch`. Full-quarter sessions span both sub-columns (`setSpan(row, day*2, 1, 2)`), while half-quarters occupy 50% left or right sub-columns, communicating timing geometrically and cleanly eliminating inline timestamp text (`[TU 1 (09:00–09:45)]`) from cell interiors.
  - **Separated Stream Rows & Break Periods (23-Row Grid):** Replaces single-cell multi-stream mixing with dedicated rows for each stream (`Str 1`, `Str 2`, `Str 3`) across Q0 through Q5, interleaved with 5 break rows (`Break 1`, `Coffee 1`, `Lunch`, `Coffee 2`, `Break 2`). Accurately captures in-break drafting (e.g., 13:00 Main Room drafting) and plenary sessions.
  - **Interactive Stream & Break Toggles:** Control strip features checkable toggle buttons for `Str 1` (soft indigo `#4338CA`), `Str 2` (soft emerald `#047857`), `Str 3` (soft amber `#B45309`), `Breaks`, and `Hide Empty Breaks`. Toggling filters the 23-row table in $<1$ ms via `isRowHidden()`.
  - **Rich HTML Hover Tooltips (`_build_rich_slot_tooltip`):** Clean cell text displays only topics and AI titles, while hovering reveals full timing (both Venue local and laptop Vienna times), assigned physical room, convener, TOHRU audio/video channel, resolved Agenda Item titles from meeting agenda data, user notes, and parallel conflict alerts.
  - **Calendar Dates in Column Headers:** Horizontal headers display day names, physical dates (`Monday (12 Oct)`), and live venue `"🟢 Today"` badge.
  - **Quick Import Drag & Drop:** Direct drag & drop of `.docx`, `.doc`, or `.zip` files onto the `SessionPlanWindow` displays an instant animated `DropOverlayWidget` and automatically ingests, stages, and switches to the new snapshot.
  - **Interactive Topic Tracking & My Topics Engine (`user_tracked_topics`):**
    - Combines followed/starred Work Item tokens (`followed_wi_tokens | starred_wi_tokens`) with per-meeting persistent tracked topics in `session_plans.db`.
    - **⭐ Track Topics Dialog (`SessionPlanTrackedTopicsDialog`):** Interactive checklist dialog with instant search, 1-click `"⭐ Track Followed WIs"` auto-matching, and select all / clear all shortcuts.
    - **1-Click Tracking in Agenda Mode:** Agenda Mode displays a dedicated `⭐` column with 1-click toggling between tracked and untracked topics.
    - **Overview Grid Context Menu:** Right-click any cell to track/untrack specific extracted topics or edit personal notes.
    - **Parallel Conflict Detection:** Automatically flags when two tracked topics are scheduled in parallel rooms during the same quarter (`⚠️ CONFLICT with Str 2`), highlighting rows in soft amber (`#FEF3C7`) and warning in the tab button.
    - **Bidirectional Double-Click Navigation:** Double-clicking any session in "My Topics" or "Drafting Radar" immediately switches to the 5-day Overview Grid, focusing and selecting that exact cell.
  - **Comprehensive Tooltips:** Rich contextual tooltips on all buttons, tabs, room guide strips, table headers, and status badges.
  - **Window Lifecycle & Application Shutdown:** Implements explicit `closeEvent` and `showEvent` handlers managing `self.clock_timer` and auto-dismissing child modal dialogs (`findChildren(QDialog)`). Synchronized with `TDocsWindow.closeEvent`, `MeetingsTab.shutdown()`, and `DragDropUI.closeEvent` (`QApplication.topLevelWidgets()` sweep) to guarantee all session plan windows terminate cleanly upon closing the parent meeting or the main application.
- **100% Manual Ingestion & Collision-Safe Caching (`SessionPlanParser`):**
  - Staged files cached in `<meeting_dir>/Agenda/SessionPlans/` (never overwrites earlier snapshots sharing the same TDoc ID, using content-hash naming `S2-2610195_snap2_...docx`).
  - Persistent SQLite database at `<meeting_dir>/Agenda/session_plans.db`.
  - Ingestion entry points:
    1. **Toolbar Button:** `🗓️ Session Plan` in `TDocsWindow` launches the viewer.
    2. **URL or TDoc#:** `SessionPlanImportDialog` downloads via `URLRouter` candidate routes.
    3. **File Browser:** OpenXML `.docx`, `.doc` (via LibreOffice), `.zip`, or HTML.
    4. **Drag & Drop:** Dropping files onto `TDocsWindow` automatically inspects whether the table is a Session Plan (`Monday..Friday`, `Q0..Q5`, `Stream 1..3`) and routes it to `Agenda/SessionPlans/`.
    5. **Row Context Menu:** Right-clicking any session plan TDoc row offers **`🗓️ Load {tdoc_id} into Session Plan...`**.
- **Parsing Heuristics & Data Modeling (`SessionPlanParser`, `SessionPlanDatabase`):**
  - Sub-slot detection: Preserves `TU 1` (first 45 mins), `TU 2` (second 45 mins), or full quarter (`colspan=2`).
  - Stream topology: Extracts Main Room (`B2 [300]`), Breakout 1 (`B3 [100]`), Breakout 2 (`D11/D12 [120]`), Conveners (*Mike*, *Wanqiang*), and TOHRU channels (`SA2_Main`, `SA2_Brk1`, `SA2_Brk2`).
  - Lunch drafting: Captures ad-hoc lunchtime drafting sessions in Main Room (e.g. 13:00 SoH questions).
  - AI range expansion: Automatically expands shorthand ranges like `KI#12 – KI#16 (20.6.12 - 16) [5,0,2,0,1]` and `Pre-Rel-19 maintenance (5.x, 6.x, 7.x) [9]`.
  - Typo resilience: Handles chair typos like `K#1.2` and unclosed parentheses.
- **Version Diffing (`SessionPlanDiffEngine`, `SessionPlanDiffDialog`):**
  - Compares any two plan snapshots slot-by-slot, detecting **🟢 Added**, **🔴 Removed**, **🟡 Rescheduled/Moved** (from origin to destination), and **🔵 Volume Changes** (`[15]` $\rightarrow$ `[5+15]`).
  - Synthesizes an executive bulleted changelog.
- **Persistent User Annotations (`user_session_annotations`):**
  - User notes, priority tags (`⭐ Must Attend`, `👥 Colleague Covering`, `👀 Monitor on TOHRU`, `✅ Done`), and comments are keyed by `(meeting, day, quarter, sub_slot, stream)`, persisting seamlessly across session plan updates.
- **Multi-Format Export Engine (`SessionPlanExporter`):**
  - **📱 Dual-View Interactive Offline HTML (`SA2_{mtg}_Session_Plan_Mobile.html`):** Standalone, 100% offline responsive single-file web application (<95 KB, zero external CDN dependencies) built for easy sharing across the delegation:
    - **Dual Presentation Views:** Segmented switcher allows 1-click toggling between **🗓️ 5-Day Spatial Overview Grid** (10-column layout mirroring desktop with half-quarters, dedicated stream rows, break rows, and stream toggles) and **📱 Mobile Day Cards** (touch-friendly day-by-day quarter view with sticky day tabs). Automatically defaults to Day Cards on mobile screens ($\le 768$px).
    - **Interactive Delegate Filter Strip:** Egalitarian team buttons (`[👥 All Team]`, `[● Anna]`, `[● Thomas]`, ...) styled with solid pastel tints. Toggling a delegate dims non-assigned cells in the Overview Grid and filters cards in Mobile Day Cards.
    - **Client-Side RFC 5545 `.ics` Export:** Single-click `[📅 Download Team .ics]` or `[📅 Anna .ics]` dynamically compiles and downloads UTC-anchored `.ics` calendar files directly in the browser via Blob URL, strictly scoped to the active delegate or full team without server round-trips.
    - **Parallel Stream Clash Warning:** Real-time client-side conflict detector scans active assignments; displays an alert badge (`⚠️ 2 Parallel Clashes`) and highlights conflicting cells with dashed red outlines.
    - **Live Venue Status Indicator:** Computes real-time venue wall-clock time from UTC and meeting timezone offset; displays an animated pulsing green banner (`🟢 LIVE NOW: Str 1 - Plenary | Next: Lunch at 13:00`) and highlights active session cells.
    - **Rich Floating Tooltips:** Hovering over any slot displays venue/laptop times, physical room, convener, TOHRU audio/video channel, resolved Agenda Item descriptions from `agenda.csv`, and personal notes.
  - **📅 Strict UTC iCalendar (`.ics`):** Exports sessions to Outlook / Apple / Google Calendar with strict UTC `DTSTART`/`DTEND` timestamps, room/floor in Location, and convener/topics/notes in Description.
  - **✉️ Colleague Email Composer:** Formats rich HTML & plain text summaries ready to paste into Outlook summarizing what changed since the previous plan, today's drafting sessions, and tracked topics.
  - Enriches TDoc rows with live `Sub-Item` topic separators, `_agenda_order`, and `_sub_item_order` extracted from `TdocsByAgenda.htm` and Word Chairman's Notes.
- **Delegation Topic Distribution & Session Clash Dispatching (`SessionPlanWindow`, `SessionPlanDatabase`, `SessionPlanExporter`):**
  - **Egalitarian Roster Model (No "Myself" Tag):** Replaced the restrictive "myself" / `(Me)` tag in delegation roster, dialogs, and pickers with an egalitarian architecture where all team members are treated equally. Every colleague knows who they are; removing the tag ensures exported HTML and schedules are clean, universal, and ready to share across the entire delegation. Auto-assigning followed WIs presents an intuitive delegate selection prompt.
  - **Two-Tier Assignment Model:** Decouples high-level strategic **Topic Ownership** (e.g. *Martin owns AI 20.6.18 and AI 5.2.1*) from tactical **Session Appointments** (ad-hoc single-slot overrides when clashes occur).
  - **Shared Boundary-Aware Resolver (`resolve_slot_effective_delegates`, `ai_number_in_text` in `session_plan_models.py`):** Single source of truth used by `SessionPlanWindow` and `SessionPlanExporter`. Raw-text fallback uses regex negative lookbehind/lookahead (`(?<![\d.])` and `(?!\d|\.\d)`) so short AI tokens like `20.6.2` or `20.6.1` never falsely match `20.6.24`, `20.6.22`, or `20.6.17`.
  - **Multi-Delegate & "All Team" Ownership (`DelegateOwnerPicker`):** Inline checkbox-based multi-select dropdown (`QToolButton` + `QWidgetAction` checkboxes) in the Topic Ownership Matrix allowing delegates to tick one or more owners, `👥 All Team`, or `✕ Clear (Unassigned)` in a single popup without modal dialog round-trips, committing once when the popup closes.
  - **Direct Overview Grid Topic Owner Assignment:** Right-clicking any cell in the desktop 5-day Overview Grid (`_on_grid_context_menu`) provides a dedicated `👤 Assign Topic Owner (AI X.Y.Z)...` submenu enabling instant 1-click assignment to `👥 All Team`, any individual delegate, `Custom Multi-Delegate`, or `✕ Clear (Unassigned)` directly from the grid without having to open the Topic Ownership Matrix.
  - **Per-Delegate Collision Evaluation:** When a general topic (like 20.6.0) or multi-delegate topic is scheduled, collisions are evaluated per individual delegate: only colleagues who are simultaneously double-booked in a parallel stream receive a clash warning, while unconflicted delegates remain marked as available and attending.
  - **Half-Quarter (TU1 / TU2) Time-Overlap Precision (`_slots_time_overlap`):**
    - Accurately models 45-minute half-quarters (`sub_slot=1` for TU1, `sub_slot=2` for TU2, and `sub_slot=0` for full quarters).
    - Sequential half-quarters (e.g., Stream 1 TU1 $\rightarrow$ Stream 2 TU2) are recognized as non-overlapping in time, completely eliminating false double-booking warnings.
    - True simultaneous collisions (e.g., Stream 1 TU1 and Stream 2 TU1, or full Q2 vs TU1/TU2) are reliably caught.
  - **Dynamic "Who is Free?" Availability Engine (`_compute_delegate_availability`):**
    - When a delegate is double-booked across parallel rooms (e.g., Stream 1 and Stream 2 at Tuesday Q2 TU1), calculates real-time room occupancy for every team member.
    - Accurately marks colleagues 🟢 *Free* for TU1 if their only assignment in that quarter is in TU2 (or another time).
    - 1-click **`⚡ Appoint Available Delegate...`** opens `SessionDispatchDialog`, appointing a free colleague for that specific session without altering master topic ownership.
  - **Multi-Stream Clash Board (`⚡ Clash & Dispatch Board`):**
    - Handles complex scenarios where $\ge 2$ parallel rooms collide for a delegate (e.g., Streams 1, 2, and 3).
    - Combines descriptions of all conflicting streams and equips the appointment action with a stream-selection popup menu (`btn_appoint.setMenu(menu)`), allowing 1-click selection of which exact room/stream to cover.
  - **Real-Time Topic Ownership Matrix Synchronization:**
    - Live status cell updating (`_update_team_topic_status_cells`) immediately reflects changes in topic coverage and collision warnings without recreating widgets or losing scroll position.
  - **Direct Grid & Agenda Context Menu Dispatching:**
    - Right-clicking any slot in either the Overview Grid (`_on_grid_context_menu`) or the Delegate Agenda (`_on_tracked_context_menu`) provides instant actions:
      - `⚡ Appoint/Change Coverage for this Session...` (pre-filled with current assignment)
      - `↺ Revert Coverage to Topic Owner` (restores default topic ownership in 1 click via `_on_revert_dispatch`).
  - **Multi-Delegate Color-Coded Heatmap & Focus Filtering:**
    - Auto-assigns distinct, accessible palette colors (`DELEGATE_PALETTE`) to each team member. Features 10 harmonious, high-contrast hues (WCAG AA compliant $\ge 4.5:1$ against white): Royal Blue (`#2563EB`), Emerald Forest (`#047857`), Deep Violet (`#7C3AED`), Warm Terracotta (`#C2410C`), Cyan Ocean (`#0E7490`), Ruby Rose (`#BE123C`), Indigo Night (`#4338CA`), Plum Orchid (`#A21CAF`), Warm Bronze (`#B45309`), and Slate Graphite (`#475569`).
    - **Solid Pastel Tint Architecture (`DelegationMember.compute_tint`):** Computes solid `#RRGGBB` pastel tints (`bg_tint` 12% and `border_tint` 35% blended with `#FFFFFF`) for delegate pills, chips, active highlight strips, tooltips, and HTML email summaries. **Strict prohibition of 8-digit hex codes (`#RRGGBBAA`)**: Qt QSS parses 8-character hex strings as `#AARRGGBB` (reversing alpha and RGB channels, turning `#2563EB1F` into dark sludge green `#1a3f07`), while Microsoft Outlook email engines ignore 8-digit hex completely. Blending directly on white produces 100% solid, predictable colors across Qt on Windows, Web, mobile, and Outlook.
    - **Compact Initials Pill Badges (`_make_delegate_badge`, `compute_unique_initials`):** Overview Grid renders space-saving circular/pill initials badges (`JO`, `DI`, `PI`, with automatic 3-letter disambiguation if two colleagues share 2-letter initials) instead of repeating full names in every cell. Solid border indicates topic ownership, dashed border indicates ad-hoc coverage override, and a 2px red border ring (`#DC2626`) + compact `⚠` pill highlights parallel-stream clashes. Full names, roles, and clashing stream details appear in rich hover tooltips.
    - **Multi-Select Delegate Focus Bar:** Top ribbon focus bar (`[🌈 All Team] [👤 Only Me]` + checkable `JO · Josep`, `DI · Dieter`, ...) supports multi-select checkbox filtering so users can highlight any combination of delegates simultaneously while dimming unrelated sessions.
  - **Delegate Agenda View (`🎯 Delegate Agenda`):**
    - Segmented switcher (`[👤 Myself ▾ | ● Anna | ● Thomas | 👥 Whole Delegation]`) renders personal timetables with dual venue/laptop times, rooms, roles, and conflict tags.
    - Direct double-click jumps to the cell on the Overview Grid; right-click allows note editing or coverage dispatching.
  - **Personalized Delegate iCalendar (.ics) Calendar Export:**
    - Main toolbar Export menu provides `📅 Export Delegate .ics (Personal / Colleague)...` to export calendar events strictly scoped to an individual delegate's assigned sessions (both owned topics and appointed coverage), strictly UTC-anchored for Outlook/Google Calendar/iPhone.
    - Full meeting `.ics` exports automatically include delegate assignment tags (`[Delegate]` / `[Delegate (cov)]`) in event summaries and descriptions.
  - **Email & Machine-Readable Interchange (Frictionless Sharing):**
    - **`📧 Export Email Summary...`** (`generate_delegation_email_summary`): Generates pre-formatted executive HTML/plain text for Outlook with color badges, topic ownership tables, ad-hoc coverage notes, and personal daily schedules. Accessible via top Export menu and Team Allocations action bar.
    - **Embedded Interchange Payload:** Automatically embeds `<!-- 3GPP-DELEGATION-PLAN: {...} -->` into the email or exports `.json`.
    - **1-Click Colleague Ingestion (`DelegationImportDialog`):** Colleagues paste the email or load the `.json` file in their 3GPP Tools, select *"Which delegate are you?"*, and instantly personalize their local Overview Grid and personal agenda.
    - **Persistence:** Roster, topic ownerships, and session appointments persist in `session_plans.db` (`delegation_roster`, `topic_owners`, `session_appointments`), surviving across intra-meeting session plan snapshot revisions (`Snap 1` $\rightarrow$ `Snap 2` $\rightarrow$ `Snap 3`).
- **Agenda & Sub-Item Ordering Engine (`tdocs_parser.py`, `tdocs_models.py`, `tdoc_delegates.py`, `tdocs_window.py`):**
  - **Live HTML Superseding Precedence:** In Working Groups like SA2 where delegates rely on session agenda order, the live order in `TdocsByAgenda.htm` strictly supersedes the static Excel list order (`"TDoc sort order within agenda item"`), as Excel lists omit late revisions or late arrivals submitted during the meeting.
  - **3-Tier Precedence Hierarchy:**
    1. *Tier 1:* Live HTML order (`_agenda_order` / `_sub_item_order`) from `TdocsByAgenda.htm` or Word Chairman's Notes.
    2. *Tier 2:* Static Excel `"TDoc sort order within agenda item"` parsed as float.
    3. *Tier 3:* Revision anchor inheritance (`base_order + rev_offset`) ensuring unindexed revisions immediately succeed their predecessor.
    4. *Fallback:* Alphanumeric TDoc identifier order.
  - **Precomputed $O(1)$ Sort Cache:** `_precompute_agenda_sort_tuples` evaluates composite keys `(is_unassigned, ai_key, tier, sub_item_order, order_val, rev_offset, tdoc_key)` during data load and merge, guaranteeing `< 1ms` instantaneous proxy model sorting across 1,800+ meeting rows.
  - **Option C Visual Delegate (`AgendaItemDelegate`):** Compact presentation rendering the Agenda Item number centered on top and a subtle rounded pill badge (`#F1F5F9`, border `#CBD5E1`, text `#334155`) for the `Sub-Item` underneath within the 48px row height, preserving the standard 14-column layout without consuming extra horizontal space.
  - **Dynamic Sub-Item Filtering & Search Integration:** Auto-reveals a `Sub-Item` `CheckableComboBox` in the filter bar whenever sub-items are present in the meeting data, participates in zero-intersection facet evaluation, and incorporates sub-item tokens into global search text filtering.
  - **Two-Way Order Synchronization:** An `Order By ▾` dropdown combobox (`📋 Agenda Order`, `🔢 TDoc Number`, `🏢 Company`, `✉️ Most Emails`, `🗂️ Default / File Order`) defaults to `📋 Agenda Order` for SA2 meetings and maintains bi-directional synchronization with table column header clicks.

### 3.12. TDoc List Excel Parsing & Empty-Export Handling (`tdocs_parser.py`, `ui_tabs.py`)
- **Problem Solved:** Portal TDoc list exports downloaded before a meeting's TDocs are allocated (e.g. SA2#178 Calgary) contain only the header row. Previously the parser cached `[]` in the `<xlsx>.json` sidecar, which was then returned forever, and the UI showed a generic *"Could not read data from the Excel file"* dialog without any log entry.
- **Rules:**
  - `TDocsParser.parse_tdocs_excel` **never persists an empty result** and **ignores/removes an empty `.json` cache**, so a re-downloaded list is always re-parsed.
  - Empty exports log a `WARNING` (header found, no TDoc rows); exceptions are logged with full traceback via `logger.exception`; successful parses log the TDoc count at `INFO`.
  - `MeetingsTab._open_tdocs_window` logs missing files / empty data and offers a 1-click **re-download** from the 3GPP portal (`_download_and_open_tdocs`) instead of a dead-end error.

### 3.13. SA2 Session Plan Parsing Engine, Conventions & Robustness (`session_plan_parser.py`, `session_plan_models.py`, `session_plan_db.py`, `session_plan_window.py`)
- **Overview:** 3GPP SA2 session plans ("Draft time allocation") are published as Microsoft Word documents (`.docx`) or circulated as Markdown/HTML tables. Chairs write these tables with significant stylistic flexibility, irregular formatting, and shorthand notations. The parsing engine converts these unstructured documents into high-fidelity structured models (`SessionPlanSnapshot`, `SessionSlot`, `AgendaSlotItem`) and persists them in SQLite (`session_plans.db`).
- **Core Real-World Parsing Quirks & Architectural Solutions:**
  1. **Physical Room Overrides (`In main room`, `In Breakout 1`, `In Breakout 2`):**
     - *Quirk:* Working Group chairs periodically reassign specific session slots to different physical rooms (e.g., shifting Stream 1 into Breakout 1, or Stream 2 into the Main Room) by writing text directly inside the cell (e.g. `In Breakout 2\nKI#8 (Net Sharing) (20.6.8)`).
     - *Fix:* Regex `ROOM_OVERRIDE_RE` isolates and normalizes the physical room, stores it in `slot.room_override` (persisted in SQLite via dynamic migration `ALTER TABLE session_slots ADD COLUMN room_override TEXT`), and cleans the string from topic tokens. All UI surfaces (Overview Grid tooltips, Drafting Radar, Delegate Agenda, Clash Board) and export formats (HTML, Outlook briefing emails, iCalendar) dynamically resolve `s.room_override or stream_room_name`.
  2. **Disambiguating "Breakout" Rooms from Coffee Breaks:**
     - *Quirk:* Room header rows (e.g., `| Main Room | Palomino | Breakout 1 | ...`) contain the word "Breakout". A naive `"break" in cell.lower()` test matched "Breakout 1", turning room header rows into fake coffee breaks across Monday–Wednesday and shifting all session timings.
     - *Fix:* `_detect_break_period` enforces strict word boundary matching `re.search(r"\bbreak\b", ...)` and evaluates active quarter context, ensuring room header rows are never misidentified as breaks.
  3. **OpenXML Vertically Merged Cells (`<w:vMerge>`) Across Parallel Streams:**
     - *Quirk:* Plenary sessions (e.g., Monday morning Q1/Q2, Friday afternoon Q4/Q5) or single-stream sessions span across all 3 stream rows using OpenXML vertical cell merges. In `python-docx`, iterating across row cells on vertically merged tables repeats the master cell's XML pointer (`_tc`) and text across Streams 1, 2, and 3.
     - *Fix:* Table parsing tracks underlying XML cell references (`seen_tcs = set()`). Vertically merged cell continuations across parallel streams are deduplicated, completely eliminating phantom duplicate sessions in Streams 2 and 3 and preventing false parallel-stream double-booking warnings for delegates.
  4. **Markdown / HTML Table `colspan` Row Expansion:**
     - *Quirk:* When session plan tables are pasted from email reflectors or Markdown notes, full-quarter sessions (spanning TU1 and TU2) frequently use `<td colspan=2/>`, resulting in fewer than 10 cells in a row (e.g., 8 or 9 cells). Naive slicing (`cells[:5]`) caused day-shifting (e.g., Tuesday TU2 shifting into Wednesday, dropping Thursday and Friday entirely).
     - *Fix:* `_process_day_cells` expands `<td colspan=2/>` and 5-cell day rows into 10 normalized half-quarter slots, maintaining exact Monday–Friday alignment.
  5. **Multi-Topic Slots & Numeric Range Expansion:**
     - *Quirk:* Cells frequently bundle multiple topics on separate lines (e.g., `In Breakout 2\nKI#10 (FWA) (20.6.10)\nKI#12 – KI#16 (20.6.12 - 16)`).
     - *Fix:* Each line is tokenized into its own specific `AgendaSlotItem`. `_expand_numeric_range` supports shorthand ranges (`20.6.12 - 16`), full prefix ranges (`20.6.12 - 20.6.16`), and implicit KI ranges (`KI#12 – KI#16`).
  6. **Dynamic Key Issue (`ki_map`) Cache for Implicit AI Resolution:**
     - *Quirk:* In SA2, study items (e.g. AI 20.6.x) define Key Issues (KI#1..KI#30). Early in the week, chairs write explicit references like `KI#8 (Net Sharing) (20.6.8)`. Later in the week, chairs omit the AI number, writing only `KI#8 (Net Sharing)` or `KI#21 (Data Framework)`.
     - *Fix:* The parser dynamically populates a meeting-wide `ki_map` cache during table traversal. Subsequent implicit KI# references automatically resolve to their full 20.6.x AI numbers.
  7. **Parenthesis Balancing & Topic Label Hygiene:**
     - *Quirk:* Stripping extracted AI numbers (e.g. removing `(20.6.8)` from `KI#8 (Net Sharing) (20.6.8)`) must not strip genuine topic brackets (e.g. `(Net Sharing)` or `(AI for Network)`), nor leave dangling closing brackets from multi-AI lists like `(5.x, 6.x )`.
     - *Fix:* `_balance_parentheses()` audits bracket parity, stripping unmatched outer delimiters while preserving valid technical qualifiers.
  8. **Timetable Administrative Note Filtering:**
     - *Quirk:* Chairs insert logistical scheduling notes into session slots (e.g., `Drafting 18:30-20:30`, `Session ends at 17:30`, `Revisions (tbd)`).
     - *Fix:* Regex `ADMIN_NOTE_RE` identifies logistical timetable annotations and suppresses them from technical Agenda Item lists.
  9. **Single-Digit Plenary Agenda Items:**
     - *Quirk:* Plenary cells frequently list `Opening (1), Agenda (2), Reports (3), Work plan (4), Outgoing LSs (5.1)`.
     - *Fix:* `single_ai_matches` extracts each numbered plenary item as an independent agenda item entity with its corresponding title.
- **UI Table Rendering Resolution Precedence (`_resolve_ai_description`):**
  - In `SessionPlanWindow`, topic titles in both Agenda Mode (`_render_agenda_mode`) and Team Topics (`_render_team_topics`) are resolved through `_resolve_ai_description(ai_num, scheduled_ais)`:
    1. Look up `ai_num` in `agenda_map` (from `agenda.csv` if available).
    2. Search scheduled slots for an exact `AgendaSlotItem` where `it.ai_number == ai_num` to retrieve its cleanly extracted `topic_label`.
    3. Fall back to the slot's cleaned non-administrative lines, omitting room override headers.
  - This guarantees that delegates viewing the Team Allocations matrix or Agenda Mode see the accurate topic description for each sub-topic, rather than the raw text of the entire multi-topic cell or a room string.
- **Multi-Source Drag & Drop Ingestion (Emails, Explorer, Archives & Text):**
  - Enables delegates to drag and drop session plan files directly onto `SessionPlanWindow` from any source, triggering identical processing to the local file browse import:
    1. *Direct Microsoft Outlook Drag & Drop:* Win32 Outlook does not provide local file paths on drag; it publishes virtual OLE streams (`FileGroupDescriptorW` / `FileContents`). `_extract_outlook_attachments` parses the binary descriptor header, decodes the Unicode filename (e.g. `S2-261nnnn_SA2-178-Session Plan.docx`), stages the raw bytes into `<meeting_dir>/Agenda/SessionPlans/Dropped/`, and triggers `_handle_manual_import("file", ...)`.
    2. *Local Files & Explorer:* Dragging `.docx`, `.doc`, `.docm`, `.zip`, `.htm`, `.html`, `.csv`, `.txt` files directly validates local file paths via `_is_valid_drop_data` and stages them with collision-free naming.
    3. *Automatic `.zip` Archive Decompression:* If a delegate drops a zipped 3GPP document archive (`.zip`), `_handle_manual_import` automatically unpacks the archive into `<meeting_dir>/Agenda/SessionPlans/Extracted/<stem>/` and targets the inner `.docx` or `.doc` file.
    4. *Email Body / Webmail Table Text:* Dragging selected Markdown or HTML table text from an email reflector message or TOHRU chat is detected by checking session plan keywords (`Stream 1`, `Main Room`, `Breakout 1`, `Q1`, `Monday`) and routed to `_handle_manual_import("paste", text)`.
    5. *Non-Blocking Visual Overlay (`DropOverlayWidget`):* Forwarding handlers (`dragEnterEvent`, `dragMoveEvent`, `dragLeaveEvent`, `dropEvent`) in `DropOverlayWidget` propagate events back to `SessionPlanWindow`, ensuring child widgets and transparent overlays never swallow or reject drops.
- **Team Allocation Ergonomics & Readability (Scheduled-Only Filter & Dynamic Sizing):**
  - **Scheduled-Only Filter Toggle (`btn_toggle_scheduled_only` / `chk_scheduled_only`):**
    - High-density Working Group agendas often declare hundreds of baseline Agenda Items (e.g. 222 items in SA2#178), while only a small subset (~40 items) are active and scheduled in parallel streams for the week.
    - Adds a checkable pill button in the Team Allocations header strip with live dynamic counts (`📅 Scheduled Only (40/222)` $\leftrightarrow$ `✓ 📅 Scheduled Only (40/222)`).
    - Toggling instantly hides unscheduled rows (`slot_count == 0`), eliminating the need to scroll through hundreds of inactive rows.
    - Seamlessly compounds with instant text search (`team_ai_search`): filtering by owner or text query operates exclusively on the visible subset.
    - Backed by $O(1)$ fast evaluation via `Qt.UserRole` integer storage on slot items (`item_slots.setData(Qt.UserRole, len(slots_for_ai))`), bypassing fragile string matching.
  - **Unclipped Table Headers & Optimized Column Geometry:**
    - Replaces clipped header text (`"Scheduled Slots in Plan"` $\rightarrow$ `"📅 Scheduled Slots"`, `"Lead Topic Owner"` $\rightarrow$ `"👤 Topic Owner"`, `"Coverage Status"` $\rightarrow$ `"🛡️ Coverage Status"`).
    - Configures explicit, content-calibrated column widths with interactive resizing: Col 0 (AI #): 65px; Col 1 (Topic Description): Stretch; Col 2 (📅 Scheduled Slots): 220px; Col 3 (👤 Topic Owner): 195px; Col 4 (🛡️ Coverage Status): 185px.
    - Expands vertical row height to 36px (`verticalHeader().setDefaultSectionSize(36)`) and injects internal cell padding (`QTableWidget::item { padding: 4px 8px; }`) to provide comfortable touch targets and prevent clipping of `DelegateOwnerPicker` dropdown chips and coverage badges.
    - Applies identical unclipped header styling, column widths, and 36px section sizing to the Clash & Dispatch Board (`["📅 Slot", "🕒 Time", "👤 Colliding Delegate", "Stream A (Assigned)", "Parallel Stream B", "🛡️ Action / Coverage"]`).


