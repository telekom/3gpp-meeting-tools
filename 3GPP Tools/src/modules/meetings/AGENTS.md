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
│   ├── exporters/                     # Multi-format export engines
│   │   ├── dossier_exporter.py        # Per-TDoc dossier exporter
│   │   ├── excel_exporter.py          # Formatted Excel (.xlsx) export
│   │   ├── llm_exporter.py            # Batch/single document extraction and export to Markdown
│   │   ├── markdown_exporter.py       # Meeting summary reports (agreed CRs, LSs, company contributions)
│   │   ├── post_meeting_exporter.py   # Post-meeting report generator (per-release & per-WI dossier)
│   │   ├── pre_meeting_exporter.py    # Pre-meeting briefing generator (active proposals, multi-reply LS, lineage dedup)
│   │   ├── tdoc2md.py                 # [PROTECTED] Standalone 3GPP TDoc to Markdown conversion core
│   │   ├── tdocs_package_exporter.py  # Packaged TDoc archive exporter
│   │   └── topic_dossier_exporter.py  # Cross-WG multi-meeting topic dossier & LLM delta feed exporter
│   ├── session_plan/                  # SA2 Session Plan parsing, diffing, DB & mobile HTML export
│   │   ├── session_plan_db.py         # SQLite persistence & cross-meeting default assignment sync
│   │   ├── session_plan_diff.py       # Slot-by-slot snapshot diff engine
│   │   ├── session_plan_exporter.py   # Interactive 4-view offline HTML, .ics & email exporter
│   │   ├── session_plan_models.py     # Domain models & boundary-aware delegate resolvers
│   │   ├── session_plan_parser.py     # Fault-tolerant Word (.docx) & Markdown session plan parser
│   │   └── session_plan_timezone.py   # Venue vs Vienna UTC timezone resolution
│   ├── tdocs/                         # TDoc & Chairman's Notes parsing and importer threads
│   │   ├── chair_notes_agreements_parser.py # RAN1/RAN2/RAN3/SA2 Chairman's Notes agreements & WA parser
│   │   ├── tdocs_parser.py            # Excel/HTML/Word TDoc list & agenda parser
│   │   └── tdocs_threads.py           # Background threads for TDoc & Word agenda loading
│   ├── agreement_tracker_worker.py    # Multi-meeting cross-WG Study Item & Topic agreement tracker worker
│   ├── chair_notes_downloader.py      # Smart FTP downloader & ZIP extractor for Final/EOM Chairman's Notes
│   ├── cohort_resolver.py             # Multi-WG meeting cohort clustering & quarterly cycle windowing
│   ├── cross_wg_ls_tracker.py         # Inter-group Liaison Statement dialogue thread tracker
│   ├── meeting_wi_enricher.py         # Asynchronous background Work Item discovery & enrichment
│   ├── meetings_data_loader.py        # 3gpp_data.db queries and meeting list loaders
│   ├── meetings_network_sync.py       # FTP sync for meetings and TDoc lists
│   ├── tdoc_file_handler.py           # Downloading and caching TDoc ZIP/Word documents
│   ├── tdoc_url_resolver.py           # Official 3GPP server HTTPS download URL resolution
│   ├── tdoc_markdown_converter.py     # Application adapter & Office wrappers for tdoc2md
│   ├── tdoc_triage_worker.py          # Local LLM technical triage generation
│   ├── tdocs_db.py                    # SQLite interface for user_tdocs.db
│   ├── topic_decision.py              # Unified technical decision model (RAN text agreements + SA/CT pCRs)
│   ├── topic_manager.py               # Persistent Meta-Topic manager & rule evaluation
│   └── stats/                         # Meeting statistics calculation and HTML report generation
└── ui/
    ├── agreement_tracker_dialog.py    # Cross-Meeting Study Item & Topic Agreement Tracker dialog
    ├── contribution_dialog.py         # Multi-meeting company contribution audit dialog
    ├── delegate_owner_picker.py       # Multi-delegate checkbox dropdown picker for topic ownership
    ├── dialogs.py                     # MeetingInfoDialog, AddMeetingDialog, MeetingsConfigDialog
    ├── meetings_actions.py            # Async download, caching, and export actions for MeetingsTab
    ├── meetings_delegates.py          # DynamicHoverMenuDelegate, TDocsButtonDelegate, WGTagDelegate
    ├── meetings_menu.py               # Right-click context menu builder for meetings table
    ├── models.py                      # MeetingsTableModel (7 consolidated columns, English date ranges)
    ├── dialogs/                       # Modal export & parameter configuration dialogs
    │   ├── post_meeting_export_dialog.py  # Post-meeting report generator dialog
    │   ├── pre_meeting_export_dialog.py   # Pre-meeting briefing generator dialog
    │   ├── tdocs_export_controller.py     # Centralized export dispatcher & controller
    │   └── topic_dossier_dialog.py        # Cross-WG Topic Dossier & Briefing Studio dialog
    ├── search_controller.py           # GlobalSearchController for inline Quick TDoc Jump & quick actions
    ├── session_plan_dialogs.py        # Timezone, Roster, Dispatch, Import & Diff dialogs for Session Plan
    ├── session_plan_window.py         # Interactive SA2 Session Plan Studio window
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
- **Official 3GU Web Portal Integration (`core.utils.portal_utils`):**
  - Canonical resolution (`get_portal_tdoc_url`, `open_portal_tdoc`) constructs and launches official 3GPP (3GU) web portal view URLs (`https://portal.3gpp.org/ngppapp/CreateTdoc.aspx?mode=view&contributionUid={uid}`).
  - Employs `contributionUid` (never obsolete `tdocId`) and automatically strips revision markers via `extract_base_tdoc()` so portal lookups consistently resolve to the base contribution.
  - Standardized across the TDoc right-click context menu, Quick Copy clipboard submenu, dossier HTML headers, meeting workspace links (`open_portal_meeting`), and contribution reservation actions (`open_portal_create_tdoc`).

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
- **UI Dialog (`PostMeetingExportDialog`):** Modeless configuration dialog accessible from `TDocsWindow`'s Export menu (`📑 Export Post-Meeting Report (.md)...`). Designed with a responsive 2-column layout (Left: Scope selection and full-height WI/AI tree; Right: Document types and conversion options) to guarantee zero clipping on laptop screens. Allows selecting export scopes (`🎯 Followed Work Items`, `⭐ Starred Work Items`, `🎯+⭐ All Tracked`, `🌐 Whole Meeting`, `👁️ Currently Visible`, or `📋 Custom Selection (Work Items & Agenda Items)...`), contribution categories (`LS in`, `pCR`, `CR`, `LS out`, `Discussion / WID`, `General LS`), and conversion options (auto-open, force reconversion, local index generation, and optional consolidated Navigable HTML/PDF dossier compilation defaulting to disabled for fast lightweight report generation).
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

### 3.6. Pre-Meeting Briefing & Markdown Dossier Generator
- **UI Dialog (`PreMeetingExportDialog`):** Modeless configuration dialog accessible from `TDocsWindow`'s Export menu (`🧭 Export Pre-Meeting Report (.md)...`). Designed with a responsive 2-column layout tailored for pre-meeting briefing and technical triage:
  - **Pre-Meeting Document Categories:** Defaults to capturing incoming liaison statements (`LS in`), proposed outgoing liaison statement replies (`LS out`), proposed pseudo-CRs (`pCR`), proposed Change Requests (`CR`), and company study/evaluation/status papers (`Discussion`, `WID`, `SID` — enabled by default).
  - **Intelligent Pre-Meeting Revision Deduplication (`deduplicate_revisions=True`):** Automatically traces revision chains (e.g. `S2-260100` $\rightarrow$ `S2-260500`) to export only the latest active non-withdrawn draft, annotating predecessor provenance `(rev of S2-260100)` to eliminate reviewing obsolete drafts twice.
  - **Unuploaded Reservation Awareness (`include_reserved_in_index=True`):** Lists agenda reservations in index tables with a `⏳ Reserved` badge while bypassing network downloads in Phase 0 to avoid 404 stalls.
  - **Contributor Focus Filter:** Optional single-click company focus selector to isolate or highlight specific companies' submissions across the meeting.
  - **Personal Triage Integration (`include_user_notes=True`):** Embeds the delegate's personal `My Status` (`🟢 Support`, `🔴 Object`, `🔵 My TDoc`) and `My Notes` into index tables and injects `> [!IMPORTANT] Delegate Pre-Meeting Stance` callouts into dossier headers.
- **Exporter Worker (`PreMeetingExporterThread`):**
  - **Structure:** Hierarchical dossier tree under `Export/PreMeeting_<WG>_<Meeting>_<timestamp>/` arranged by Release (`Rel-XX/`) and Work Item (`<WI>/`), containing subfolders (`LS_IN/`, `LS_OUT/`, `pCR/`, `CR/`, and `Discussion/`).
  - **Multi-Proposal Candidate LS Reply Clustering:** Supports 1-to-many candidate reply mappings where a single `LS in` has competing draft replies from multiple companies. Generates reciprocal Markdown callouts (`Proposed Reply Liaison Statements (N candidates)` on `LS in` and `Competing Reply Proposals (M other(s))` on `LS out`) enabling 1-click side-by-side navigation across competing company proposals.
  - **Work Item Inheritance for Draft LS Replies:** When a candidate `LS out` proposal leaves `Related WIs` unassigned, it automatically inherits the Work Item of the parent `LS in` it answers, ensuring competing replies co-locate inside the parent topic directory.
  - **Contributor Activity Landscape Matrix:** Synthesizes master and local WI activity matrices aggregating document volume per company across `pCR`, `CR`, `Discussion`, and `LS out`.

### 3.7. Meeting Analytics & Statistics Dashboard
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

### 3.8. Opportunistic Work Item Auto-Enrichment (`meeting_wi_enricher.py`)
- **Decoupled Asynchronous Architecture:** `MeetingWIEnricherThread` is fired fire-and-forget in the background when `TDocsWindow` is initialized or when a meeting list is refreshed. It never blocks or slows down initial TDocs table loading.
- **Bounded Hyperlink Extraction:** Scans populated rows of the "Related WIs" column for 3GPP Portal hyperlinks matching `workitemId=(\d+)`. Breaks boundedly when consecutive blank TDoc rows are encountered.
- **Database Cross-Checking:** Queries SQLite `work_items` via `WorkItemsDatabase.get_existing_wi_codes()` to isolate only unindexed WIs.
- **Background Scraping:** Dispatches polite concurrent requests (`fetch_wi_details_from_portal`) and atomically persists newly discovered WIs with their metadata, responsible groups, and remarks (`WorkItemsDatabase.upsert_single_work_item()`).
- **Audit Marker Caching:** Writes `<excel_path>.wi_checked.json` (recording timestamp, file mtime, discovered count, and scraped codes). On subsequent opens, `is_enrichment_needed` checks `mtime(marker) >= mtime(excel)` and exits in ~0 ms.
- **Dynamic Refresh & Config:** Quietly updates tracked tokens and filter badges if any newly scraped WI matches user watchlist tokens. Can be toggled on/off in `MeetingsConfigDialog`.

### 3.9. Meetings Browser Modern Architecture (`MeetingsTab`)
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

### 3.10. High-Speed Parallel Pre-Download Pipeline for Markdown Exports (`llm_exporter.py`, `post_meeting_exporter.py`, `pre_meeting_exporter.py`, `tdoc_file_handler.py`)
- **Decoupled Two-Phase Processing:** Both visible-table Markdown exports (`TDocMarkdownExporterThread` in `llm_exporter.py`), post-meeting report dossiers (`PostMeetingExporterThread` in `post_meeting_exporter.py`), and pre-meeting briefing dossiers (`PreMeetingExporterThread` in `pre_meeting_exporter.py`) employ a two-phase architecture to eliminate serial network latency when exporting unindexed or uncached TDocs:
  - **Phase 1 (Concurrent Pre-Download):** Sub-millisecond pre-scan identifies unindexed or missing source documents (`.docx`/`.doc`) across the export scope, skipping already-cached and fresh dossiers (and skipping unuploaded 'Reserved' TDocs). Dispatches missing archives to a polite worker pool (`ThreadPoolExecutor(max_workers=5..6)`). Downloads stream concurrently, cutting total elapsed network time by 5x–6x (~60s down to ~10s for 40 contributions) while staying within 3GPP server rate-limits.
  - **Phase 2 (In-Memory Markdown Conversion):** Sequentially converts the now locally-cached documents via direct OpenXML parsing without intermittent network pauses. Guarantees 100% single-threaded COM isolation for Word fallbacks (`win32com.client`) and prevents system-freezing multi-process churn from concurrent PowerShell vector metafile conversions.
- **Atomic File Integrity & Self-Healing Cache (`TDocFileHandler`):**
  - Downloads stream to temporary `.zip.part` files prior to atomic `os.replace` renaming, guaranteeing that interrupted network transfers or user cancellations never leave corrupt, 0-byte ZIP files on disk.
  - Automatically verifies local ZIP integrity with `zipfile.is_zipfile`; corrupted or truncated archives are automatically purged and re-fetched fresh.
  - Extracted `.docx`/`.doc` documents are similarly written via atomic `.part` rename.
  - Bidirectional route fallback: Automatically tries primary and secondary routes (`revisions_url` and `docs_ftp_url`) with a generous 15s timeout and non-blocking `cancel_check()` responsiveness.

### 3.11. Chairman's Notes & Cross-Meeting Study Item Agreement Tracker (F-19)
- **Multi-Surface Input:**
  - **Drag & Drop (`TDocsWindow.dropEvent`):** Dragging `.docx`, `.doc`, `.zip`, `.htm`, or `.html` session documents onto `TDocsWindow` renders an interactive frosted blue drop overlay (`TDocsDropOverlay`) with dashed drop borders.
  - **Menu Action:** **🔄 Refresh ▾ $\rightarrow$ 📝 Import Word Document (.docx / .doc)...** provides a manual file chooser.
  - **Downloader Bridge (`ChairNotesDownloaderThread`):** Unlocked across all Working Groups (`RAN1`, `RAN2`, `RAN3`, `SA2`, etc.). Probes `/Inbox/Chairs_Notes`, `/Inbox/Chair_Notes` (case-insensitive variants) into `Agenda/ChairNotes/`, smartly selects the Final/EOM combined notes and per-session EOM files (preferring `_clean` over `_tracked_changes`), automatically extracts `.zip` packages (such as `R2_XXX_Chair_Notes_..._final.zip`), and merges TDoc updates and extracted agreements into `Agenda/chair_notes_cache.json`.
- **Background Conversion & Staging Worker (`WordAgendaImporterThread`):**
  - Copies and extracts `.docx`, `.doc`, and `.zip` archives safely to `{meeting_dir}/Agenda/ChairNotes/`.
  - Strips NTFS `Zone.Identifier` stream metadata via `_unblock_file`.
  - Headless LibreOffice sanitizes legacy Word 97–2003 `.doc` files (which often embed VBA macros) to `.docx` without executing macros or freezing the GUI via Word COM.
- **Dual-Mode Word Document & Agreements Parser (`ChairNotesAgreementsParser`, `TDocsParser.parse_tdocs_from_docx`):**
  - **Mode 1: Color- & Banner-Aware Table Parser (RAN3, SA2, CT):**
    - **Multi-Row Header Scanning (Rows 0–4):** Robustly locates table headers (`TDoc | Title | Comments` in RAN3, multi-column in SA2) even when session chairs place meeting banners or date/time headers on rows 0–1.
    - **Shaded Agenda Item Banners & WI Inheritance:** Extracts shaded AI banners and `WID [code]` / `SID [code]` tags, propagating parent AI Work Item mappings (e.g. `10` `FS_6G_Radio` $\rightarrow$ `10.2.2`) to sub-clauses.
    - **OpenXML Color & Prefix Semantic Classification:** Distinguishes `From RAN3#<prev>:` historical recaps from current-meeting decisions, and classifies OpenXML green (`#008000`, `#00B050`, `#538135`) paragraphs as **Agreements** or **Working Assumptions** (`WA:`) and blue (`#0000FF`, `#3333FF`) paragraphs as **Open Issues / FFS**.
  - **Mode 2: Document-Order Block Parser (RAN1, RAN2):**
    - Walks paragraphs and 1x1/2x1 boxed tables in true document order (`iter_block_items`).
    - Maps headings (`Heading 1..4`) to `Agenda Item`, extracts parenthetical WI tags (`(FS_6G_Radio, ...)`), 5th/6th tab columns on `Doc-title` lines (`Rel-20\tFS_6G_Radio`), and `6GR` / `Study on 6G Radio` heading indicators.
    - Captures 1-column boxed consensus/guidance/terminology tables (Pattern A, e.g. Table 7 in RAN2#133 with *"Terminology and guidance"*, defining Data consumers and Termination points) as well as explicitly titled agreement boxes, while filtering out antitrust and IPR boilerplate.
    - Captures multi-paragraph `Agreements` blocks (Pattern B) with blank-line resilience: maintains agreement blocks across empty paragraphs in Word documents without premature flushing, using consecutive empty line tracking and section breaker detection (such as new Agenda Item headings, new TDoc lines, or substantive offline discussion / proposal headers).
    - Substantive `style == 'Agreement'` paragraphs are captured while filtering out procedural dispositions (`"The CR is postponed"`, `"Noted"`, etc.).
  - **Cache Versioning & Auto-Invalidation (`CHAIR_NOTES_CACHE_VERSION = 2`):**
    - `chair_notes_cache.json` stores an explicit `"version": 2` tag.
    - When `TopicDossierExporter` or `ChairNotesAgreementsParser.load_cached_chair_notes_data` detects older caches (`version < 2` or missing), it automatically triggers a background re-scan and cache update, ensuring enriched multi-paragraph agreements and boxed guidance tables are surfaced seamlessly without requiring manual cache wipes.
- **Live Model Synchronization & Post-Meeting Export Integration:**
  - `TDocsTableModel.merge_agenda_data` populates in-room discussions and `[⚖️ AGREEMENT]: ...` blocks directly into the `Secretary Remarks` column, updates `TDoc Status`, and synthesizes `[AI X.Y Agreement]` rows for agenda-level agreements not tied to a single TDoc.
  - `PostMeetingExporterThread` & `PostMeetingExportDialog` include a `⚖️ Chairman's Notes Agreements & Working Assumptions` option (enabled by default) that renders dedicated per-WI and master index agreement tables alongside TDoc dossiers.
- **Cross-Meeting Study Item & Topic Agreement Tracker (`AgreementTrackerDialog`, `AgreementTrackerWorker`):**
  - Accessible from `MeetingsTab` (**📊 Reports ▾ $\rightarrow$ ⚖️ Track Study Item & Meeting Agreements...**) and `TDocsWindow` (**🔄 Refresh ▾ $\rightarrow$ ⚖️ Track Study Item & Meeting Agreements...**).
  - Scans multiple Working Groups (`RAN2`, `RAN3`, `RAN1`, `SA2`, etc.) and date ranges in parallel (`ThreadPoolExecutor`), auto-downloading missing TDoc lists and Final/EOM Chairman's Notes from `/Inbox/Chairs_Notes`.
  - Filters chronologically across meetings by **Work Item / Study Item** tokens (with 1-click `🎯 Followed` and `⭐ Starred` loaders), **Agenda Item prefix** (`9.3` matching `9.3.1`), **Topic Keyword**, and **Decision Type** (`🟢 Agreements`, `🟡 Working Assumptions`, `🔵 Conclusions / Way Forward`, `🟣 Open Issues / FFS`, `📄 Agreed TPs / pCRs / CRs`, and optional historical recaps).
  - Provides interactive table exploration, rich detail preview with 1-click TDoc/Notes/Meeting openers, and exports to **Study Item Markdown + Interactive HTML Dossier (`.md` & `.html`)** and **Multi-Sheet Excel (`.xlsx`)**.

### 3.12. SA2 Meeting Session Plan Management, Version Diffing & Mobile Export
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
  - **📱 4-View Interactive Offline HTML (`SA2_{mtg}_Session_Plan_Mobile.html`):** Standalone, 100% offline responsive single-file web application (<110 KB, zero external CDN dependencies) built for easy sharing across the delegation (intentionally omitting purely personal Starred Topics):
    - **4-View Segmented Switcher:** Allows 1-click toggling between:
      1. **🗓️ 5-Day Spatial Overview Grid (`#overviewView`):** 10-column layout mirroring desktop with half-quarters, dedicated stream rows, break rows, and stream toggles. Features **2D Sticky Headers** inside a bounded scroll container (`.overview-grid-container { overflow: auto; max-height: calc(100vh - 175px); }`) locking both the `Period / Stream` left column (`left: 0; z-index: 10`), the `Day` header row (`.hdr-days-row th { position: sticky; top: 0; z-index: 20; }`), the `1st / 2nd Half` TU row (`top: var(--day-hdr-h, 29px); z-index: 19`), and the top-left corner cell (`.corner-hdr { position: sticky; top: 0; left: 0; z-index: 30; }`) via `updateStickyHeaderOffsets()`.
      2. **📱 Mobile Day Cards (`#cardsView`):** Touch-friendly day-by-day quarter view with sticky day tabs. Automatically defaults to Day Cards on mobile screens ($\le 768$px).
      3. **⚠️ Conflicts Tab (`#conflictsView`, `#btnViewConflicts`):** Dedicated parallel-stream collision view (also reachable by clicking `#clashWarningBadge`). Lists every delegate clash with exact half-quarter (`TU 1` / `TU 2`) timing, side-by-side conflicting streams, role badges (`[Owner]`, `[Override]`, `[All Team]`), a **`🟢 Available / Free Colleagues`** strip showing unbooked teammates who can cover, and a **`🔍 Show in Grid`** action (`jumpToSlotInGrid`). Respects active delegate filtering.
      4. **❓ Unassigned AIs Tab & Filter (`#unassignedView`, `#btnViewUnassigned`, and `#btnUnassigned` filter chip):** Dedicated gap-analysis view listing all scheduled session slots with no assigned delegate (paired with free colleagues at that slot and `🔍 Show in Grid` links) alongside a complete catalog table of unassigned Agenda Items (`AI #`, `Topic Description`, `Scheduled Slots`). In Grid/Card views, toggling the `❓ Unassigned AIs (N)` filter pill highlights unassigned cells (`unassigned-slot`) and dims assigned ones.
    - **📍 Highlight Now & Live Team Dispatch Board (`#btnHighlightNow`, `#nowDispatchPanel`, clickable `#liveVenueBanner`):** Answers *"Where are we right now, and where should each delegate go?"*:
      - Evaluates active sessions from the real-time venue wall-clock (or falls back to the first scheduled slot with a manual **Time Unit Stepper** `◀ Prev`, `<select id="nowSlotPicker">`, `Next ▶`, and `🕒 Live Now` reset button so delegates can step through any half-quarter of the week).
      - Renders per-delegate dispatch cards (**"👥 Where Each Delegate Should Go"**) showing each colleague's assigned **Stream**, **Physical Room**, **Role**, and **Topic** (or warning if double-booked across parallel rooms, or marking them `🟢 Free / Available for Coverage`), plus an **Uncovered Active Streams** alert banner and animated `.now-focus-pulse` highlighting in the Overview Grid and Day Cards.
    - **Interactive Delegate Filter Strip:** Egalitarian team buttons (`[👥 All Team]`, `[● Anna]`, `[● Thomas]`, ...) styled with solid pastel tints. Toggling a delegate dims non-assigned cells in the Overview Grid and filters cards in Mobile Day Cards and Conflicts. Filter actions use DOM attribute binding (`this.getAttribute('data-del')`) and sanitized arrays (`.filter(Boolean)`), ensuring robust handling of arbitrary names.
    - **Client-Side RFC 5545 `.ics` Export:** Single-click `[📅 Download Team .ics]` or `[📅 Anna .ics]` dynamically compiles and downloads UTC-anchored `.ics` calendar files directly in the browser via Blob URL, strictly scoped to the active delegate or full team without server round-trips.
    - **Rich Floating Tooltips:** Hovering over any slot displays venue/laptop times, physical room, convener, TOHRU audio/video channel, resolved Agenda Item descriptions from `agenda.csv`, and personal notes.
    - **Robust Script & Template Sanitization:** Python HTML synthesis strictly escapes JavaScript regex newlines (`/[\\r\\n]+/g`), calendar line joins (`'\\r\\n'`), string quotes, and serializes variables via `json.dumps()`, eliminating Python string interpolation syntax errors. Day tabs use explicit element passing (`switchDay(day, this)`) with DOM fallbacks. Verified by automated Node.js syntax and runtime execution tests in CI.
  - **📅 Strict UTC iCalendar (`.ics`):** Exports sessions to Outlook / Apple / Google Calendar with strict UTC `DTSTART`/`DTEND` timestamps, room/floor in Location, and convener/topics/notes in Description.
  - **✉️ Colleague Email Composer:** Formats rich HTML & plain text summaries ready to paste into Outlook summarizing what changed since the previous plan, today's drafting sessions, and tracked topics.
  - Enriches TDoc rows with live `Sub-Item` topic separators, `_agenda_order`, and `_sub_item_order` extracted from `TdocsByAgenda.htm` and Word Chairman's Notes.
- **Delegation Topic Distribution, Cross-Meeting Persistence & Session Clash Dispatching (`SessionPlanWindow`, `SessionPlanDatabase`, `SessionPlanExporter`):**
  - **Egalitarian Roster Model (No "Myself" Tag):** Replaced the restrictive "myself" / `(Me)` tag in delegation roster, dialogs, and pickers with an egalitarian architecture where all team members are treated equally. Every colleague knows who they are; removing the tag ensures exported HTML and schedules are clean, universal, and ready to share across the entire delegation. Auto-assigning followed WIs presents an intuitive delegate selection prompt.
  - **Two-Tier Assignment Model with Cross-Meeting Persistent Defaults:**
    - **Tier 1 — Default Topic Ownership (Cross-Meeting Persistent):** High-level strategic AI topic ownership (e.g. *Martin owns AI 20.6.18 and AI 5.2.1*) and team roster (`default_delegation_roster` and `default_topic_owners` SQLite tables in `session_plans.db`, automatically synchronized with `<app_data_root>/delegation_defaults.json`) persist across meetings (`SA2#177` $\rightarrow$ `SA2#178` $\rightarrow$ `SA2#179`). Opening a new meeting automatically seeds its default roster and AI topic owners from the global defaults, while any updates to the roster or topic owners propagate forward as the new default assignment.
    - **Tier 2 — Session-Level Overrides (Per-Meeting `session_appointments`):** Tactical single-slot coverage overrides (e.g. appointing a free colleague to cover Tuesday Q2 TU1 Stream 2 during a collision) are scoped strictly to the specific meeting and `(day, quarter, sub_slot, stream)`.
    - **Override Visibility & 1-Click Reset (`clear_session_appointments`):** The Topic Ownership Matrix displays active session override counters (`⚡ N session override(s)`) in the `🛡️ Coverage Status` column and provides a **`↺ Reset Session Overrides`** button in the Team Allocations toolbar to clear all per-meeting ad-hoc appointments back to default topic owners in 1 click.
  - **Shared Boundary-Aware Resolver (`resolve_slot_effective_delegates`, `ai_number_in_text` in `session_plan_models.py`):** Single source of truth used by `SessionPlanWindow` and `SessionPlanExporter`. Raw-text fallback uses regex negative lookbehind/lookahead (`(?<![\d.])` and `(?!\d|\.\d)`) so short AI tokens like `20.6.2` or `20.6.1` never falsely match `20.6.24`, `20.6.22`, or `20.6.17`.
  - **Multi-Delegate & "All Team" Ownership (`DelegateOwnerPicker`):** Inline checkbox-based multi-select dropdown (`QToolButton` + `QWidgetAction` checkboxes) in the Topic Ownership Matrix allowing delegates to tick one or more owners, `👥 All Team`, `⛔ Mark as Not Followed`, or `✕ Clear (Unassigned)` in a single popup without modal dialog round-trips—committing checkbox toggles when the popup closes (`aboutToHide`) and committing `QAction` resets (`_reset_selection` $\to$ `_commit`) immediately so allocations and clashes recalculate without being masked by Qt's `aboutToHide`-before-`triggered` signal ordering.
  - **Direct Overview Grid Topic Owner Assignment:** Right-clicking any cell in the desktop 5-day Overview Grid (`_on_grid_context_menu`) provides a dedicated `👤 Assign Topic Owner (AI X.Y.Z)...` submenu (expanding `.x` wildcards via `expand_ai_number` into concrete `agenda_map` sub-AIs) enabling instant 1-click assignment to `👥 All Team`, any individual delegate, `Custom Multi-Delegate`, `⛔ Mark as Not Followed`, or `✕ Clear (Unassign)` directly from the grid without having to open the Topic Ownership Matrix.
  - **Per-Delegate Collision Evaluation:** When a general topic (like 20.6.0) or multi-delegate topic is scheduled, collisions are evaluated per individual delegate across both the Overview Grid, Clash Board, and `🎯 Delegate Agenda` (including `👥 Whole Delegation (ALL)` view): only colleagues who are simultaneously double-booked in a parallel stream at an overlapping half-quarter receive a clash warning, while unconflicted delegates remain marked as available and attending.
  - **Half-Quarter (TU1 / TU2) Time-Overlap Precision (`_slots_time_overlap`):**
    - Accurately models 45-minute half-quarters (`sub_slot=1` for TU1, `sub_slot=2` for TU2, and `sub_slot=0` for full quarters).
    - Sequential half-quarters (e.g., Stream 1 TU1 $\rightarrow$ Stream 2 TU2) are recognized as non-overlapping in time, completely eliminating false double-booking warnings.
    - True simultaneous collisions (e.g., Stream 1 TU1 and Stream 2 TU1, or full Q2 vs TU1/TU2) are reliably caught.
  - **Dynamic "Who is Free?" Availability Engine (`_compute_delegate_availability`):**
    - When a delegate is double-booked across parallel rooms (e.g., Stream 1 and Stream 2 at Tuesday Q2 TU1), calculates real-time room occupancy for every team member.
    - Accurately marks colleagues 🟢 *Free* for TU1 if their only assignment in that quarter is in TU2 (or another time).
  - **Enhanced Multi-Stream Clash & Dispatch Board (`⚡ Clash & Dispatch Board`, `btn_conflicts_quick`):**
    - Top navigation ribbon features a live **`⚠️ Conflicts (N)`** quick-jump pill button (`self.btn_conflicts_quick`) that jumps straight to `👥 Team Allocations` and scrolls to the Clash & Dispatch Board.
    - 7-column layout (`["📅 Slot", "🕒 Time", "👤 Colliding Delegate", "Stream A (Assigned)", "Parallel Stream B", "🟢 Free Colleagues", "🛡️ Action / Coverage"]`) displaying explicit `(TU 1)` / `(TU 2)` half-quarter badges, assignment source tags (`[Owner]`, `[Override]`, `[All Team]`), and all unbooked colleagues in the **`🟢 Free Colleagues`** column.
    - **1-Click Quick-Appoint Menu (`_quick_appoint_colleague`):** Clicking **`⚡ Appoint Coverage ▾`** on any conflict row lists instant 1-click actions (`⚡ Quick-Appoint <Free Colleague> → Cover Str X (<Room>)`) for every free colleague and conflicting stream, alongside full dialog dispatch options (`📋 Open Dispatch Dialog for Str X...`).
    - **Double-Click Jump to Grid (`_on_clash_cell_double_clicked`):** Double-clicking any conflict row switches directly to the 5-day Overview Grid and focuses the exact colliding slot.
  - **Real-Time Topic Ownership Matrix Synchronization:**
    - Live status cell updating (`_update_team_topic_status_cells`) immediately reflects changes in topic coverage, session override counts, and collision warnings without recreating widgets or losing scroll position.
  - **Direct Grid & Agenda Context Menu Dispatching:**
    - Right-clicking any slot in either the Overview Grid (`_on_grid_context_menu`) or the Delegate Agenda (`_on_tracked_context_menu`) provides instant actions:
      - `⚡ Appoint/Change Coverage for this Session...` (pre-filled with current assignment)
      - `↺ Revert Coverage to Topic Owner` (restores default topic ownership in 1 click via `_on_revert_dispatch`).
  - **Multi-Delegate Color-Coded Heatmap & Focus Filtering:**
    - Auto-assigns distinct, accessible palette colors (`DELEGATE_PALETTE`) to each team member. Features 10 harmonious, high-contrast hues (WCAG AA compliant $\ge 4.5:1$ against white): Royal Blue (`#2563EB`), Emerald Forest (`#047857`), Deep Violet (`#7C3AED`), Warm Terracotta (`#C2410C`), Cyan Ocean (`#0E7490`), Ruby Rose (`#BE123C`), Indigo Night (`#4338CA`), Plum Orchid (`#A21CAF`), Warm Bronze (`#B45309`), and Slate Graphite (`#475569`).
    - **Solid Pastel Tint Architecture (`DelegationMember.compute_tint`):** Computes solid `#RRGGBB` pastel tints (`bg_tint` 12% and `border_tint` 35% blended with `#FFFFFF`) for delegate pills, chips, active highlight strips, tooltips, and HTML email summaries. **Strict prohibition of 8-digit hex codes (`#RRGGBBAA`)**: Qt QSS parses 8-character hex strings as `#AARRGGBB` (reversing alpha and RGB channels, turning `#2563EB1F` into dark sludge green `#1a3f07`), while Microsoft Outlook email engines ignore 8-digit hex completely. Blending directly on white produces 100% solid, predictable colors across Qt on Windows, Web, mobile, and Outlook.
    - **Compact Initials Pill Badges (`_make_delegate_badge`, `compute_unique_initials`):** Overview Grid renders space-saving circular/pill initials badges (`JO`, `DI`, `PI`, with automatic 3-letter disambiguation if two colleagues share 2-letter initials) instead of repeating full names in every cell. Solid border indicates topic ownership, dashed border indicates ad-hoc coverage override, and a 2px red border ring (`#DC2626`) + compact `⚠` pill highlights parallel-stream clashes. Full names, roles, and clashing stream details appear in rich hover tooltips.
    - **Multi-Select Delegate Focus Bar:** Top ribbon focus bar (`[🌈 All Team]` + checkable `JO · Josep`, `DI · Dieter`, ...) supports multi-select checkbox filtering so users can highlight any combination of delegates simultaneously while dimming unrelated sessions.
  - **Delegate Agenda View (`🎯 Delegate Agenda`):**
    - Segmented switcher (`[👥 Whole Delegation | ● Anna | ● Thomas]`) renders personal and team timetables with dual venue/laptop times, rooms, roles, and half-quarter (`TU 1`/`TU 2`) accurate per-delegate conflict tags.
    - Direct double-click jumps to the cell on the Overview Grid; right-click allows note editing or coverage dispatching.
  - **Personalized Delegate iCalendar (.ics) Calendar Export:**
    - Main toolbar Export menu provides `📅 Export Delegate .ics (Personal / Colleague)...` to export calendar events strictly scoped to an individual delegate's assigned sessions (both owned topics and appointed coverage), strictly UTC-anchored for Outlook/Google Calendar/iPhone.
    - Full meeting `.ics` exports automatically include delegate assignment tags (`[Delegate]` / `[Delegate (cov)]`) in event summaries and descriptions.
  - **Email & Machine-Readable Interchange (Frictionless Sharing):**
    - **`📧 Export Email Summary...`** (`generate_delegation_email_summary`): Generates pre-formatted executive HTML/plain text for Outlook with color badges, topic ownership tables, ad-hoc coverage notes, and personal daily schedules. Accessible via top Export menu and Team Allocations action bar.
    - **Embedded Interchange Payload:** Automatically embeds `<!-- 3GPP-DELEGATION-PLAN: {...} -->` into the email or exports `.json`.
    - **1-Click Colleague Ingestion (`DelegationImportDialog`):** Colleagues paste the email or load the `.json` file in their 3GPP Tools, select *"Which delegate are you?"*, and instantly personalize their local Overview Grid and personal agenda.
    - **Persistence:** Roster and topic ownerships persist across meetings (`default_delegation_roster`, `default_topic_owners`, and `<app_data_root>/delegation_defaults.json`) as well as per-meeting in `session_plans.db` (`delegation_roster`, `topic_owners`, `session_appointments`), surviving across intra-meeting session plan snapshot revisions (`Snap 1` $\rightarrow$ `Snap 2` $\rightarrow$ `Snap 3`).
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
  3. **OpenXML Vertically Merged Cells (`<w:vMerge>`) Across Parallel Streams & Breaks:**
     - *Quirk:* Plenary sessions (e.g., Monday morning Q0/Q1/Q2, Friday afternoon Q4/Q5) or single-stream sessions span across all 3 stream rows (and sometimes into adjacent break rows like `08:50–09:00 Break 1`) using OpenXML vertical cell merges. In `python-docx`, iterating across row cells on vertically merged tables repeats the master cell's XML pointer (`_tc`) and text across Streams 1, 2, 3, and break rows.
     - *Fix:* Table parsing tracks underlying XML cell references (`seen_tcs = set()`) across both stream rows and break rows. Vertically merged cell continuations are deduplicated, eliminating phantom duplicate sessions in Streams 2/3, preventing `Opening of meeting at 0900` from leaking into `Short Break`, and preventing false parallel-stream double-booking warnings for delegates.
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
  10. **OpenXML Track Changes (`<w:ins>`, `<w:del>`, `<w:moveTo>`, `<w:moveFrom>`):**
      - *Quirk:* Revised session plans (e.g., `r1.docx`, `r2.docx`) are distributed with Word Track Changes active rather than accepted. In `python-docx`, `cell.text` only reads direct `./w:r/w:t` children of `<w:p>`, completely ignoring `<w:ins><w:r><w:t>` insertions and `<w:moveTo>` blocks (causing newly scheduled or moved sessions and revised time headers to disappear).
      - *Fix:* `_extract_xml_text(xml_elem)` recursively walks `<w:p>` and `<w:tc>` XML elements, extracting `<w:t>` inside normal runs, `<w:ins>`, `<w:moveTo>`, and hyperlinks while skipping `<w:del>` and `<w:moveFrom>` deletions.
  11. **Document-Specified Quarter Times & Midpoint Calculation (`_resolve_quarter_times`):**
      - *Quirk:* Working Group chairs occasionally shift quarter boundaries (e.g., moving Q2 from `1100 - 1230` to `1230 - 1400` in Calgary SA2#178 `r1.docx`). Hardcoding `QUARTER_TIMES[quarter]` ignored the document's time column.
      - *Fix:* `_resolve_quarter_times(quarter, current_times)` parses the document's time column (including Track Changes) and dynamically computes the exact midpoint for half-quarters (`TU 1`: `12:30–13:15`, `TU 2`: `13:15–14:00`), propagating effective times to desktop grid row headers, hover tooltips, and HTML exports.
  12. **`.x` / `.X` Wildcard Agenda Items (`30.x`, `5.x`, `20.6.x`):**
      - *Quirk:* Chairs routinely schedule wildcard Agenda Items such as `Work planning (30.x), etc (Rel-21 5G-A revisions)` or unparenthesized `Work planning 30.x` / uppercase `20.6.X`, whereas `agenda.csv` enumerates concrete sub-items (`30`, `30.1`, `30.2`). Previously standalone `.x` and uppercase `.X` tokens failed extraction, and `30.x` did not match concrete `30.*` ownership or agenda entries.
      - *Fix:* `SessionPlanParser._extract_agenda_items` and `_extract_room_override_and_tokens` parse both parenthesized and standalone `.x`/`.X` identifiers (including multi-AI groups like `Rel-21 (30.x) and Rel-20 (20.x)` and comma-separated plenary lists `Opening (1), Work planning (30.x)`), normalizing suffixes to lowercase `.x`. In `session_plan_models.py`, `expand_ai_number()` and `expand_slot_ai_numbers()` expand `.x` wildcards into their concrete `agenda_map` child AIs (preferring strict `<prefix>.*` children over bare chapter headers like `30` when children exist, falling back to `<prefix>` or the raw `.x` token otherwise), powering `_matching_topic_ownerships()`, `resolve_slot_effective_delegates()`, `get_slot_unassigned_ais()`, `is_slot_not_followed()`, `build_slot_ai_bullets_html()`, `_is_slot_tracked()`, desktop/HTML hover tooltips, Overview Grid context menus, mobile AI chips, and HTML `UNASSIGNED_AIS` gap analysis, while `SessionPlanWindow._combined_ai_keys()` surfaces scheduled wildcard items even when absent from `agenda.csv`.
- **Desktop Overview Grid Widget Lifecycle (`_render_overview_grid_impl`):**
  - When switching between session plan snapshots or importing a revised `.docx`, `self.grid_table.setRowCount(0)` is called before populating the 23 rows. In Qt, calling `setItem(row, col, item_e)` on a cell that previously held a `cellWidget` does not remove the `cellWidget` (which overlays `QTableWidgetItem`), and changing `setSpan` stretches stale widgets across sub-columns. Resetting row count to `0` destroys all previous `cellWidget`s and resets column spans cleanly.
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
    - Applies identical unclipped header styling, column widths, and 36px section sizing to the 7-column Clash & Dispatch Board (`["📅 Slot", "🕒 Time", "👤 Colliding Delegate", "Stream A (Assigned)", "Parallel Stream B", "🟢 Free Colleagues", "🛡️ Action / Coverage"]`).

### 3.14. Three-State Delegation Model (`Assigned` / `Unassigned` / `Not Followed`) & Export Gap Hygiene (`session_plan_models.py`, `delegate_owner_picker.py`, `session_plan_window.py`, `session_plan_exporter.py`)
- **Problem Solved:** Working delegations do not follow every Agenda Item declared by a 3GPP Working Group. When an agenda item is intentionally out of scope, leaving it unassigned caused it to trigger false alerts in the HTML export's `❓ Unassigned AIs` view, `#btnUnassigned` highlight filter, and `📍 Highlight Now` active stream dispatch board, obscuring actual delegation coverage gaps.
- **Architectural Implementation:**
  - **Data Model & Normalization (`session_plan_models.py`):**
    - Defines `NOT_FOLLOWED_TOKENS = ("NOT_FOLLOWED", "NOT FOLLOWED", "NOT_FOLLOW", "NOT FOLLOW", "⛔ NOT FOLLOWED", "⛔ NOT_FOLLOWED", "⚪ NOT FOLLOWED", "EXCLUDE", "IGNORED")`.
    - `TopicOwnership.is_not_followed()` detects when an item is explicitly marked not followed.
    - `TopicOwnership.is_all()` and `TopicOwnership.get_owners()` treat not-followed items as having zero individual delegate owners.
    - `resolve_slot_effective_delegates()` ignores not-followed topics so they never assign phantom coverage or trigger false colleague collisions.
    - `is_slot_not_followed(slot, plan)` returns `True` when all agenda items scheduled in a session slot are explicitly marked not followed.
  - **Cross-Meeting Persistence:**
    - Marking a topic as not followed saves `owner_name = "NOT_FOLLOWED"` to `topic_owners` and `default_topic_owners` (and `<app_data_root>/delegation_defaults.json`), automatically seeding the `Not Followed` status across all subsequent meetings without requiring schema alterations.
  - **Desktop UI Integration (`delegate_owner_picker.py`, `session_plan_window.py`):**
    - `DelegateOwnerPicker` provides a 1-click `⛔ Mark as Not Followed` action in its dropdown menu, displaying `⛔ Not Followed ▾` with distinct neutral/muted slate styling (`#64748B`, `#F1F5F9`, border `#CBD5E1`).
    - The Topic Ownership Matrix in `SessionPlanWindow` renders a `⚪ Not Followed` badge in the `🛡️ Coverage Status` column.
    - The Overview Grid cell context menu provides a 1-click `⛔ Mark as Not Followed` action under the topic assignment submenu (`👤 Assign Topic Owner (AI X.Y.Z)...`), with `(Not Followed)` status indication.
    - Overview Grid cells for not-followed slots are rendered with neutral dashed borders and muted text (`border: 1px dashed #CBD5E1; color: #64748B; background: #F8FAFC;`), and rich hover tooltips indicate `⛔ Delegation: Not Followed (Out of Scope)`.
  - **HTML Export & Gap Analysis Hygiene (`session_plan_exporter.py`):**
    - `_get_slot_unassigned_ais()` and `_is_slot_unassigned()` exclude not-followed topics and slots.
    - The `❓ Unassigned AIs` catalog view and summary metric (`unassigned_ai_count`, `unassigned_slot_count`) only report genuine unassigned gaps, completely excluding intentionally ignored topics.
    - Grid cells and day cards tag not-followed slots with `is-not-followed`, `data-not-followed="1"`, `data-unassigned="0"`, and `⛔ Not Followed` pill badges.
    - The `📍 Highlight Now` dispatch board (`#nowDispatchPanel`) suppresses `❓ Unassigned Active Stream(s)` warnings for active streams whose topics are marked Not Followed.
    - The executive email summary generator tags not-followed items with `⛔ Not Followed` badges in the Master Topic Ownership table.
  - **Unified Compact Grid Cards, Right-Side Pills & Color Hierarchy (`session_plan_models.py`, `session_plan_window.py`, `session_plan_exporter.py`):**
    - **Shared Display Normalization (`format_slot_display_text`):** Extracts leading `In main room` / `In breakout 1/2` directives into compact `📍` room-override badges, merges standalone trailing `[TDocCount]` lines inline onto the preceding topic line, and flags empty `"Drafting:"` placeholders (`is_empty_drafting`) so they render in muted slate rather than drafting red.
    - **Desktop Grid (`_create_slot_cell_widget`):** Wraps slot text and pills inside a single bordered `QFrame#slotCard`, positioning delegate initials (`[JO]`, `[DI]`, `[👥 ALL]`), clash alerts (`[⚠]`), `❓`, and `⛔` in a slim top-aligned right rail (`pill_col`) rather than stacking a separate header row above the card. Uses subtle ~7% `DelegationMember.compute_tint` backgrounds on assigned slots even in `All Team` view and **neutral slate borders** (`BORDER_STRONG`) on unassigned slots (eliminating stream-color collisions with delegate colors). Shortens stream toggle labels (`🏢 Str 1: <ShortRoom>`) with full chair/capacity/TOHRU metadata in tooltips, moves the badge legend into an `ℹ️ Legend` tooltip pill, formats Day headers across 2 compact lines (`Mon 16 Nov\n1st Half`), and triggers deferred `resizeRowsToContents()` to prevent vertical text clipping.
    - **HTML Overview Grid (`render_overview_cell`):** Wraps each slot inside a padded white `<td>` enclosing a desktop-aligned rounded `.slot-card` flex container (body left, `.cell-del-float` right-rail pills stacked vertically with compact 2-letter initials and no leading dot), renders `👥 ALL` cleanly without multi-line yellow `❓ AI N` chip overflow (using a compact `❓` right-rail indicator on partial non-ALL slots while preserving full AI breakdown in tooltips and `#cardsView`), highlights parallel clashes with Desktop's warm amber card fill (`#FFFBEB` / `#F59E0B`), `⚠` right-rail pill, and per-delegate red pill ring (`.clash-pill-ring` applied only to the colliding `data-del-name` pill by `markClashCells`), aligns the header/toolbar with `🕒 UTC` + `📍`/`💻` timezone pills and non-wrapping `1st Half` / `2nd Half` column headers, and mirrors Desktop's subtle ~7% assigned tints and neutral/dashed unassigned/not-followed cards.

### 3.15. Unconfirmed / Tentative Source Company Handling (`?`, `(?)`, `[...]`) (`company_sanitizer.py`, `tdoc_delegates.py`, `tdocs_models.py`, `tdocs_window.py`, `contribution_worker.py`, `contribution_dialog.py`)
- **Problem Solved:** In 3GPP Working Groups (especially SA2), revisions of CRs and pCRs must carry over the original contribution's `Source` list even when co-signing companies have not confirmed co-signing the revision (or may actively oppose the revised text). Delegates indicate unconfirmed carry-over companies by appending a question mark (`Huawei?`, `Nokia (?)`) or wrapping company names in square brackets (`[Apple, AT&T, Nokia]`, `[Samsung]`). Counting these unconfirmed carry-over companies in contribution statistics or default company filters artificially inflates company metrics and misattributes authorship, while stripping them completely would lose useful historical provenance.
- **Architectural Implementation:**
  - **Lossless Segment Tokenizer (`CompanySanitizer.tokenize_source_segments`):**
    - Parses raw `Source` strings into `(segment_text, is_tentative)` pairs with 100% text preservation.
    - Treats square-bracketed company blocks (`[...]`) and question-marked company names (`?`, `(?)`, `？`) identically (`is_tentative=True`).
    - Distinguishes role annotations in square brackets (`ROLE_BRACKET_RE`: `[pen holder]`, `[Rapporteur]`, `[Editor]`, `[Moderator]`, `[WI rapporteur]`, `[to be SA2]`) from tentative companies, attaching the role annotation to the preceding confirmed company.
    - Recombines comma-separated corporate legal suffixes (`CORP_SUFFIX_RE`: `, Inc?`, `, Ltd.?`, `, GmbH?`) and splits consecutive question-marked companies where authors omitted commas (`Huawei? Samsung? China Mobile?` or `ETRI?NEC?`).
  - **Confirmed vs. Tentative Contributor Resolution (`CompanySanitizer.parse_contributors_detailed`):**
    - Returns `(confirmed_set, tentative_set)`. If a company appears in both confirmed and tentative segments within the same string, it is normalized to confirmed (`tentative -= confirmed`).
    - `CompanySanitizer.get_matching_contributors(original_source, include_tentative=False)` and `CompanySanitizer.get_primary_contributors(original_source, include_tentative=False)` exclude tentative companies by default (`include_tentative=False`), ensuring single-meeting statistics (`StatisticsExporterThread`), cross-meeting contribution audits (`ContributionSearchWorker`, `ContributionStatsExporterThread`), and Markdown company reports (`MarkdownExporterThread`) never count unconfirmed sources.
  - **Inline Visual Rendering (`SourceCompanyDelegate` in `tdoc_delegates.py`, `tdoc_info_dialog.py`):**
    - Installed on the `Source` column in `TDocsWindow` (col 3) and `ContributionReportDialog` (col 4), as well as the metadata form in `TDocInfoDialog`.
    - When `CompanySanitizer.has_tentative_sources(source_str)` is `True`, renders confirmed contributors in standard text color and unconfirmed segments (`?`, `(?)`, `[...]`) in muted italic slate (`#94A3B8`, `font-style: italic`) via `QTextDocument` (`CompanySanitizer.format_source_html`), introducing zero extra columns, icons, or UI clutter.
    - When no tentative markers are present, falls back directly to `super().paint()` for zero rendering overhead.
  - **Self-Explanatory Hover Tooltip (`CompanySanitizer.build_source_tooltip_html`):**
    - Hovering over any `Source` cell with tentative markers displays a structured rich HTML tooltip separating **✅ Confirmed Contributors** from **❔ Unconfirmed / Carry-over**, paired with a concise explanation of the 3GPP CR/pCR carry-over convention and statistical exclusion rule.
  - **Optional Filter Inclusion Toggles:**
    - **`TDocsWindow`:** `TDocsTableModel` precomputes `_Sanitized_Companies` (confirmed only), `_Tentative_Companies`, and `_All_Sanitized_Companies`. Right-clicking the Company filter dropdown (`self.company_combo`) provides a context menu action `❔ Include unconfirmed (?, [...]) sources in filter` (persisted via `MeetingsSettings.get_include_tentative_sources()`).
    - **`ContributionReportDialog`:** `CompanySelectorWidget` includes an `Include unconfirmed (?, [...]) sources` checkbox (`chk_include_tentative`, unchecked by default). When enabled, `ContributionSearchWorker` includes tentative matches and tags them with `Company (?)` in the `Matched Company` column.
  - **Clean Markdown Export & LLM De-Confusion (Two-Line Header Pattern, Approach B):**
    - **Document Metadata Headers (`tdoc_markdown_converter.py`, `llm_exporter.py`):** Authoritative `- **Source:**` lists *only* confirmed co-sponsoring companies (`CompanySanitizer.get_confirmed_source_text`). If tentative sources exist, a dedicated secondary line is appended: `- **Unconfirmed / Carry-over Sources:** <tentative> *(carry-over from earlier revision; unconfirmed co-signers)*`. Standard extractors and LLMs cleanly attribute the document to confirmed sources while preserving historical predecessor provenance without hallucination.
    - **Markdown Summary Tables (`post_meeting_exporter.py`, `markdown_exporter.py`):** In master `index.md`, local WI `index.md`, and meeting summary tables (`Company.md`), `CompanySanitizer.format_source_markdown_table_cell` formats the Source cell cleanly with confirmed sources primary and carry-overs italicized: `Confirmed *(carry-over: Tentative)*` (with markdown pipes and newlines safely escaped).
    - **Local AI Triage Context (`tdoc_triage_worker.py`):** Prompts sent to Ollama format `Source` using `CompanySanitizer.format_source_triage_context`, providing an explicit note that carry-over companies have not confirmed co-signing the revision, preventing 7B/8B local models from misattributing proposals in technical summaries.
    - **Navigable Dossier HTML & PDF (`dossier_exporter.py`):** Uses `CompanySanitizer.format_source_html` in interactive HTML tables and metadata strips, and `format_source_triage_context` in companion PDF evaluation prompt spans.

### 3.16. TDocs Window Parallel Meeting Caching (`tdocs_cacher.py`, `tdocs_agenda_controller.py`, `tdocs_window.py`, `tdocs_menus.py`)
- **Background & Need:** Delegates accessing a meeting in `TDocsWindow` often require all official TDoc archives cached locally on disk prior to traveling, offline sessions, or rapid document comparison. Previously, caching was only accessible from the parent Meetings tab context menu (`📥 Cache TDocs (Docs/)`).
- **Parallel Retrieval Core (`TDocsCacherThread` in `tdocs_cacher.py`):**
  - Scrapes the meeting's official 3GPP `Docs/` directory, extracts valid TDoc `.zip` archives, and downloads missing files concurrently using `ThreadPoolExecutor(max_workers=5)`.
  - Automatically identifies base TDocs from revision names (e.g. `S2-2601234` from `S2-2601234r01.zip`), storing archives into canonical `{local_meeting_dir}/{base_tdoc}/{filename}.zip` folders matching downstream extractors and openers.
  - Emits real-time progress (`progress = pyqtSignal(str)`) throttled during skips to preserve GUI responsiveness.
  - Skips already-downloaded archives (`os.path.exists`) and safely handles non-QObject parents.
- **TDocs Window Integration (`tdocs_agenda_controller.py`, `tdocs_window.py`):**
  - **Dual Menu Entry Points:** Available under **`🔄 Refresh ▾` $\rightarrow$ `📥 Cache All TDocs (Docs/)`** and **`🗂️ Resources ▾` $\rightarrow$ `📥 Cache TDocs (Docs/)`**.
  - **Thread-Safe Controller:** `start_window_tdocs_caching` blocks duplicate runs, manages button state (`⏳ Caching TDocs...` with live count feedback), and displays completion summaries (`Downloaded: X, Skipped: Y, Total: Z`).
  - **Window Lifecycle Cleanliness:** Registered in `TDocsWindow.closeEvent` (`all_threads`) to ensure immediate cooperative cancellation if the window is closed during active downloads.
- **Cache Detection (`is_tdoc_cached` in `tdocs_menus.py`):**
  - Inspects both direct meeting directories (`meeting_dir / base_tdoc`) and legacy cache subdirectories (`meeting_dir / TDocs`), ensuring cached files are immediately recognized as `(Local)` in right-click menus and instant document opening.

### 3.17. Cross-WG, Multi-Meeting Topic Dossier & Briefing Studio (`meta_topics.json`, `topic_manager.py`, `cohort_resolver.py`, `topic_decision.py`, `cross_wg_ls_tracker.py`, `topic_dossier_exporter.py`, `topic_dossier_dialog.py`)
- **Problem Solved:** Major 3GPP initiatives (Ambient IoT, AI/ML Phase 2, 6G Architecture, Network Energy Savings) span across multiple Working Groups (SA2 architecture, RAN2/RAN3 radio & protocols, CT1 NAS/URSP, SA3 security) and across 6–10 consecutive quarterly meeting cycles. Individual single-meeting exports fail to capture the multi-WG evolution or inter-group dialogue.
- **Persistent Meta-Topics (`src/config/meta_topics.json`, `src/config/meta_topics_user.json`, `topic_manager.py`):**
  - Architecture cleanly separates factory presets (`src/config/meta_topics.json`, tracked in Git) from active user modifications (`src/config/meta_topics_user.json`, Git-ignored, mirrored to `~/3GPP_Delegate_Helper/meta_topics_user.json`).
  - Additions, edits, and deletions permanently persist in the user configuration files, completely immune to Git updates, checkouts, or repository wipes.
  - Built-in factory presets can be restored non-destructively in 1 click (`🔄 Restore Built-in Topics...`), safely preserving any custom topics.
  - User and built-in semantic topic definitions (`MetaTopicDefinition`, `MetaTopicRules`).
  - Correlates contributions using: (1) normalized Work Item / Study Item tokens (`ambientiot`, `fs_ambientiot`), (2) WG-specific Agenda Item wildcard patterns (`SA2: 5.14*`, `RAN2: 8.7*`, `CT1: 6.18*`), (3) title/abstract keywords, and (4) manual pinned/excluded TDoc IDs.
  - Features CRUD operations (`list_topics`, `save_topic`, `clone_topic`, `delete_topic`) with atomic JSON persistence and built-in templates.
- **Meeting Cohort Resolution (`cohort_resolver.py`):**
  - Groups multi-WG meetings occurring on staggered dates across quarters into standardized quarterly cohorts (e.g. `2025-Q4 Cycle` aligning SA2 #160 in October with RAN2 #129 and CT1 #146 in November).
  - Supports flexible windowing (`Past N Quarters`, `Quarterly Cycle Range`, or `Date Window`).
- **Unified Decision & Document Model (`topic_decision.py`):**
  - Elevates **RAN Chairman's Notes text agreements & Working Assumptions** and **SA/CT Approved pCRs to TRs & CRs to TSs** into identical first-class standing (`TopicDecision`).
  - Classifies outcomes into `🟢 Text Agreement`, `🟢 Approved pCR`, `🟢 Approved CR`, `🔵 Endorsed WF`, `🟡 Working Assumption`, `🟣 Open Issue / FFS`, and `📬 Agreed LS`.
  - Generates unified chronological Markdown matrix rows (`to_markdown_row`).
- **Cross-WG Liaison Statement (LS) Dialogue Tracker (`cross_wg_ls_tracker.py`):**
  - Traces inter-group questions, guidance, and reply chains (e.g. `SA2 LS OUT` $\rightarrow$ `RAN2 LS IN` $\rightarrow$ `RAN2 reply LS OUT`).
  - Synthesizes `ls_dialogue.md` featuring dialogue thread trees and comprehensive inventory tables with bidirectional dossier hyperlinks.
- **Two-Phase Parallel Exporter & Incremental LLM Delta Feed (`topic_dossier_exporter.py`):**
  - **Comprehensive 3GPP FTP Auto-Download Pipeline (Enabled by Default):** Automatically downloads missing TDoc contributions, revisions, Chairman's Notes archives, and meeting agenda lists directly from 3GPP FTP (`allow_network_download=True`, default checked in `TopicDossierDialog` and `InputsCollectorThread`) to guarantee 100% complete dossiers.
  - Two-phase parallel processing: Phase 1 concurrent pre-download (`ThreadPoolExecutor(max_workers=5)`) probing official candidate routes (`resolve_tdoc_download_url`, `INBOX/Revisions`, `Docs`, `Revisions`) with 12s failover timeouts + Phase 2 in-memory OpenXML conversion to Markdown with figures and Visio diagrams.
  - Multi-candidate resolution checks base TDoc directories, pre-existing `.md` caches (instant copy), `.zip` archives, and 3GPP server URLs. Auto-downloads missing Chairman's Notes for RAN meetings into `Agenda/ChairNotes/` to extract all agreements, and generates structured metadata stubs only for unuploaded/unavailable documents that are seamlessly upgraded on subsequent runs.
  - **Incremental Manifest Tracking (`.export_manifest.json`):** Re-running on existing directories skips already converted files in $<1$ second without clobbering filesystem modification timestamps (preserving vector embeddings in RAG caches).
  - **LLM Delta Feed (`deltas/delta_session_N.md`):** Generates dedicated incremental diff files containing strictly the new decisions, approved pCRs, and incoming LSs added in that session, accompanied by explicit AI prompting guidance.
  - Master index synthesis: `index.md`, `agreements_timeline.md`, `ls_dialogue.md`, and optional `consolidated_briefing.md`.
- **Topic Dossier Studio UI (`topic_dossier_dialog.py`, `ui_tabs.py`):**
  - Accessible via **`MeetingsTab` $\rightarrow$ `📊 Reports ▾` $\rightarrow$ `🌐 Cross-WG Topic Dossier && Briefing Studio...`**.
  - Fully independent, modeless top-level window (`Qt.Window`, `Qt.NonModal`) with window-modal sub-dialogs, allowing delegates to freely interact with `MeetingsTab`, open `TDocsWindow`, and inspect documents concurrently.
  - Modeless 2-column studio:
    - Left: Saved Meta-Topics list with `+ New`, `✏️ Edit`, `📋 Clone`, `🗑️ Delete`, and `🔄 Restore Built-in Topics...`. Automatically manages action button states when topics list is empty, and preserves selection index across deletes.
    - Right: Topic Runner, Scope Mode segmented buttons (`🧭 Pre-Meeting Briefing`, `📑 Post-Meeting Report`, `📚 Complete Feature Dossier`), Cohort Window picker, Incremental folder status badge (`🟢 Existing Corpus` vs `🆕 New Directory`), live progress bar, and 1-click `📂 Open Output Folder`.

### 3.18. Cross-WG Agenda Item (AI) Auto-Discovery & Suggestion Engine (`ai_suggester.py`, `ai_suggestion_dialog.py`, `topic_dossier_dialog.py`)
- **Problem Solved:** When creating cross-WG topic dossiers (e.g. for "Data Framework & Data Transfer"), delegates frequently know the relevant Agenda Item for their primary Working Group (e.g. SA2 `20.6.21`), but have no visibility into how other Working Groups (RAN2, RAN3, SA5, CT4) structure their agendas for the same topic.
- **Scanning & Ranking Engine (`AgendaItemSuggester` in `ai_suggester.py`):**
  - **Single-AI Seeding & Topic Inference (`infer_topic_from_known_ais`):** If a delegate enters an AI for just one Working Group (e.g. `SA2: 20.6.21`), the engine automatically queries cached agendas, resolves the official 3GPP procedure/sub-topic description (`Study on Architecture for 6G System (FS_6G_ARC) - 6G data framework in SA WG2`), extracts Work Item tokens (`fs_6g_arc`), and identifies recurring technical keywords (`data framework`, `data transfer`), seamlessly propagating these across target WGs (RAN2, RAN3, SA5, CT4) without requiring manual keyword typing.
  - Analyzes recent meeting agendas across target WGs by loading pre-parsed JSON caches (`Agenda/*.xlsx.json`) or Excel workbooks.
  - Automatically correlates topic keywords (`clean_kws`) and Work Item tokens (`clean_wis`) against Agenda Item descriptions, TDoc titles, and related WI fields.
  - **Umbrella Noise Suppression:** Broad parent study items (e.g. `FS_6G_ARC`, `FS_6G_Radio`) touch thousands of documents across dozens of sub-topics. If keywords are provided, candidate AIs must have keyword relevance in either the AI description or in TDoc titles, preventing generic Release-20 items (voice, charging, emergency) from polluting topic suggestions.
  - **Generic AI Protection (`GENERIC_AI_CATEGORIES`):** Protects administrative and miscellaneous Agenda Items (e.g., `Corrections`, `Review of incoming LS`, `Drafting`, `Miscellaneous`, `AOB`) from receiving "High" confidence even if papers mention topic keywords, preventing dangerous wildcard contamination (`7.2.2*`, `5.3*`) in topic rules.
  - **Confidence Classification:**
    - `🟢 High`: AI description directly matches keywords (`desc_match`), or high concentration of matching TDoc titles (`title_matches >= 3`) with Work Item correlation on non-generic AIs.
    - `🟡 Medium`: Secondary correlations where contribution titles contain topic keywords, or generic categories with title hits.
    - `⚪ Low`: Low-frequency matches or administrative items.
- **Interactive Multi-WG Suggestion Dialog (`AiSuggestionDialog` in `ai_suggestion_dialog.py`):**
  - **Non-Freezing Background Scanner (`AiSuggestionScanWorker`):** Asynchronous background scanning (`QThread`) with progress bar and real-time status reporting (`Scanning recent RAN2 agendas (2/5)...`), keeping the UI responsive and movable.
  - Multi-tabbed layout per Working Group showing candidate Agenda Items, confidence badges, hit counts (`X/Y TDocs across Z mtgs`), reasons, and sample contribution titles on hover tooltips.
  - Inferred metadata banner: Surfaces discovered procedural context (e.g. `💡 Inferred from known AI: SA2 20.6.21 ('Study on Architecture for 6G System - 6G data framework')`).
  - Quick action shortcuts: `✅ Select High Confidence`, `Select All`, `Clear Selection`.
  - Live preview bar generating canonical rule strings (e.g. `SA2: 20.6.21*; RAN2: 9.3.3.1*, 10.3.3.1*; RAN3: 10.2.3*; SA5: 6.20.6.3*; CT4: 21.5*`).
- **Topic Editor Integration (`TopicEditorDialog` in `topic_dossier_dialog.py`, `topic_pills.py`):**
  - Features dedicated **`🔍 Auto-Suggest AIs...`** action button next to the Agenda Item pattern field.
  - Automatically detects known AI patterns from `edit_ai`, infers topic keywords and Work Item codes, passes them to `AiSuggestionDialog`, and non-destructively merges accepted recommendations into existing rules while auto-populating empty keyword and Work Item fields.
  - **Pill & Tag UI Architecture (`topic_pills.py`):**
    - `WgTogglePill`: Checkable pill chips for Working Groups with group-specific color palettes (SA blue `#DBEAFE`/`#1E40AF`, RAN purple `#EDE9FE`/`#5B21B6`, CT teal `#CCFBF1`/`#0F766E`), dynamically toggling between `✓ WG` and `+ WG`.
    - `TagCloudWidget`: Interactive token input displaying Work Items and Keywords as removable pill chips (`[ token ✕ ]`) with inline entry, supporting comma/Enter ingestion, copy-paste parsing, and backspace removal.
    - `AgendaItemPillsWidget`: Per-WG grouped Agenda Item pattern manager displaying active Working Groups, removable pattern chips (`[ 20.6.21* ✕ ]`), inline `+ Add AI` quick-entry, and dynamic visibility synchronization with Target WG toggles.

### 3.19. Considered Inputs & Decisions Viewer (`TopicInputsViewerDialog` in `topic_inputs_dialog.py`) & Flexible Meeting Timeframe Scope
- **Problem Solved:** Delegates need to inspect, filter, verify, and cross-reference all candidate technical decisions (Chairman's Notes text agreements, Approved pCRs, Approved CRs, Endorsed Way Forwards, Working Assumptions), contributions, and Liaison Statements before or instead of running a full Markdown dossier export. Furthermore, 3GPP delegates operate in synchronized **"Meeting Rounds"** / **"Meeting Cycles"** (e.g. Feb 2026, May 2026, Aug 2026, Nov 2026) and specific calendar date ranges rather than artificial calendar quarters or epidemiological "cohorts".
- **Natural 3GPP Domain Terminology & Flexible Scope Windowing (`cohort_resolver.py`, `topic_dossier_dialog.py`):**
  - Replaces "cohort" across all user-facing dialogs, reports, and summaries with natural standards terminology: **"Meeting Scope && Timeframe"**, **"Synchronized Meeting Rounds"**, **"Meeting Cycles"**, and **"Custom Date Range"**.
  - Provides intuitive preset timeframes:
    - `Past 4 Meeting Rounds (~1 Year)`
    - `Past 6 Meeting Rounds (~1.5 Years)`
    - `Past 8 Meeting Rounds (~2 Years)`
    - `Release 20 Timeframe (2025–2027)`
    - `Release 19 Timeframe (2023–2025)`
    - `Release 18 Timeframe (2021–2023)`
    - `📅 Custom Date Range...` (revealing inline `From:` and `To:` `QDateEdit` calendar controls)
    - `All Available Meetings`
  - `MeetingRoundResolver` (`MeetingCohortResolver`) supports `date_range` filtering against meeting start/end dates.
- **Modeless Visual Inspector (`TopicInputsViewerDialog` in `topic_inputs_dialog.py`):**
  - **Modeless Architecture (`Qt.Window`, `Qt.NonModal`):** Runs independently alongside `MeetingsTab` and the main window, allowing delegates to keep the inspector open while browsing other meetings or reading TDocs in parallel.
  - **Asynchronous Data Harvesting (`InputsCollectorThread`):** Scans all target Working Groups across selected meeting rounds in the background with progress reporting, keeping the GUI responsive and buttery smooth.
  - **Executive KPI Strip:** Live metric counters for Total Inputs, Technical Decisions && Agreements, Inter-WG Liaisons, Discussion Papers, and Participating WGs.
  - **Multi-Faceted Search & Quick-Filter Pill Bar:** Instant full-text search across TDoc IDs, titles, companies, agenda items, and decision texts, combined with an interactive 2-row Quick-Filter Pill Bar (`FilterPillButton`):
    - *Row 1 (Working Groups):* `[ All WGs ]`, `[ SA2 (N) ]`, `[ RAN2 (N) ]`, `[ RAN3 (N) ]`, `[ CT4 (N) ]`, `[ SA5 (N) ]` with group family color badges and live document count indicators.
    - *Row 2 (Categories):* `[ All Inputs ]`, `[ 🟢 Decisions (N) ]`, `[ 📝 Agreements (N) ]`, `[ 📑 Approved pCRs (N) ]`, `[ ✅ Approved CRs (N) ]`, `[ 📬 Liaisons (N) ]`, `[ 📄 Discussions (N) ]`.
    - Single-click toggle immediately refreshes the active tab's table, with `↺ Reset` for 1-click filter clearing.
  - **Tabbed Visualization Panels:**
    - `⚖️ Technical Decisions && Agreements`: Unified chronological register of approved pCRs, CRs, text agreements, and endorsed WFs with status color badges and hover summaries.
    - `📄 All Considered Contributions`: Complete inventory of all scoped TDocs with TDoc ID, WG, Meeting, Date, AI, Type, Status, Title, and normalized Source Company.
    - `📬 Inter-WG Liaison Statements`: Scoped incoming and outgoing LS tracker with direction badges (`📥 LS IN`, `📤 LS OUT`), reply-to/reply-in correlation.
    - `📊 WG && Round Distribution`: High-level multi-group matrix comparing document and decision volume across Working Groups and meeting rounds.
  - **Action Controls:** 1-click `📄 Open Document` (local cache), `🌐 Open on 3GPP Portal`, `📋 Copy Markdown Table` to clipboard, `💾 Export to CSV...`, and `🚀 Proceed to Full Dossier Export...`.
  - **Ultra-Fast Batch Population & Reflow Suppression:** Prevents Qt table layout stutter by batching updates with `setUpdatesEnabled(False)` and `blockSignals(True)`, populating columns in `Interactive` mode, and executing `resizeColumnsToContents()` once at the end. Eliminates per-cell $O(N^2)$ font metric reflows when filtering (reducing response time from ~9.4s down to <0.12s across 1,500+ items). Auto-switches active tabs to match selected category filters (e.g. Decisions $\rightarrow$ Tab 0, Liaisons $\rightarrow$ Tab 2).

### 3.20. Standalone RAN Agreement Markdown Exporter & Decision Provenance Engine (`topic_dossier_exporter.py`, `topic_decision.py`, `topic_inputs_dialog.py`)
- **Standards Architecture Context (RAN vs SA/CT Decision Models):**
  - **SA & CT Working Groups:** Decisions are formal Change Requests (CRs) or Pseudo-CRs (pCRs) where the *full document attachment* is approved into a baseline specification or Technical Report.
  - **RAN Working Groups (RAN1, RAN2, RAN3):** Working group decisions are text agreements hammered out live in the meeting room or offline drafting sessions. *Only the verbatim text recorded in the Chairman's Notes represents formal, binding consensus.* The cited discussion paper is merely marked "Noted" in the meeting report and serves as technical provenance and context.
- **Standalone RAN Agreement Markdown Exporter (`_write_agreement_markdown` in `topic_dossier_exporter.py`):**
  - Synthesizes dedicated, standalone Markdown documents for each Chairman's Notes text agreement or working assumption, located under `{round_cycle}/{wg}_{mtg}/Agreements/Agreement_{tdoc_or_ai}.md`.
  - Structured for both human delegates and LLM digestion:
    - **Metadata Frontmatter Table:** Working Group, Meeting, Date, Agenda Item, Decision Type badge, Status, Target Spec, Work Item, and Chairman's Notes source.
    - **Background Contribution Link:** Relative bidirectional hyperlink to the background discussion paper in `{round_cycle}/{wg}_{mtg}/Discussion/{tdoc}.md` (`[R2-2600942](../Discussion/R2-2600942.md) *(Noted Proposal)*`), accompanied by direct 3GPP Portal download URLs.
    - **Verbatim Consensus Text:** Formatted blockquotes and raw code fences isolating the exact agreed wording from the meeting record.
    - **Topic & Context Traceability:** Traceability section connecting the decision back to the overarching meta-topic and meeting cycle.
  - Master tables (`agreements_timeline.md`, `index.md`, `delta_session_N.md`) and `TopicDecision.to_markdown_row()` automatically link to the standalone agreement document.
- **Decision Provenance Badges & Tooltips (`TopicInputsViewerDialog`):**
  - Visual badges in the Decisions table distinguish between:
    - `🟢 Text Agreement (Notes)`: Tooltip explicitly states: *Binding Consensus Text (Chairman's Notes AI X.Y) — In RAN WGs, only the verbatim text in Chairman's Notes is binding consensus; background proposal was Noted.*
    - `🟢 Approved pCR (Doc)` / `🟢 Approved CR (Doc)`: Tooltip highlights: *Approved Specification Change (TDoc ID) — Full document attachment approved as baseline change for TS/TR X.*

### 3.21. Collapsible Agreement & Decision Detail Reader Drawer (`TopicInputsViewerDialog` in `topic_inputs_dialog.py`)
- **Problem Solved:** Technical decisions from Chairman's Notes often span multiple complex paragraphs, numbered agreement lists, and conditional scoping statements that are truncated inside compact table rows. Delegates need to read the full text with clear typography, inspect its procedural context, and quickly copy the verbatim wording for drafting contributions or sending meeting reports.
- **Embedded Vertical Splitter & Ergonomics:**
  - Integrates a bottom reader drawer (`self.drawer_widget`) coupled to the tabbed tables via `QSplitter(Qt.Vertical)` (`setChildrenCollapsible(False)`), allowing delegates to adjust the vertical boundary smoothly according to content length.
  - Features a collapsible drawer toggle (`▲ Collapse` $\leftrightarrow$ `▼ Expand`) and 1-click **`📋 Copy Verbatim Text`** action with an ephemeral confirmation badge (`📖 Decision & Agreement Detail Reader — 📋 Copied!`).
- **Reactive Selection Synchronization:**
  - **Decisions Table Selection (`_on_decision_selection_changed`):** Renders the decision provenance badge (`🟢 Text Agreement (Chairman's Notes)` or `🟢 Approved pCR (Approved Document)`), metadata attribute line (WG, Meeting, Date, AI, Spec, WI, Background Proposal), consensus nature callout, and the full multi-line verbatim text in a styled monospace container (`Consolas`).
  - **Contributions Table Selection (`_on_contribution_selection_changed`):** Displays document type/status badges, metadata, contribution title in bold, and formatted abstract / secretary remarks.
  - **Liaison Statements Table Selection (`_on_ls_selection_changed`):** Displays direction badge (`📥 LS IN` / `📤 LS OUT`), title, source company, status, and full inter-WG reply threading metadata (`Reply To: S2-2601999`, `Reply In: R2-2602100`).
  - Automatically synchronizes display when switching between tabs (`_on_tab_changed`), or clears when no row is selected.






