# AGENTS.md — 3GPP Tools Developer & AI Agent Guide

This document is the authoritative technical reference for human developers and AI coding agents working on the **3GPP Meeting Tools & Diagram Converter** repository. It outlines project architecture, subsystem responsibilities, coding conventions, threading and execution rules, and security boundaries.

---

## 1. Project Overview

**3GPP Tools** is a comprehensive Python/PyQt5 desktop engineering environment tailored for telecommunications engineers and 3GPP standards delegates. It integrates:
- **PlantUML & Office Integration:** Live PlantUML editing with debounced preview and conversion to native editable Microsoft Visio (`.vsdx`) and PowerPoint shapes.
- **Specification Search & Evolution:** SQLite FTS5 trigram substring search across 3GPP specification releases with "First Added" chronological tracking, cutoff date detection, and OpenXML parsing of multi-part documents.
- **Protocol Inspection & Evolution:** Message and Information Element (IE) diffing across releases for 5GS/EPS NAS, NR/LTE RRC, NGAP, GTPv1-U, PDU Session UP (TS 38.415), and PFCP (TS 29.244).
- **Meeting & Work Item Management:** Offline synchronization of 3GPP meetings, TDocs, dynamic agendas, email discussion summarization, and company contribution audits.
- **Local AI Intelligence (Ollama):** Zero-cost, 100% offline, privacy-guaranteed AI integration for TDoc technical triage, call-flow sequence generation, email stance summarization, and an evidence-grounded 3GPP Research Assistant.

---

## 2. Environment & Execution

### Prerequisites
- **Operating System:** Windows (primary environment, using Win32 COM APIs and Windows shell hooks).
- **Python:** `>= 3.8` (typically executed with Python 3.10 to 3.14 64-bit).
- **Java JRE / JDK:** 64-bit Java 11+ recommended (or legacy Java 8) for PlantUML JAR rendering. Portable JREs are supported via zero-admin configuration.
- **Microsoft Office (Optional but required for Visio/Word features):** Microsoft Visio and Word desktop installations for COM automation. PlantUML editing and specification tools function without Office installed.
- **Local Ollama Daemon (Optional for AI features):** Running locally at `http://127.0.0.1:11434`.

### Installation & Launching
Dependencies are managed via `requirements.txt`:
```powershell
pip install -r requirements.txt
```

Launch the application:
```powershell
# Using the Windows batch launcher:
.\src\Launch.cmd

# Or directly via Python with proper PYTHONPATH:
$env:PYTHONPATH = ".;.\src;$env:PYTHONPATH"
python -m main_tools
```

---

## 3. Repository Layout

```text
3GPP Tools/
├── README.md                      # Comprehensive user & architectural documentation
├── requirements.txt               # Direct runtime dependencies
├── AGENTS.md                      # Root developer and AI agent guide (this file)
├── output/                        # Default export directory for converted diagrams/docs
├── src/
│   ├── Launch.cmd                 # Startup script with PYTHONPATH configuration
│   ├── main_tools.py              # Application entry point, logging, faulthandler setup
│   ├── main_window.py             # DragDropUI main window, tab instantiation, system bar
│   ├── assets/                    # PlantUML JAR, icons, and bundled runtime assets
│   ├── config/                    # JSON user preferences & externalized prompts
│   │   ├── prompts/               # Hot-reloadable prompt templates (.txt)
│   │   └── *.json                 # Subsystem configurations (Ollama, specs, meetings, etc.)
│   ├── core/                      # Shared cross-cutting infrastructure
│   │   ├── ai/                    # OllamaClient, health monitoring, OllamaConfigDialog
│   │   ├── config/                # Global application constants and paths
│   │   ├── network/               # NetworkSession singleton, RFC 1918 proxy bypassing
│   │   ├── ui/                    # Base components, global QSS style, console/queue panels
│   │   ├── utils/                 # Path resolution, Java runtime detection, helpers
│   │   └── queue_manager.py       # Multi-threaded background conversion & sync task queue
│   ├── db/                        # Persistent SQLite databases (3gpp_data.db, etc.)
│   └── modules/                   # Independent feature subsystems (tabs)
│       ├── assistant/             # Evidence-grounded local 3GPP Research Assistant (see AGENTS.md)
│       ├── emails/                # WG reflector email scraping, quote stripping, stance audit
│       ├── meetings/              # Meetings calendar, TDoc browser, AI triage (see AGENTS.md)
│       ├── nas/                   # Protocol evolution matrix, bit-level rendering, parsers (see core/parsing/AGENTS.md)
│       ├── puml2visio/            # PlantUML editor, Live Preview, PPTX/Visio conversion, Java UI
│       ├── spec_search/           # FTS5 trigram search, release matrix, multipart doc parser
│       ├── specifications/        # 3GPP FTP spec archive synchronization & catalogue
│       ├── word_tools/            # Docx splitting by headings, Word COM comparison
│       └── work_items/            # 3GPP Work Item database, scraper, and TDoc correlation
```

---

## 4. Subsystems & Module Guide

### 4.1. Visio Tools (`src/modules/puml2visio`)
- **PlantUML IDE & Editor Ergonomics:** Debounced live SVG rendering via `LivePreviewManager` utilizing `plantuml.jar`. `CodeDropTextEdit` features `PlantUMLHighlighter`, inline `FindReplaceBar`, line numbering, soft tab indentation, comment toggles (`Ctrl+/`), and live error line pinpointing.
- **Snippet & Element Inserter Engine:** Declarative catalog in `plantuml_snippets.py` with dynamic participant resolution (`extract_participants`), selection-aware wrapping, and dual-surface UI (toolbar quick menu & collapsible `ElementsPaletteWidget`).
- **Dual Live Preview Architecture:** Side-by-side in-app embedded SVG viewer (`EmbeddedSvgPreview` with `QSvgRenderer`/`QGraphicsView`, wheel zoom, and pan) and external browser live preview via `LivePreviewManager`.
- **Java Maintenance:** `JavaMaintenanceDialog` inspects Java runtime versions (Legacy vs Modern) and auto-updates PlantUML JARs.
- **AI Call Flow Generator:** Modeless dialog streaming sequence diagram synthesis from textual 3GPP procedure descriptions into the editor via chunked Ollama HTTP streaming. Features customizable 3GPP architectural style directives (box groupings, message IEs, state transitions, alt/opt branches, autonumbering), embedded procedural presets from TS 23.502/38.300, live in-dialog SVG diagram preview tab, automated background syntax validation with 1-click error auto-fix, and multi-turn conversational refinement.
- **Conversion Pipeline:** PPTX to Visio conversion uses an Enhanced Metafile (EMF) bridge to preserve native editable vector shapes with text shrink-wrapping.

### 4.2. Word Tools (`src/modules/word_tools`)
- **Document Slicing:** Splits massive specification `.docx` files by heading level (`Heading 1`, `Heading 2`, etc.) using `python-docx`.
- **Document Comparison:** Automates Microsoft Word COM (`win32com.client`) to generate legal redline comparisons between two `.docx` or `.doc` revisions.

### 4.3. Specifications & Search (`src/modules/specifications`, `src/modules/spec_search`)
- **Unified Inline Filter Engine:** Instant filtering directly above the specifications table with multi-series selection (`CheckableComboBox`), dynamic TSG-to-WG scoping (e.g. SA/RAN/CT), radio tech filters (`5G`, `LTE`, `6G`), document type toggles (`TS`/`TR`), and dismissible active filter pills with 1-click reset.
- **FTS5 Trigram Substring Search:** Utilizes SQLite FTS5 configured with `tokenize="trigram"` in `3gpp_spec_search.db` for arbitrary substring, field, and acronym lookups.
- **Release Matrix & "First Added":** Visual matrix displaying text evolution (`🟢 Added`, `✓ Present`, `🔴 Removed`) across chronological releases.
- **Multi-Part Document Parser:** Automatically chains and sequences multi-part split documents (e.g., `_s00_s04.docx`, `_s05_s08.docx`) into a unified release model using direct `lxml` OpenXML parsing without Word COM overhead.
- **Specification Document Resolution & Clipboard Assist (`spec_doc_resolver.py`):** Centralizes on-disk discovery of specification files (`.docx`, `.doc`, `.pdf`, `.zip`), automatic on-the-fly unzipping (skipping `-rm` revision-mark files), clause-to-split-file targeting (`match_multipart_document_for_clause`), clipboard assistance (copying `Clause X.Y.Z` for fast navigation), and non-blocking background downloads (`SpecDownloadWorker`).
- **Clause-Level Comparison & Word Track Changes (`spec_clause_word_comparator.py`, `spec_clause_diff_dialog.py`):** Unified comparison engine providing an in-app color-coded Visual Diff tab (`+` additions, `-` deletions, line stats), multi-tier AI/LLM analysis prompt synthesis, and background worker (`SpecClauseWordCompareThread`) extracting targeted clause excerpts and automating Word COM `CompareDocuments` with native Track Changes.
- **Cross-Tab Deep Linking & Matrix Navigation:** Full bidirectional signaling between `SpecSearchTab` and `SpecificationsTab` (`navigate_to_specs_tab_requested`, `search_in_spec_search_requested`, `index_in_spec_search_requested`, `batch_index_in_spec_search_requested`), cell double-click / context menu actions, and chronological version navigation (`get_first_added_version`, `get_preceding_version`).
- **Direct Specification Version Indexing into Spec Search:** 1-click ingestion of specific specification releases directly from the specifications table row context menu, Column 0 action button menu (`⋮`), and Column 2 `Get Word ▾` document menu (`Index vX.Y.Z into Spec Search DB` / `Re-index vX.Y.Z`). Features batch multi-specification indexing (`Index selected (N) into Spec Search DB`), sub-millisecond status checking via `is_version_indexed`, automatic detection of existing local Word documents to avoid redundant FTP downloads, and seamless auto-focus/prompt in `SpecSearchTab.focus_spec` when focusing unindexed releases.
- **Specification User Annotations & Tag Search:** Persistent, backwards-compatible custom domain tags and multi-line notes stored in `spec_user_annotations` within `3gpp_data.db`. Supported by full text matching across user tags/comments, interactive tag chips and note indicators, modal editor `EditSpecAnnotationsDialog`, and batch-tagging `BatchTagDialog`.
- **Specification Quick Copy & Dual-Format Clipboard (`spec_menu.py`, `ui_tabs.py`, `spec_info_dialog.py`):** Multi-surface clipboard export (context menu, action menus, dialog footer, `Ctrl+C`) for single and batch-selected specifications. Sets dual-format `QMimeData` (plain text fallback alongside rich HTML hyperlinks for Outlook, Word, and Teams) with non-blocking floating tooltip feedback.
- **Working Group Domain Badges & 1-Click Scoping:** Color-coded domain badges by TSG (SA: indigo, RAN: emerald, CT: amber, default: slate) in the Specification column. Clicking a badge triggers instant filtering to that Working Group without manual dropdown navigation.
- **Specification Markdown Export (`spec_markdown_converter.py`):** Converts 3GPP specifications into high-fidelity Markdown with version-isolated media directories. Supports both Mode 1 (unified document) and Mode 2 (clause-split dossier with `index.md` TOC), automatically sequencing multi-part specs and handling legacy `.doc` archives.
- **Specification Clause Extraction & Multi-Format Clipboard Export (`spec_clause_exporter.py`):** Extracts individual specification clauses on demand directly from the Evolution Matrix context menu or `SpecClauseInspector` control bar:
  - **Native Word Excerpt (.docx):** Non-blocking background worker (`SpecClauseWordExtractorThread`) locates specification documents via `SpecDocResolver`, automatically unzips archives, scans heading hierarchies with `WordHeadingScanner`, routes multi-part split documents (`match_multipart_document_for_clause`), and extracts standalone `.docx` excerpts preserving native formatting and tables via `WordExcerptEngine` (with `python-docx` fallback).
  - **High-Fidelity Markdown Export (.md with Pictures):** Exports standalone Markdown files integrating the repository's `tdoc_markdown_converter` pipeline (`convert_document_to_markdown`). Automatically extracts all pictures, figures, and diagrams into an isolated media directory (`{filename}_media/`), converts vector EMF/WMF drawings to PNG, links embedded Visio sources, renders native Markdown tables, and adds formal 3GPP release metadata headers.
  - **Multi-Format Clipboard Copying:** `copy_clause_to_clipboard` sets dual-format `QMimeData` providing styled rich HTML for Outlook, Microsoft Word, and Teams (including citation header banners and styled paragraphs) alongside plain text fallbacks, GitHub-flavored Markdown quote callouts (`> [!NOTE]`), plain text, and formal 3GPP citations.
- **Spec Search UI Architecture & Ergonomics (`spec_search_tabs.py`, `spec_search_components.py`, `spec_search_models.py`):**
  - **Compact Toolbar & Live Database Health:** Consolidates administrative actions into compact dropdowns (`📥 Import Specs ▾`, `⚙️ Database ▾`) alongside a live database capacity badge (`📦 X Specs · Y Releases Indexed`).
  - **Indexed Specs Tree Card & Instant Filter:** Left card container with instant client-side filtering (`filter_tree`), 1-click `[All]` / `[None]` toggle pills, and responsive sizing avoiding date clipping.
  - **Unified Search Console Card:** High-contrast search input with embedded icon, clear button, fixed date cutoff edit (`125px`, eliminating digit truncation), and executive KPI badge chips (`Results: [ 🔍 query ] [ N Clauses ] [ M Spec(s) ]`).
  - **Sleek Rounded Status Badges (`SpecEvolutionStatusDelegate`):** Custom delegate rendering clean centered rounded pill badges (`🟢 Added`, `⚡ Post-Cutoff`, `✓ Present`, `—`) and soft sky-blue row selection (`#EBF3FC` with `#0F172A` text), eliminating harsh Windows solid blue selection and grid cell rectangular color fills.
  - **Two-Tier Clause Content Inspector:** Tier 1 provides context title, clickable first-added pill, match navigation (`◀ N Match(es) ▶`), 1-click **Layout Toggle** (`⬒ Bottom Split` <-> `◫ Right Pane`), and **Maximize Button** (`⤢` <-> `⤡`). Tier 2 provides a unified, harmonized action toolbar (`📋 Copy ▾`, `✂️ Extract ▾`, `📖 Full Doc ▾`, `🔍 Specs Tab`, `🔍 Diff Prev`, `🤖 LLM Diff`). Includes intelligent heading deduplication preventing repeated clause titles.

### 4.4. Protocol Matrix & Inspector (`src/modules/nas`)
- **Multi-Protocol Support:** Covers 5GS NAS (TS 24.501), EPS NAS (TS 24.301), NR RRC (TS 38.331), LTE RRC (TS 36.331), NGAP (TS 38.413), GTPv1-U (TS 29.281), PDU Session UP (TS 38.415), and PFCP (TS 29.244).
- **Parsing Architecture:** Unified under `ProtocolDocxDispatcher`. For in-depth parsing guidelines, refer to [`src/modules/nas/core/parsing/AGENTS.md`](file:///src/modules/nas/core/parsing/AGENTS.md).
- **Inspection UI & Ergonomics:** Modern Card Frame architecture with live database KPI stats ribbon (`📦 Specs`, `🏷️ Versions`, `📨 Messages · IEs`), safe maintenance dropdown (`⚙️ Manage ▾`), and welcoming empty-state placeholders for both the matrix and inspector panels.
- **Specification Clues & Working Group Context (`nas_dialogs.py`):** Centralized `PROTOCOL_SPEC_CATALOG` resolves 3GPP specification numbers to Working Group badges (`[CT1]`, `[RAN2]`, `[CT4]`), acronyms (`5GS NAS`, `NR RRC`, `PFCP`), full formal 3GPP titles, and plain-English architectural roles. Features release-tagged tree items (`Rel-20 · v20.0.0`), multi-token search by acronym/spec number, quick selection helpers (`All` | `Latest` | `Clear`), and a descriptive banner in `NASVersionSelectDialog`.
- **Protocol Library Item Delegate (`NASMessageItemDelegate`):** Custom styled delegate rendering message names with right-aligned WG/Spec badges and rounded clause pills (`Cl. X.Y.Z`).
- **Evolution Matrix & 1-Click Clipboard (`nas_models.py`, `nas_tabs.py`):** Multi-release comparison table with visual diff coloring (`🟢 Added`, `🟡 Modified`, `🔴 Removed`), inline diff legend, column headers formatted with 3GPP Release tags (`Rel-20\nv20.0.0`), and 1-click **📋 Copy TSV** matrix export for Excel/Word/Email.
- **Inspection UI (`NASInspectorWidget`):** Renders bit-level octet layouts, OpenXML tables with `gridSpan`/`vMerge`, and ASN.1 syntax blocks with field description tables, cross-spec reference detection (TS 24.501 <-> TS 24.301), and a 1-click **📋 Copy Definition** clipboard action.

### 4.5. Meetings & TDocs (`src/modules/meetings`)
- **Subsystem Guide:** For complete module architecture, conversion pipeline details, and filter proxy design, consult [`src/modules/meetings/AGENTS.md`](file:///src/modules/meetings/AGENTS.md).
- **Full-Width Meetings Browser & 2-Row Top Ribbon:** Complete modernization of `MeetingsTab` replacing the cramped vertical right sidebar with a 100% full-width table and clean 2-row horizontal top ribbon:
  - **Row 1 (Primary Navigation & Actions):** Quick launch button (`🚀 Last Meeting`), inline `⚡ TDoc:` jump box with dynamic `📄 Doc` and `🗓️ Mtg` action buttons, right-aligned `🔄 Sync ▾` menu (*Sync All*, *Add / Fetch Meeting...*, *Preferences*), `📊 Reports ▾` menu (*Contributions Report...*, *Export Merged TDocs*), and `⚙️ Settings`.
  - **Row 2 (Filter Engine):** `🔍 Search` (number, city, title), TSG quick-toggle pills (`[All] [SA] [RAN] [CT]`), `Working Group` checkable multi-select combo, Nature (`Regular` / `Ad-Hoc`), Format (`In-Person` / `Electronic`), and Date presets dropdown (`All Dates`, `Upcoming`, `This Year`, `Past 12 Months`, `Custom Range...`) with dynamic calendar pickers.
- **7-Column Table & Domain Visual Hierarchy:**
  - **Working Group Badges (`WGTagDelegate`):** Color-coded domain pills by TSG (Soft Indigo for SA, Soft Emerald for RAN, Soft Amber for CT) using `core.ui.ui_components.get_wg_colors`.
  - **Consolidated English Dates:** Formatted into human-readable ranges (`12 – 16 Oct 2026`) with temporal tooltips (`🟢 Ongoing`, `🔵 Upcoming`, `⚪ Past`).
  - **Preserved Allocated TDoc Range:** Formatted with full WG prefixes (`C4-264009 → C4-264023`, `—` when not yet allocated).
  - **Location Icons:** Clean icon prefixes (`📍 City, Country` vs `🌐 Online (eMeeting)`).
  - **Double-Click Navigation:** Double-clicking any meeting row directly launches the meeting viewer or opens the metadata dialog.
- **Floating Comparison Cart:** Converted from the bulky full-width bottom frame into a compact lower-right status bar pill that **auto-hides completely when empty** and slides in automatically when items are staged into Slot A or Slot B.
- **Destructive Action Relocation:** Moved `Clear All Meetings...` from the main UI into `MeetingsConfigDialog` under a dedicated *⚠️ Database Maintenance (Danger Zone)* section.
- **Three-Phase Sync:** Asynchronously synchronizes meeting lists, TDoc metadata, and dynamic agenda reservations from 3GPP FTP servers.
- **AI TDoc Triage:** Modeless dialog extracting CR cover sheets and redlines to synthesize structured executive summaries (*Objective*, *Key Changes*, *Potential Impacts*).
- **Personal Notes & Stance:** Integrates local SQLite user notes (`user_tdocs.db`) with stance tagging (`⚪ Neutral`, `🔵 My TDoc`, `🟢 Support`, `🔴 Object`, `🟡 Monitor`).
- **Tracked Work Items Quick Filters:** Integrates with the Work Items module via `TDocsFilterProxyModel` and `TDocsWindow` quick-toggle chips (`🎯 Followed WIs`, `⭐ Starred WIs`). Filters meeting contributions by followed/starred WIs with dynamic badge counts and full multi-facet `AND` combination (e.g. Followed WIs + Approved/Agreed TDocs). Includes default support (`include_unassigned_ls_in = True`, toggleable via right-click) to preserve incoming Liaison Statements (LS in) without an assigned Work Item. Resolves acronym variations, WG suffixes (`-ARC`, `-CT`, etc.), and `FS_` prefixes, with fallback to Agenda Item description for discussion papers.
- **Opportunistic Work Item Auto-Enrichment (`meeting_wi_enricher.py`):** Fully decoupled, non-blocking background thread (`MeetingWIEnricherThread`) automatically launched when opening or refreshing any meeting. Performs an off-thread openpyxl bounded scan of the "Related WIs" column for 3GPP Portal hyperlinks (`workitemId=\d+`), checks SQLite (`WorkItemsDatabase.get_existing_wi_codes()`), and asynchronously scrapes missing/unindexed Work Items directly from the 3GPP Portal without freezing or delaying the TDocs table. Employs a persistent JSON audit marker (`<excel>.wi_checked.json`) to guarantee zero overhead on subsequent opens (~0 ms), and dynamically updates tracked tokens and filter badges upon completion. Configurable via `MeetingsConfigDialog`.
- **Post-Meeting Report Generator:** Modeless `PostMeetingExportDialog` and `PostMeetingExporterThread` compiling meeting contributions into a structured, per-release and per-Work Item Markdown dossier (`Rel-XX/<WI>/{LS_IN, pCR, CR, LS_OUT}/`) with collision-free media copying, cross-WI notices, multi-AI listing, executive `index.md` tables, and a default option (`Include unassigned & general LS in`) ensuring incoming LSs without an assigned Work Item are safely preserved under `General/LS_IN/`. Features automated bidirectional LS reply tracking (`Reply in`/`Reply to` with `Revised to` chain resolution) that hyperlinks outgoing reply dossiers directly from the `LS in` status table (`replied to in [S2-XXXX](...)`) and incoming references in the `LS out` table (`approved (reply to [S2-YYYY](...))`), injects reciprocal Markdown callouts (`[!TIP] Outgoing Reply` / `[!NOTE] Reply Liaison Statement`) into dossier headers, and guarantees that any approved reply `LS out` document answering an included `LS in` is exported to disk (routed to `General/LS_OUT/` if its Work Item was not in the export scope) without cluttering the report with unwanted WI sections.
- **Meeting Analytics & Statistics Engine (`exporter_thread.py`, `contribution_stats.py`, `html_dashboard_compiler.py`):** Multi-scope interactive analytics compiling single-meeting dashboards (`Statistics_Report.html`), cross-meeting contribution audits, and multi-tab Excel workbooks (`Statistics_Data.xlsx`). Features unified semantic 3GPP outcome palettes (`Agreed`/`Approved` in emerald/deep green, `Not Treated` in slate grey, `Revised` in indigo), a 6-metric responsive KPI strip (Total TDocs, CR Agreement Rate %, Meeting Treatment Rate %, Document Composition, Active Entities with Joint %, Active Work Items/Topics), document category segmentation (CR vs Discussion vs LS), Louvain alliance network graphs, and dual Chart/Data toggle tables on every card.
- **TDoc Markdown Conversion & Provenance Preservation (`tdoc_markdown_converter.py`, `tdoc_url_resolver.py`):** High-fidelity conversion of 3GPP Word documents (`.docx`, `.doc`) into Markdown with track changes, tables, and diagrams. Features automatic document type revision mode resolution (`resolve_revision_mode_for_type`): CR and pCR documents retain specification change markup (`revisions="markup"` with `[INSERTED]`/`[DELETED]` tags and markup legend), whereas Liaison Statements (`LS in`, `LS out`, `LS`) are recognized as plain correspondence documents and exported in clean `accept` mode (`revisions="accept"`, omitting change marks and markup legends). Multi-tier stale checking automatically upgrades previously cached markup dossiers of LS documents. Automatically resolves and embeds official 3GPP server HTTPS download URLs (`resolve_tdoc_download_url`) into standalone Markdown dossier headers (`- **3GPP Server URL:**`), Post-Meeting report index tables (`[Dossier](...) · [📥 3GPP](...)`), AI TDoc Triage dialog badges/clipboard summaries, and Markdown report tables, preserving permanent document provenance across mixed downstream archives and external readers.
  > [!CAUTION]
  > **DO NOT MODIFY `src/modules/meetings/core/tdoc2md.py`**. This file is a drop-in core imported directly from an external standalone repository (`3gpp-tdoc-markdown-exporter`). All application-specific integrations, office adapters, or custom pipelines MUST reside exclusively in `tdoc_markdown_converter.py` or calling modules.

### 4.6. Research Assistant (`src/modules/assistant`)
- **Subsystem Guide:** For complete architecture, semantic tool registry, and execution rules, consult [`src/modules/assistant/AGENTS.md`](file:///src/modules/assistant/AGENTS.md) (and [`README.md`](file:///src/modules/assistant/README.md) for full design rationale).
- **Architecture:** Bounded, evidence-grounded agent architecture connecting a local LLM to application knowledge services.
- **Authority Boundary:** Strictly **`KNOWLEDGE_READ`**. The model has no SQL access, shell execution, or database write permissions.
- **Semantic Tool Set:** Model interaction is restricted to predefined semantic tools (`find_specifications`, `get_specification`, `search_spec_text`, `get_spec_clause`, `list_protocols`, `list_protocol_messages`, `find_protocol_message`, `find_protocol_ie`).

### 4.7. Work Items (`src/modules/work_items`)
- **Parallel Scraper & Storage:** Concurrently synchronizes 3GPP Work Items across all TSGs/WGs from dynamic reports into `3gpp_data.db` (`work_items`, `wi_group_map`, `wi_remarks`).
- **Direct WI Addition & Portal Scraper:** `AddWorkItemDialog` and background worker `ManualWIFetcherThread` enable delegates to add and scrape any individual Work Item by ID or 3GPP Portal URL without a full database sync. Utilizes `extract_wi_code` for robust format normalization, `fetch_wi_details_from_portal` for comprehensive metadata and remarks extraction (including shorthand WG mapping via `PORTAL_TO_WG_MAP`), and atomic persistence via `WorkItemsDatabase.upsert_single_work_item()`.
- **Targeted Updating:** `TargetedWIScraperThread` performs fast, targeted updates for a specific list of WI codes without triggering a full multi-WG scrape.
- **Personal Tracking (Starred & Following):** Decoupled from `3gpp_data.db` and persisted in `src/config/work_items_config.json` via `WorkItemsSettings`. Supports dual independent flags:
  - `starred`: Personal interest / watchlist.
  - `following`: Assigned delegation responsibility (team work-split).
- **Interactive UI & Filtering:** `WorkItemsTab` incorporates a dedicated 1-click `TrackingDelegate`, filter chips (`⭐ Starred`, `🎯 Following`), `➕ Add WI` toolbar and context menu actions, batch context menu actions, and in-dialog toggling inside `WorkItemInfoDialog`. `WorkItemsDatabase` optimizes tracking filters using SQLite primary-key `IN` clauses with safe empty-set evaluations (`1=0`).

---

## 5. Architectural Rules & Constraints for AI Agents

When modifying or extending this codebase, adhere to the following architectural rules:

### 5.1. Threading & UI Responsiveness
- **Never block the Qt GUI thread:** Never execute network requests, database scraping, large document parsing, or LLM inference on the main thread.
- **Worker Conventions:** Use `QThread` subclasses or submit tasks to `QueueManager`.
- **Cooperative Cancellation:** Workers must periodically check `self.isInterruptionRequested()` or application `CancellationToken` instances. Never call `QThread.terminate()` as it can corrupt SQLite locks and C runtime states.
- **Safe Signal Delivery:** Pass data back to UI widgets via Qt signals (`pyqtSignal`) rather than directly manipulating widgets from background threads.

### 5.2. Win32 COM Automation Safety
- **Apartment Initialization:** Any background thread using `win32com.client` must call `pythoncom.CoInitialize()` before accessing COM objects and `pythoncom.CoUninitialize()` in a `finally` block.
- **Resource Cleanup:** Always wrap Word/Visio COM operations in `try...finally` blocks to explicitly call `app.Quit()` and delete references (`del app`), preventing orphaned background `WINWORD.EXE` or `VISIO.EXE` processes.
- **Legacy Conversion:** Prefer direct `lxml` extraction for reading DOCX files; only invoke COM when native Word layout/rendering is strictly necessary.

### 5.3. Local AI & Corporate Proxy Bypassing
- **Strictly Offline & Local:** Do not add cloud LLM fallbacks. All inference must target the local Ollama instance (`127.0.0.1:11434`).
- **Loopback Proxy Bypass:** Use `core.network.session.get_ai_session()` for AI HTTP calls. This factory automatically detects RFC 1918 loopback addresses (`is_local_address()`) and configures `trust_env=False` to prevent corporate intercepting proxies (e.g., Zscaler, BlueCoat) from hijacking local inference traffic.
- **Prompt Decoupling:** Decouple prompt text into `src/config/prompts/` to allow live hot-reloading based on file `mtime` without application restarts.

### 5.4. Database Integrity & Multi-Store Isolation
The project intentionally separates domain data into distinct SQLite databases located in `src/db/`:
1. `3gpp_data.db`: Meeting schedules, TDocs, specification catalogue metadata, work items.
2. `3gpp_protocol_data.db`: Structured protocol models, messages, and IE hierarchies.
3. `3gpp_spec_search.db`: Extracted clause text with FTS5 trigram indexes.
4. `user_tdocs.db`: User notes, personal review status, and triage archives.

- **Cross-Database Queries:** Do not couple queries across these databases using SQLite `ATTACH` unless strictly isolated inside a dedicated maintenance script.
- **WAL Compaction:** After massive writes or deletions, perform compaction via `PRAGMA wal_checkpoint(TRUNCATE)` and `VACUUM` (available via `DatabaseMaintenanceDialog`).

### 5.5. Logging & Diagnostics
- Use standard `logging.getLogger(__name__)`.
- Log records are mirrored to `3gpp_tools.log` (with timestamps/levels) and to the in-app `ConsolePanel` (clean message text via `GuiLogHandler`).
- Avoid dumping massive raw payloads (e.g., complete multi-megabyte specification clauses) into logs; log lengths and concise identifiers instead.

### 5.6. UI Consistency, Styling & Code Reuse
- **Reuse Common Styles:** Prefer shared design tokens and style constants in `src/core/ui/ui_components.py` (e.g., `BUTTON_STYLE_TOOLBAR_SECONDARY`, `CARD_FRAME_STYLE`, `BADGE_STYLE_*`, `TABLE_STYLE_CLEAN`) over ad-hoc inline stylesheets.
- **Support Interactive States:** Checkable buttons or toggleable controls must properly define interactive pseudo-states (such as `QPushButton:checked` and `QPushButton:checked:hover`) using the application's unified palette (`#EBF3FC` background, `#1E5C99` / `#395396` border and text).
- **Functionality & Component Reuse:** Always reuse existing core infrastructure, dialogs, network managers, and utility functions across feature modules rather than duplicating logic.

### 5.7. Protected Drop-In Files & External Components
- **Strictly Protected File (`src/modules/meetings/core/tdoc2md.py`):** Under **NO** circumstances should `src/modules/meetings/core/tdoc2md.py` be modified. It is maintained externally in a separate repository (`3gpp-tdoc-markdown-exporter`) and dropped into this project verbatim. If application-specific adaptations, office tools, wrappers, or custom pipelines are required, they must reside exclusively in `tdoc_markdown_converter.py` or surrounding application code.

### 5.8. Qt UI Text & Ampersand Mnemonic Escaping
- **Literal Ampersands Must Be Escaped (`&&`):** In Qt/PyQt, controls that support keyboard shortcuts and mnemonics (`QPushButton`, `QToolButton`, `QCommandLinkButton`, `QCheckBox`, `QRadioButton`, `QGroupBox`, `QAction`, `QMenu`, `QMenuBar`, `QTabBar`, and `QTabWidget`) interpret a single `&` as an accelerator mnemonic prefix.
  - When a single `&` is used in a string literal (e.g., `"Save & Close"`, `"Followed & Starred"`, `"Work Items & Enrichment"`), Qt suppresses the character and either renders an underscore `_` or underbars the succeeding character on Windows.
  - **Always escape literal ampersands as `&&`** across all user-facing labels, button texts, group titles, menu items, action labels, and tab titles (e.g. `"Save && Close"`, `"Followed && Starred"`, `"🎯 Work Items && Enrichment"`, `"Tags && Notes"`).
  - Never use a single `&` in UI labels unless deliberately establishing an `Alt+<Key>` accelerator shortcut.
  - This rule is enforced automatically across all `src/` modules via `tests/test_ui_ampersand_escaping.py`.

### 5.9. Documentation Synchronization & Commit Hygiene
- **Documentation as Code:** Never treat documentation as an afterthought. When adding new capabilities, refactoring architectural patterns, modifying configurations, or updating UI controls, update corresponding documentation (`README.md`, `AGENTS.md`, and module guides) in the same work unit.
- **Commit Completed Work Units:** Do not accumulate large batches of uncommitted, disparate changes across multiple subsystems. Once a feature or bugfix is verified and documented, commit it cleanly with a concise, descriptive commit message.
- **Never Execute `git push` (Manual User Review Boundary):** AI agents must **never** run `git push`. Pushing commits to the remote repository is reserved exclusively for the human developer. This guarantees the user has full opportunity to inspect commit logs, review diffs, test changes, and amend commits before publishing upstream.

---

## 6. Development & Verification Workflow

When adding new features or fixing bugs:
1. **Verify Python Syntax & Imports:** Ensure imports resolve cleanly with `PYTHONPATH=".;src"`.
2. **Preserve Comments & Docstrings:** Do not strip existing architectural comments or docstrings from untouched code.
3. **Check Submodule Guides & Protected Files:** When touching protocol parsing, consult [`src/modules/nas/core/parsing/AGENTS.md`](file:///src/modules/nas/core/parsing/AGENTS.md). When touching the Assistant, consult [`src/modules/assistant/AGENTS.md`](file:///src/modules/assistant/AGENTS.md). When touching meetings or TDoc markdown conversion, consult [`src/modules/meetings/AGENTS.md`](file:///src/modules/meetings/AGENTS.md) (and **NEVER edit `src/modules/meetings/core/tdoc2md.py`**).
4. **UI Mnemonic Escaping Verification:** When adding or modifying buttons, actions, checkboxes, menus, or groupboxes containing ampersands, always use `&&` (not `&`). Verify by running `python -m unittest tests/test_ui_ampersand_escaping.py`.
5. **Manual Verification:** Because the application is a graphical Qt desktop environment, test modified tabs or workers by running `python -m main_tools` and verifying that the Qt event loop starts and stops cleanly without console tracebacks or thread leaks.
6. **Synchronize Documentation (Docs-as-Code):** Whenever introducing new features, configuration parameters, data models, or UI workflows, immediately update the relevant documentation:
   - Root [`AGENTS.md`](file:///AGENTS.md) (architecture overview, module guides, and safety constraints).
   - Root [`README.md`](file:///README.md) (feature lists, user workflows, and usage instructions).
   - Subsystem guides (e.g. [`src/modules/meetings/AGENTS.md`](file:///src/modules/meetings/AGENTS.md), [`src/modules/assistant/AGENTS.md`](file:///src/modules/assistant/AGENTS.md), [`src/modules/nas/core/parsing/AGENTS.md`](file:///src/modules/nas/core/parsing/AGENTS.md)).
   - In-code docstrings, type annotations, and module header comments.
   Keeping documentation synchronized is vital: in this repository, documentation acts as the ground-truth system context and guardrails for both human developers and AI coding agents.
7. **Commit Changes Atomically (Local Only):** Once a logical feature or bugfix is implemented, verified, and documented:
   - Create an atomic git commit grouping the implementation, relevant tests, and documentation updates together.
   - Use clear conventional commit messages (e.g., `feat(module): ...`, `fix(module): ...`, `docs(agents): ...`).
   - Avoid leaving lingering uncommitted changes across unrelated modules or lumping disparate tasks into monolithic commits.
   - **Do NOT execute `git push`:** Pushing to the remote repository is strictly left to the developer, providing a final review and verification boundary before publishing changes.
