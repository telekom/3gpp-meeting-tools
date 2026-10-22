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
Dependencies are managed via `pyproject.toml` and `requirements.txt`:
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
├── pyproject.toml                 # Package definition & CLI entry points
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
│       ├── assistant/             # Evidence-grounded local 3GPP Research Assistant
│       ├── emails/                # WG reflector email scraping, quote stripping, stance audit
│       ├── meetings/              # Meetings calendar, TDoc browser, AI triage, sync workers
│       ├── nas/                   # Protocol evolution matrix, bit-level rendering, parsers
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
- **Unified Inline Filter Engine:** Instant, zero-modal filtering directly above the specifications table. Features a structured 4-row hierarchy (Administration, Primary Search & Scope, Faceted Dimensions, Active Dismissible Badges).
- **Multi-Series & Faceted Filtering:** Integrates `CheckableComboBox` for multi-series selection (e.g. Series 23 and 24 simultaneously) supported by comma-separated query handling in `SpecsDatabase.search_files()`. Features dynamic TSG-to-WG scoping (selecting `SA`, `RAN`, or `CT` automatically filters relevant working groups), radio tech dropdown (`5G`, `LTE`, `6G`, etc.), and document type toggles (`All`, `TS`, `TR`).
- **Active Filter Pills & 1-Click Removal:** Dynamically renders dismissible tag chips for each active constraint (`[Group: SA5 ✕]`, `[Tech: 5G ✕]`, `[Series: 23, 24 ✕]`) with hover visual feedback, allowing delegates to drop individual criteria without resetting the entire view, alongside a 1-click `❌ Reset All` button.
- **FTS5 Trigram Substring Search:** Utilizes SQLite FTS5 configured with `tokenize="trigram"` in `3gpp_spec_search.db` for arbitrary substring, field, and acronym lookups.
- **Release Matrix & "First Added":** Visual matrix displaying text evolution (`🟢 Added`, `✓ Present`, `🔴 Removed`) across chronological releases.
- **Multi-Part Document Parser:** Automatically chains and sequences multi-part split documents (e.g., `_s00_s04.docx`, `_s05_s08.docx`) into a unified release model using direct `lxml` OpenXML parsing without Word COM overhead.
- **Specification Document Resolution & Clipboard Assist (`spec_doc_resolver.py`):** Centralizes on-disk discovery of specification files (`.docx`, `.doc`, `.pdf`, `.zip`), automatic on-the-fly unzipping (skipping `-rm` revision-mark files), clause-to-split-file targeting (`match_multipart_document_for_clause`), clipboard assistance (copying `Clause X.Y.Z` for fast navigation), and non-blocking background downloads (`SpecDownloadWorker`).
- **Clause-Level Comparison & Word Track Changes (`spec_clause_word_comparator.py`, `spec_clause_diff_dialog.py`):** Unified comparison engine providing an in-app color-coded Visual Diff tab (`+` additions, `-` deletions, line stats), multi-tier AI/LLM analysis prompt synthesis, and background worker (`SpecClauseWordCompareThread`) that extracts targeted clause excerpts via `WordExcerptEngine`, applies corporate sensitivity labels, and automates Microsoft Word COM `CompareDocuments` to generate focused comparisons with native Track Changes between releases.
- **Cross-Tab Deep Linking & Matrix Navigation:** Full bidirectional signaling between `SpecSearchTab` and `SpecificationsTab` (`navigate_to_specs_tab_requested`, `search_in_spec_search_requested`, `index_in_spec_search_requested`), cell double-click / context menu actions, and chronological version navigation (`get_first_added_version`, `get_preceding_version`).
- **Specification User Annotations & Tag Search:** Persistent, backwards-compatible custom domain tags and multi-line notes stored in dedicated table `spec_user_annotations` within `3gpp_data.db`. Supported by `SpecsDatabase` methods (`get_spec_annotations`, `set_spec_annotations`, `add_tag_to_specs`, `get_all_unique_tags`), full text matching across user tags and comments in `search_files()`, interactive tag chips and note indicators in `SpecificationsTab` table, modal editor `EditSpecAnnotationsDialog`, batch-tagging `BatchTagDialog`, and dedicated annotations card in `SpecInfoDialog`.
- **Specification Markdown Export (`spec_markdown_converter.py`):** Converts 3GPP specifications into high-fidelity Markdown with version-isolated media directories (`<stem>_media/` or `<stem>_clauses/media/`). Supports both Mode 1 (unified document `<stem>.md`) and Mode 2 (clause-split dossier `<stem>_clauses/` with `index.md` table of contents). Automatically sequences multi-part specifications (`_s00_s04`, etc.) and seamlessly handles legacy `.doc` archives.

### 4.4. Protocol Matrix & Inspector (`src/modules/nas`)
- **Multi-Protocol Support:** Covers 5GS NAS (TS 24.501), EPS NAS (TS 24.301), NR RRC (TS 38.331), LTE RRC (TS 36.331), NGAP (TS 38.413), GTPv1-U (TS 29.281), PDU Session UP (TS 38.415), and PFCP (TS 29.244).
- **Parsing Architecture:** Unified under `ProtocolDocxDispatcher`. For in-depth parsing guidelines, refer to [`src/modules/nas/core/parsing/AGENTS.md`](file:///src/modules/nas/core/parsing/AGENTS.md).
- **Inspection UI:** Renders bit-level octet layouts, OpenXML tables with `gridSpan`/`vMerge`, and ASN.1 syntax blocks with field description tables.

### 4.5. Meetings & TDocs (`src/modules/meetings`)
- **Three-Phase Sync:** Asynchronously synchronizes meeting lists, TDoc metadata, and dynamic agenda reservations from 3GPP FTP servers.
- **AI TDoc Triage:** Modeless dialog extracting CR cover sheets and redlines to synthesize structured executive summaries (*Objective*, *Key Changes*, *Potential Impacts*).
- **Personal Notes & Stance:** Integrates local SQLite user notes (`user_tdocs.db`) with stance tagging (`⚪ Neutral`, `🔵 My TDoc`, `🟢 Support`, `🔴 Object`, `🟡 Monitor`).
- **Tracked Work Items Quick Filters:** Integrates with the Work Items module via `TDocsFilterProxyModel` and `TDocsWindow` quick-toggle chips (`🎯 Followed WIs`, `⭐ Starred WIs`). Filters meeting contributions by followed/starred WIs with dynamic badge counts and full multi-facet `AND` combination (e.g. Followed WIs + Approved/Agreed TDocs). Resolves acronym variations, WG suffixes (`-ARC`, `-CT`, etc.), and `FS_` prefixes, with fallback to Agenda Item description for discussion papers.
- **Post-Meeting Report Generator:** Modeless `PostMeetingExportDialog` and `PostMeetingExporterThread` compiling meeting contributions into a structured, per-release and per-Work Item Markdown dossier (`Rel-XX/<WI>/{LS_IN, pCR, CR, LS_OUT}/`) with collision-free media copying, cross-WI notices, multi-AI listing, and executive `index.md` tables.
- **TDoc Markdown Conversion:** High-fidelity conversion of 3GPP Word documents (`.docx`, `.doc`) into Markdown with track changes, tables, and diagrams.
  > [!CAUTION]
  > **DO NOT MODIFY `src/modules/meetings/core/tdoc2md.py`**. This file is a drop-in core imported directly from an external standalone repository (`3gpp-tdoc-markdown-exporter`). All application-specific integrations, office adapters, or custom pipelines MUST reside exclusively in `tdoc_markdown_converter.py` or calling modules.

### 4.6. Research Assistant (`src/modules/assistant`)
- **Architecture:** Bounded, evidence-grounded agent architecture connecting a local LLM to application knowledge services.
- **Authority Boundary:** Strictly **`KNOWLEDGE_READ`**. The model has no SQL access, shell execution, or database write permissions.
- **Semantic Tool Set:** Model interaction is restricted to predefined semantic tools (`find_specifications`, `get_specification`, `search_spec_text`, `get_spec_clause`, `list_protocols`, `list_protocol_messages`, `find_protocol_message`, `find_protocol_ie`).
- **Detailed Design:** See [`src/modules/assistant/README.md`](file:///src/modules/assistant/README.md).

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

---

## 6. Development & Verification Workflow

When adding new features or fixing bugs:
1. **Verify Python Syntax & Imports:** Ensure imports resolve cleanly with `PYTHONPATH=".;src"`.
2. **Preserve Comments & Docstrings:** Do not strip existing architectural comments or docstrings from untouched code.
3. **Check Submodule Guides & Protected Files:** When touching protocol parsing, consult [`src/modules/nas/core/parsing/AGENTS.md`](file:///src/modules/nas/core/parsing/AGENTS.md). When touching the Assistant, consult [`src/modules/assistant/README.md`](file:///src/modules/assistant/README.md). When working with meetings or Markdown conversion, **NEVER edit `src/modules/meetings/core/tdoc2md.py`**.
4. **Manual Verification:** Because the application is a graphical Qt desktop environment, test modified tabs or workers by running `python -m main_tools` and verifying that the Qt event loop starts and stops cleanly without console tracebacks or thread leaks.
