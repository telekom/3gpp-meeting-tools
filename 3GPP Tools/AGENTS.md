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
├── docs/                          # Technical references, audits & porting roadmaps
│   └── LEGACY_FUNCTIONALITIES.md  # Legacy Meeting Helper functionality audit & porting guide
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
- **Snippet & Element Inserter Engine (`plantuml_snippets.py`, `preset_dialog.py`):** Declarative catalog with dynamic participant resolution (`extract_participants`), selection-aware wrapping, and dual-surface UI (toolbar quick menu & collapsible `ElementsPaletteWidget`). Fully externalized to JSON (`src/config/plantuml_presets.json`) with Git-ignored user customization support (`src/config/plantuml_presets_user.json`). Features standalone "UE" presets (`participant "UE" as UE`, `actor "UE" as UE`, `Dual UEs`, `UE Subcomponents`) and individual 3GPP Network Functions (gNB, AMF, SMF, UPF, UDM, AUSF, PCF, NRF, NSSF, NEF, UDR, AF, SEPP, eNB, MME, SGW, PGW, HSS), in-app `UserPresetDialog` for code-free preset creation/editing, and hot reload without application restarts.
- **Dual Live Preview Architecture:** Side-by-side in-app embedded SVG viewer (`EmbeddedSvgPreview` with `QSvgRenderer`/`QGraphicsView`, wheel zoom, and pan) and external browser live preview via `LivePreviewManager`.
- **Java Maintenance:** `JavaMaintenanceDialog` inspects Java runtime versions (Legacy vs Modern) and auto-updates PlantUML JARs.
- **AI Call Flow Generator:** Modeless dialog streaming sequence diagram synthesis from textual 3GPP procedure descriptions into the editor via chunked Ollama HTTP streaming. Features customizable 3GPP architectural style directives (box groupings, message IEs, state transitions, alt/opt branches, autonumbering), embedded procedural presets from TS 23.502/38.300, live in-dialog SVG diagram preview tab, automated background syntax validation with 1-click error auto-fix, and multi-turn conversational refinement.
- **PowerPoint Direct Presentation Append (`powerpoint_converter.py`, `select_presentation_dialog.py`, `ui_tabs.py`):** Enables delegates to append exported PlantUML diagrams directly as a new slide at the end of their currently open PowerPoint slide deck via **"Add to Open PowerPoint Presentation..."** in the Export menu. Automatically discovers running PowerPoint presentations via COM (`get_open_powerpoint_presentations`), shows `SelectPresentationDialog` when multiple presentations are open (with active deck preselected and detailed slide count / path metadata), provides graceful fallback confirmation if no presentation is open, groups unpacked shapes into native editable Office vector shapes, centers and scales according to slide aspect ratio (16:9 widescreen or 4:3), and embeds PlantUML source code into the slide's Speaker Notes.
- **Visio & Office Conversion Studio (`batch_convert_tab.py`):** Multi-card architecture replacing the legacy single-box drop target with:
  - **System Diagnostics & Configuration Ribbon:** Real-time COM registration status for Microsoft Visio and PowerPoint, clickable PlantUML/Java runtime engine badge, and active export directory controls (`output/`).
  - **Card 1: Batch PlantUML Converter:** Multi-format conversion (`.vsdx`, `.pptx`, `.svg`, `.txt`), options (source code embedding, auto-open), compact drop target with file/folder browse buttons, and staging queue with automatic diagram type detection (`Sequence`, `Class`, `Activity`, `State`, `Component`).
  - **Card 2: Office Format Bridge:** Lossless bidirectional PowerPoint (`.pptx`) ⇄ Visio (`.vsdx`) conversion via Enhanced Metafile (EMF) bridging and text shrink-wrapping.
  - **Card 3: Visio PlantUML Extractor & Reverse Engineering:** Instant headless OpenXML extraction (<5ms) of embedded PlantUML source code from `.vsdx` files with syntax-highlighted editor preview, clipboard copy, and 1-click **"📝 Open in PlantUML IDE Tab"** bidirectional routing.
  - **Card 4: 3GPP Stencils Reference:** Interactive catalog of standardized 5GC, NG-RAN, EPS, and IMS Network Function presets with category filter, search bar, user-defined presets support (`⭐ User`), 1-click preset editor, copy, and editor insertion.
  - **Recent Exports Strip:** Persistent history drawer showing recently generated diagrams with 1-click `👁️ Open`, `📂 Show in Folder`, and `🔗 Copy Path` actions.

### 4.2. Word Tools (`src/modules/word_tools`)
- **Modern 2-Tier Header Ribbon & Segmented Navigation (`word_tabs.py`, `word_settings_dialog.py`):** Modernized `WordExtractorTab` replacing the permanent LibreOffice warning banner and cramped dropdown with a 2-tier ribbon: Row 1 features a live active Word monitor pill (`🪟 Active Word`), 3GPP macro template badge (`3gpp_70.dotx`), LibreOffice status pill, and a `⚙️ Word Settings ▾` dropdown menu (with `WordSettingsDialog` configuring corporate sensitivity labels, default company, and templates). Row 2 features styled segmented navigation pills (`Compare Documents`, `Word → Markdown`, `3GPP CR && Templates`, `Sanitizer && Revisions`, `Slice && Excerpt`, `Format Converter`, `Visio && Media Extractor`).
- **Standardized Document Selector (`DocumentSelectorPane 2.0`):** Reusable multi-source document input widget supporting local files, active foreground Word COM documents, and remote/SharePoint URLs. Features click-to-browse drop labels, dedicated file browsing buttons, selection reset (`✕ Clear`), and live file size/metadata indicators.
- **Document Comparison & Slot Swap (`word_comparator.py`, `word_tabs.py`):** Automates Microsoft Word COM (`win32com.client`) to generate native Track Changes document comparisons between two `.docx` or `.doc` revisions. Features 1-click **"⇄ Swap A ↔ B"** slot swapping between original and revised documents, live foreground Word inspection (`get_active_word_document`), non-destructive snapshotting (`SaveCopyAs`), and configurable comparison options (formatting change detection, keep source documents open).
- **3GPP Contribution & Template Workbench (`cr_pcr_dialog.py`, `styles_3gpp.py`, `word_tabs.py`):** First-class dedicated workbench in `WordExtractorTab` empowering delegates to launch formal Change Request (CR) or Pseudo-CR (pCR) authoring from any baseline specification document or active Microsoft Word draft. Features heuristic spec/version detection, dynamic 3GPP macro ribbon management (`3gpp_70.dotx` 1-click install/uninstall, Microsoft Templates folder navigation), and direct links to official authoritative forms on 3GPP FTP (`Template_3GPP_CR.docx`, `Pseudo-CR-Form-v0-1.zip`).
- **Interactive Heading Excerpt & Chapter Slicing (`docx_splitter.py`, `word_excerpt_extractor.py`, `word_tabs.py`):** Hierarchical clause tree browser displaying visual heading levels (`[H1]`, `[H2]`, `[H3]`), instant filter matching with live match counter (`N match(es)`), quick selection shortcuts (`Select All`, `Select Level 1 Only`, `Clear Selection`), and subtractive clause slicing by chapter prefix and depth.
- **Visio Extraction & Format Conversion Panels (`word_extractor.py`, `word_converter.py`, `word_tabs.py`):** Modernized Visio extraction and multi-format conversion cards equipped with `DocumentSelectorPane`, automatic engine detection badges (`🟢 Microsoft Word COM` vs `🟠 LibreOffice`), and completion result frames with 1-click **"📂 Open Output Folder"** shortcuts.
- **Word to Markdown Conversion (`word_converter.py`, `word_tabs.py`):** Fast, headless conversion of arbitrary Word documents (`.docx`, `.docm`, legacy binary `.doc`) to high-fidelity Markdown with revision tracking (`Inline Markup`, `Accept All Changes`, `Reject All Changes`), automatic extraction of figures and embedded drawings into companion media folders (`{doc_name}_media/`), rasterization of vector EMF/WMF drawings to PNG, and extraction/linking of native Visio (`.vsdx`) diagrams. Supported through both a dedicated operation card in `WordExtractorTab` (with `DocumentSelectorPane` supporting 1-click **"🪟 Use Active Word"** live drafting snapshots, post-conversion folder/file launch, clipboard copy, and collapsible in-app preview) and general format conversion (`Target Format: MD`).
- **Document Sanitizer, Proofing Language & Revision Hub (F-07, `document_sanitizer.py`, `author_renamer.py`, `word_tabs.py`):** Pure Python OpenXML (`.docx` / `.docm`) ZIP package level engine providing unified pre-flight document sanitation for 3GPP delegates. Features: (1) **Proofing Language Unification:** Normalizes language across all document stories (body, tables, headers, footers, footnotes, endnotes, comments, and DrawingML canvas shapes) and styles (`styles.xml`, `docDefaults`), preventing newly typed text from reverting; (2) **Surgical Language Revision Purge:** Automatically detects and purges trivial Track Changes language formatting balloons (`<w:rPrChange>` tracking `<w:lang>`), while cleanly synchronizing compound revisions so text insertions/deletions and technical edits remain pristine; (3) **Revision Author Consolidation & Anonymizer:** Replaces disparate reviewer identities with a unified corporate signature (`'Deutsche Telekom'`, `'Editor'`, `'Rapporteur'`, `'Anonymous'`), synchronizes comment initials, and updates Dublin Core metadata (`dc:creator`, `cp:lastModifiedBy`); (4) **Executive Audit & KPI Dashboard:** Single-pass (<150ms) audit displaying detected language breakdown, language revision count, total insertions/deletions, and comment metrics; (5) **1-Click Presets:** Instant configuration presets for `3GPP Contribution Prep`, `Fix Language Only`, and `Full Anonymize`. Integrated into `WordExtractorTab` with the `🛠️ Sanitizer && Revisions` navigation pill, `DocumentSelectorPane` multi-source input (with automatic active Word COM detection & safe close/reopen), and automatic `.bak` backups.

- **Enterprise Microsoft Information Protection (MIP) / Purview Sensitivity Labeling (F-06, `sensitivity_label.py`, `batch_label_dialog.py`, `word_settings_dialog.py`):** Full compliance engine preventing corporate Data Loss Prevention (DLP) blocks when saving or emailing 3GPP Word contributions. Features multi-level sensitivity preset configuration (`OFFEN`, `INTERN`, `VERTRAULICH - unverschlüsselt`, or custom enterprise tenant GUIDs), in-UI preset management (`SensitivityLabelEditDialog` for adding, editing, and deleting presets), sub-millisecond headless OpenXML classification inspection (`docProps/custom.xml`) with live badges on `DocumentSelectorPane` drop targets, active Word COM live verification, and a dual-mode batch applicator (`BatchSensitivityLabelDialog`) for processing either an entire folder (with recursive subfolder search) or an arbitrary set of selected documents. Automatically applies configured labels across all Word opening and export pipelines (`document_opener.py`, `word_comparator.py`, `cr_pcr_generator.py`, `docx_splitter.py`).

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
  - **Zero-Waste Header & Panel Consolidation:** Completely eliminates the standalone top toolbar strip to gain ~45px vertical screen space, migrating `📥 Import ▾` and `⚙️ ▾` dropdown menus directly into the `Indexed Specs` card header alongside an inline 4px import progress bar.
  - **Specification Domain Names & Keyword Search:** Enriches indexed specifications with formal titles and domain acronyms across all surfaces. Tree items render `TS {spec_number}: {title} ({count})` with full formal titles in tooltips. Instant tree filtering (`filter_tree`) matches both spec numbers and title keywords (e.g. typing "UAS", "architecture", "drone", "RRC", or "NAS" matches immediately even without knowing the specification number).
  - **Evolution Matrix Domain Tabs:** Tab headers display clean specification domain labels (`TS 23.256 · UAS`, `TS 23.501 · 5GS Architecture`, `TS 24.501 · 5GS NAS`, `TS 29.244 · PFCP`) with full formal document titles in tab tooltips.
  - **Unified Search Console Card:** High-contrast search input with embedded icon, clear button, fixed date cutoff edit (`125px`, eliminating digit truncation), and executive KPI badge chips (`Results: [ 🔍 query ] [ N Clauses ] [ M Spec(s) ]`).
  - **Sleek Rounded Status Badges (`SpecEvolutionStatusDelegate`):** Custom delegate rendering clean centered rounded pill badges (`🟢 Added`, `⚡ Post-Cutoff`, `✓ Present`, `—`) and soft sky-blue row selection (`#EBF3FC` with `#0F172A` text), eliminating harsh Windows solid blue selection and grid cell rectangular color fills.
  - **Two-Tier Clause Content Inspector:** Tier 1 provides context title (with specification domain indicator), clickable first-added pill, match navigation (`◀ N Match(es) ▶`), 1-click **Layout Toggle** (`⬒ Bottom Split` <-> `◫ Right Pane`), and **Maximize Button** (`⤢` <-> `⤡`). Tier 2 provides a unified, harmonized action toolbar (`📋 Copy ▾`, `✂️ Extract ▾`, `📖 Full Doc ▾`, `🔍 Specs Tab`, `🔍 Diff Prev`, `🤖 LLM Diff`, `📜 Version Changes`). Includes intelligent heading deduplication preventing repeated clause titles.
- **Specification Change History & Integrated TDocs (`spec_history_db.py`, `spec_history_extractor.py`, `spec_tdoc_opener.py`, `spec_history_dialog.py`, `spec_tdoc_report.py`, `spec_tdoc_report_dialog.py`):**
  - **Isolated Storage (`3gpp_spec_history.db`):** Dedicated SQLite database decoupling historical CR/p-CR ledgers from core `3gpp_data.db`. Features master ledger `spec_change_history` and $O(1)$ cross-reference index `spec_tdoc_cross_ref`. Registered in `DatabaseMaintenanceDialog` for real-time size tracking and 1-click database wiping with zero risk to personal notes or specifications catalog.
  - **High-Speed OpenXML Ingestion (<0.35s / spec):** Direct headless `lxml` streaming parser extracting full change history annexes across modern and legacy templates (TS 24.501, TS 38.331, TS 29.244, TS 31.102). Handles multi-part documents (`_sAnnexes.docx`), chained multi-table annexes (TS 24.007), paragraph-aware cell joining, `<w:gridSpan>` expansion, date carry-forward across blank rows, meeting repairs (`#112` -> `CT#112`), and draft version inference.
  - **Working Group Pseudo-CR (p-CR) Tokenizer:** Audited regex `(?<![A-Za-z-])([A-Z][A-Z0-9]{1,4}-\d{5,7}(?:[rR]\d{1,2}(?!-))?)(?![0-9])` tolerant of concatenated rapporteur typos (e.g. `R3-261549R3-261553`), extracting 150+ p-CRs per draft row into first-class indexed entities with revision preservation.
  - **Direct TDoc Document Opening & Meetings Cross-Linking:** 1-click opening of referenced TDoc documents (`.docx`, `.doc`, `.zip`) from `SpecVersionHistoryDialog`, `SpecClauseInspector`, and `SpecClauseDiffDialog`. Resolves against local meeting cache, automatically downloads missing archives from 3GPP servers via `MeetingsDatabase` (`docs_folder_url`) and `URLRouter`, provides 1-click **"🚀 Jump to TDoc in Meetings Tab"**, and formats plain text / rich HTML clipboard citations.
  - **Interactive Notes & Inline Clickable TDoc Badges (`spec_history_dialog.py`):** Merges separate text fields and redundant button scroll areas into a unified, resizable `QTextBrowser`. Automatically parses rapporteur notes and renders all integrated TDocs, p-CRs, and revisions as clickable inline pill badges (`#E0F2FE`). Left-click instantly opens target Word/ZIP documents; right-click provides quick actions (open on 3GPP Portal 3GU, jump to Meetings tab, copy citation, copy identifier).
  - **On-Demand TDoc Metadata Dossier & Report Generator (`spec_tdoc_report.py`, `spec_tdoc_report_dialog.py`):** Resolves full technical metadata (Title, Submitting Organizations/Companies, Agenda Item, Type, Approval Status, Technical Abstract / Proposed Solution, Work Item, and Secretary Remarks) directly from meeting agenda spreadsheets (`.xlsx`) and cached JSON without downloading the actual TDoc documents. Features an interactive modeless viewer dialog (`SpecTDocReportDialog`) with real-time multi-facet filtering (by company, type, status, keyword), technical solution drawer, Markdown dossier export, and CSV spreadsheet generation.
  - **Modeless Multi-Window Usability:** `SpecVersionHistoryDialog` and `SpecTDocReportDialog` configure `Qt.Window` and `Qt.NonModal` window modality, enabling delegates to freely browse, search clauses, inspect meetings, and compare documents across the application while keeping history and dossier windows open side-by-side.
  - **Real-Time Terminal Output Logging:** All change history operations, TDoc document resolutions, background extractions, and report generations emit descriptive, emoji-badged log records directly streaming to the main window's bottom Terminal Output console.
  - **Zero-Lag UI Performance:** Single aggregate pre-fetch in `SpecificationsTab.refresh_table()` eliminating $N+1$ query overhead while badging version dropdown items (`📜 {cnt} CRs/p-CRs`). Accessible from Column 2 `Get Word ▾` menu, Column 0 context menu, `SpecInfoDialog`, `SpecSearchTab` Tier 2 toolbar, Evolution Matrix context menu, and clause diff dialog. All Qt actions escape literal ampersands (`&&`) to prevent Windows accelerator mnemonic underlines.
- **3GPP Clause-to-TDoc Attribution Engine & OpenXML Fingerprinting (`spec_clause_attribution.py`, `spec_clause_diff_dialog.py`, `spec_search_components.py`):**
  - **Problem Statement & Workflow Grounding:** Pinpoints which specific Change Request (CR) or Working Group Pseudo-CR (p-CR) introduced or modified a clause across specification releases, authoring companies, and technical rationale.
  - **Mode A vs Mode B Dichotomy:**
    - *Mode A (Normative TS Plenary CRs):* Individual CR rows in `spec_change_history` mapped to formal CR numbers, categories, and multi-file CR packs (`SP-xxxxxx.zip`, extracting only target `*CR<num>*.docx`).
    - *Mode B (Study TR Working Group p-CRs):* Resolves omnibus p-CR lists in study TRs where contributions originally used placeholder numbering (`6.x`, `Solution #X`) before final numbers were assigned by rapporteurs.
  - **Ancestor Context Inheritance & The Key Issue Bridge:**
    - Generic leaf headings (`6.20.2.1 General`) automatically inherit full architectural titles by traversing `parent_clause` in `3gpp_spec_search.db` (`6.20 Solution #20: Protocol Architecture for Ambient IoT`).
    - The Key Issue Bridge links TR solution body declarations (`"addresses Key Issue #4"`) to candidate p-CRs via meeting agenda items (`AI 20.6.4`) and title/abstract citations (`KI #4`).
  - **Multi-Stage Progressive Attribution Pipeline:**
    - *Stage 1 (Zero-Download Metadata & Structural Ranking, <50ms):* Scans rapporteur notes, matches Key Issue / Solution patterns, computes cleaned semantic title overlap, and scores salient diff keywords against persistent SQLite agenda metadata (`tdoc_metadata_cache` in `3gpp_spec_history.db`).
    - *Stage 2 (Track-Changes & Placeholder-Agnostic Content Fingerprinting):* Inspects candidate Word OpenXML documents (`word/document.xml`), excludes deleted `<w:del>` subtrees when matching added (`+`) diff lines, normalizes numbers (`6.x` / `6.20` $\to$ `<NUM>`), and computes exact long-sentence and 6-word shingle containment for deterministic ground truth (`✅ Verified in DOCX (98% Text Match)`).
    - *Stage 3 (Local Ollama Grounding):* Feeds local Ollama with clause diffs and matched excerpts for multi-pCR attribution and synthesis.
  - **UI Integration & Multi-Surface Usability:**
    - *`SpecClauseDiffDialog` Tab 3 ("🎯 Modifying TDoc"):* Top candidate hero banner with 1-click `📄 Open TDoc`, `🚀 Jump to Meetings`, and `📋 Copy Citation`, ranked candidate table with confidence pills (`Verified`, `Very High`, `High`, `Moderate`), details drawer, and background `🔬 Deep Verify Top 3 (DOCX)` worker.
    - *`SpecClauseInspector` Tier 2 Action Toolbar:* Dedicated `🎯 Origin TDoc` action button opening the clause diff dialog directly focused on attribution.
    - *Modeless Usability:* `SpecClauseDiffDialog` runs modelessly (`super().__init__(None, Qt.Window)`) so delegates can keep diff and attribution analyses open while navigating the main window.
- **Feature-Level Origin CR Ranking & Non-Exclusive Affected Clause Filter (`spec_clause_attribution.py`, `spec_history_extractor.py`, `spec_history_dialog.py`, `spec_search_tabs.py`):**
  - **The Feature-Level Origin Challenge:** Beyond attributing single clause diffs, delegates need to discover which foundational Change Request (CR) originally introduced an entire feature, technology, or architecture concept across all specification releases (e.g. discovering that `CR 2679 (Cat B)` in `SP-210083` introduced `NSAC`/`NSACF` into `TS 23.501` at `v17.0.0` / `SP#91E` citing study phase `KI#1`).
  - **Headless OpenXML Cover Sheet Extractor (`extract_cr_cover_sheet_metadata`):** High-speed stream extraction from `.docx` and `.zip` CR packages parsing `"Clauses affected"` tables, formal Category (`Cat B`, `Cat C`, `Cat F`), CR number, spec number, and work item codes.
  - **Multi-Signal Origin CR Ranking Matrix (`rank_origin_crs_for_query`):**
    - *Signal 1 (Title / Subject Query Match & Origin Intent, Up to 45 pts):* Exact acronym and term matches, boosted by feature introduction keywords (`"definition"`, `"introduction"`, `"adding"`, `"KI#"`) and penalizing maintenance keywords (`"correction"`, `"clean-up"`, `"alignment"`).
    - *Signal 2 (3GPP Category Weighting, Up to 25 pts):* `Category B` (Addition of feature) heavily weighted (+25) over `Category C` (+15) and `Category F` (+5).
    - *Signal 3 (Additive, Non-Exclusive Affected Clauses Alignment, Up to 30 pts):* Rewards candidate CRs matching affected clauses from the search hit list (+25 exact, +15 prefix), while strictly ensuring that CRs with omitted cover sheet metadata are NEVER penalized or excluded.
    - *Signal 4 (Chronological Release Primacy, Up to 20 pts):* Rewards the earliest release baseline (`vX.0.0`, `vX.1.0`) and earliest plenary meeting in the release cycle.
    - *Role Assignment:* Categorizes candidates into `🎯 Foundational Origin CR`, `🏗️ Architecture Expansion`, and `🔧 Maintenance / Correction`.
  - **Spec Search UI Connection (`🎯 Find Origin CR`):** 1-click launch button in the search console bar pre-setting the active search query, clause, and enabling origin ranking across all releases.
  - **`SpecVersionHistoryDialog` Enhancements:** Pre-populates search and clause filters, adds dual-mode clause filtering (`Rank by Clause (Additive)` vs `Strict Clause Filter`), dynamically switches to 9 columns (`Relevance` with role/score badges and tooltips), and renders an executive role & evidence banner in the interactive details drawer.

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
  - **Row 1 (Primary Navigation & Actions):** Quick launch button (`🚀 Last Meeting`), inline `⚡ TDoc:` jump box with dynamic `📄 Doc`, `🗓️ Mtg`, and `⚖️ vs Word` action buttons, right-aligned `🔄 Sync ▾` menu (*Sync All*, *Add / Fetch Meeting...*, *Preferences*), `📊 Reports ▾` menu (*Contributions Report...*, *Export Merged TDocs*), and `⚙️ Settings`.
  - **Row 2 (Filter Engine):** `🔍 Search` (number, city, title), TSG quick-toggle pills (`[All] [SA] [RAN] [CT]`), `Working Group` checkable multi-select combo, Nature (`Regular` / `Ad-Hoc`), Format (`In-Person` / `Electronic`), and Date presets dropdown (`All Dates`, `Upcoming`, `This Year`, `Past 12 Months`, `Custom Range...`) with dynamic calendar pickers.
- **7-Column Table & Domain Visual Hierarchy:**
  - **Working Group Badges (`WGTagDelegate`):** Color-coded domain pills by TSG (Soft Indigo for SA, Soft Emerald for RAN, Soft Amber for CT) using `core.ui.ui_components.get_wg_colors`.
  - **Consolidated English Dates:** Formatted into human-readable ranges (`12 – 16 Oct 2026`) with temporal tooltips (`🟢 Ongoing`, `🔵 Upcoming`, `⚪ Past`).
  - **Preserved Allocated TDoc Range:** Formatted with full WG prefixes (`C4-264009 → C4-264023`, `—` when not yet allocated).
  - **Location Icons:** Clean icon prefixes (`📍 City, Country` vs `🌐 Online (eMeeting)`).
  - **Double-Click Navigation:** Double-clicking any meeting row directly launches the meeting viewer or opens the metadata dialog.
- **Floating Comparison Cart & "Active Word vs. TDoc" Direct Comparison (`ComparisonCartWidget`, `ComparisonManager`):** Shared, synchronized comparison cart embedded in the bottom status bar of both `MeetingsTab` and `TDocsWindow`. Auto-hides into a compact `⚖️ Stage Active Word ▾` launcher when empty and expands into interactive `[Slot A ▾]` and `[Slot B ▾]` pill buttons, a `⇄` 1-click slot swap button, and a `🪟 Active Word` quick-stage button. Delegates can stage any TDoc, the currently active Microsoft Word document, any open Word document, or a local file into Slot A (Original/Baseline) or Slot B (Revised), or trigger an instant 1-click **Active Word vs. TDoc** legal redline comparison from the `MeetingsTab` Quick TDoc Jump bar (`⚖️ vs Word`), `TDocsWindow` Instant Fetch bar (`⚖️ vs Word`), table row context menu, Column 0 `Open` menu, or Related TDoc hyperlink menu. Automatically re-snapshots modified open Word documents right before launching `WordComparatorThread`.
- **Destructive Action Relocation:** Moved `Clear All Meetings...` from the main UI into `MeetingsConfigDialog` under a dedicated *⚠️ Database Maintenance (Danger Zone)* section.
- **Three-Phase Sync:** Asynchronously synchronizes meeting lists, TDoc metadata, and dynamic agenda reservations from 3GPP FTP servers.
- **AI TDoc Triage:** Modeless dialog extracting CR cover sheets and redlines to synthesize structured executive summaries (*Objective*, *Key Changes*, *Potential Impacts*).
- **Personal Notes & Stance:** Integrates local SQLite user notes (`user_tdocs.db`) with stance tagging (`⚪ Neutral`, `🔵 My TDoc`, `🟢 Support`, `🔴 Object`, `🟡 Monitor`).
- **Agenda & Sub-Item Ordering Engine (`tdocs_parser.py`, `tdocs_models.py`, `tdoc_delegates.py`, `tdocs_window.py`):** Tailored for session tracking in Working Groups such as SA2 where session progress follows agenda order. Supports live `TdocsByAgenda.htm` and Word Chairman's Notes precedence over static Excel lists (`"TDoc sort order within agenda item"`), automatic extraction of topic separator rows as `Sub-Item`s, precomputed $O(1)$ sort caching, a compact Option C visual badge in `AgendaItemDelegate` (Sub-Item rendered as a rounded `#F1F5F9` pill within the Agenda Item cell), dynamic `Sub-Item` checkable filtering with facet-awareness, full-text search integration, and an `Order By ▾` dropdown combobox (`📋 Agenda Order`, `🔢 TDoc Number`, `🏢 Company`, `✉️ Most Emails`, `🗂️ Default / File Order`) synchronized with column header clicks.
- **Company Review Comments Matrix & Delegation Exporter (`excel_export_dialog.py`, `excel_exporter.py`, `tdocs_export_controller.py`):** Dedicated Review Matrix generation enabling delegates to export customized company review spreadsheets directly from in-app notes and stances. Features a 1-click **📋 Review Matrix** preset in `ExcelExportDialog` (automatically staging standard agenda columns alongside `My Status` and `My Notes`), dynamic corporate attribution (`Status (<Company>)` and `Comments (<Company>)` configured via user preferences or auto-discovered from `WordConfig`), visual status cell highlighting in Excel with soft color tints (`🟢 Support` = emerald `#E6F4EA`, `🔴 Object` = soft red `#FCE8E6`, `🟡 Monitor` = soft amber `#FEF7E0`, `🔵 My TDoc` = soft blue `#E8F0FE`), auto-expanded multiline comment widths (42 chars), and smart default filename resolution (`Comments_<Company>_<WG>_<Meeting>_<Scope>.xlsx`).
- **Tracked Work Items Quick Filters:** Integrates with the Work Items module via `TDocsFilterProxyModel` and `TDocsWindow` quick-toggle chips (`🎯 Followed WIs`, `⭐ Starred WIs`). Filters meeting contributions by followed/starred WIs with dynamic badge counts and full multi-facet `AND` combination (e.g. Followed WIs + Approved/Agreed TDocs). Includes default support (`include_unassigned_ls_in = True`, toggleable via right-click) to preserve incoming Liaison Statements (LS in) without an assigned Work Item. Resolves acronym variations, WG suffixes (`-ARC`, `-CT`, etc.), and `FS_` prefixes, with fallback to Agenda Item description for discussion papers.
- **Opportunistic Work Item Auto-Enrichment (`meeting_wi_enricher.py`):** Fully decoupled, non-blocking background thread (`MeetingWIEnricherThread`) automatically launched when opening or refreshing any meeting. Performs an off-thread openpyxl bounded scan of the "Related WIs" column for 3GPP Portal hyperlinks (`workitemId=\d+`), checks SQLite (`WorkItemsDatabase.get_existing_wi_codes()`), and asynchronously scrapes missing/unindexed Work Items directly from the 3GPP Portal without freezing or delaying the TDocs table. Employs a persistent JSON audit marker (`<excel>.wi_checked.json`) to guarantee zero overhead on subsequent opens (~0 ms), and dynamically updates tracked tokens and filter badges upon completion. Configurable via `MeetingsConfigDialog`.
- **Post-Meeting Report Generator:** Modeless `PostMeetingExportDialog` and `PostMeetingExporterThread` compiling meeting contributions into a structured, per-release and per-Work Item Markdown dossier (`Rel-XX/<WI>/{LS_IN, pCR, CR, LS_OUT, Discussion}/`) with collision-free media copying, cross-WI notices, multi-AI listing, executive `index.md` tables, and a default option (`Include unassigned & general LS in`) ensuring incoming LSs without an assigned Work Item are safely preserved under `General/LS_IN/`. Features:
  - **Ergonomic 2-Column Split Layout:** Eliminates vertical screen overflow and tree clipping on laptop displays by organizing the dialog into two balanced columns (Left: Scope selection grid & full-height dual-mode WI/AI tree with search; Right: Document categories & output options), fitting comfortably within 650px height with zero clipping.
  - **Custom Work Item & Agenda Item Scoping:** Replaces flat list selection with a synchronized dual-view tree switcher (`🗂️ Work Items` vs `📋 Agenda Items`), automatic pre-selection of followed/starred WIs and child AIs, tri-state check cascading, and instant multi-token filtering across WI names, natural numeric AI numbers (e.g. `5.1` before `5.10`), and topics.
  - **Fine-Grained Scoping:** Allows delegates to export specific Agenda Items within a Work Item (e.g. exporting only Stage 2 architecture under `TEI18` while excluding Stage 3 protocol AIs).
  - **Live "Estimated Scope" Strip & Dynamic Badges:** Real-time KPI banner and live document count badges on all category checkboxes updated on every click.
  - **Discussion Papers & WID Inclusion (`include_disc`):** Optional inclusion of noted/agreed discussion papers and Work Item status reports under `Discussion/`.
  - **Liaison Statement Reply Tracking:** Automated bidirectional LS reply tracking (`Reply in`/`Reply to` with `Revised to` chain resolution) that hyperlinks outgoing reply dossiers directly from the `LS in` status table (`replied to in [S2-XXXX](...)`) and incoming references in the `LS out` table (`approved (reply to [S2-YYYY](...))`), injects reciprocal Markdown callouts (`[!TIP] Outgoing Reply` / `[!NOTE] Reply Liaison Statement`) into dossier headers, and guarantees that any approved reply `LS out` document answering an included `LS in` is exported to disk (routed to `General/LS_OUT/` if its Work Item was not in the export scope) without cluttering the report with unwanted WI sections.
- **Meeting Analytics & Statistics Engine (`exporter_thread.py`, `contribution_stats.py`, `html_dashboard_compiler.py`):** Multi-scope interactive analytics compiling single-meeting dashboards (`Statistics_Report.html`), cross-meeting contribution audits, and multi-tab Excel workbooks (`Statistics_Data.xlsx`). Features unified semantic 3GPP outcome palettes (`Agreed`/`Approved` in emerald/deep green, `Not Treated` in slate grey, `Revised` in indigo), a 6-metric responsive KPI strip (Total TDocs, CR Agreement Rate %, Meeting Treatment Rate %, Document Composition, Active Entities with Joint %, Active Work Items/Topics), document category segmentation (CR vs Discussion vs LS), Louvain alliance network graphs, and dual Chart/Data toggle tables on every card.
- **TDoc Markdown Conversion & Provenance Preservation (`tdoc_markdown_converter.py`, `tdoc_url_resolver.py`):** High-fidelity conversion of 3GPP Word documents (`.docx`, `.doc`) into Markdown with track changes, tables, and diagrams. Features automatic document type revision mode resolution (`resolve_revision_mode_for_type`): CR and pCR documents retain specification change markup (`revisions="markup"` with `[INSERTED]`/`[DELETED]` tags and markup legend), whereas Liaison Statements (`LS in`, `LS out`, `LS`) are recognized as plain correspondence documents and exported in clean `accept` mode (`revisions="accept"`, omitting change marks and markup legends). Multi-tier stale checking automatically upgrades previously cached markup dossiers of LS documents. Automatically resolves and embeds official 3GPP server HTTPS download URLs (`resolve_tdoc_download_url`) into standalone Markdown dossier headers (`- **3GPP Server URL:**`), Post-Meeting report index tables (`[Dossier](...) · [📥 3GPP](...)`), AI TDoc Triage dialog badges/clipboard summaries, and Markdown report tables, preserving permanent document provenance across mixed downstream archives and external readers.
- **High-Speed Parallel Pre-Download Pipeline for Markdown Exports (`llm_exporter.py`, `post_meeting_exporter.py`, `tdoc_file_handler.py`):** Accelerates bulk Markdown exports ("Export TDocs to Markdown (.md)" and Post-Meeting Dossier) by introducing a decoupled two-phase pipeline:
  - **Phase 1 (Concurrent Network Fetch):** Fast pre-scan identifies unindexed or missing source documents in `tdocs_list` / `unique_tdocs`, bypassing already-cached fresh dossiers. Dispatches missing archives to a polite worker pool (`ThreadPoolExecutor(max_workers=5..6)`). Downloads stream concurrently, cutting network wait time by 5x–6x (~60s down to ~10s for 40 contributions) while staying within 3GPP server rate-limits.
  - **Phase 2 (In-Memory Conversion):** Sequentially converts the locally-cached documents to Markdown using direct OpenXML parsing with zero network stalls, while guaranteeing single-threaded COM isolation for Word fallbacks and eliminating multi-process churn from vector metafile conversions.
  - **Atomic File Integrity & Self-Healing Cache (`TDocFileHandler`):** Downloads stream into temporary `.zip.part` / `.part` files prior to atomic OS renaming, eliminating corrupt/truncated archives on cancellation or network loss. Automatically validates archive integrity via `zipfile.is_zipfile` and purges 0-byte corrupt files, while supporting bidirectional route failover between `revisions_url` and `docs_ftp_url`.
  > [!CAUTION]
  > **DO NOT MODIFY `src/modules/meetings/core/tdoc2md.py`**. This file is a drop-in core imported directly from an external standalone repository (`3gpp-tdoc-markdown-exporter`). All application-specific integrations, office adapters, or custom pipelines MUST reside exclusively in `tdoc_markdown_converter.py` or calling modules.

### 4.6. Research Assistant (`src/modules/assistant`)
- **Subsystem Guide:** For complete architecture, semantic tool registry, and execution rules, consult [`src/modules/assistant/AGENTS.md`](file:///src/modules/assistant/AGENTS.md) (and [`README.md`](file:///src/modules/assistant/README.md) for full design rationale).
- **Architecture:** Bounded, evidence-grounded agent architecture connecting a local LLM to application knowledge services.
- **Authority Boundary:** Strictly **`KNOWLEDGE_READ`**. The model has no SQL access, shell execution, or database write permissions.
- **Semantic Tool Set:** Model interaction is restricted to predefined semantic tools (`find_specifications`, `get_specification`, `search_spec_text`, `get_spec_clause`, `list_protocols`, `list_protocol_messages`, `find_protocol_message`, `find_protocol_ie`).
- **Cockpit UI & Ergonomics (`AssistantTab`, `assistant_components.py`):** Features responsive dual-mode layout (`◫ Side Drawer` <-> `⬒ Bottom Split`), active Ollama model status pill with 1-click configuration launcher, session management (`➕ New Chat`, `📋 Copy Chat`, `💾 Export Markdown`), auto-wrapping multi-line prompt editor (`MultilinePromptEdit` with `Shift+Enter` for newlines), starter prompt suggestion chips (`StarterChipsWidget`), and clean card frame styling.
- **Rich Chat Cards & Clickable Citations:** Card-framed user and assistant message bubbles with timestamps and KPI metadata (`⏱️ Latency · Steps · Sources`), formatted technical Markdown rendering (bold, lists, code blocks, tables), interactive citations (`<a href="evidence:E#">E#</a>`) that auto-focus cited evidence in the Sources tree and preview pane, and per-message `📋 Copy Answer` links.
- **Embedded Evidence Inspector & Cross-Tab Deep Linking:** Non-blocking embedded excerpt viewer (`EvidenceInspectorWidget`) with source badges (`[Spec Text]`, `[Protocol]`, `[Catalogue]`) and 1-click jumps to other application tabs (`🔎 In Spec Search`, `🔬 In Protocols`, `📚 In Specs`, `📋 Copy Excerpt`).
- **Research Cockpit Tabs & Diagnostics:** Tabbed drawer organizing Sources (`📚 Sources (N)`), a real-time activity stream (`⚡ Activity (N)` with semantic tool icons and pulsing progress indicator), and an Ollama trace cockpit (`🔬 Diagnostics` with latency, token counts, generation speed in tok/s, and raw REST exchange trace).

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
- **WAL Compaction & Storage Hub (`DatabaseMaintenanceDialog`):** Provides an executive storage dashboard featuring top KPI summary cards (Total Storage, Reclaimable Freelist Space, WAL Log Health, User Table Counts), two-tier table rows with domain-specific icons and filenames, and animated progress feedback for `PRAGMA wal_checkpoint(TRUNCATE)` and `VACUUM` batch compaction. Destructive wiping actions (`🗑️ Wipe`) are visually isolated with warning styles to prevent misclicks. Features 1-click **📂 Open DB Folder** in Windows File Explorer.

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
- **Mnemonic-Aware Controls vs. Non-Mnemonic Widgets:**
  - **Mnemonic-Aware Controls (`QPushButton`, `QToolButton`, `QCommandLinkButton`, `QCheckBox`, `QRadioButton`, `QGroupBox`, `QAction`, `QMenu`, `QMenuBar`, `QTabBar`, and `QTabWidget`):**
    - In Qt/PyQt, these interactive controls interpret a single `&` as a keyboard accelerator mnemonic prefix.
    - Passing a single `&` (e.g. `"Save & Close"`, `"Work Items & Enrichment"`) causes Qt to swallow the ampersand and underline the following character on Windows (e.g. rendering `Work Items _Enrichment`).
    - **Literal Ampersands Must Be Escaped (`&&`):** For these controls, always write literal ampersands as `&&` in static code (e.g. `QPushButton("Save && Close")`, `QGroupBox("🎯 Work Items && Enrichment")`, `QAction("Tags && Notes")`, `QCheckBox("Followed && Starred")`).
    - **Dynamic Data Ingestion:** Strings loaded dynamically from external sources (`useful_links.json`, databases, FTP scrapes, network APIs) that contain `&` must be sanitized with `escape_qt_mnemonic(text)` from `core.ui.ui_components` before passing them to `menu.addAction()`, `btn.setText()`, `tab.addTab()`, etc. `escape_qt_mnemonic()` is idempotent.
  - **Non-Mnemonic Widgets & Text Surfaces (`QLabel`, `QLineEdit`, `QTextEdit`, `QTextBrowser`, `QToolTip`, `QStatusBar`, `setWindowTitle`, table/tree cells, and HTML markup):**
    - These controls and surfaces do **NOT** parse keyboard mnemonics (unless a buddy is explicitly assigned to a `QLabel`).
    - **CRITICAL RULE — Never Use `&&` Here:** If you write `&&` in a `QLabel` (e.g. `QLabel("Field && Structure Inspector")`), `label.setText()`, `setWindowTitle()`, `setToolTip()`, or HTML strings (`<b>...</b>`), Qt or the HTML engine will literally render `"&&"` on screen to the user!
    - In `QLabel`, window titles, tooltips, and placeholders, always write a single literal `&` (or `&amp;` in HTML).
- **Automated Regression Testing:** Both sides of this rule are enforced automatically by `tests/test_ui_ampersand_escaping.py`:
  - `test_no_unescaped_ui_ampersands_in_src`: Flags missing `&&` in mnemonic-aware controls.
  - `test_no_double_ampersands_in_non_mnemonic_widgets`: Flags improper `&&` in `QLabel`, window titles, tooltips, and non-mnemonic `setText()` calls.
  - Always verify compliance by running: `py -3.14 -m unittest tests/test_ui_ampersand_escaping.py`.

### 5.9. Documentation Synchronization & Commit Hygiene
- **Documentation as Code:** Never treat documentation as an afterthought. When adding new capabilities, refactoring architectural patterns, modifying configurations, or updating UI controls, update corresponding documentation (`README.md`, `AGENTS.md`, and module guides) in the same work unit.
- **Commit Completed Work Units:** Do not accumulate large batches of uncommitted, disparate changes across multiple subsystems. Once a feature or bugfix is verified and documented, commit it cleanly with a concise, descriptive commit message.
- **Never Execute `git push` (Manual User Review Boundary):** AI agents must **never** run `git push`. Pushing commits to the remote repository is reserved exclusively for the human developer. This guarantees the user has full opportunity to inspect commit logs, review diffs, test changes, and amend commits before publishing upstream.

### 5.10. Dialog Modality, Floating Windows & High-DPI UI Rules
- **NEVER Use `Qt.WindowStaysOnTopHint` on Tool/Utility Dialogs:**
  - Forcing dialogs or floating windows to always stay on top pins them above *all* desktop applications (Microsoft Word, Teams, Outlook, web browsers). This severely impairs multitasking and delegate workflows.
  - Windows must be created with standard `self.setWindowFlags(Qt.Window)` without the `Qt.WindowStaysOnTopHint` flag.
- **Modeless Multi-Window Architecture (`Qt.NonModal`):**
  - Utility and inspector tools (e.g. `DatabaseMaintenanceDialog`, `DatabaseInspectorDialog`, `ProcessManagerDialog`, `SpecVersionHistoryDialog`, `SpecTDocReportDialog`) must be non-blocking floating windows (`self.setWindowModality(Qt.NonModal)`).
  - Launch with `.show()` rather than `.exec_()`. Blocking modal dialogs prevent delegates from interacting with other tabs, searching specs, or cross-referencing documents.
  - Retain a persistent reference on the spawning parent (e.g. `self.db_maintenance_dialog`) to prevent Python garbage collection, and invoke `.raise_()` + `.activateWindow()` if triggered when already open.
  - Modal execution (`.exec_()`) is strictly restricted to brief confirmation prompts (e.g. `QMessageBox.question` before destructive actions) or modal input pickers.
- **High-DPI Scaling & Dynamic Button Sizing:**
  - Never set hardcoded tight pixel widths (e.g. `setFixedWidth(72)`) on buttons containing text and icons. Under Windows display scaling (125%, 150%, 200%) or custom system fonts, text will be clipped (e.g. `Inspect` -> `Insne`, `Compact` -> `Compa`, `Wipe` -> `Win`).
  - Rely on CSS padding (e.g. `padding: 2px 10px;`), `min-width` rules, and allocate generous container column bounds (`>= 280px` for multi-button table action cells).
- **Vertical Clearance & Font Descenders:**
  - Buttons and table row section heights must provide sufficient vertical clearance (e.g. button `min-height: 26-28px`, table row `defaultSectionSize: 46-48px`) so that character descenders (`p`, `g`, `y`, `q`, `j`) are never sliced off at the bottom border.


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
