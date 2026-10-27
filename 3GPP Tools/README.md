# 📊 3GPP Meeting Tools & Diagram Converter

An advanced, component-based desktop IDE designed to bridge the gap between text-based diagramming (`PlantUML`) and corporate enterprise environments (`Microsoft Visio` and `PowerPoint`). 

Built specifically with telecommunications and 3GPP standards workflows in mind, this tool allows you to write highly efficient PlantUML sequence, activity, and network diagrams, instantly export them as fully editable native Office shapes, rapidly slice massive specification documents into manageable chapters, track NAS, ASN.1 (RRC / NGAP / XnAP / F1AP / E1AP / S1AP), GTP-U, PDU Session UP (TS 38.415), and PFCP (TS 29.244) protocol message evolutions, search arbitrary substrings across specification releases using FTS5 trigram indexing with "First Added" and cutoff date detection, conduct multi-meeting company contribution audits with interactive KPI analytics and co-signer tracking, inspect and compact local SQLite databases, safely manage background COM Office processes, track emails across working groups linked to specific TDocs and their revision families, consult an evidence-grounded local 3GPP Research Assistant with semantic tool calling, connect seamlessly to local LLMs via Ollama for zero-cost offline intelligence, automated call-flow diagram synthesis, and real-time technical triage of meeting TDocs, and navigate, filter, and synchronize the vast 3GPP meeting, specification, and work item archives locally.

---

## 📑 Table of Contents
1. [✨ Features](#features)
2. [🏛️ Architectural Evolution & Deprecated Legacy Features](#deprecated-features)
3. [🏗️ Architecture & Data Flow](#architecture)
4. [⚙️ Prerequisites](#prerequisites)
5. [🚀 Installation](#installation)
6. [📖 How to Use the GUI](#usage)
7. [🛠️ Known Quirks / Troubleshooting](#troubleshooting)

---

## <a id="features"></a>✨ Features

### 🦙 Local AI & LLM Integration (Ollama Core Service)
* **100% Offline & Private Intelligence:**
  * **Confidentiality & Zero Cost:** Connects directly to a locally hosted Ollama daemon (`http://127.0.0.1:11434`). Sensitive company draft positions, unreleased Change Requests (CRs), and meeting notes never leave your machine or cross external APIs.
  * **Airplane & Meeting-Room Ready:** Operates completely offline during flights and face-to-face 3GPP sessions without requiring external internet connectivity.
* **🤖 AI Technical Triage for Meeting TDocs:**
  * **Rapid Technical Extraction:** Right-click any TDoc or click **🤖 AI Triage** in the details dialog to stream an executive summary of Change Requests (CRs), Discussion Papers, or Liaison Statements.
  * **Optimized for 7B Laptop Models:** Employs a focused context truncation strategy that isolates the 3GPP CR cover sheet table (*Reason for change*, *Summary of change*, *Consequences if not approved*, *Clauses affected*) and modified clauses (`[INSERTED: ...]`, `[DELETED: ...]`), bounding context to ~16,000 characters for sub-10-second summaries on models like `qwen2.5:7b`.
  * **Structured 3-Part Schema:** Automatically synthesizes proposals into:
    * **🎯 Objective & Problem Statement** (architectural gaps, ambiguities, race conditions).
    * **🔑 Key Technical Changes** (normative clause edits, new IEs, state transitions, timers).
    * **⚠️ Potential Impacts & Contentious Points** (side effects, backwards compatibility, vendor objections).
  * **1-Click Integration with Personal Notes & Status:** Pick your stance directly from the embedded **My Status** dropdown (`⚪ Neutral`, `🔵 My TDoc`, `🟢 Support`, `🔴 Object`, `🟡 Monitor`) and click **💾 Save to My Notes** to append the AI triage output to your private SQLite notes with model, timestamp, and official 3GPP server download link headers.
  * **3GPP Server Link & Provenance Copying:** Features an interactive **📥 3GPP Server ZIP** download link directly on the triage metadata card. Clicking **📋 Copy** automatically prefixes the executive summary with the TDoc title and 3GPP download URL, guaranteeing permanent document provenance when pasting summaries into emails, Teams, or reports.
  * **Modeless, Multi-Window Experience:** Operates as an independent top-level window (`Qt.Window`), allowing delegates to continue browsing meetings or reading Word redlines while tokens stream.
  * **Hot-Reloadable Prompt Templates:** Prompts are decoupled into `config/prompts/triage_system.txt` and `config/prompts/triage_user.txt` for instant tuning without restarting the app.
* **🤖 AI Email Discussion Summarizer for TDoc Families:**
  * **Zero-Context-Bloat Quote Stripper:** Advanced quotation truncation engine (`clean_email_body_for_summary`) detects and removes historical reply chains, inline quotation headers, and disclaimers. Strips 80–90% of redundant context while preserving the delegate's direct response.
  * **Full Audit Logging:** Detailed per-email logs in `3gpp_tools.log` reporting raw versus sanitized character lengths, percentage reduction, and warnings on unusually large non-delimited bodies to verify context cleanliness.
  * **Chronological Argument Reconstruction:** Re-orders messages chronologically (oldest to newest) to let the LLM map how arguments evolved across multiple companies.
  * **Structured Stance Breakdown:** Categorizes discussion into **🎯 Discussion Topic & Core Contention**, **🏢 Company Positions & Alliances**, **📝 Proposed Solutions & Revisions Mentioned**, and **📌 Current Consensus & Next Steps**.
  * **1-Click Append to Personal Notes:** Directly appends discussion summaries to your SQLite personal notes (`user_tdocs.db`) with timestamp and model metadata.
* **🤖 Evidence-Grounded 3GPP Research Assistant (Assistant Tab):**
  * **Bounded Agent Architecture:** Integrates a local LLM with application-side knowledge services via a bounded controller (`AgentController`). The model autonomously decides which semantic research tools to call, while Python retains strict authority over validation, execution limits, and application state.
  * **Strict `KNOWLEDGE_READ` Authority Boundary:** Strictly bounded to read-only knowledge inspection. The agent has zero arbitrary SQL access, zero filesystem write access, zero process/shell execution permissions, and zero ability to alter application configurations or download unindexed files.
  * **8-Tool Validated Semantic Registry:** Equips the assistant with validated, schema-enforced semantic tools:
    * *Specification Catalogue:* `find_specifications` (search catalogue by number/keyword), `get_specification` (retrieve known versions and metadata).
    * *Specification Text Index:* `search_spec_text` (FTS5 trigram substring matching across indexed clauses), `get_spec_clause` (retrieve exact extracted clause text).
    * *Structured Protocol Knowledge:* `list_protocols` (enumerate parsed protocols), `list_protocol_messages` (list messages/SIBs in a release), `find_protocol_message` (retrieve field/IE structures), `find_protocol_ie` (retrieve Clause 9 bit layouts and coding tables).
  * **Progressive Multi-Pane Research UI (`AssistantTab`):**
    * **Conversation Pane:** Streams grounded markdown answers with bracketed citations (e.g., `[E1, E2]`).
    * **Research Activity Pane:** Real-time log of tool calls displaying invoked tool names, input arguments, durations, and output statuses.
    * **Application-Owned Evidence Cards:** Displays concrete evidence items with clear provenance (source type, specification, version, clause/message, and preview), distinguishing local application coverage from model hallucinations.
  * **Hot-Reloadable System Prompt:** System instructions are externalized to `config/prompts/assistant_system.txt`, enabling real-time prompt tuning with automatic timestamp detection without application restarts.
* **3GPP Call Flow Sequence Generation:**
  * Directly transforms textual procedure steps (e.g., from TS 23.502 or TS 38.300) into standard PlantUML sequence diagrams.
  * Integrates real-time token streaming, multi-turn conversational refinement, in-dialog SVG diagram preview, automated syntax validation, and prompt hot-reloading within Visio Tools.
* **Persistent Status Bar Pill & Health Monitor:**
  * **Interactive Status Widget:** Permanent status button integrated directly into the bottom status bar displaying active connectivity and model selection (🟢 `🦙 <model_name>` when connected, 🔴 `🦙 Offline` when disconnected).
  * **Instant Socket Pre-Flight:** Worker thread performs sub-millisecond TCP socket pre-checks to `127.0.0.1:11434` prior to issuing HTTP requests, preventing UI stalls and timeout freezes when the Ollama server is stopped.
  * **Log Transition Memoization:** Heartbeat monitoring runs silently in the background, only emitting Qt signals and writing informational log records when connection states or installed model inventories actually change.
* **Dedicated AI Network Session & Corporate Proxy Bypassing:**
  * **Bypass Corporate Proxy Traps:** Routes through a dedicated AI session factory (`get_ai_session()`) equipped with RFC 1918 loopback detection (`is_local_address()`), guaranteeing local prompts never get routed into or blocked by enterprise proxies (e.g., Zscaler, BlueCoat).
  * **Zero Scraper Delays:** AI sessions are completely decoupled from 3GPP FTP scraper sessions, bypassing humanness sleep delays to stream inference tokens with maximum throughput.
* **Centralized Configuration & Dynamic Model Switching:**
  * One-click configuration dialog (`OllamaConfigDialog`) to update host URLs, switch between proxy modes (*Direct/Bypass*, *Auto-Detect*, *App Proxy*), test connectivity, and hot-swap default models (e.g., `qwen2.5:7b`, `llama3.1:8b`, `deepseek-r1:14b`) with settings persisted across sessions in `ollama_config.json`.

---

### 🎨 Visio Tools (PlantUML & PowerPoint Converter)
* **PlantUML Mini-IDE & Ergonomics:**
  * **Syntax Highlighting & Line Numbers:** Real-time high-contrast syntax highlighting (`PlantUMLHighlighter`) distinguishing directives (`@startuml`, `skinparam`), control blocks (`alt`, `opt`, `loop`), declarations (`participant`, `actor`, `database`), arrows (`->`, `-->`, `<->`, `->x`), hex colors, stereotypes, strings, and single/multi-line comments.
  * **Dynamic Element & Snippet Inserter (`🧩 Insert Element ▼`):** Over 55 categorized snippets with dynamic participant detection (`{p1}`, `{p2}`) and selection-aware block wrapping. Automatically detects declared entities (e.g. `UE`, `gNB`, `AMF`) and injects them into arrows and notes instead of generic placeholders.
  * **Searchable Syntax & Elements Palette (`📚 Syntax Palette`):** Collapsible sidebar featuring instant substring filtering, diagram-type scope filtering (Sequence, State, Activity, Timing, Packetdiag), monospace syntax preview cards, and one-click insertion.
  * **3GPP Architecture Presets:** Pre-styled, enterprise-ready entity skeletons for 5GS Core & RAN, 5GS Basic, EPS / 4G, 5G Roaming (VPLMN/HPLMN with SEPP), O-RAN split nodes, and IMS VoNR.
  * **Editor Ergonomics:** Inline non-intrusive Find & Replace bar (`Ctrl+F`), multi-line comment toggling (`Ctrl+/` or `Ctrl+'`), soft tabs (2 spaces), block indent/unindent with `Tab` / `Shift+Tab`, and smart auto-indentation on Enter.
  * **Template Overwrite Guard:** Confirms before replacing non-empty editor buffers when clicking `Insert` next to boilerplate templates.
* **Dual Live Preview & Error Pinpointing:**
  * **Embedded In-App Split Preview (`👁️ Split Preview`):** Side-by-side SVG viewer dock using `PyQt5.QtSvg` with smooth mouse-wheel zoom, drag-to-pan, 1:1 reset, and fit-to-window.
  * **External Browser Preview (`🌐 Browser Preview`):** Standalone browser live view auto-updating on debounced changes.
  * **Live Syntax Error Pinpointing:** When PlantUML reports a syntax error (`Error line N`), line $N$ is immediately highlighted in soft red with a wavy underline in the code gutter and editor.
* **Decoupled Architecture & Visio-Independent PlantUML:**
  * Operates fully independently of Microsoft Visio. PlantUML editing, live SVG/PNG previews, and automated JAR updates function properly even when Microsoft Visio is not installed on the system.
* **☕ Centralized Java Runtime & PlantUML Maintenance Center:**
  * Single, dedicated maintenance dialog (`JavaMaintenanceDialog`) accessible via the `☕ Java` button in the Console Panel.
  * Live runtime inspection displaying version, vendor, architecture (64-Bit), executable path, and active profile (`Modern` vs `Legacy`).
  * **Zero-Admin Portable Java:** Choose between automatic discovery and a custom portable JRE folder. Easily use Java on locked-down corporate PCs without administrator privileges or `.msi` installers.
  * **Integrated PlantUML Updater:** Inspect local JAR versions and trigger background online updates directly inside the dialog.
* **🤖 AI Call Flow to PlantUML Sequence Generator:**
  * **Modeless, Multi-Window Experience:** Runs as an independent, non-blocking window (`Qt.Window`) with its own taskbar presence. Delegates can browse specifications, search meetings, and edit diagrams simultaneously while generation takes place.
  * **Real-Time Token Streaming & Cancellation:** Leverages chunked HTTP streaming (`stream_chat`) from local Ollama instances to stream code into the preview editor as it is synthesized, with an instant cancel button to abort runaway outputs.
  * **3GPP Architectural Style Toggles:** Configure generation preferences via dedicated checkboxes: box groupings by domain/network slice, message Information Elements (IEs) and parameter payloads, participant state notes (`hnote`), alternative and failure branches (`alt` / `else`), and sequence step autonumbering (`autonumber`).
  * **Built-In 3GPP Procedure Presets:** One-click menu (`📚 Sample Procedures ▼`) to load normative procedure texts from TS 23.502 and TS 38.300 (Initial Registration, PDU Session Establishment, Xn Handover, Service Request) or paste directly from the system clipboard.
  * **Dual Code & Diagram Preview Tabs:** Toggle between the full `PlantUMLHighlighter` code editor and an embedded interactive SVG diagram preview (`👁️ Diagram Preview`) rendered automatically upon generation completion.
  * **Automated Syntax Validation & 1-Click Auto-Fix:** Runs background PlantUML validation on completed code; flags syntax errors with exact line numbers on a warning banner, and provides a `🪄 Auto-Fix Error` button that instructs the LLM to repair the exact syntax fault.
  * **Conversational Multi-Turn Refinement:** Interactive prompt bar (`💬 Refine:`) allowing iterative adjustments ("add UPF data path", "wrap in box Core Network", "add timeout branch") while maintaining full generation history across turns.
  * **3GPP Template & Skin Harmony:** Automatically strips markdown wrapper fences and enforces project-standard 3GPP lifeline styling, monochrome skin parameters, and box padding from `plantuml_templates.py`.
  * **Dynamic Hot-Reload Prompts:** Prompt instructions are decoupled into external configuration files (`config/prompts/callflow_system.txt` and `config/prompts/callflow_user.txt`). Modifications made in external text editors are detected via file timestamps (`mtime`) and reloaded instantly on the next inference click without restarting the application.
  * **One-Click Editor & Clipboard Integration:** Provides dedicated buttons to **📥 Replace Editor**, **➕ Append to Editor**, or **📋 Copy Code** with visual confirmation.
* **Batch Conversion Engine:** Drag and drop hundreds of `.puml`, `.txt`, or `.pptx` files to queue them for multi-threaded background conversion.
* **PowerPoint to Visio Pipeline:** Seamlessly convert entire PowerPoint presentations into multi-page Visio documents (`.vsdx`). Uses Enhanced Metafile (EMF) bridging to perfectly preserve editable native Office shapes, automatically aggressively ungroup them, and shrink wrap their text boundaries.
* **Custom Visio Stencil Engine:** Converts standard PlantUML shapes into grouped Visio shapes (`.vsdx`) mapped directly to custom 3GPP node stencils.

---

### 🔎 3GPP Specification Full-Text & Substring Search Engine
* **FTS5 Trigram Substring Search & Chronological Tracking:**
  * **Arbitrary Substring Matching:** Powered by an embedded SQLite Full-Text Search (FTS5) engine configured with a 3-character Trigram tokenizer (`tokenize="trigram"`). Enables near-instantaneous search for exact phrases, field substrings, acronyms, or protocol constants across millions of words without full-table scan delays.
  * **Targeted Release & Clause Filtering:** Filter queries by specific clause patterns (e.g., `5.2`, `8.1.4`, `Annex A`) or execute cross-specification queries across all active releases simultaneously.
* **Release Evolution Matrix & "First Added" Text Tracking:**
  * **Per-Specification Tabbed Matrix Visualization:** Automatically isolates search results into dedicated per-specification tabs (e.g., `TS 23.501 (32)`, `TS 23.502 (20)`). This prevents sparse empty matrices, eliminates colliding clause numbers, and preserves clean chronological column ordering per document.
  * **"First Added" Identification:** Automatically determines the exact earliest release where matching text was introduced, rendering clear visual indicators:
    * 🟢 **`🟢 Added`**: Highlighted in soft green to indicate the exact version where text first appeared in that clause.
    * ⚪ **`✓ Present`**: Retained and present in subsequent releases.
    * 🔴 **`✗ Removed`**: Highlighted in soft red when text present in a previous release was deleted in that version.
    * ➖ **`-`**: Clause not matching or not present in that release.
* **Date Cutoff Analysis:**
  * **Official Release Date Storage:** Tracks official 3GPP portal upload and publication dates across all indexed specification releases.
  * **Post-Cutoff Date Additions Filter:** Toggle the **🎯 Date Cutoff** selector to highlight text introduced after a target date:
    * ⚡ **`⚡ Post-Cutoff Added`**: Highlighted in soft amber/yellow to clearly identify text additions introduced after cutoff dates.
    * **Exclusive Filter Mode:** Check **Show Only Post-Cutoff Additions** to hide clauses where matching text was already present prior to the selected priority date (filtering out prior art).
* **Universal Specification Ingestion Dialog:**
  * **Unrestricted Catalog Access:** Master-detail browser spanning all ~1,500+ specifications across Series 01 through 55 and all Working Groups (RAN1–4, SA1–6, CT1–4).
  * **Live Search & Presets:** Filter by keyword, topic, or specification number with built-in quick presets for core 3GPP specifications.
  * **Explicit Checkbox Selection:** Dedicated checkbox column for unambiguous selection tracking with dynamic count badges (`Selected: N version(s)`).
  * **Smart Batch Selectors:** One-click helpers including **`⚡ Select All Unindexed`**, **`⭐ Select Latest per Release`** (supporting both decimal and 3-digit lettered versions like `i40` / `g30`), **`☑️ Select All`**, and **`◻️ Clear`**.
  * **Revision-Mark Filtering:** Automatically discards 3GPP Word change-mark files (`-rm` / `_rm`) during unzipping and local imports, ensuring only clean (`-cl`) specification text is indexed.
* **Multi-Part Split Document Parsing:**
  * **Split Document Sequencing:** Automatically detects, sequences, and unifies modern multi-part specification archives (e.g., `_s00_s04.docx`, `_s05_s08.docx`, `_s09_s14.docx`) into a single consolidated release model in SQLite.
  * **High-Performance XML Extraction:** Direct `lxml` parsing extracts document structure directly from OpenXML without Microsoft Word COM runtime overhead.
  * **Intelligent Heading & TOC Sanitization:** Distinguishes genuine 3GPP headings from numbered procedure call flow steps and strips Table of Contents (TOC) dot leaders and stub entries.
* **Rich Clause Content Inspector:**
  * **💡 Key Match Excerpt:** Dedicated callout banner at the top of the inspector displaying the exact matching paragraph and surrounding sentence context.
  * **Interactive Term Navigation:** **`[ ◀ Prev ]`** and **`[ Next ▶ ]`** buttons with an active match counter (`🎯 N Match(es)`) and auto-scroll to match locations.
  * **One-Click Citation Copy:** **`[ 📋 Copy Citation ]`** button formats and copies complete clause text with official 3GPP document headers, versions, and release dates.
  * **One-Click Chronological Navigation:** Clickable **`🟢 First Added in v...`** chip jumps directly to the matrix release where the clause was first introduced; **`[ 🔍 Diff Prev ]`** opens an instant side-by-side textual diff against the immediately preceding release.
* **Direct Document Dispatch & Clipboard Clause Assist:**
  * **Direct Word & PDF Launch:** Open the exact specification document directly from the Spec Search tab via the **`📝 Open Word ▾`** dropdown menu (Word, PDF, or containing folder), double-clicking any matrix cell, or right-clicking a cell.
  * **Multi-Part Document Targeting:** For split specifications (e.g. `TS 23.501` Rel-17/18/19), `SpecDocResolver` automatically selects and opens the exact sub-part file covering the selected clause (e.g., Clause 4.4.1 opens `_s00_s04.docx`).
  * **Clipboard Clause Assist:** When launching Word, the system automatically copies the clause reference (`Clause X.Y.Z`) to the Windows clipboard with a status alert, allowing delegates to jump directly to the target clause using <kbd>Ctrl+F</kbd> + <kbd>Ctrl+V</kbd>.
  * **On-Demand Retrieval & Extraction:** Automatically detects cached `.zip` files and extracts the clean specification (`-cl`, skipping revision-mark `-rm` files) on the fly, or offers 1-click background downloading with progress tracking if the document is not yet cached.
* **Clause-Level Comparison & Word Track Changes:**
  * **Unified In-App Diff & AI Analysis Dialog:** Interactive modal providing instant color-coded visual text diffs (`+` additions, `-` deletions, chunk headers), change statistics, and multi-tier prompt generation for LLM standards analysis.
  * **Programmatic Word Comparison (Track Changes):** Generate a native Microsoft Word comparison (`CompareDocuments` legal compare with tracked revisions) between any two clause versions directly from within the unified dialog or UI.
  * **Targeted Excerpting & Security Labeling:** Scans document headings with `WordHeadingScanner`, extracts only the targeted clause excerpts via `WordExcerptEngine`, applies corporate sensitivity labels, and executes Word COM comparison to produce a clean, focused Word document with native Track Changes.
  * **Universal Access:** Launch the unified comparison directly from the matrix cell right-click menu (**`🔍 Compare Clause (Diff & AI Analysis)...`**), or click **`📝 Open in Word (Track Changes)`** from inside the comparison dialog.
* **Two-Way Cross-Tab Synergy with Specifications Tab:**
  * **Deep Link to Specifications Tab:** Click the **`🔍 Specs Tab`** button in the Inspector or select **`Show in Specifications Tab`** from the matrix context menu to switch to the **📚 Specifications** tab with the exact document and release selected.
  * **Deep Link to Spec Search:** From the **📚 Specifications** tab, right-click any specification or version to select **`🔎 Track Evolution in Spec Search`** (instantly filtering the search matrix to that spec and focusing the active version).
  * **1-Click Version Indexing into Spec Search DB:** In the **📚 Specifications** tab, right-click any row to select **`📥 Index vX.Y.Z into Spec Search DB`** (or **`🔄 Re-index vX.Y.Z`** if already present), or select multiple specifications to batch-index active versions (**`📥 Index selected (N) into Spec Search DB`**). The engine detects locally extracted Word documents to avoid re-downloading, switches to Spec Search, and automatically activates and focuses the newly indexed release in the tree.
* **Persistent Search Configuration:**
  * Active search queries, clause filters, cutoff dates, and exact checked specification versions are automatically persisted to `spec_search_config.json` and restored across application sessions.
* **Non-Blocking Database Maintenance:**
  * Background `SpecSearchWipeWorker` thread enables fast, freeze-free database resets with automatic checkpointing and SQLite schema reconstruction.

---

### 🔬 3GPP Protocol Evolution Matrix & Inspector (NAS, ASN.1 & PFCP)
* **Comprehensive Multi-Protocol Ingestion:**
  * **NAS Protocols:** Complete support for **5GS NAS (TS 24.501)** and **EPS NAS (TS 24.301)**.
  * **ASN.1 Protocols:** Native support for **NR RRC (TS 38.331)**, **LTE RRC (TS 36.331)**, and RAN3 Application Protocols: **NGAP (TS 38.413)**, **XnAP (TS 38.423)**, **F1AP (TS 38.473)**, **E1AP (TS 38.463)**, and **S1AP (TS 36.413)**.
  * **GTP-U Protocol:** Native support for **GTPv1-U (TS 29.281)** covering user plane tunnels across 5GS (`N3`, `N9`, `N19`, `F1-U`, `Xn-U`, `W1-U`), EPS (`S1-U`, `X2-U`, `S5/S8`), and legacy interfaces. Parses signalling messages (Echo Request/Response, Error Indication, End Marker, Tunnel Status) and synthesizes G-PDUs with Clause 5.2 Extension Headers (PDU Session Container, NR RAN Container, PDU Set Information Container) with per-interface filtering.
  * **PDU Session User Plane Protocol:** Native support for **TS 38.415** (PDU Session & PDU Set Information User Plane Protocols) over `NG-U`, `Xn-U`, `F1-U`, and `N9` interfaces. Parses DL/UL PDU Session Information and DL PDU Set Information frames into fine-grained bit-level fields, tracking delay measurement flags, QoS Monitoring timestamps, and PDU Set sequence numbering across releases.
  * **PFCP Protocol:** Native support for **PFCP (TS 29.244)** spanning both 5GC (`N4`, `N4mb`) and EPC (`Sxa`, `Sxb`, `Sxc`) reference points. Parses top-level Node-Related (Clause 7.4) and Session-Related (Clause 7.5) PDU messages alongside the master Information Element Type registry (Table 8.1.2-1).
  * **Hierarchical Grouped IE Unrolling:** Recursively traverses and unrolls nested PFCP Grouped IEs (e.g., `Create PDR └─ PDI └─ SDF Filter`, `Create FAR`, `Create URR`, `Usage Report`) into the Evolution Matrix with tree indentation and depth tracking.
  * **Interface Applicability Metadata & Filtering:** Automatically indexes per-IE interface applicability tags (`Sxa`, `Sxb`, `Sxc`, `N4`, `N4mb`) and provides a dynamic UI filter dropdown (`[All Interfaces | N4 | N4mb | Sxa | Sxb | Sxc]`) that appears whenever a PFCP message is selected.
  * **Release 20+ Multi-Part Document Ingestion:** Automatically detects, sequences, and parses modern split 3GPP specifications by aggregating all clause sub-documents into a single unified release model.
  * **Automated Legacy `.doc` Conversion:** Automatically converts older binary Word 97–2003 `.doc` specifications to `.docx` via headless COM automation or LibreOffice with Protected View bypass and NTFS Zone Identifier unblocking before parsing.
  * **High-Performance XML Parsing:** Direct `lxml` extraction parses message definition tables, ASN.1 syntax blocks, PFCP Grouped IE tables, and field description tables directly from `.docx` archives without requiring Word runtime overhead.
* **Evolution Matrix & Visual Diffing:**
  * **Dual Browse Modes (`[📨 Messages | 🧩 All IEs / Types]`):** Seamlessly toggle between browsing standard Protocol Messages/SIBs and browsing all parsed Clause 9 / Annex D / ASN.1 / PFCP Information Elements directly. In **`🧩 All IEs / Types`** mode, you can inspect standalone handover containers (e.g., `Intra N1 mode NAS transparent container 9.11.2.6`), security identifiers (e.g., `Serving network name 9.12.1`), and nested containers without needing an associated Clause 8 message, viewing their release-evolution matrix and full Clause 9 octet/bit diagrams.
  * **Hierarchical Sequence, Group & Container Unrolling:** Recursively unrolls nested ASN.1 sequence/choice fields, PFCP grouped structures, and encapsulated NAS container payloads (e.g., unrolling `Payload container 9.11.3.39` to reveal `└─ UE parameters update transparent container 9.11.3.53A`, `└─ SOR transparent container 9.11.3.51`, `└─ CIoT small data container 9.11.3.18B`, and `└─ Service-level-AA container 9.11.2.10` with its sub-parameters `9.11.2.11`–`9.11.2.18`), allowing you to track high-level and deep parameter changes across releases simultaneously.
  * **Visual Release Diffing:** Color-coded matrix cells immediately highlight field additions (🟢 Green), removals (🔴 Red), and format/type modifications (🟡 Yellow) between chronological 3GPP releases.
  * **Hierarchical Specification Tree:** Interactive tree view (`QTreeWidget`) grouping releases under collapsible specification parents (`TS 38.331`, `TS 29.244`, `TS 24.501`, `TS 24.301`, etc.) with master toggles, specification-wide selection, and right-click deletion context menus.
  * **Persistent Filter Configuration:** Tree expansion states, active releases, browse modes, message/IE selections, and search terms are automatically persisted in `nas_config.json` across sessions.
* **Dual-Layer, Tokenized & Extended Description Search:**
  * **Multi-Word Tokenized Matching:** Search queries automatically split into whitespace-delimited tokens across names, field paths, and definitions, allowing flexible searches (e.g., searching `"UE parameters transparent container"` seamlessly matches `"UE parameters update transparent container"`).
  * **Debounced Filtering:** Dedicated 200ms debounced search bars for Message/IE Names and Information Elements / Fields.
  * **Clean TOC & Stub Sanitization:** Automatically detects and strips Word Table of Contents (TOC) page numbers and empty structural/Void section headers to keep the catalog clean and accurate.
  * **Extended Description Search (`📖 Desc`):** Toggle deep text search across underlying Clause 8/9 IE descriptions and ASN.1 field description tables (e.g., searching `"emergency"` or `"slicing"` highlights matching messages even if the keyword is not in the field name).
* **Structure & Field Descriptions Inspector:**
  * **Bit-Level Structure Rendering:** Renders bit-level octet diagrams (Figure 8.x/9.x) and value coding tables (Table 8.x/9.x) with full OpenXML `gridSpan` (colspan) and `vMerge` (rowspan) support.
  * **ASN.1 Syntax & Descriptions View:** Renders syntax-highlighted ASN.1 definition blocks accompanied by formatted 3GPP Field Description tables.
  * **Reverse Field/IE Lookup:** Interactive header badge (`Used in: N messages ▾`) and right-click matrix context menu to trace and jump to all messages referencing a given IE or ASN.1 type across active releases.

---

### 🗄️ Database Maintenance, Inspection & Compaction
* **System-Level Database Manager:** Integrated `🗄️ Database` tool accessible directly from the bottom Console Panel.
* **Read-Only SQLite Schema & Data Inspector (`DatabaseInspectorDialog`):** Click **Inspect** on any tracked database to launch a dedicated browser displaying all tables and views, live row counts, SQL DDL creation schemas, column data types, PRAGMA diagnostics (page size, freelist page count, journal mode), and interactive paginated table data without needing external database tools.
* **Automatic Freelist & WAL Compaction:** Inspects on-disk database sizes and active Write-Ahead Logs (`-wal`), executing `PRAGMA wal_checkpoint(TRUNCATE)` and `VACUUM` to defragment pages and reclaim megabytes of disk space after heavy scraping or database wiping.
* **Batch Maintenance:** One-click **Compact All Databases** to optimize `3gpp_data.db`, `3gpp_protocol_data.db`, `3gpp_spec_search.db`, and auxiliary cache databases concurrently.

---

### 🖥️ COM Process & Task Manager (`ProcessManagerDialog`)
* **Dedicated Office Process Management:** Integrated `🖥️ Task Manager` tool accessible directly from the bottom Console Panel.
* **Active Process & Document Inspection:** Tracks live instances of Microsoft Word (`WINWORD.EXE`), PowerPoint (`POWERPNT.EXE`), and Visio (`VISIO.EXE`), displaying process IDs and all currently open document titles.
* **Safe Document-Level Closing:** Inspect open documents in Word and PowerPoint and cleanly close specific files via COM automation without killing the host application.
* **Session Preservation & Selective Keep-Open:** Uncheck documents you wish to keep open during automated comparison or slicing workflows; the system respects user exclusions.
* **Emergency Recovery & Orphan Cleanup:** Single-click **Kill Process** or **Kill All Office Processes** to immediately terminate frozen or orphaned background COM instances, unlocking lingering file locks.

---

### 📖 Help & Configurable Useful Links Menu
* **Tab-Bar Dropdown Menu (`📖 Help ▾`):** Integrated directly into the top-right corner of the main tab bar, providing instant 1-click access to curated 3GPP developer tools, specification viewers, and official portals.
* **Curated Community Tools & Portals:**
  * **3GPP Guru Viewer (F-17):** Fast web-based 3GPP specification viewer with syntax highlighting and cross-clause hyperlinks (`https://3gpp.guru/`).
  * **SpecViewer:** Release-by-release 3GPP specification browser with clause search and diffing (`https://specviewer.org/`).
  * **TDocHamster:** 3GPP TDoc search, comparison, and meeting contribution visualizer (`https://tdochamster.com/`).
  * **3GPP Forge OpenAPI Specs (F-18):** Direct launcher to official 3GPP OpenAPI 3.0 YAML definitions for 5G Core, NWDAF, and SBA on GitLab (`https://forge.3gpp.org/rep/all/apis`).
  * **Official 3GPP Portals:** Instant access to 3GPP Portal (3GU) and 3GPP Work Plan / GANTT release schedules.
  * **Application Documentation (<kbd>F1</kbd>):** Retains standard <kbd>F1</kbd> keyboard shortcut to immediately open local application documentation.
* **100% Configurable via JSON (`config/useful_links.json`):** Add, remove, re-order, or customize categories, titles, URLs, and emoji visual cues (🌐, 🔍, 🐹, 🛠️, 🏛️, 📅) by simply editing `config/useful_links.json`. Updates are detected dynamically when opening the menu without restarting the application, with a 1-click **⚙️ Configure Useful Links (JSON)...** editor shortcut.

---

### 📡 3GPP Meeting, Specification & Work Items Database
* **Asynchronous Three-Phase Syncing Engine:** 
  * **Phase 1 (FTP Directory Mapping):** Scrapes the 3GPP FTP archives in parallel to instantly populate your database with all available meeting numbers, gracefully handling hidden RAN Ad-Hoc (`TSGR_AHs`) subdirectories.
  * **Phase 2 (Deep Document Scrape):** Crawls the `Docs/` folder of every meeting. Uses smart regex stripping to ignore file extensions and revisions, mathematically sorting the files to determine the first and last TDocs of the meeting.
  * **Phase 3 (DynaReport Upserting):** Injects metadata (Location, Start/End Dates, Ad-Hoc/Electronic status) by fetching the legacy 3GPP Portal HTML tables.
* **Targeted Quick Fetch:** Instantly sync individual specifications (e.g., `23.801-01`) or entire specification series (e.g., `23`) directly from the FTP server without needing to run a lengthy full database sync.
* **Unified Specifications Filter Bar & Multi-Series Selection:**
  * **Zero-Modal Live Filtering:** Replaces legacy modal filter dialogs with instant, inline filtering directly above the specifications table with debounced auto-refresh.
  * **Multi-Series Checkable ComboBox:** Filter across multiple specification series simultaneously (e.g., `23` and `24` or `38` and `36`) using `CheckableComboBox`.
  * **Dynamic TSG & WG Synergy:** Selecting a TSG chip (`SA`, `RAN`, `CT`) dynamically scopes the working group dropdown choices to relevant groups (e.g., `SA1`..`SA6`) without state collisions.
  * **Faceted Attributes:** Real-time dropdown and segmented controls for Radio Technology (`5G`, `LTE`, `6G`, `3G`, `2G`), Document Type (`All`, `TS`, `TR`), Release Version (`Ver:`), and Scope Toggles (`💾 Downloaded Only`, `⭐ Favorites`).
  * **Dismissible Active Filter Badges:** Visible pill badges for every active constraint (`[Group: SA5 ✕]`, `[Tech: 5G ✕]`, `[Series: 23, 24 ✕]`, `[Tag: data framework ✕]`) featuring 1-click single-constraint removal with hover feedback, alongside a unified `❌ Reset All` button.
  * **Dual-Tier Debouncing & High-Performance Rendering:** Employs twin single-shot debounce timers (300ms for text search and version inputs, 150ms for faceted dropdowns, chips, and toggles) alongside Qt batch update suppression (`setUpdatesEnabled`), deferred on-demand menu population via `aboutToShow`, active tag diffing, and conditional SQLite query joins to deliver instantaneous, stutter-free filtering across 84,000+ specifications.

* **🏷️ Specification Tags, Personal Notes & Universal Search Integration:**
  * **Custom Domain Tags & Multi-Line Notes:** Assign custom domain tags (e.g. `data framework`, `analytics`, `charging`, `NWDAF`) and personal remarks/comments to any specification.
  * **Backwards-Compatible Storage (`spec_user_annotations`):** Stored in a dedicated SQLite table within `3gpp_data.db`, completely isolated from core specification rows so your tags and notes survive all DynaReport updates, catalog re-syncs, and schema migrations.
  * **Normal Search Integration:** Searching in the primary search bar of the Specifications tab (or universal Ingestion dialogs) automatically queries tags and comments alongside specification numbers and titles.
  * **In-Table Badges & Hover Previews:** The specifications table dynamically renders custom tag pills (`[🏷️ data framework]`) directly beneath the specification title, alongside a `📝 Note` pill showing full multi-line notes on hover.
  * **Convenient Multi-Surface Editing:** Double-click any table row, select `🏷️ Edit Tags & Notes...` from the `⋮` action menu or table context menu, or manage annotations directly inside `SpecInfoDialog` (`ℹ️ View Details`).
  * **Batch Tagging:** Select multiple specifications (e.g., an entire set of SA5 specs) and right-click $\rightarrow$ **`🏷️ Add Tag to selected (N)...`** to tag all of them at once.
  * **Tag Filter Dropdown & Pills:** Filter the table by any user tag using the dedicated `🏷️ Tag:` filter dropdown or dismissible filter badges.

* **📋 Specification Quick Copy & Rich Clipboard Engine:**
  * **Dual-Format Clipboard (`QMimeData`):** Seamless multi-target clipboard export setting both plain text and rich HTML. Rich-text applications (Outlook, Word, Teams, OneNote) paste formatted text with active hyperlinks, while plain-text editors (Notepad, git commit, terminal, IDE) paste clean, unadorned text.
  * **Flexible Link Targeting (Web, Direct ZIP, or Both):**
    * *🔗 Dual Links (Web + ZIP):* Integrates both destinations cleanly into a single citation line—the specification number links to the official 3GPP DynaReport status overview, while a dedicated version badge (e.g., `[ZIP v20.2.0]`) links directly to the ZIP download file on 3GPP FTP (`<a href="...DynaReport/23501.htm">TS 23.501</a> (<a href="...23501-k20.zip">ZIP v20.2.0</a>): Title (WG)`).
    * *🌐 3GPP Web Links (DynaReport):* Standard specification citations where the number links directly to the 3GPP DynaReport status page.
    * *📦 Direct ZIP Download Links:* Citations where the specification number links directly to the target version's ZIP archive file on the 3GPP FTP server (`<a href="...23501-k20.zip">TS 23.501 (v20.2.0)</a>: Title (WG)`).
  * **Multi-Surface Access & Version Context:**
    * *Right-Click Table Context Menu:* Copies using the currently selected version from that row's dropdown (or latest release).
    * *Column 0 Action Menu (`⋮`):* Fast 1-click citation copying and granular submenus.
    * *Column 2 Document Menu (`📝 Open Word ▾` / `📦 Download ▾`):* Dedicated `📋 Quick Copy (vX.Y.Z)` submenu targeting that exact version's download link.
    * *Specification Details Dialog (`SpecInfoDialog`):* Bottom footer `📋 Quick Copy ▾` dropdown button.
  * **Configurable Default Mode (`Ctrl+C`):** Select one or more rows in the table and press `Ctrl+C` to instantly copy complete citations. Delegates can configure the default link style (`Both Links`, `Web Link`, or `Direct ZIP`) in 1 click under **`⚙️ Default Link Mode (Ctrl+C)`**, persisting to `specs_config.json`.
  * **Batch Multi-Select Copying:** Select multiple rows to copy comma-separated numbers, multi-line title lists, full citations with working groups, bulleted Markdown lists (`- [TS 23.501](web) ([ZIP v20.2.0](zip)): Title (WG)`), or URL lists.
  * **Non-Blocking Tooltip Feedback:** Every copy action displays an instant floating cursor tooltip (`📋 Copied {label} to clipboard!`) without interrupting delegate workflow.

* **✂️ 3GPP Specification Clause & Excerpt Extractor:**
  * **Instant Excerpt Dialog from Version Menu:** Select any specification release in the Specifications tab and click **`✂️ Extract Clauses (Excerpt)...`** to launch a modeless clause extractor window.
  * **Automated Background Fetch & Smart Unzipping:** Automatically downloads missing archives from 3GPP FTP, skips YAML, OpenAPI, and PDF files, and extracts clean Word documents while filtering out revision marks (`-rm`).
  * **Multi-Part & Legacy Support:** Naturally sequences multi-part split specifications (e.g., `_s00_s04`, `_s05_s08`) and automatically converts legacy binary `.doc` files to `.docx` via COM automation or headless LibreOffice.
  * **Interactive Tree Browser with Search Filter:** Filter clauses by number or keyword (e.g., `5.2`, `registration`, `NWDAF`) in an interactive tree view with auto-cascading checkboxes.
  * **Formatting & Corporate Label Preservation:** Uses the isolated fragment engine to build clean DOCX excerpts, merge multi-part sections, restore source typography and styles, and apply corporate sensitivity labels.
  * **One-Click Auto-Open:** Automatically prompts for a save destination (defaulting to the local specification directory) and opens the generated excerpt in Microsoft Word upon completion.

* **3GPP Work Items (WIs) Synchronizer:**
  * **Direct Work Item Addition & Live Portal Scrape (`➕ Add WI`):** Add and inspect individual Work Items instantly without needing to run a full multi-WG sync. Enter a pure Work Item Code (e.g. `1121123`), prefixed ID (`#1121123`), or full 3GPP Portal URL (`https://portal.3gpp.org/desktopmodules/WorkItem/WorkItemDetails.aspx?workitemId=1121123`). The background worker (`ManualWIFetcherThread`) automatically fetches live metadata from the 3GPP Portal—extracting the WI title, acronym, release, status, start/end dates, latest WID version, mapped working groups (translating shorthand codes like `R3` $\to$ `RAN3`), and secretary remarks—with real-time editable preview and 1-click ⭐ Star and 🎯 Follow assignment upon saving.
  * **Parallel Multi-WG Scraper:** Concurrently scrapes active Work Items across all 19 Technical Specification Groups and Working Groups (SA, SA1-6, RAN, RAN1-6, CT, CT1-6) from official 3GPP dynamic report pages using multi-threaded execution (5 workers).
  * **High-Performance Bulk Upsert:** Utilizes atomic SQLite bulk transactions (`executemany` with `ON CONFLICT DO UPDATE`) to instantly sync thousands of work items and map them to their respective working groups via relational sidecar tables (`work_items`, `wi_group_map`, `wi_remarks`).
  * **Interactive UI Tab:** Features a dedicated tab with a real-time progress bar, status feedback, and helpful button tooltips. Includes debounced multi-select CheckableComboBox filters (Release, WG) with persistent state-saving, chronologically sorted historical remarks via a custom interactive UI bubble, and clickable WID hyperlinks that automatically route through the global TDoc fetcher or 3GPP Portal.
  * **Personal & Team Work Item Tracking ("Starred" & "Following"):**
    * **Independent Dual-Flag Tracking:** Supports distinct tracking workflows for personal interest (⭐ Starred) and assigned delegation responsibility (🎯 Following) without altering the shared SQLite database schema.
    * **1-Click In-Row Toggling (`Track` Column):** Features dedicated left/right hit targets for one-click star and follow toggling directly within the table view.
    * **Multi-Chip Quick Filters:** Dedicated `🎯 Following (N)` and `⭐ Starred (M)` filter chips support single and union multi-filtering across thousands of records.
    * **Targeted Pre-Meeting Sync (`🎯 Update Tracked`):** Asynchronously fetches latest WIDs and Secretary Remarks for only tracked Work Items in seconds.
    * **Inspector Dialog & Batch Actions:** Toggle tracking status directly inside `WorkItemInfoDialog` or via right-click multi-selection context menus.
    * **Cross-Module Meeting Integration:** Followed and Starred Work Items seamlessly project into the Meeting TDocs Viewer as instant filter chips (`🎯 Followed WIs`, `⭐ Starred WIs`), enabling delegates to isolate contributions related to their assigned topics in one click.

* **3GPP Work Items (WIs) & Specification Linkage:**
  * **Relational Mapping (`spec_wi_map`):** Bi-directionally maps 3GPP Specifications to Work Items during Pass 2 DynaReport scraping without requiring rigid locks on un-synced WIs.
  * **Specification Inspector Chips:** Details dialogs display interactive primary (⭐) and secondary Work Item chips with direct 3GPP portal navigation.
  * **Work Items Table & Local Specs Inspector:** The Work Items tab features dedicated **WG** and **Linked Specs** columns, local specification inspectors (`LinkedSpecsDialog`), and one-click citation copy actions.

* **Intelligent TDocs Manager:**
  * **Smart Global TDoc Search:** Instantly locate and download any document across the entire database. Just type a TDoc number (e.g., `S2-2605740r11`) and the UI will dynamically reveal minimalist quick-actions to download the specific file or open its parent meeting context—all without leaving the main dashboard.
  * **🤖 Instant AI Technical Triage:** Right-click any proposal row or open its details card to launch the AI Triage assistant. Streams a focused summary covering the problem statement, proposed normative changes, and contentious impact points, with 1-click status assignment and note persistence into SQLite.
  * **Persistent Personal Notes & Status (Sidecar Database):** Keep a private, local SQLite database that "overlays" your data onto the 3GPP list. Double-click any TDoc to assign a color-coded status (🟢 Support, 🔴 Object, 🟡 Monitor) and save personal notes. Your data survives perfectly even when downloading fresh 3GPP Excel updates.
  * **Tracked Work Items Quick Filters (🎯 Followed & ⭐ Starred):** Instant 1-click filter bar toggle chips (`🎯 Followed WIs`, `⭐ Starred WIs`) that isolate meeting contributions associated with your assigned delegation responsibilities and personal watchlist. Dynamically displays badge counts of matching TDocs, automatically reflects newly tracked items from the Work Items tab on focus, and seamlessly combines via `AND` logic with other filters (e.g., TDoc Status like "Approved/Agreed", Contributing Company, Type). By default, includes unassigned incoming Liaison Statements (`LS in`) without an assigned Work Item so crucial inter-group communications are never missed (toggleable via right-click context menu on either chip). Features robust token normalization stripping WG suffixes (`-ARC`, `-CT`, etc.) and `FS_` prefixes, with fallback to parenthesized acronyms in Agenda Item descriptions for discussion papers lacking explicit metadata.
  * **⚡ Opportunistic Work Item Auto-Enrichment:** When loading or refreshing any meeting's TDocs list, a dedicated background worker (`MeetingWIEnricherThread`) automatically scans the "Related WIs" column off-thread for 3GPP Portal hyperlinks (`workitemId=\d+`). It performs a sub-millisecond check against SQLite (`work_items`), and quietly scrapes missing/unindexed Work Items from the 3GPP Portal in the background with zero impact on UI responsiveness or TDoc table loading speed. Employs a persistent JSON audit marker (`<excel>.wi_checked.json`) ensuring subsequent opens bypass checking (~0 ms), and dynamically updates tracked tokens and filter badges upon completion. Fully configurable in Meetings Configuration.
  * **📑 Post-Meeting Report & Dossier Markdown Generator (`Export Post-Meeting Report (.md)...`):**
    * **Hierarchical Release & Work Item Organization:** Exports meeting contributions into an organized dossier under `{meeting_dir}/Export/PostMeeting_<WG>_<Meeting>_<timestamp>/`, grouped chronologically by 3GPP Release (e.g., `Rel-19`, `Rel-20`) and then by Work Item (e.g., `FS_6G_ARC`).
    * **Contribution Type Categorization:** Within each Work Item, documents are cleanly routed into dedicated type subfolders: `LS_IN/` (incoming liaison statements), `pCR/` (agreed pseudo-CRs), `CR/` (agreed Change Requests), and `LS_OUT/` (agreed outgoing liaison statements). General non-WI incoming liaison statements are collected under `General/LS_IN/` (enabled by default via `Include unassigned & general LS in`).
    * **Liaison Statement Reply Tracking & Guaranteed Dossier Inclusion:** Automatically traces bidirectional reply relationships between incoming and outgoing liaison statements (`Reply in` / `Reply to` with `Revised to` revision chain resolution). The `LS in` status table directly links to reply documents (`replied to in [S2-XXXX](...)`), while `LS out` status tables reference incoming requests (`approved (reply to [S2-YYYY](...))`), supplemented by reciprocal `[!TIP]` and `[!NOTE]` header callouts in individual dossiers. Crucially, any approved reply `LS out` answering an included `LS in` is guaranteed to be exported to disk (routed under `General/LS_OUT/` if its Work Item was not in the export scope) so that all reply links resolve without creating unwanted non-followed Work Item sections in the report.
    * **Document-Type-Aware Revision Mode:** Automatically resolves proper revision rendering modes (`resolve_revision_mode_for_type`): specification change proposals (`CR` and `pCR`) retain Word track changes (`revisions="markup"` with `[INSERTED]`/`[DELETED]` tags and markup legend), while liaison statements (`LS in`, `LS out`, `LS`) are recognized as plain correspondence documents without specification change marks and exported in clean `revisions="accept"` mode without change markup or legends.
    * **3GPP Server URL & Provenance Preservation:** Resolves official 3GPP server HTTPS download URLs (`https://www.3gpp.org/ftp/.../{tdoc_id}.zip`) via `tdoc_url_resolver.py`, routing revisions (`r01`, `rev2`) to `/INBOX/Revisions/` and base contributions to `/Docs/`. Embeds permanent provenance URLs into standalone dossier headers (`- **3GPP Server URL:**`) and renders dual offline/online links (`[Dossier](...) · [📥 3GPP](...)`) in executive and Work Item index tables so readers can immediately download original archives even after dossiers are aggregated or distributed downstream.
    * **Strict Post-Meeting Quality Filter:** Automatically filters for normative meeting outputs—capturing all received incoming liaison statements and approved/agreed contributions (`agreed`, `approved`), while strictly excluding discussion papers, non-agreed proposals, and withdrawn CRs.
    * **Multi-Agenda Item Coverage:** For Work Items spanning multiple Agenda Items (such as SA2's `FS_6G_ARC` covering AI 5.2.1, 5.2.2, etc.), dynamically aggregates and displays all associated AIs with their meeting descriptions in local and master indices.
    * **Cross-WI Contribution Management:** Contributions linked to multiple Work Items are seamlessly included across all associated WI dossiers, embedding a prominent callout notice (`> [!NOTE] Cross-WI Contribution: This document is also associated with ...`) with clickable reciprocal cross-references.
    * **Executive Master Index & Local Dossier Summaries:** Generates a comprehensive master `index.md` containing executive summary tables, outcome breakdown matrices, and direct relative Markdown hyperlinks to individual TDoc dossiers, along with localized `index.md` files for every Work Item.
    * **Collision-Free Media Isolation:** Replicates embedded figures and diagrams into local `media/` subdirectories with `{tdoc_id}_` filename prefixes, guaranteeing image isolation without cross-document file collisions.
    * **Flexible Scope & Options Dialog (`PostMeetingExportDialog`):** Choose between exporting `🎯 Followed Work Items`, `⭐ Starred Work Items`, `🎯+⭐ All Tracked`, `🌐 Whole Meeting`, `👁️ Currently Visible`, or a custom Work Item checklist, with toggles for contribution types and automated folder reveal upon completion.
  * **Smart Revision Inheritance:** When a TDoc gets a new revision during a meeting, the new child document automatically inherits a "Ghost" version of the personal notes and status you assigned to the base document!
  * **Interactive Secretary Remarks:** TDocs mentioned in the Secretary Remarks are automatically identified and converted into hyperlinks. Left-click a link to instantly jump to that row (intelligently wiping active filters if necessary), or right-click to download it, trigger AI Triage, or add it to your Comparison Cart.
  * **Natural Sorting, Dynamic AI Labels & Agenda CSV Support:** Bulletproof multi-select dropdowns and natural numerical sorting for complex multi-level Agenda Items (e.g., AI 20.6.2 sorts correctly before 20.6.11). Natively parses both **SA2 and RAN2 agenda CSV formats** (`agenda.csv`), intelligently extracting valid acronyms and subtopics while ignoring non-acronym parenthetical words (e.g., `(un)`, `(re)`), falling back gracefully to clean descriptions, and providing full-title hover tooltips across tables and statistical exports.
  * **Comprehensive Analytics Dashboards:** Generate interactive offline HTML Plotly reports detailing TDoc outcomes, top contributing companies, and complex strategic alliance network graphs (co-signing clusters) using Louvain community detection algorithms.
  * **SA2 Electronic Revisions & Agenda Parsing:** Automatically parses `TdocsByAgenda.htm` to extract comments, inject on-the-fly revisions directly into your table, and provides a "No Comments Only" filter. For eMeetings, it automatically scrapes the `INBOX/Revisions/` FTP folder.
  * **SA2 Chairman's Notes & Session List Ingestion (`.doc` / `.docx`):**
    * **Frosted Drop Overlay:** Drag and drop `.doc`, `.docx`, `.htm`, or `.html` session documents onto the TDocs window; a visual frosted-blue drop overlay appears with dashed borders and instant drop targets.
    * **Non-Blocking Background Worker (`WordAgendaImporterThread`):** Copies imported files to `{meeting_dir}/Agenda/`, unblocks NTFS Zone Identifiers, converts legacy macro-bearing `.doc` files via headless LibreOffice, and parses table data in the background without freezing the UI.
  * **Multi-Action Resources & Terminal Menu:** Instantly jump to local cache directories, fetched HTML Agenda files, Main FTP folders, Docs/ folders, or Revisions/ folders directly from the UI. Select **💻 Open Terminal in Meeting Cache** from the meeting context menu or the TDocs folder menu to launch Windows Terminal (`wt.exe`) or PowerShell directly in the meeting's local storage folder.
  * **Quick Launch History:** Remembers your active working group session, allowing you to bypass the database table and jump back into your last opened meeting with a single click.

* **📊 Cross-Meeting Company Contribution Audit & Reporting:**
  * **Multi-Meeting Filtering Engine:** Filter and audit contributions across arbitrary date ranges, multiple Working Groups (e.g., SA2, SA1, RAN2), specific 3GPP Work Items, and canonical company entities.
  * **Canonical Company Sanitization & Co-Author Rules:** Integrates directly with `CompanySanitizer` to normalize diverse corporate subsidiaries and brand spellings. Easily toggle between **"Any co-author (joint contribution)"** and **"Primary source only (first listed)"**.
  * **Work Item (WI) Auto-Completing Token Box:** Search and add multiple Work Item acronyms/codes using a tag/chip input backed by `WorkItemsDatabase` autocompletion.
  * **High-Performance Bounded Concurrency & O(1) Cache:** Operates a laptop-safe `ThreadPoolExecutor` capped at 3 worker threads to prevent CPU throttling, memory bloat, and 3GPP server rate-limiting. Bypasses expensive full-table regex scanning by resolving `Related WIs` directly from 3GPP document sheets and agenda descriptions. Employs `TDocsParser` sidecar JSON caches (`.xlsx.json`) for sub-second cache hits on previously opened meetings, with a one-click toggle to **"Bypass local cache (re-fetch from 3GPP)"**.
  * **Modeless, Multi-Window Experience:** Operates as an independent, non-blocking top-level window (`Qt.Window`), enabling full simultaneous access to the main meetings tab, individual meeting tables, and email inspectors.
  * **Clean Application Aesthetic:** Renders in neutral monochrome text without distracting colored cells or heavy badge backgrounds, matching the application's clean design theme.
  * **Dedicated Working Group (WG) & Abstract Columns:** Dedicated **WG** column at the start of the table for instant committee attribution, and an **Abstract** column with multi-line tooltip wrapping and popup reading support.
  * **Direct Row Navigation & Context Menu:**
    * **Smart Double-Click:** Double-clicking `WG` or `Meeting` launches that meeting's full table in `TDocsWindow`; double-clicking `Abstract` opens the clean pop-up reader; double-clicking `TDoc` (or any other cell) downloads (if missing), extracts, and opens the actual document in Word via `TDocActionThread`.
    * **Right-Click Command Menu:** Context actions to **🗓️ Open Meeting Table**, **📂 Open Docs Folder** (opens the meeting's public FTP `Docs/` directory in your browser), **📄 Open TDoc Document** (downloads, extracts, and launches the Word document directly), **🌐 View on 3GU Portal**, **🔗 Copy URL** (copies direct ZIP download link to clipboard), **📋 Copy TDoc / Title**, and **📝 View Abstract...**.
  * **In-App KPI Summary Bar:** Real-time summary metric cards displaying **Total Contributions** (with WG coverage count), **Agreed / Approved** (total count and percentage of total submissions), **Collaboration** (% joint vs. solo), **Top Co-signer** (most frequent partner organization with joint count), and **Top Work Item** (accounting for multi-entry WIs while filtering out placeholders like `DUMMY` and `Unspecified`).
  * **Interactive HTML Analytics Dashboard (`📊 Statistics Dashboard...`):** Generates an offline Plotly HTML report with native integer formatting featuring:
    * **Monthly Contribution Trend:** Stacked bar chart of outcomes grouped chronologically by Meeting End Date / Month (`YYYY-MM`).
    * **Overall Contribution Outcomes:** Donut chart with accurate counts and percentage tooltips using semantic 3GPP colors (Agreed/Approved in green, Not Treated in grey, Revised in indigo, Noted in amber, etc.).
    * **Top Co-signers:** Horizontal bar chart displaying partner organizations segmented by Working Group in stacked colors.
    * **Working Group Activity:** Stacked bar chart showing total contributions per Working Group categorized by outcome status.
    * **Work Item & Study Item Allocation:** Ranking of top 20 active Work Items, splitting multi-entry rows and filtering placeholders.
  * **Corporate Deutsche Telekom Magenta Excel Export:** Exports the aggregated, filtered view into a styled `.xlsx` workbook featuring zebra striping, centered alignments, and active download hyperlinks pointing directly to each TDoc's specific meeting Docs folder.

* **Fast Native Wi-Fi & 3GPP Network Detection:**
  * Automatically detects when you are connected to the official "3GPPWIFI" network during live standardization meetings.
  * **Zero-Subprocess In-Memory Detection:** Queries Windows Network List Manager (NLM) via native COM (`netprofm.dll`) and `netsh wlan`, resolving active SSIDs in ~2ms with zero shell overhead or AMSI security locks.
  * **Instant TCP Server Reachability:** Performs non-blocking TCP socket probes directly to the internal meeting server (`10.10.10.10:80/443`), bypassing slow `ping.exe` subprocesses.
  * **Decoupled Status Display:** Persistent status indicator in the bottom system bar displaying `🟢 <SSID> (Local Server Active)` or `🌐 <SSID>`. Enables dynamic local caching, proxy bypassing, and direct FTP routing through the high-speed local meeting network.

* **3GPP FTP Session Manager:** Automatically injects randomized User-Agents and HTTP Keep-Alive headers. Features a configurable **Humanness Delay** engine to bypass aggressive 3GPP server throttling and "Too Many Requests" blocks, which can be dialed down to 0.0 for maximum scraping speed.

* **🔒 Secure Corporate Proxy Profiles & DPAPI Credential Store:**
  * **Full Profile Encryption:** Protects corporate network topology and credentials by encrypting all proxy parameters (HTTP/HTTPS hosts, ports, Active Directory usernames, and passwords) into an isolated payload blob using Windows DPAPI (`core.utils.dpapi.SecureCredentialStore`).
  * **Zero-Plaintext On-Disk Invariant:** Plaintext hostnames, IP addresses, usernames, and passwords are permanently banned from `network_config.json`. If encryption fails or DPAPI is unavailable, the engine refuses to persist unencrypted credentials to disk.
  * **Authenticode & Windows Security Catalog Integrity:** Dynamically verifies that `crypt32.dll` and `win32crypt` are genuine Microsoft system binaries using `WinVerifyTrust` with Windows Security Catalog (`CatRoot`) fallback, 64-bit module address checks in `System32`, and workspace anti-shadowing guards.
  * **Application-Specific Entropy Salt:** Applies a secondary cryptographic salt so that other scripts running under the same Windows account cannot by chance decrypt the profile blob without applying the exact application salt.
  * **Instant Profile Toggle:** One-click activation/deactivation enables seamless switching between corporate proxy routing and direct public internet without discarding saved credentials.

---

### 📧 Universal TDoc Email Tracker & Inspection Dialog
* **Working Group-Agnostic Ingestion:** Indexes emails across any 3GPP Working Group (SA2, RAN2, CT1, etc.) directly from your Outlook folders without moving emails or touching server-side folders. Operates independently of the dedicated eMeeting logic to prevent regressions.
* **WG-Dependent Multi-Folder Profiles & Custom Tag Colors:**
  * Configure specific Outlook folders per Working Group (saved globally in `emails_config.json`).
  * Assign custom tags (e.g., `[WG]`, `[Disc]`, `[Offline]`, `[Inbox]`) and pick personalized badge colors using an interactive `QColorDialog`. Tags render in the conversation stream with custom contrasting colors.
* **Smart Quotation Boundary & Direct Message Detection:**
  * Differentiates whether a TDoc was cited in the **Subject**, the **Direct Body** of the message, or an inherited historical reply chain (**Quoted**).
  * Eliminates false-positive cascades where casual replies (`"ok, danke"`, `"+1"`) cite TDocs buried in older email footers.
  * Toggle **`☑️ Include Quoted Matches`** to hide or reveal conversational thread citations on demand.
* **Exchange Internal Senders & DMARC Resolution:**
  * Automatically resolves listserv rewrites (`LIST.ETSI.ORG`) and internal Exchange X.500 addresses (`/o=...` / `EX`) to primary SMTP addresses to ensure company sanitization recognizes internal colleagues.
* **Modeless, Multi-Window Architecture:**
  * The inspection dialog operates as an independent, modeless top-level window (`Qt.Window`). It never freezes or blocks the main TDocs list or background downloads, allowing you to snap windows side-by-side.
  * Multiple TDocs can be inspected concurrently without duplicate window spawning.
* **Interactive TDoc Linkifier:**
  * Automatically converts every detected 3GPP TDoc number in the Subject line, Match Excerpt banner, and Body text into a clickable link.
  * Current document family numbers are highlighted in amber (`#FFF176`), while cross-referenced TDocs appear with interactive links (e.g., `🔗 S2-2608457`).
  * Clicking any referenced TDoc instantly launches an inspection window for that document.
* **Reading Pane Controls & Standalone Viewer:**
  * **Interactive Vertical Splitter:** Drag the splitter bar between the email list and the reading pane to adjust viewing proportions.
  * **Unicode Whitespace Compression:** Automatically strips invisible non-breaking spaces (`\xa0`) and collapses excessive blank lines from Word/Outlook formatting into clean, readable text.
  * **💡 Match Found Callout:** Displays an excerpt banner directly above the message showing the exact surrounding sentence context where the TDoc was found.
  * **`⧉ Pop Out View`:** Detaches the message preview into an independent, fully resizable viewer (`StandaloneEmailReaderWindow`), ideal for laptop screens or secondary monitors.
* **Read / Unread Lifecycle & Ignore Engine:**
  * Track local read states in SQLite (`general_emails.is_read`).
  * Selecting an email marks it as read after an 800ms debounce.
  * Multi-select rows with `Ctrl` or `Shift` to batch Mark Read, Mark Unread, Ignore, or Delete.
  * **`🚫 Ignore` Action:** Suppresses high-volume distribution list announcements or rapporteur compilation emails from all document counts without deleting them. Ignored flags are preserved across re-syncs. Toggle **`Show Ignored`** to review or un-ignore them.
* **TDocs Window Integration:**
  * **`Emails` Column:** Displays aggregate family email counts with unread badges (e.g., `✉️ 5 (🔵 2)`).
  * **Context Menu:** Right-click any row to view related emails or toggle all emails for that TDoc's revision family between read and unread.
  * **`📧 Emails ▾` Header Menu:** One-click menu to sync related emails, configure folders, mark all as read, or execute a high-speed wipe of the generic emails database.

---

### 📧 eMeeting Email Manager (Dedicated SA2 eMeeting Dashboard)
* **High-Performance Sync Engine:** Connects directly to your local Microsoft Outlook via COM automation. Pulls, parses, and indexes thousands of eMeeting mailing list emails in milliseconds using SQLite chunked batching (`executemany`) with zero memory spikes.
* **Master-Detail Thread Architecture:** Bypasses broken Outlook reply chains by logically grouping emails purely by parsed TDoc numbers. The UI features a split-screen design: a Left Panel displaying active TDoc threads and a Right Panel displaying the isolated, chronological conversation for the selected topic.
* **Intelligent 3GPP Parser:** Uses smart regex to extract TDoc numbers (6-8 digits), Agenda Items, Revisions, and free text directly from standard 3GPP bracketed subject lines and email bodies.
* **DMARC Listserv Bypass:** Automatically detects when 3GPP mailing lists rewrite the sender address to `LIST.ETSI.ORG`. It parses the actual sender's name and email address from the email body and maps them to known telecommunication companies.
* **Advanced Dual-Layer Filtering:** 
  * **Macro-Filters (Thread Level):** Use Star (⭐) and Follow (👀) buttons, or the global search bar, to instantly filter the left-hand thread list down to specific topics or Agenda Items of interest.
  * **Micro-Filters (Conversation Level):** Once a thread is selected, use the Company dropdown, Sender dropdown, or Text search boxes to isolate specific replies strictly within that single conversation.
* **Interactive Email Analytics:** Click the **Statistics** button to instantly generate an interactive, offline HTML Plotly dashboard visualizing Agenda Item volumes, company activity rankings, timeline histograms, and top delegate leaderboards.
* **Automated Archiving:** Safely extracts physical `.msg` files to your hard drive and dynamically builds a clean target folder hierarchy in Outlook (e.g., `Archive/SA2_175/9.1.1/`) to permanently organize your inbox.

---

### 📝 Word Document Manipulation & AI Integration
* **🤖 AI/LLM Corpus Exporter:**
  * **Smart Automation:** Automatically downloads missing TDocs from the 3GPP FTP and extracts the underlying Word documents in the background.
  * **Intelligent Parsing:** Uses a custom Regex State Machine to handle complex 3GPP formatting, including extracting Track Changes and parsing tricky "all new text" placeholder clauses (e.g., `6.4.5.X`).
  * **Mega-File Compilation:** Compiles and groups the extracted text into clean, Agenda Item-specific Markdown files tailored specifically for LLM context windows (Gemini, Claude, GPT, Ollama).
* **Global Comparison Cart:** A persistent, round-robin state dashboard that bridges multiple meeting windows. Intelligently push any Base TDoc or specific Revision into alternating slots, then launch a native Word comparison instantly.
* **Isolated Word Diff Engine:** Uses COM `DispatchEx` to spawn an invisible, isolated instance of Microsoft Word. It safely opens files as Read-Only, auto-accepts tracked changes purely in RAM, assigns proper document names for the comparison pane, and generates a visual diff without freezing your active Word sessions or locking local files.
* **LibreOffice Integration Engine:**
  * **Macro-Free & Sandboxed Conversion:** Built-in adapter leveraging headless LibreOffice with isolated user profiles (`-env:UserInstallation`) to suppress network printer hangs and bypass macro security restrictions.
  * **Installed & Portable Support:** Seamless auto-detection of system-installed LibreOffice and single-click integration for portable distributions (`LibreOfficePortable.exe`).
* **Corporate IT Bypass (Sensitivity Labels):** Automatically injects configurable Microsoft Purview Sensitivity Labels (e.g., "OFFEN") directly into COM objects to bypass blocking corporate IT popup dialogs during automated saves.
* **Intelligent DocxSplitter:** Safely slices massive 3GPP TS/TR specifications into individual Word documents based on Heading 1 or Heading 2 boundaries, perfectly preserving styles, images, and Visio objects.
* **Selective Heading Excerpt Extractor:** Load one or multiple `.docx` files, scan their heading hierarchy, select arbitrary clauses/sections, and combine the selected excerpts into a single Word document. Multi-part specifications can be naturally ordered by filename or kept in drag-and-drop order; parent selections automatically include their complete subtree without duplicating explicitly selected child headings. The workflow preserves Word formatting and embedded content belonging to the retained sections while excluding unrelated document content.
* **Background Word-to-PDF Converter:** A headless Word automation thread that silently converts generated files to PDFs or XPS without interrupting your workflow.
* **Native Visio Extractor:** Parses the raw XML (`document.xml`) of a `.docx` file, identifies embedded `OLEObject` bins, and extracts raw `.vsdx` Visio diagrams straight out of the Word document to your local disk.

---

## <a id="deprecated-features"></a>🏛️ Architectural Evolution & Deprecated Legacy Features

**3GPP Tools** represents a comprehensive, clean-sheet architectural rewrite of the legacy **3GPP Meeting Helper** desktop suite. Over years of 3GPP standards participation, several specialized desktop scripts were created. During the modernization to PyQt5, each legacy capability was audited for delegate value, performance, and stability.

While high-value workflows (such as offline AI triage, specification FTS5 trigram search, protocol bit-level matrices, and PlantUML-to-Visio conversion) were modernized and expanded, **four legacy features were deliberately deprecated and retired**. They are not recommended for porting due to severe architectural shortcomings, data-loss risks, or complete obsolescence by superior modern subsystems:

### 🔴 1. Word Process Watchdog & Brute-Force Killer (Legacy F-08: `taskkill /f /im WINWORD.EXE`)
* **Legacy Behavior:** Provided an indiscriminate "Close Word" button that executed Windows shell `os.system("taskkill /f /im WINWORD.EXE")` to eliminate hung COM processes.
* **Why Deprecated:** Executing a blind `taskkill` destroys active, unsaved drafting documents across all open Microsoft Word windows on the operating system, creating an unacceptable risk of data loss for delegates actively drafting Change Requests.
* **Modern Replacement:** Fully superseded by the built-in **`🖥️ COM Task Manager` ([`ProcessManagerDialog`](src/core/ui/process_manager_dialog.py))** and strict COM apartment lifecycle management (`pythoncom.CoUninitialize()`). The modern task manager provides non-destructive, document-by-document COM closing, user-controlled exclusion sets (`Keep Open`), and PID-isolated emergency recovery without killing unrelated interactive user sessions.
* **Reference:** Detailed technical audit in [`docs/legacy_features/F08_word_process_watchdog_killer.md`](docs/legacy_features/F08_word_process_watchdog_killer.md).

### 🔴 2. Word Meeting Contribution Summary Report (.docx Table) (Legacy F-09)
* **Legacy Behavior:** Automated Microsoft Word COM to build a static, unformatted 2,000-row Word table summarizing meeting contributions.
* **Why Deprecated:** Static Word tables are rigid, unsearchable, provide no interactive data exploration, take minutes to build cell-by-cell over COM, and frequently crashed on non-German Windows systems due to hardcoded Word style names (`'Überschrift 1'`).
* **Modern Replacement:** Fully superseded by the **Interactive HTML5 Analytics Dashboard (`Statistics_Report.html`)**, responsive Chart.js visual metrics (treatment rate, agreement rate, document composition, Louvain alliance networks), structured multi-tab Excel workbooks (`Statistics_Data.xlsx`), and the AI-driven **Post-Meeting Export Dossier Generator** (`post_meeting_export_dialog.py`) which compiles organized Markdown dossiers (`.md`) with embedded media.
* **Reference:** Detailed technical audit in [`docs/legacy_features/F09_word_meeting_contribution_report.md`](docs/legacy_features/F09_word_meeting_contribution_report.md).

### 🔴 3. Bi-directional Excel AutoFilter Synchronization (Legacy F-11)
* **Legacy Behavior:** Mirrored active filter selections from the desktop GUI into an external, open Microsoft Excel spreadsheet via COM AutoFilter arrays (`ws.Range("A1").AutoFilter(...)`).
* **Why Deprecated:** In the legacy tool, this was an awkward workaround for the limitations of Tkinter, which could not render badges, colors, or fast filtering. Automating Excel COM while a user is actively typing in a cell raises fatal COM errors (`0x800AC472 VBA_E_IGNORE`), and pushing thousands of strings over COM is slow and brittle.
* **Modern Replacement:** Obsolete. Modern 3GPP Tools features an in-app, high-performance PyQt5 filter proxy engine (`TDocsFilterProxyModel`) with instant multi-facet filtering (TSG pills, Working Group multi-select combo, date presets, ad-hoc toggles, and followed/starred Work Item chips) rendering at 60 FPS directly within the application, eliminating the need to keep external Excel workbooks open in parallel.
* **Reference:** Detailed technical audit in [`docs/legacy_features/F11_excel_bidirectional_autofilter_sync.md`](docs/legacy_features/F11_excel_bidirectional_autofilter_sync.md).

### 🔴 4. Native Excel COM Pivot Tables & Charts (Legacy F-13)
* **Legacy Behavior:** Automated Excel VBA COM (`PivotCaches().Create()`) to construct pivot tables and embedded charts directly inside `.xlsx` workbooks.
* **Why Deprecated:** Strictly required Microsoft Office desktop to be installed and running, introduced fragile VBA COM version dependencies (`xlPivotTableVersion15`), and froze when Excel encountered modal dialogs.
* **Modern Replacement:** Fully superseded by the **HTML Analytics Dashboard Compiler ([`html_dashboard_compiler.py`](src/modules/meetings/core/stats/html_dashboard_compiler.py))**. Generates rich, interactive Chart.js visualizations, KPI strips, and alliance network graphs that open instantly in any web browser without Office dependencies. For delegates who still desire Excel pivot tables, the clean data workbooks generated via `openpyxl` allow creating native pivot tables with just 2 clicks (`Insert ➔ PivotTable`) without risking script crashes.
* **Reference:** Detailed technical audit in [`docs/legacy_features/F13_native_excel_pivot_tables_charts.md`](docs/legacy_features/F13_native_excel_pivot_tables_charts.md).

> [!TIP]
> For a complete roadmap of all 19 legacy functionalities, their modern equivalents, and business value ratings, see the [**Strategic Value & ROI Analysis (VALUE_ANALYSIS.md)**](docs/legacy_features/VALUE_ANALYSIS.md) and the [**Legacy Features Catalog (README.md)**](docs/legacy_features/README.md).

---

## <a id="architecture"></a>🏗️ Architecture & Data Flow

This application strictly adheres to the **Model-View-Controller (MVC)** and **Event-Driven Architecture (EDA)** paradigms using `PyQt5`. 

1. **The UI Layer (`src/modules/*/ui/` & `src/main_window.py`):** Contains Qt Widgets, `QAbstractTableModel` implementations, modeless top-level inspection dialogs (`CallFlowDialog`, `TDocTriageDialog`), and status bar permanent action widgets. The UI never performs synchronous network I/O or blocks the main event loop.
2. **The Core Layer (`src/core/` & `src/modules/*/core/`):** Contains domain logic. All database transactions (`sqlite3` with FTS5 trigrams), REST AI communications (`core/ai/ollama_client.py`), prompt management with timestamp hot-reloading (`modules/puml2visio/core/prompt_manager.py`), FTP network scraping (`requests`), COM automation (`win32com` & `pythoncom`), headless LibreOffice conversions, and direct XML manipulation (`lxml` & `python-docx`) are isolated here.
3. **The Threading Bridge:** Independent worker tasks inherit from `QThread` (e.g., `OllamaMonitorThread`, `CallFlowGeneratorThread`, `TDocTriageWorker`, `WifiMonitorThread`, `GeneralEmailSyncThread`, `WordAgendaImporterThread`, `ContributionSearchWorker`, `TDocsDownloaderThread`, `LLMExporterThread`). Workers operate completely decoupled with their own error boundaries and communicate with the main thread strictly through thread-safe `pyqtSignals`.
4. **The Singleton Managers & Security Utilities:** Global network state (`NetworkState`), specialized AI HTTP session creation (`session.get_ai_session()`), cryptographic credential protection (`core.utils.dpapi.SecureCredentialStore`), and Comparison Cart states are managed by thread-safe singletons and dynamic JSON config loaders.

---

## <a id="prerequisites"></a>⚙️ Prerequisites

To run this application natively or build it from source, you must have the following installed on your Windows machine:

1. **Python 3.10+**
2. **Microsoft Word (Desktop App)** (Required for native COM Automation Splitter, Converter, Diff Engine, and TDoc text extraction)
3. **Microsoft Outlook (Desktop App)** (Required for the eMeeting and General Email Managers)
4. **Java Runtime Environment (JRE) 11+ or 21 (LTS)** (Required for modern PlantUML rendering and automatic updates. Recommended: [Eclipse Temurin](https://adoptium.net/temurin/releases/?version=21). Both standard `.msi` installers and zero-admin `.zip` portable archives are fully supported)
5. *(Optional but Recommended for Local AI)* **Ollama** (Install from [ollama.ai](https://ollama.ai) to enable private, offline document summaries, AI technical triage, call-flow to sequence conversion, and AI assistants).
6. *(Optional but Recommended)* **LibreOffice (Installed or Portable)** (Required for safe, macro-free conversion of legacy Word 97–2003 `.doc` files, including SA2 Chairman's Notes and older specifications. If using portable LibreOffice, link `LibreOfficePortable.exe` using the **📂 Locate Executable** button in the Word Tools tab.)
7. *(Optional)* **Microsoft Visio** (To view and edit generated `.vsdx` files)
8. *(Optional)* **Microsoft PowerPoint** (For `.pptx` to `.vsdx` conversions)

---

## <a id="installation"></a>🚀 Installation

### 1. Clone the Repository
```bash
git clone [https://github.com/telekom/3gpp-meeting-tools.git](https://github.com/telekom/3gpp-meeting-tools.git)
cd 3gpp-meeting-tools/3GPP\ Tools
```

### 2. Install Python Dependencies
```bash
pip install -r requirements.txt
```
*Note: This installs `PyQt5`, `requests`, `python-docx`, `beautifulsoup4`, `openpyxl`, `pandas`, `plotly`, `networkx`, `lxml`, `Pillow`, and `pywin32`.*

### 3. Launch the Application
```bash
# Using the Windows batch launcher (recommended):
.\src\Launch.cmd

# Or directly via Python with proper PYTHONPATH:
$env:PYTHONPATH = ".;.\src;$env:PYTHONPATH"
python -m main_tools
```
*Upon first launch, the app will automatically download the latest `plantuml.jar` from GitHub if it is not present in your assets folder.*

---

## <a id="usage"></a>📖 How to Use the GUI

### 🦙 Configuring Ollama & Local LLMs
1. **Status Bar Quick Glance:** Observe the **bottom right status bar**. The pill will display `🦙 <selected_model>` in green if Ollama is running, or `🦙 Offline` in soft red if the server is stopped.
2. **Launching Settings:** Click the `🦙` button in the status bar to open the **Ollama LLM Configuration** dialog.
3. **Connection & Testing:**
   * Enter your Ollama server URL (defaults to `http://127.0.0.1:11434`).
   * Select your **Proxy Routing Mode**:
     * **Direct (Bypass Proxy - Recommended):** Ensures loopback calls never get routed into corporate firewalls or proxy walls.
     * **Auto-Detect:** Automatically uses direct connections for local/private IPs and the application proxy for external domains.
     * **Route via App Proxy:** Enforces routing through your configured corporate proxy profile.
   * Click **🔌 Test** to verify connectivity and refresh your list of locally installed models.
4. **Selecting Default Model:** Choose your preferred model from the **Active Model** dropdown (e.g., `qwen2.5:7b`, `llama3.1:8b`) and click **Save**. The status bar pill updates immediately.

---

### 🤖 Instant AI Technical Triage of Meeting TDocs
1. Open any meeting in the **Meetings** tab to view the TDocs table.
2. **Triggering Triage:**
   * **Right-Click Context Menu:** Right-click any proposal row and choose **🤖 AI Technical Triage for <TDoc>...**.
   * **Details Card:** Double-click the `TDoc` column to open the details view and click the **🤖 AI Triage** button.
   * **Secretary Remarks Links:** Right-click any hyperlinked TDoc number in Secretary Remarks or Related TDocs and choose **🤖 AI Technical Triage (<TDoc>)**.
3. **Streaming Executive Summary:**
   * The modeless triage window opens immediately. The background worker checks local markdown caches, extracts Word document text via COM if uncached, filters out unchanged boilerplate text, and streams a structured summary from Ollama.
   * The response provides 3 core sections:
     1. **🎯 Objective & Problem Statement**
     2. **🔑 Key Technical Changes**
     3. **⚠️ Potential Impacts & Contentious Points**
4. **Saving to Personal Notes & Assigning Stance:**
   * Adjust your stance via the **My Status** dropdown (`⚪ Neutral`, `🔵 My TDoc`, `🟢 Support`, `🔴 Object`, `🟡 Monitor`).
   * Click **💾 Save to My Notes**. The summary is appended beneath any existing notes with an explicit header, timestamp, and 3GPP server download URL (`--- 🤖 AI Triage (<model> @ YYYY-MM-DD HH:MM) ---`). The table and sidecar database update immediately.
   * Click **📋 Copy** to copy the markdown text to your clipboard, automatically prepended with the TDoc title and official 3GPP server download URL for full downstream provenance. Click the **📥 3GPP Server ZIP** header link at any time to directly download the original document.

---

### 🤖 AI Summarizing TDoc Email Discussions
1. In the TDocs window, double-click the **Emails** badge or right-click a row and select **📧 View Related Emails...**.
2. **Triggering Discussion Summary:**
   * **All Visible Emails:** Click the **`🤖 AI Summarize Discussion`** button on the toolbar to summarize the entire conversation thread shown in the table.
   * **Selected Subset:** Select specific messages using `Ctrl` or `Shift` click, then click **`🤖 AI Summarize Discussion`** to summarize only that branch of the conversation.
3. **Clean Streamed Analysis:**
   * An independent modeless window opens immediately.
   * The worker thread strips all nested quotation chains, headers, and listserv disclaimers, logging the exact character reductions in `3gpp_tools.log`.
   * It sorts the debate chronologically and streams the structured technical synthesis:
     1. **🎯 Discussion Topic & Core Contention**
     2. **🏢 Company Positions & Alliances**
     3. **📝 Proposed Solutions & Revisions Mentioned**
     4. **📌 Current Consensus & Next Steps**
4. **Notes Integration:**
   * Click **💾 Save to My Notes** to append the discussion summary directly beneath your TDoc's personal notes in `user_tdocs.db`.

---

### 🤖 3GPP Research Assistant (Evidence-Grounded Question Answering)
1. Navigate to the **🤖 Assistant** tab.
2. **Connectivity Check:** Observe the Assistant header status (`Ready`) and ensure your local Ollama instance is active via the bottom status bar pill (🟢 `🦙 <model_name>`).
3. **Posing 3GPP Engineering Queries:**
   * Enter your technical question into the prompt input at the bottom (e.g., *"What information elements are contained in the Registration Request message in TS 24.501?"*, *"Explain ATSSS steering functionality and rules in TS 23.501"*, or *"Find all protocol messages related to PDU Session Establishment"*).
   * Press **Enter** or click **🚀 Ask**.
4. **Inspecting Live Research Activities:**
   * The **Research Activity** pane displays the bounded agent's tool-calling progress step-by-step in real time.
   * Inspect each semantic tool invoked (e.g., `find_specifications`, `search_spec_text`, `get_spec_clause`, `find_protocol_message`), the input arguments passed, the elapsed duration, and the execution status (`SUCCESS` / `ERROR`).
5. **Evaluating Application-Owned Evidence:**
   * The **Evidence** pane populates with concrete evidence items retrieved strictly from local SQLite databases.
   * Each item displays its assigned evidence tag (`[E1]`, `[E2]`, etc.), source domain (*Specification Catalogue*, *Spec Search*, or *Protocol Knowledge*), specification number, version, clause or message name, and a text snippet.
6. **Reviewing Grounded Answers & Citations:**
   * The model streams its synthesized technical analysis into the **Conversation** pane, citing exact supporting sources with bracketed identifiers (e.g., `[E1]`, `[E2]`).
   * The assistant explicitly distinguishes locally covered knowledge from model knowledge and refuses to hallucinate specifications or protocol fields absent from your databases.
7. **Control & Customization:**
   * Click **⏹️ Stop** at any time to abort an in-flight tool call or token streaming sequence.
   * Click **🗑️ Clear** to wipe the conversation history and start a fresh research session.
   * Custom instructions can be tuned without restarting the application by editing `config/prompts/assistant_system.txt`.

---

### 🎨 Visio Tools & Java / PlantUML Maintenance
1. **PlantUML Mini-IDE Workspace:**
   * **Writing Code:** Type or paste PlantUML code into the central editor. Syntax highlighting and line numbers update dynamically.
   * **Inserting Elements & Snippets:** Click **`🧩 Insert Element ▼`** on the top toolbar to insert messages, arrows, logic fragments, or 3GPP presets at the cursor. If you highlight multiple message lines first, block fragments (`alt`, `opt`, `loop`, `box`, `note`, `activate`) automatically indent and enclose your selection.
   * **Syntax Cheat-Sheet & Palette:** Click **`📚 Syntax Palette`** to open the searchable sidebar. Filter by keyword (e.g., `"timer"`, `"hnote"`, `"oran"`) or current diagram type to inspect syntax previews and insert elements.
   * **Live Previewing:** Click **`👁️ Split Preview`** to view your diagram side-by-side inside the app with mouse-wheel zoom and drag-pan, or toggle **`🌐 Browser Preview`** to open the live view in an external browser.
   * **Finding & Replacing:** Press **`Ctrl+F`** or click **`🔍 Find`** to toggle the inline search and replace bar.
   * **Commenting & Indenting:** Press **`Ctrl+/`** to toggle single-quote comments (`' `) across selected lines. Use **`Tab`** and **`Shift+Tab`** to indent or unindent multi-line blocks.
2. **Java Runtime & PlantUML Maintenance Center:**
   * Click the **☕ Java** button on the bottom Console Panel at any time.
   * The dialog displays the active Java version, architecture, vendor, and PlantUML JAR version.
   * **Using Portable Java Without Admin Rights:**
     1. Download the `.zip` archive of Eclipse Temurin 21 (x64) from adoptium.net.
     2. Extract it anywhere (e.g., `C:\Users\<user>\tools\jdk-21` or directly as a `jre` folder in your `3GPP Tools` root).
     3. Select **Custom Java Runtime**, click **📁 Browse...**, and choose the extracted folder.
     4. Click **⚡ Test** and **Save & Apply**.
   * Click **🔄 Check Online & Update JAR** inside the dialog to immediately update PlantUML to the latest release.
3. **Generating Sequence Diagrams from Call Flow Text (AI):**
   * On the editor toolbar, click **🤖 Generate from Call Flow...**.
   * The generator dialog opens as an independent, modeless window. You can continue working in other tabs or searching specifications while it remains open.
   * Paste 3GPP procedural steps (e.g., from TS 23.502 or TS 38.300) into the left text box, or pick a normative reference from **📚 Sample Procedures ▼**.
   * Toggle 3GPP generation style preferences as desired: **Box Groupings**, **Message IEs**, **State Transitions**, **Alt / Failure Branches**, and **Autonumber**.
   * Select your installed Ollama model and target diagram type (default: *Sequence*).
   * Click **🚀 Generate Diagram**. The model will stream PlantUML syntax token-by-token directly into the preview area with syntax highlighting.
   * Once finished, the tool automatically validates the diagram. If errors occur, an alert banner pinpoints the problem line and offers **🪄 Auto-Fix Error**.
   * Switch to the **👁️ Diagram Preview** tab to review the rendered SVG diagram directly in the dialog without touching your editor.
   * Use the **💬 Refine:** prompt bar at the bottom to have a multi-turn conversation with the model to adjust, re-group, or expand the diagram.
   * Click **📥 Replace Editor** to overwrite the editor contents, **➕ Append to Editor** to insert the code into your existing diagram, or **📋 Copy Code** to copy to clipboard.
   * *(Prompt Engineering)* You can customize generation prompts without restarting the app by editing `config/prompts/callflow_system.txt` or `config/prompts/callflow_user.txt` in any text editor. The tool reloads updated prompt files automatically upon clicking generate.
4. **Exporting Diagrams:** Click **Export Diagram ▼** and select **To Visio (.vsdx)** to generate a native Visio file, or use other options like PowerPoint, SVG, or ASCII.
5. **Batch Process & PowerPoint Conversion:** Navigate to the **📂 Visio Tools** tab and drag-and-drop `.puml`, `.txt`, or `.pptx` (PowerPoint) files into the drop zone. The system will detect the file type and process it into an editable Visio file in the background.

---

### 🔎 3GPP Specification Full-Text & Evolution Search
1. Navigate to the **🔎 Spec Search** tab.
2. **Importing & Managing Specifications (Zero-Waste Layout):**
   * Action menus (**`📥 Import ▾`** and **`⚙️ ▾`**) are integrated directly into the header of the **Indexed Specs** card, eliminating top-bar clutter and maximizing vertical reading space.
   * Ingest specifications from the specifications catalogue (**`📥 Import from Specs DB...`**) or local single/multi-part split Word documents (**`📁 Import Local .docx...`**).
   * Monitor indexed specifications and releases at a glance via the live capacity chip in the card header.
   * Manage indexed records via **`⚙️ ▾`** (`Clear Specific Version...`, `⚠️ Wipe Entire Search Database...`).
   * **Specification Domain Names & Smart Filter:** Tree items display concise specification titles (e.g. `TS 23.256: Support of Uncrewed Aerial Systems (UAS)... (1)`) with full formal titles in tooltips. Instant tree filtering (`Filter specs...`) searches both specification numbers and title keywords (e.g., typing `"UAS"`, `"architecture"`, `"drone"`, or `"NAS"` matches immediately even without remembering the 3GPP document number).
3. **Executing Substring Searches:**
   * Type any exact phrase or keyword into the search bar inside the unified **Search Console Card** (e.g., `"slice replacement"`, `"ATSSS"`, `"emergency"`). Search queries with 3 or more characters automatically execute across the FTS5 trigram index.
   * Optionally enter a clause number in the **Clause** field (e.g., `5.2`, `4.3.2`) to focus on specific sections.
   * Review instant executive KPI result pills above the matrix: `Results: [ 🔍 'query' ] [ N Clauses ] [ M Spec(s) ]`.
4. **Date Cutoff & "First Added" Text Analysis:**
   * Review the **Release Evolution Matrix** displayed in per-specification tabs with clear domain clues (e.g., `TS 23.256 · UAS (5)`, `TS 23.501 · 5GS Architecture (32)`) and full document titles in tab tooltips.
   * Status cells are styled as sleek, centered, rounded pill badges (**`🟢 Added`**, **`⚡ Post-Cutoff`**, **`✓ Present`**, or muted **`—`**). Selecting any row highlights it in soft sky-blue (`#EBF3FC`) without masking or inverting badge colors.
   * Toggle **🎯 Date Cutoff** and select a cutoff date. Text introduced after that date will be highlighted with ⚡ **`⚡ Post-Cutoff`**.
   * Check **Show Only Post-Cutoff Additions** to filter out older prior art and show only clauses containing post-cutoff date modifications.
5. **Inspecting Matching Clause Content & Flexible Reading Layout:**
   * Click any cell in the matrix to load the clause into the **Clause Content Inspector**.
   * **Flexible Layout & Maximize:** Click **`◫ Right Pane`** to switch the inspector to a tall vertical side-by-side reading pane (layout preference is saved automatically across sessions), or click **`⤢`** to maximize the reading pane to full height.
   * The **💡 Key Match Excerpt** callout at the top highlights the matching paragraph with surrounding sentence context without duplicating clause headings.
   * Use **`[ ◀ ]`** and **`[ ▶ ]`** to cycle between match occurrences in long clauses.
   * Click **`📋 Copy ▾`** to copy the clause in your preferred format:
     * **📋 Copy for Email / Outlook (Rich Text):** Dual-format clipboard copy setting styled rich HTML with a formal 3GPP citation header banner and formatted paragraphs alongside clean plain text fallback.
     * **📝 Copy as Markdown (Quote Block):** Formatted GitHub-style callout quote block (`> [!NOTE]`) citing specification release and clause metadata.
     * **📄 Copy Plain Text:** Full clause text with citation.
     * **🔖 Copy Citation Only:** Pure citation string (`3GPP TS 23.256 v19.5.0, Clause 5.15 (Title)`).
   * Click **`🟢 First Added in v...`** to jump immediately to the matrix release where this text was introduced, or click **`[ 🔍 Diff Prev ]`** to view the in-app side-by-side textual diff.
6. **Extracting Individual Clauses to Word (.docx) & Markdown (.md):**
   * Click **`✂️ Extract ▾`** on the inspector toolbar (or right-click any matrix cell):
     * **✂️ Extract Clause in Word (.docx):** Non-blocking background worker locates the Word document (auto-extracting or downloading if needed), scans heading hierarchies with `WordHeadingScanner`, routes multi-part split documents (`_s00_s04`, `_s05_s08`), and compiles a clean, standalone Word document containing that clause with native styles, tables, and figures, immediately opening it in Word.
     * **💾 Save Clause as Word (.docx)...:** Prompts for a custom destination folder/filename before generating the standalone Word excerpt.
     * **📘 Export Clause to Markdown (.md)...:** Converts the clause into high-fidelity Markdown powered by the application's Markdown conversion engine. All figures, diagrams, and pictures are automatically extracted into a version-isolated media directory (`{filename}_media/`), EMF/WMF vector drawings are converted to crisp PNG, embedded Visio source files are hyperlinked, and native Markdown tables are generated.
7. **Opening Full Specifications Directly & Clipboard Assistance:**
   * Click **`📖 Full Doc ▾`** in the Inspector (or double-click any matrix cell) to launch the full specification release in Microsoft Word, PDF, or open its local folder.
   * For multi-part documents (`_s00_s04.docx`, etc.), the resolver automatically opens the specific sub-part containing that clause.
   * The clause identifier (`Clause X.Y.Z`) is automatically copied to your clipboard so you can jump to it in Word immediately via <kbd>Ctrl+F</kbd> + <kbd>Ctrl+V</kbd>.
8. **Clause-Level Comparison & Word Track Changes:**
   * Right-click any matrix cell and select **`🔍 Compare Clause (Diff & AI Analysis)...`**, or click **`📝 Open in Word (Track Changes)`** from inside the comparison dialog.
   * The background engine extracts only the selected clause from both releases, applies corporate sensitivity labels, and drives Microsoft Word COM to generate an executive tracked-changes comparison document. In addition, you can inspect the color-coded textual diff directly in the app without launching Word.
9. **Cross-Tab Deep Linking:**
   * Click **`🔍 Specs Tab`** in the Inspector to jump directly to that specification and release in the **📚 Specifications** tab.
   * In the **📚 Specifications** tab, right-click any specification or release to select **`🔎 Track Evolution in Spec Search`** or **`📥 Index vX.Y.Z into Spec Search DB`** (or batch-index multiple selected specifications via **`📥 Index selected (N) into Spec Search DB`**).

---

### 🔬 3GPP Protocols Evolution Matrix (NAS, ASN.1 & PFCP)
1. Navigate to the **🔬 Protocols** tab.
2. **Action Bar & Database Stats Ribbon:**
   * Review local database capacity at a glance via the top metrics ribbon: `📦 X Specs`, `🏷️ Y Releases`, `📨 Z Messages · W IEs`.
   * Click **📥 Import from Specs DB** to select specification releases across **TS 38.331 (NR RRC)**, **TS 36.331 (LTE RRC)**, **TS 38.413 (NGAP)**, **TS 29.244 (PFCP)**, **TS 24.501 (5GS NAS)**, or **TS 24.301 (EPS NAS)**. The dialog includes a descriptive summary banner explaining the domain and role of each specification. Missing versions download and convert automatically from the 3GPP FTP archive.
   * Click **📁 Import Local .docx** to ingest local single-file or multi-part split specification documents directly.
   * Administrative and maintenance operations (`Clear Selected Version`, `Wipe Database`) are safely accessible via the **⚙️ Manage ▾** dropdown menu.
3. **Selecting Releases with Built-In Specification Clues:**
   * Specifications in the tree are annotated with their responsible Working Group badges (`[CT1]`, `[RAN2]`, `[CT4]`) and friendly acronyms (`5GS NAS`, `PFCP`, `NR RRC`), while versions clearly denote their 3GPP generation (`Rel-20 · v20.0.0`, `Rel-19 · v19.7.0`).
   * Use quick-toggle buttons in the tree header (**All** | **Latest** | **Clear**) to instantly compare only the latest releases across all specifications without manually managing dozens of checkboxes.
   * Multi-line tooltips describe the formal title, responsible Working Group, and architectural role for every specification.
4. **Browsing & Searching Protocol Messages / IEs:**
   * **Browse Mode Toggle (`[📨 Messages | 🧩 All IEs / Types]`):**
     * **`📨 Messages` Mode (Default):** Each item in the library is formatted with its bold name, source specification/WG chip, and clause pill (e.g. `AUTHENTICATION REQUEST [CT1 · 24.501] Cl. 8.2.4.1.1`).
     * **`🧩 All IEs / Types` Mode:** Browse all parsed Clause 9 / Annex D / ASN.1 / PFCP Information Elements directly.
   * **Search by Acronym, Spec or Name:** Type acronyms (`"PFCP"`, `"NAS"`, `"RRC"`, `"NGAP"`), spec numbers (`"24.501"`, `"38.331"`), or message names into the search bar. The library live-updates with an active match counter (`Showing X of Y items`).
   * Click the **`📖 Desc`** button to toggle extended description search, matching keywords located deep inside Clause 8/9 IE definitions and ASN.1 field description tables.
   * For PFCP messages, use the **Interface Selector Dropdown** (`All Interfaces`, `N4`, `N4mb`, `Sxa`, `Sxb`, `Sxc`) positioned above the matrix table to filter parameters by target reference point.
5. **Evolution Matrix & 1-Click TSV Export:**
   * When no message is selected, an illustrated empty state (`🔀`) guides delegates.
   * When populated, the matrix header displays the source specification clue and an inline diff legend: `🟢 Added`, `🟡 Modified`, `🔴 Removed`.
   * Column headers are tagged with 3GPP release numbers (e.g. `Rel-20\nv20.0.0`).
   * Click **📋 Copy TSV** to export the entire multi-release comparison matrix directly to the system clipboard for immediate pasting into Microsoft Excel, Word, or emails.
6. **Inspecting Structure, Syntax & Reverse Lookup:**
   * Click any row in the matrix to inspect its bit-level octet layout diagram or ASN.1 syntax block and field description callout in the **Inspector** card below.
   * Click **📋 Copy Definition** to copy the formatted ASN.1 syntax or parameter definition.
   * Click the **Used in: N msgs ▾** badge in the inspector header (or right-click any row in the matrix) to find all other messages referencing that parameter across active releases.

---

### 🗄️ Database Maintenance, Inspection & Compaction
1. Click the **🗄️ Database** button located in the bottom Console Panel next to Task Manager and Proxy.
2. The dialog displays all SQLite database files (`3gpp_data.db`, `3gpp_protocol_data.db`, `3gpp_spec_search.db`), their current on-disk sizes, and Write-Ahead Log (`-wal`) statuses.
3. **Database Compaction:** Click **Compact** on an individual database or **🧹 Compact All Databases** to flush WAL logs, execute SQLite `VACUUM`, optimize indices, and instantly reclaim free disk space.
4. **Read-Only Schema & Data Inspection:** Click **Inspect** on any database entry to open `DatabaseInspectorDialog`. Browse tables, check exact schema definitions and column types, inspect diagnostic PRAGMAs, and review paginated rows with real-time text searching without locking the database.

---

### 🖥️ Managing Active COM Processes & Office Documents (Task Manager)
1. Click the **🖥️ Task Manager** button in the bottom Console Panel.
2. **Process Inspection:** The dialog displays all active instances of Microsoft Word (`WINWORD.EXE`), Microsoft PowerPoint (`POWERPNT.EXE`), and Microsoft Visio (`VISIO.EXE`) along with their Windows Process IDs (PIDs).
3. **Document-Level Visibility & Safe Closing:**
   * For Word and PowerPoint, the tree view lists each currently open document title.
   * Check or uncheck documents to control scope: uncheck any files you wish to keep open during batch operations.
   * Click **Close Selected Documents** to cleanly save and close target documents via COM automation without killing the underlying application.
4. **Emergency Orphan Process Termination:**
   * If a background conversion or comparison hangs due to a modal prompt or blocked file lock, select the hanging process and click **Kill Process**.
   * Click **Kill All Office Processes** to immediately clean up all orphaned background Word, PowerPoint, and Visio processes in one click.

---

### 📊 3GPP Meetings & Specifications
1. Navigate to the **Meetings** tab.
2. Click **Sync All Meetings** to trigger the 3-Phase scraper. You can also use **Open Last Meeting** to instantly resume your previous working group session.
3. Use the **Global TDoc Search** input to instantly find a specific document. Type a valid TDoc number (e.g., `S2-2605740`), and press **Enter** (or click **📄 Doc**) to fetch and open it immediately, or click **🗓️ Mtg** to launch its parent meeting table.
4. Right-click any meeting to access its FTP folders, view its info, open its cached **TDocs List**, or select **💻 Open Terminal in Meeting Cache** to launch a command terminal directly in that meeting's local files directory.
5. In the TDocs Window, use the **Search** bar or dropdown filters to find specific documents. The Agenda Item filter automatically presents structured labels (e.g., `AI 20.6: Acronym - Subtopic`) derived from SA2 and RAN2 `agenda.csv` files, and hovering over any row in the Agenda Item column displays full title tooltips. Double-click any cell to open the Notes editor and assign a color-coded status to a document. (You can also select **📁 Resources ▾ $\rightarrow$ 💻 Open Terminal in Meeting Cache** to open a terminal in the active meeting folder).
6. **Importing SA2 Session Documents & Chairman's Notes:**
   * **Drag & Drop:** Drag any `.docx`, `.doc`, or `.htm` session document anywhere onto the TDocs window. A visual frosted drop overlay will highlight the window.
   * **Menu Import:** Alternatively, click the **🔄 Refresh** menu and select **📝 Import Word Document (.docx / .doc)...**.
   * The file is automatically copied to `{meeting_dir}/Agenda/`, converted in the background via LibreOffice (if `.doc`), parsed, and merged into the table without freezing the UI.
7. Click the Action column to automatically download, unzip, and open documents, or use the **⚖️ Add to Comparison Cart** submenu to select base versions or revisions for diffing.
8. Under the Specifications tab, use **🎯 Quick Fetch** to surgically inject single specifications or series into the database without a full sync.
9. **Extracting Clauses (Word Excerpt):**
   * In the **Specifications** tab, click the **Documents** dropdown for any specification release and choose **`✂️ Extract Clauses (Excerpt)...`**.
   * The modeless **Extract Clauses** dialog opens immediately. Missing specification ZIPs download and unpack in the background, multi-part documents (`_s00_s04`, etc.) are sequenced, legacy `.doc` files are converted to `.docx`, and change-mark files (`-rm`) are filtered out.
   * Filter clauses by keyword or number, select desired clauses in the hierarchy (parent check states automatically cascade to all subclauses), and click **`✂️ Extract Selected Clauses...`** to generate a clean, formatted Word document with source typography and corporate sensitivity labels.
10. **Exporting Specifications to Markdown:**
    * In the **Specifications** tab, click the **Documents** dropdown for any release to export directly into Markdown:
      * **`📘 Open / Convert to Markdown (.md)`**: Generates a unified, continuous document (`<stem>.md`) with a specification metadata table and an isolated `<stem>_media/` image folder.
      * **`📑 Open / Convert to Markdown Clauses`**: Generates a chapter-split dossier (`<stem>_clauses/`) with individual clause files (`01_Scope.md`, `04_Architecture.md`, etc.), a master `index.md` table of contents with links, and a private `media/` folder.
    * Multi-part specifications (`_s00_s04`, `_s05_s08`, etc.) are automatically detected, chronologically ordered, and stitched.
    * Version isolation guarantees images and diagrams across different releases (e.g. Rel-17 vs. Rel-18) never overwrite each other.
11. **Generating a Post-Meeting Report & Markdown Dossier:**
    * In the TDocs Window, click **`📤 Export ▾`** and select **`📑 Export Post-Meeting Report (.md)...`**.
    * Select your export scope (e.g., **`🎯 Followed Work Items`** to focus on your delegation's assignments, or **`🌐 Whole Meeting`** for a complete report).
    * Configure included categories (Incoming LS, agreed pCRs, agreed CRs, agreed Outgoing LS, and default unassigned/general LS in preservation) and click **Export Dossier**.
    * The tool compiles all relevant documents into structured `Rel-XX/<WI>/{LS_IN, pCR, CR, LS_OUT}/` folders with collision-free media, cross-WI notices, and master/local `index.md` summaries, automatically revealing the generated export folder in Windows Explorer.
12. **Adding Tags, Notes & Searching Custom Annotations:**
    * **Editing Annotations:** Double-click any specification row, click `⋮` $\to$ **`🏷️ Edit Tags & Notes...`**, or right-click any row to open the editor.
    * **Adding Tags:** Type custom tags (e.g., `data framework`, `analytics`, `charging`) and press Enter or click suggestions.
    * **Adding Notes:** Write multi-line technical context or delegation stance in the notes box.
    * **Normal Search Integration:** Type your custom tag or note terms into the primary search bar (e.g., typing `data framework`); all tagged specifications will match immediately.
    * **Filtering by Tag:** Select any tag from the **`🏷️ Tag:`** dropdown on the filter bar to isolate only specifications matching that domain tag.
    * **Batch Tagging:** Select multiple specifications in the table, right-click $\to$ **`🏷️ Add Tag to selected (N)...`** to tag all of them simultaneously.

---

### 📊 Cross-Meeting Company Contribution Audit & Reporting
1. Navigate to the **Meetings** tab and click **`📊 Reports ▾`** $\to$ **`📊 Contributions Report...`**.
2. The audit dialog opens as a modeless, non-blocking window, allowing you to use other app tabs and meeting windows concurrently.
3. **Configuring Search Filters:**
   * **Working Groups:** Select one or multiple Working Groups via the checkable dropdown (e.g., `SA2`, `SA1`, `RAN2`).
   * **Date Range:** Define the meeting start and end date boundaries.
   * **Work Items:** Type a WI acronym or code (e.g., `FS_6G_ARC`, `AmbientIoT_Ph2-ARC`) into the token box and press Enter or click **Add**. Autocompletion guides your input using the `WorkItemsDatabase`. Multiple tokens can be added as removable chips.
   * **Target Companies:** Filter and check companies from the list populated from `CompanySanitizer`.
   * **Author Match Rule:** Select **Any co-author** to capture joint contributions, or **Primary source only** to match documents where the company is the first listed submitter.
   * **Bypass Cache:** Check **Bypass local cache (re-fetch from 3GPP)** to force downloading fresh TDoc lists from the 3GPP portal.
4. Click **🔍 Search Contributions**. The background engine searches matching meetings concurrently (max 3 worker threads), leveraging O(1) sidecar JSON caches to display results without freezing your PC.
5. **Reviewing Results & In-App KPIs:**
   * Review the real-time KPI strip: **Total Contributions** (with WG coverage count), **Agreed / Approved** (total count and percentage of total submissions), **Collaboration** (% joint vs. solo), **Top Co-signer**, and **Top Work Item**.
   * Inspect the neutral, uncolored results table containing `WG`, `Meeting`, `TDoc`, `Title`, `Source`, `Matched Company`, `Type`, `For`, `Agenda Item`, `TDoc Status`, `Related WIs`, and `Abstract`.
6. **Row Actions & Direct Navigation:**
   * **Open Meeting Table:** Double-click the **WG** or **Meeting** cell (or right-click and select **🗓️ Open Meeting Table**) to launch that meeting's full table in `TDocsWindow`; double-clicking `Abstract` opens the clean pop-up reader; double-clicking `TDoc` (or any other cell) downloads (if missing), extracts, and opens the actual document in Word via `TDocActionThread`.
   * **Open Docs Folder:** Right-click and choose **📂 Open Docs Folder** to open the meeting's public FTP documents directory in your browser.
   * **View on 3GU Portal:** Right-click and select **🌐 View on 3GU Portal** to open the official portal page for that TDoc.
   * **Copy URL:** Right-click and choose **🔗 Copy URL** to copy the direct ZIP download link to your clipboard.
   * **Read Abstract:** Double-click the **Abstract** cell (or right-click and select **📝 View Abstract...**) to open the popup text reader.
7. **Exporting Data & Analytics:**
   * Click **📥 Export to Excel...** to generate a styled Deutsche Telekom magenta `.xlsx` spreadsheet with active download hyperlinks pointing directly to each TDoc's specific meeting Docs folder.
   * Click **📊 Statistics Dashboard...** to compile and open an interactive offline HTML Plotly report displaying monthly contribution trends, overall decision outcomes, top co-signers segmented by Working Group, Working Group activity breakdowns, and Work Item allocations.
8. **Filtering Meeting TDocs by Tracked Work Items:**
   * In any open `TDocsWindow`, use the quick-toggle chips on the filter bar:
     * Click **`🎯 Followed WIs (N)`** to isolate contributions related to Work Items assigned to your delegation responsibility.
     * Click **`⭐ Starred WIs (M)`** to isolate contributions related to your personal watchlist.
   * **Multi-Facet Combinations:** Combine tracked WI chips with status, company, or type dropdowns (e.g. click `🎯 Followed WIs` and set `TDoc Status` to `agreed` to see only approved papers under your followed topics).
   * **Dual Toggles:** When both chips are active, the viewer displays contributions belonging to either category (Followed $\cup$ Starred).
   * **Real-World Acronym Matching:** Automatically handles `FS_` study prefixes, strips Working Group suffixes (`-ARC`, `-CT`, etc.), and resolves discussion papers lacking explicit metadata via Agenda Item descriptions.
   * **Instant Cross-Tab Sync:** Tracking selections made in the Work Items tab are dynamically updated as soon as the meeting window is focused.

---

### 📧 Tracking Related Emails for TDocs (Universal Meeting Support)
1. **Configuring Folders & Tag Colors:**
   * In any open TDocs window, click the **📧 Emails ▾** header menu and select **⚙️ Configure Outlook Folders...**.
   * Click **➕ Add Folder via Outlook...** to browse and map your Working Group distribution list folders (e.g., `SA2_WG`, `SA2_DISC`, `RAN2_List`).
   * Enter a short Tag (e.g., `WG`, `Disc`, `Offline`) and click the color button to assign a distinct visual badge color using the color picker. Configurations are saved globally per Working Group.
2. **Syncing Outlook Emails:**
   * Click **📧 Emails ▾ $\rightarrow$ 🔄 Sync Related Emails...**.
   * Confirm the date range (defaults to meeting start/end dates $\pm 3$ days buffer) and click **🚀 Start Sync**.
   * The background engine indexes all mentions of TDocs in both Subject lines and Message bodies without downloading physical `.msg` files.
3. **Inspecting TDoc Conversation Threads:**
   * Review the **Emails** column in the main TDocs table. Cells display total family counts and blue unread badges (e.g., `✉️ 4 (🔵 2)`).
   * Double-click any cell in the **Emails** column (or right-click a row and select **📧 View Related Emails...**) to open the modeless inspection dialog.
   * **Family Breadcrumbs:** The top card displays the complete document revision lineage (e.g., `S2-2601000 ➔ S2-2601234 ➔ S2-2601555`).
   * **Quotation Filter:** Uncheck **Include Quoted Matches** to filter out reply chains that only mentioned the TDoc in historical quoted text.
4. **Navigating & Reading Emails:**
   * Drag the interactive **vertical splitter** to expand the reading pane.
   * Click **⧉ Pop Out View** to detach the reading pane into an independent viewer window (`StandaloneEmailReaderWindow`), ideal for laptop screens or secondary monitors.
   * **Interactive TDoc Links:** Every 3GPP document number cited in the Subject line, Match Excerpt banner, or Body is rendered as an interactive link. Click any link (e.g., `🔗 S2-2608457`) to open that document's related emails immediately.
   * Click **🚀 Open in Outlook** to view the original message live in native Microsoft Outlook.
5. **Managing Read & Ignored Statuses:**
   * Selecting an email automatically marks it as read.
   * Select multiple rows using `Ctrl` or `Shift` to batch **Mark Read**, **Mark Unread**, **Ignore**, or **Delete**.
   * **`🚫 Ignore`:** Suppresses high-volume distribution list announcements or rapporteur compilation emails from badge counts across all referenced TDocs without deleting them from the database. Ignored flags are preserved across re-syncs. Toggle **`Show Ignored`** to review or un-ignore them.
   * Right-click any row in the main TDocs table to mark all emails for that document family as read or unread in one click.
   * To reset generic meeting email records, click **📧 Emails ▾ $\rightarrow$ 🗑️ Wipe Generic Emails Database...**.

---

### 📋 3GPP Work Items (WIs)
1. Navigate to the **3GPP Work Items** tab.
2. Click the **🔄 Sync 3GPP WIs** button (hover over it for tooltip details) to trigger the parallel multi-threaded scraper across all 19 Technical Specification Groups and Working Groups.
3. Monitor the real-time progress bar and status messages as records are fetched and bulk upserted into the shared database.
4. Use the **Local Search** bar and multi-select **Checkable Dropdowns** to debounce-filter the table by Acronym, Name, Code, Release, or Working Group. Your selected filters are automatically saved and restored between application sessions.
5. **Track Work Items (⭐ Star & 🎯 Follow):**
   * Single-click the **Left Star (`☆`/`⭐`)** in the **Track** column to bookmark topics of personal technical interest.
   * Single-click the **Right Target (`🎯`)** to mark Work Items you are actively assigned to follow for your delegation.
   * Toggle the **`🎯 Following`** and **`⭐ Starred`** chips above the table to isolate your assigned responsibilities and watchlists.
   * Select multiple rows and right-click to batch star/unstar or follow/unfollow Work Items.
   * Click **`🎯 Update Tracked`** in the header to run a fast, targeted refresh of only your followed and starred items prior to meeting sessions.
6. **Add Individual Work Items (`➕ Add WI`):** Click **`➕ Add WI`** in the header toolbar (or right-click anywhere on the table) to open the **Add / Fetch Work Item** dialog. Paste a Work Item ID (e.g., `1121123`, `#1121123`) or a direct 3GPP Portal URL and click **🔍 Fetch Details**. The tool will scrape the title, acronym, release, dates, latest WID, mapped groups, and secretary remarks from the portal, allowing you to review or edit fields, optionally set ⭐ Star or 🎯 Follow tracking, and save directly into the local database without running a full sync.
7. **Interactive Columns:** Click any blue **Latest WID** hyperlink to download the document via the global search engine (or fall back to the 3GPP Web Portal). Click the interactive **💬 Remarks** button to view a chronologically sorted history of secretary remarks for that specific work item. Double-click any row to inspect complete details, impacted specs, and toggle tracking status inside `WorkItemInfoDialog`.

---

### 📧 eMeeting Email Manager (SA2 Electronic Sessions)
1. Open a specific electronic meeting from the database, click the **📧 Emails ▾** menu, and choose **📊 Open eMeeting Email Manager (Dashboard)**.
2. Click **⚙️ Folders** to browse your Outlook directory and safely map your Source (Inbox) and Target (Archive) folders.
3. Click **🔄 Sync Source** to download and index all eMeeting emails.
4. Select a TDoc thread from the **Left Panel** to view its chronological email history in the **Right Panel**.
5. Use the **⭐ Star** and **👀 Follow** buttons in the reading pane to track specific documents or entire topics. Use the left-side filters to isolate these threads, and the right-side dropdowns to filter by Company or Sender strictly within a thread.
6. Select rows and click **➡️ Move Selected** (or **⏭️ Move All**) to organize emails into dynamic Agenda Item subfolders inside your Outlook archive.
7. Click **📊 Statistics** to generate and open an interactive visual analytics dashboard of the meeting's email traffic.

---

### 📝 Slicing, Extracting & Comparing Word Documents
1. **Compare documents:** In the **Comparison Cart** pill at the bottom-right of the Meetings Tab (which automatically appears as soon as items are staged), sequentially select documents. The round-robin queue automatically populates Slot A and Slot B with local files or fetched 3GPP revisions. Click **Compare in Word** to generate a native Word redline comparison in the background.
2. **Split a large specification:** In **Word Tools**, select **Subtractive Slicing (Split by Clause)**, choose the target clause prefix and heading depth, and drop a `.docx` file. The splitter creates individual chapter/section documents while retaining the source document's formatting and embedded content.
3. **Build a selective excerpt from headings:** Select **Extract Sections by Heading** and drop one or more `.docx` files. This is intended both for single large specifications and for specifications split across several Word files.
   * Choose **Filename order** (default) to naturally sort multi-part specification files by filename, or **Drop order** when the order in which the files were added should be preserved.
   * The **Source Documents** list shows every file currently included. Select one or more entries and use **Remove Selected** (or the `Delete` key) to remove them; **Clear All** resets the complete input set. Changing the source set or ordering refreshes the heading structure.
   * Scan/browse the combined hierarchical heading tree, use the heading filter plus **Expand All / Collapse All** for large documents, and check the sections to retain. Selecting a parent heading includes its complete subtree; selecting both a parent and one of its children does not duplicate the child content.
   * Non-contiguous selections are combined directly into **one output `.docx`** in document order. For example, selecting `4.2` and `7.3` produces an excerpt containing those two complete sections one after another. The original Table of Contents is not added automatically.
   * Extraction operates on copies of the source documents and uses native Word processing where required for multi-file assembly, with the goal of retaining styles, tables, figures, equations, and embedded/OLE content used by the selected sections without carrying unrelated excerpt content into the result. Original source files are not modified.

---

### ⚙️ Configuring Corporate Proxies & Networking
If you are behind a corporate firewall:
1. **Network Status Indicator:** Check the **bottom right status bar** to observe active network detection (Public Internet vs. 3GPP Local Network).
2. **Configuring Secure Proxy Profiles:**
   * Click the **📡 Proxy** button in the Console Panel.
   * Check **Activate this proxy profile** to route outgoing HTTP/HTTPS traffic through the proxy.
   * Enter your HTTP and HTTPS proxy addresses (e.g., `proxy.company.com:8080`). Use the **Use the same proxy address for HTTPS** checkbox to sync them automatically.
   * For corporate networks requiring authentication (HTTP 407), enter your username (`Domain\Username` or standard `Username`) and password.
   * Ensure **Save proxy profile securely (Windows DPAPI)** is checked. All parameters—including hosts, ports, username, and password—are encrypted via `core.utils.dpapi.SecureCredentialStore` into an unreadable binary blob before saving to `network_config.json`. No plaintext proxy topology or credentials ever touch the disk.
   * Click **⚡ Test Connection** to verify end-to-end connectivity to `www.3gpp.org` via a non-blocking background thread (`ProxyTestWorker`). Sensitive credentials are automatically redacted from error popups and logs.
   * Click **Save & Apply** to apply the configuration immediately to the active `NetworkSession`.
3. **Deactivating Corporate Proxy (Direct Mode):**
   * When disconnecting from the corporate network, click **📡 Proxy** and select **Deactivate (Direct)**. The session immediately routes directly to the public internet while retaining your encrypted profile for easy reactivation.
4. **Scraper Humanness & Request Delays:**
   * Click the **⚙️ Network** button in the Console Panel.
   * Adjust **Min Delay** and **Max Delay** to throttle network requests to mimic human browsing behavior, or set them to `0.0` for maximum download speed.
   * Toggle **Rotate Modern User-Agents** to cycle browser headers across requests.

---

## <a id="troubleshooting"></a>🛠️ Known Quirks / Troubleshooting

* **Running Without Microsoft Visio:**
  * Microsoft Visio is completely optional. Diagramming in PlantUML, live SVG previews, specification full-text searches, and protocol matrix features work normally without Visio.
  * When Visio is not installed, the tool sets the status bar to `🟡 System Ready (Visio Unavailable)` and disables the Visio conversion batch tab while allowing automated PlantUML updates to download and function properly.
* **Portable Java (Zero-Admin Workaround):**
  * If corporate IT prevents running `.msi` installers, download the **`.zip`** archive of OpenJDK / Temurin 21.
  * Extract the zip into any user directory or directly into `<project_root>/jre`.
  * Open the **☕ Java** dialog from the bottom bar, select **Custom Java Runtime**, and browse to that folder.
* **Java 8 Legacy Pinned Profile & Upgrading:**
  * If the **☕ Java** dialog reports `Java 8` and `Profile Mode: Legacy`, PlantUML updates are pinned to `v1.2023.13` because upstream PlantUML dropped Java 8 compatibility after 2023.
  * Upgrading to Java 11 or 21 (LTS) unlocks automatic GitHub updates to the latest modern PlantUML engine.
* **Corporate IT "Aktion blockiert" on Drag & Drop:**
  * If Windows Defender Attack Surface Reduction (ASR) blocks dragging downloaded `.doc` files directly from your `Downloads` folder, either:
    1. Use the **🔄 Refresh $\rightarrow$ 📝 Import Word Document...** file picker menu.
    2. Unblock the file via Right Click $\rightarrow$ Properties $\rightarrow$ **Zulassen (Unblock)**.
* **Legacy Word 97–2003 Macro Permissions:**
  * Legacy `.doc` files containing VBA macros (like SA2 Chairman's Notes) are blocked by Word COM security settings. Ensure LibreOffice is installed or point the app to portable LibreOffice (`LibreOfficePortable.exe`) in the Word tab to enable automated, macro-free conversion.
* **Sensitivity Label Dialogs (Microsoft Purview / Azure Information Protection):**
  * If automated Word conversions or comparisons trigger corporate classification popups, configure your default sensitivity label string (e.g., `OFFEN` or `INTERNAL`) in `word_config.json` to allow silent headless saves.
* **Ollama Connection Refused or Showing Offline:**
  * Verify that the local daemon is active by running `ollama list` in PowerShell or Windows Terminal.
  * If running Ollama on another machine on your local network (e.g., a GPU workstation), ensure `OLLAMA_HOST=0.0.0.0` is set in the daemon environment and configure the target IP in the **🦙 Ollama Configuration** dialog.
* **Tuning AI Diagram, Assistant & Triage Prompts:**
  * All local AI prompts are decoupled into external text files in `config/prompts/` and hot-reloaded automatically via file modification timestamps (`mtime`) on the next click without restarting the app:
    * **3GPP Research Assistant:** `config/prompts/assistant_system.txt`
    * **TDoc Technical Triage:** `config/prompts/triage_system.txt` and `config/prompts/triage_user.txt`
    * **Email Discussion Summarizer:** `config/prompts/email_summary_system.txt` and `config/prompts/email_summary_user.txt`
    * **Call-Flow Sequence Diagram Generator:** `config/prompts/callflow_system.txt` and `config/prompts/callflow_user.txt`
* **Windows MAX_PATH (260 Character Limit) & Deep Folder Paths:**
  * When downloading extensive TDoc packages or extracting deep specification hierarchies, path lengths can exceed the standard Windows 260-character limit. The application core automatically canonicalizes paths using Windows Extended-Length Path prefixes (`\\?\` via `to_extended_path`), preventing `[WinError 3]` or file-not-found exceptions without requiring system-wide registry modifications.
