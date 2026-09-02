# 📊 3GPP Meeting Tools & Diagram Converter

An advanced, component-based desktop IDE designed to bridge the gap between text-based diagramming (`PlantUML`) and corporate enterprise environments (`Microsoft Visio` and `PowerPoint`)[cite: 14, 15, 23]. 

Built specifically with telecommunications and 3GPP standards workflows in mind, this tool allows you to write highly efficient PlantUML sequence, activity, and network diagrams, instantly export them as fully editable native Office shapes, rapidly slice massive specification documents into manageable chapters, track NAS and ASN.1 (RRC / NGAP) protocol message evolutions, search arbitrary substrings across specification releases using FTS5 trigram indexing with "First Added" and cutoff date detection, manage local SQLite databases with built-in compaction tools, track emails across working groups linked to specific TDocs and their revision families, and seamlessly navigate, filter, and synchronize the vast 3GPP meeting, specification, and work item archives locally[cite: 14, 15, 18, 23].

---

## 📑 Table of Contents
1. [✨ Features](#features)
2. [🏗️ Architecture & Data Flow](#architecture)
3. [⚙️ Prerequisites](#prerequisites)
4. [🚀 Installation](#installation)
5. [📖 How to Use the GUI](#usage)
6. [🛠️ Known Quirks / Troubleshooting](#troubleshooting)

---

## <a id="features"></a>✨ Features

### 🔎 3GPP Specification Full-Text & Substring Search Engine
* **FTS5 Trigram Substring Search & Chronological Tracking:**
  * **Arbitrary Substring Matching:** Powered by an embedded SQLite Full-Text Search (FTS5) engine configured with a 3-character Trigram tokenizer (`tokenize="trigram"`)[cite: 14, 15, 23]. Enables near-instantaneous search for exact phrases, field substrings, acronyms, or protocol constants across millions of words without full-table scan delays[cite: 14, 15, 23].
  * **Targeted Release & Clause Filtering:** Filter queries by specific clause patterns (e.g., `5.2`, `8.1.4`, `Annex A`) or execute cross-specification queries across all active releases simultaneously[cite: 14, 15, 23].
* **Release Evolution Matrix & "First Added" Text Tracking:**
  * **Per-Specification Tabbed Matrix Visualization:** Automatically isolates search results into dedicated per-specification tabs (e.g., `TS 23.501 (32)`, `TS 23.502 (20)`)[cite: 14, 15, 23]. This prevents sparse empty matrices, eliminates colliding clause numbers, and preserves clean chronological column ordering per document[cite: 14, 15, 23].
  * **"First Added" Identification:** Automatically determines the exact earliest release where matching text was introduced, rendering clear visual indicators[cite: 14, 15, 23]:
    * 🟢 **`🟢 Added`**: Highlighted in soft green to indicate the exact version where text first appeared in that clause[cite: 14, 15, 23].
    * ⚪ **`✓ Present`**: Retained and present in subsequent releases[cite: 14, 15, 23].
    * 🔴 **`✗ Removed`**: Highlighted in soft red when text present in a previous release was deleted in that version[cite: 14, 15, 23].
    * ➖ **`-`**: Clause not matching or not present in that release[cite: 14, 15, 23].
* **Date Cutoff Analysis:**
  * **Official Release Date Storage:** Tracks official 3GPP portal upload and publication dates across all indexed specification releases[cite: 14, 15, 23].
  * **Post-Cutoff Date Additions Filter:** Toggle the **🎯 Date Cutoff** selector to highlight text introduced after a target date[cite: 14, 15, 23]:
    * ⚡ **`⚡ Post-Cutoff Added`**: Highlighted in soft amber/yellow to clearly identify text additions introduced after cutoff dates[cite: 14, 15, 23].
    * **Exclusive Filter Mode:** Check **Show Only Post-Cutoff Additions** to hide clauses where matching text was already present prior to the selected priority date (filtering out prior art)[cite: 14, 15, 23].
* **Universal Specification Ingestion Dialog:**
  * **Unrestricted Catalog Access:** Master-detail browser spanning all ~1,500+ specifications across Series 01 through 55 and all Working Groups (RAN1–4, SA1–6, CT1–4)[cite: 14, 15, 23].
  * **Live Search & Presets:** Filter by keyword, topic, or specification number with built-in quick presets for core 3GPP specifications[cite: 14, 15, 23].
  * **Explicit Checkbox Selection:** Dedicated checkbox column for unambiguous selection tracking with dynamic count badges (`Selected: N version(s)`)[cite: 14, 15, 23].
  * **Smart Batch Selectors:** One-click helpers including **`⚡ Select All Unindexed`**, **`⭐ Select Latest per Release`** (supporting both decimal and 3-digit lettered versions like `i40` / `g30`), **`☑️ Select All`**, and **`◻️ Clear`**[cite: 14, 15, 23].
  * **Revision-Mark Filtering:** Automatically discards 3GPP Word change-mark files (`-rm` / `_rm`) during unzipping and local imports, ensuring only clean (`-cl`) specification text is indexed[cite: 14, 15, 23].
* **Multi-Part Split Document Parsing:**
  * **Split Document Sequencing:** Automatically detects, sequences, and unifies modern multi-part specification archives (e.g., `_s00_s04.docx`, `_s05_s08.docx`, `_s09_s14.docx`) into a single consolidated release model in SQLite[cite: 14, 15, 23].
  * **High-Performance XML Extraction:** Direct `lxml` parsing extracts document structure directly from OpenXML without Microsoft Word COM runtime overhead[cite: 14, 15, 23].
  * **Intelligent Heading & TOC Sanitization:** Distinguishes genuine 3GPP headings from numbered procedure call flow steps and strips Table of Contents (TOC) dot leaders and stub entries[cite: 14, 15, 23].
* **Rich Clause Content Inspector:**
  * **💡 Key Match Excerpt:** Dedicated callout banner at the top of the inspector displaying the exact matching paragraph and surrounding sentence context[cite: 14, 15, 23].
  * **Interactive Term Navigation:** **`[ ◀ Prev ]`** and **`[ Next ▶ ]`** buttons with an active match counter (`🎯 N Match(es)`) and auto-scroll to match locations[cite: 14, 15, 23].
  * **One-Click Citation Copy:** **`[ 📋 Copy Citation ]`** button formats and copies complete clause text with official 3GPP document headers, versions, and release dates[cite: 14, 15, 23].
* **Persistent Search Configuration:**
  * Active search queries, clause filters, cutoff dates, and exact checked specification versions are automatically persisted to `spec_search_config.json` and restored across application sessions[cite: 14, 15, 23].
* **Non-Blocking Database Maintenance:**
  * Background `SpecSearchWipeWorker` thread enables fast, freeze-free database resets with automatic checkpointing and SQLite schema reconstruction[cite: 14, 15, 23].

---

### 🔬 3GPP Protocol Evolution Matrix & Inspector (NAS & ASN.1 / RRC / NGAP)
* **Comprehensive Multi-Protocol Ingestion:**
  * **NAS Protocols:** Complete support for **5GS NAS (TS 24.501)** and **EPS NAS (TS 24.301)**[cite: 2, 14, 15, 23].
  * **ASN.1 Protocols:** Native support for **NR RRC (TS 38.331)**, **LTE RRC (TS 36.331)**, and **NGAP (TS 38.413)**[cite: 14, 15, 23].
  * **Release 20+ Multi-Part Document Ingestion:** Automatically detects, sequences, and parses modern split 3GPP specifications by aggregating all clause sub-documents into a single unified release model[cite: 14, 15, 23].
  * **Automated Legacy `.doc` Conversion:** Automatically converts older binary Word 97–2003 `.doc` specifications to `.docx` via headless COM automation or LibreOffice with Protected View bypass and NTFS Zone Identifier unblocking before parsing[cite: 10, 11, 14, 15, 23].
  * **High-Performance XML Parsing:** Direct `lxml` extraction parses message definition tables, ASN.1 syntax blocks, and field description tables directly from `.docx` archives without requiring Word runtime overhead[cite: 14, 15, 23].
* **Evolution Matrix & Visual Diffing:**
  * **Hierarchical ASN.1 Sequence & CHOICE Unrolling:** Recursively unrolls nested sequence parameters and critical extensions (e.g., `└─ radioBearerConfig`, `└─ masterCellGroup`), allowing you to track high-level and deep field changes across releases simultaneously[cite: 14, 15, 23].
  * **Visual Release Diffing:** Color-coded matrix cells immediately highlight field additions (🟢 Green), removals (🔴 Red), and format/type modifications (🟡 Yellow) between chronological 3GPP releases[cite: 14, 15, 23].
  * **Hierarchical Specification Tree:** Interactive tree view (`QTreeWidget`) grouping releases under collapsible specification parents (`TS 38.331`, `TS 24.501`, `TS 24.301`, etc.) with master toggles, specification-wide selection, and right-click deletion context menus[cite: 14, 15, 23].
  * **Persistent Filter Configuration:** Tree expansion states, active releases, message selections, and search terms are automatically persisted in `nas_config.json` across sessions[cite: 14, 15, 23].
* **Dual-Layer & Extended Description Search:**
  * **Debounced Filtering:** Dedicated 250ms debounced search bars for Message Names and Information Elements / Fields[cite: 14, 15, 23].
  * **Extended Description Search (`📖 Desc`):** Toggle deep text search across underlying Clause 9 IE descriptions and ASN.1 field description tables (e.g., searching `"emergency"` or `"slicing"` highlights matching messages even if the keyword is not in the field name)[cite: 14, 15, 23].
* **Structure & Field Descriptions Inspector:**
  * **NAS Bit-Level Structure Rendering:** Renders bit-level octet diagrams (Figure 9.x) and value coding tables (Table 9.x) with full OpenXML `gridSpan` (colspan) and `vMerge` (rowspan) support[cite: 14, 15, 23].
  * **ASN.1 Syntax & Descriptions View:** Renders syntax-highlighted ASN.1 definition blocks accompanied by formatted 3GPP Field Description tables[cite: 14, 15, 23].
  * **Reverse Field/IE Lookup:** Interactive header badge (`Used in: N messages ▾`) and right-click matrix context menu to trace and jump to all messages referencing a given IE or ASN.1 type across active releases[cite: 14, 15, 23].

---

### 🗄️ Database Maintenance & Compaction
* **System-Level Database Manager:** Integrated `🗄️ Database` tool accessible directly from the bottom system bar[cite: 14, 15, 23].
* **Automatic Freelist & WAL Compaction:** Inspects on-disk database sizes and active Write-Ahead Logs (`-wal`), executing `PRAGMA wal_checkpoint(TRUNCATE)` and `VACUUM` to defragment pages and reclaim megabytes of disk space after heavy scraping or database wiping[cite: 14, 15, 23].
* **Batch Maintenance:** One-click **Compact All Databases** to optimize `3gpp_data.db`, `3gpp_protocol_data.db`, `3gpp_spec_search.db`, and auxiliary cache databases concurrently[cite: 14, 15, 23].

---

### 📡 3GPP Meeting, Specification & Work Items Database
* **Asynchronous Three-Phase Syncing Engine:** 
  * **Phase 1 (FTP Directory Mapping):** Scrapes the 3GPP FTP archives in parallel to instantly populate your database with all available meeting numbers, gracefully handling hidden RAN Ad-Hoc (`TSGR_AHs`) subdirectories[cite: 14, 15, 23].
  * **Phase 2 (Deep Document Scrape):** Crawls the `Docs/` folder of every meeting[cite: 14, 15, 23]. Uses smart regex stripping to ignore file extensions and revisions, mathematically sorting the files to determine the first and last TDocs of the meeting[cite: 14, 15, 23].
  * **Phase 3 (DynaReport Upserting):** Injects metadata (Location, Start/End Dates, Ad-Hoc/Electronic status) by fetching the legacy 3GPP Portal HTML tables[cite: 14, 15, 23].
* **Targeted Quick Fetch:** Instantly sync individual specifications (e.g., `23.801-01`) or entire specification series (e.g., `23`) directly from the FTP server without needing to run a lengthy full database sync[cite: 14, 15, 23].

* **3GPP Work Items (WIs) Synchronizer:**
  * **Parallel Multi-WG Scraper:** Concurrently scrapes active Work Items across all 19 Technical Specification Groups and Working Groups (SA, SA1-6, RAN, RAN1-6, CT, CT1-6) from official 3GPP dynamic report pages using multi-threaded execution (5 workers)[cite: 14, 15, 23].
  * **High-Performance Bulk Upsert:** Utilizes atomic SQLite bulk transactions (`executemany` with `ON CONFLICT DO UPDATE`) to instantly sync thousands of work items and map them to their respective working groups via relational sidecar tables (`work_items`, `wi_group_map`, `wi_remarks`)[cite: 14, 15, 23].
  * **Interactive UI Tab:** Features a dedicated tab with a real-time progress bar, status feedback, and helpful button tooltips[cite: 14, 15, 23]. Includes debounced multi-select CheckableComboBox filters (Release, WG) with persistent state-saving, chronologically sorted historical remarks via a custom interactive UI bubble, and clickable WID hyperlinks that automatically route through the global TDoc fetcher or 3GPP Portal[cite: 14, 15, 23].

* **3GPP Work Items (WIs) & Specification Linkage:**
  * **Relational Mapping (`spec_wi_map`):** Bi-directionally maps 3GPP Specifications to Work Items during Pass 2 DynaReport scraping without requiring rigid locks on un-synced WIs[cite: 1, 3, 4, 23].
  * **Specification Inspector Chips:** Details dialogs display interactive primary (⭐) and secondary Work Item chips with direct 3GPP portal navigation[cite: 1, 6, 23].
  * **Work Items Table & Local Specs Inspector:** The Work Items tab features dedicated **WG** and **Linked Specs** columns, local specification inspectors (`LinkedSpecsDialog`), and one-click citation copy actions[cite: 1, 5, 23].

* **Intelligent TDocs Manager:**
  * **Smart Global TDoc Search:** Instantly locate and download any document across the entire database[cite: 14, 15, 23]. Just type a TDoc number (e.g., `S2-2605740r11`) and the UI will dynamically reveal minimalist quick-actions to download the specific file or open its parent meeting context—all without leaving the main dashboard[cite: 14, 15, 23].
  * **Persistent Personal Notes & Status (Sidecar Database):** Keep a private, local SQLite database that "overlays" your data onto the 3GPP list[cite: 14, 15, 23]. Double-click any TDoc to assign a color-coded status (🟢 Support, 🔴 Object, 🟡 Monitor) and save personal notes[cite: 4, 11, 14, 15, 23]. Your data survives perfectly even when downloading fresh 3GPP Excel updates[cite: 4, 11, 14, 15, 23].
  * **Smart Revision Inheritance:** When a TDoc gets a new revision during a meeting, the new child document automatically inherits a "Ghost" version of the personal notes and status you assigned to the base document[cite: 3, 14, 15, 23]!
  * **Interactive Secretary Remarks:** TDocs mentioned in the Secretary Remarks are automatically identified and converted into hyperlinks[cite: 3, 14, 15, 23]. Left-click a link to instantly jump to that row (intelligently wiping active filters if necessary), or right-click to download it or add it to your Comparison Cart[cite: 2, 4, 11, 14, 15, 23].
  * **Natural Sorting & Smart Filtering:** Bulletproof multi-select dropdowns and natural numerical sorting for complex multi-level Agenda Items (e.g., AI 20.6.2 sorts correctly before 20.6.11)[cite: 3, 14, 15, 23].
  * **Comprehensive Analytics Dashboards:** Generate interactive offline HTML Plotly reports detailing TDoc outcomes, top contributing companies, and complex strategic alliance network graphs (co-signing clusters) using Louvain community detection algorithms[cite: 1, 4, 11, 14, 15, 23].
  * **SA2 Electronic Revisions & Agenda Parsing:** Automatically parses `TdocsByAgenda.htm` to extract comments, inject on-the-fly revisions directly into your table, and provides a "No Comments Only" filter[cite: 3, 4, 11, 14, 15, 23]. For eMeetings, it automatically scrapes the `INBOX/Revisions/` FTP folder[cite: 4, 11, 12, 14, 15, 23].
  * **SA2 Chairman's Notes & Session List Ingestion (`.doc` / `.docx`):**
    * **Frosted Drop Overlay:** Drag and drop `.doc`, `.docx`, `.htm`, or `.html` session documents onto the TDocs window; a visual frosted-blue drop overlay appears with dashed borders and instant drop targets[cite: 11, 23].
    * **Non-Blocking Background Worker (`WordAgendaImporterThread`):** Copies imported files to `{meeting_dir}/Agenda/`, unblocks NTFS Zone Identifiers, converts legacy macro-bearing `.doc` files via headless LibreOffice, and parses table data in the background without freezing the UI[cite: 10, 11, 12, 23].
  * **Multi-Action Resources Menu:** Instantly jump to local cache directories, fetched HTML Agenda files, Main FTP folders, Docs/ folders, or Revisions/ folders directly from the UI[cite: 4, 11, 14, 15, 23].
  * **Quick Launch History:** Remembers your active working group session, allowing you to bypass the database table and jump back into your last opened meeting with a single click[cite: 14, 15, 23].

* **Smart Network Detection:** Automatically detects when you are connected to the official "3GPPWIFI" network during live meetings[cite: 14, 15, 23]. It runs a lightweight background thread to ping the internal local server (e.g., `10.10.10.10`) and displays a persistent visual indicator in the status bar[cite: 4, 11, 14, 15, 23]. This enables dynamic features like bypassing public internet firewalls and routing downloads directly through the high-speed local meeting network[cite: 4, 11, 14, 15, 23].

* **3GPP FTP Session Manager:** Automatically injects randomized User-Agents and HTTP Keep-Alive headers[cite: 12, 13, 14, 15, 23]. Features a configurable **Humanness Delay** engine to bypass aggressive 3GPP server throttling and "Too Many Requests" blocks, which can be dialed down to 0.0 for maximum scraping speed[cite: 12, 13, 14, 15, 23].

---

### 📧 Universal TDoc Email Tracker & Inspection Dialog
* **Working Group-Agnostic Ingestion:** Indexes emails across any 3GPP Working Group (SA2, RAN2, CT1, etc.) directly from your Outlook folders without moving emails or touching server-side folders[cite: 18, 20]. Operates independently of the dedicated eMeeting logic to prevent regressions[cite: 18, 19].
* **WG-Dependent Multi-Folder Profiles & Custom Tag Colors:**
  * Configure specific Outlook folders per Working Group (saved globally in `emails_config.json`)[cite: 21].
  * Assign custom tags (e.g., `[WG]`, `[Disc]`, `[Offline]`, `[Inbox]`) and pick personalized badge colors using an interactive `QColorDialog`[cite: 21]. Tags render in the conversation stream with custom contrasting colors[cite: 21].
* **Smart Quotation Boundary & Direct Message Detection:**
  * Differentiates whether a TDoc was cited in the **Subject**, the **Direct Body** of the message, or an inherited historical reply chain (**Quoted**)[cite: 20, 21].
  * Eliminates false-positive cascades where casual replies (`"ok, danke"`, `"+1"`) cite TDocs buried in older email footers[cite: 20, 21].
  * Toggle **`☑️ Include Quoted Matches`** to hide or reveal conversational thread citations on demand[cite: 21].
* **Exchange Internal Senders & DMARC Resolution:**
  * Automatically resolves listserv rewrites (`LIST.ETSI.ORG`) and internal Exchange X.500 addresses (`/o=...` / `EX`) to primary SMTP addresses to ensure company sanitization recognizes internal colleagues[cite: 20].
* **Modeless, Multi-Window Architecture:**
  * The inspection dialog operates as an independent, modeless top-level window (`Qt.Window`)[cite: 21]. It never freezes or blocks the main TDocs list or background downloads, allowing you to snap windows side-by-side[cite: 18, 21].
  * Multiple TDocs can be inspected concurrently without duplicate window spawning[cite: 18].
* **Interactive TDoc Linkifier:**
  * Automatically converts every detected 3GPP TDoc number in the Subject line, Match Excerpt banner, and Body text into a clickable link[cite: 21].
  * Current document family numbers are highlighted in amber (`#FFF176`), while cross-referenced TDocs appear with interactive links (e.g., `🔗 S2-2608457`)[cite: 21].
  * Clicking any referenced TDoc instantly launches an inspection window for that document[cite: 18, 21].
* **Reading Pane Controls & Standalone Viewer:**
  * **Interactive Vertical Splitter:** Drag the splitter bar between the email list and the reading pane to adjust viewing proportions[cite: 21].
  * **Unicode Whitespace Compression:** Automatically strips invisible non-breaking spaces (`\xa0`) and collapses excessive blank lines from Word/Outlook formatting into clean, readable text[cite: 21].
  * **💡 Match Found Callout:** Displays an excerpt banner directly above the message showing the exact surrounding sentence context where the TDoc was found[cite: 21].
  * **`⧉ Pop Out View`:** Detaches the message preview into an independent, fully resizable viewer (`StandaloneEmailReaderWindow`) with live selection synchronization across multi-monitor or laptop setups[cite: 21].
* **Read / Unread Lifecycle & Ignore Engine:**
  * Track local read states in SQLite (`general_emails.is_read`)[cite: 19].
  * Selecting an email marks it as read after an 800ms debounce[cite: 21].
  * Multi-select rows with `Ctrl` or `Shift` to batch Mark Read, Mark Unread, Ignore, or Delete[cite: 21].
  * **`🚫 Ignore` Action:** Suppresses high-volume distribution list announcements or rapporteur compilation emails from all document counts without deleting them[cite: 19, 21]. Ignored flags are preserved across re-syncs[cite: 19, 20]. Toggle **`Show Ignored`** to review or un-ignore them[cite: 21].
* **TDocs Window Integration:**
  * **`Emails` Column:** Displays aggregate family email counts with unread badges (e.g., `✉️ 5 (🔵 2)`)[cite: 18].
  * **Context Menu:** Right-click any row to view related emails or toggle all emails for that TDoc's revision family between read and unread[cite: 18, 22].
  * **`📧 Emails ▾` Header Menu:** One-click menu to sync related emails, configure folders, mark all as read, or execute a high-speed wipe of the generic emails database[cite: 18, 22].

---

### 📧 eMeeting Email Manager (Dedicated SA2 eMeeting Dashboard)
* **High-Performance Sync Engine:** Connects directly to your local Microsoft Outlook via COM automation[cite: 14, 15, 23]. Pulls, parses, and indexes thousands of eMeeting mailing list emails in milliseconds using SQLite chunked batching (`executemany`) with zero memory spikes[cite: 14, 15, 23].
* **Master-Detail Thread Architecture:** Bypasses broken Outlook reply chains by logically grouping emails purely by parsed TDoc numbers[cite: 14, 15, 23]. The UI features a split-screen design: a Left Panel displaying active TDoc threads and a Right Panel displaying the isolated, chronological conversation for the selected topic[cite: 14, 15, 23].
* **Intelligent 3GPP Parser:** Uses smart regex to extract TDoc numbers (6-8 digits), Agenda Items, Revisions, and free text directly from standard 3GPP bracketed subject lines and email bodies[cite: 14, 15, 23].
* **DMARC Listserv Bypass:** Automatically detects when 3GPP mailing lists rewrite the sender address to `LIST.ETSI.ORG`[cite: 14, 15, 23]. It parses the actual sender's name and email address from the email body and maps them to known telecommunication companies[cite: 14, 15, 23].
* **Advanced Dual-Layer Filtering:** 
  * **Macro-Filters (Thread Level):** Use Star (⭐) and Follow (👀) buttons, or the global search bar, to instantly filter the left-hand thread list down to specific topics or Agenda Items of interest[cite: 14, 15, 23].
  * **Micro-Filters (Conversation Level):** Once a thread is selected, use the Company dropdown, Sender dropdown, or Text search boxes to isolate specific replies strictly within that single conversation[cite: 14, 15, 23].
* **Interactive Email Analytics:** Click the **Statistics** button to instantly generate an interactive, offline HTML Plotly dashboard visualizing Agenda Item volumes, company activity rankings, timeline histograms, and top delegate leaderboards[cite: 14, 15, 23].
* **Automated Archiving:** Safely extracts physical `.msg` files to your hard drive and dynamically builds a clean target folder hierarchy in Outlook (e.g., `Archive/SA2_175/9.1.1/`) to permanently organize your inbox[cite: 14, 15, 23].

---

### 📝 Word Document Manipulation & AI Integration
* **🤖 AI/LLM Corpus Exporter:**
  * **Smart Automation:** Automatically downloads missing TDocs from the 3GPP FTP and extracts the underlying Word documents in the background[cite: 1, 4, 11, 14, 15, 23].
  * **Intelligent Parsing:** Uses a custom Regex State Machine to handle complex 3GPP formatting, including extracting Track Changes and parsing tricky "all new text" placeholder clauses (e.g., `6.4.5.X`)[cite: 1, 14, 15, 23].
  * **Mega-File Compilation:** Compiles and groups the extracted text into clean, Agenda Item-specific Markdown files tailored specifically for LLM context windows (Gemini, Claude, GPT)[cite: 1, 4, 11, 14, 15, 23].
* **Global Comparison Cart:** A persistent, round-robin state dashboard that bridges multiple meeting windows[cite: 4, 11, 14, 15, 23]. Intelligently push any Base TDoc or specific Revision into alternating slots, then launch a native Word comparison instantly[cite: 2, 4, 11, 14, 15, 23].
* **Isolated Word Diff Engine:** Uses COM `DispatchEx` to spawn an invisible, isolated instance of Microsoft Word[cite: 14, 15, 23]. It safely opens files as Read-Only, auto-accepts tracked changes purely in RAM, assigns proper document names for the comparison pane, and generates a visual diff without freezing your active Word sessions or locking local files[cite: 14, 15, 23].
* **LibreOffice Integration Engine:**
  * **Macro-Free & Sandboxed Conversion:** Built-in adapter leveraging headless LibreOffice with isolated user profiles (`-env:UserInstallation`) to suppress network printer hangs and bypass macro security restrictions[cite: 10, 23].
  * **Installed & Portable Support:** Seamless auto-detection of system-installed LibreOffice and single-click integration for portable distributions (`LibreOfficePortable.exe`)[cite: 9, 10, 23].
* **Corporate IT Bypass (Sensitivity Labels):** Automatically injects configurable Microsoft Purview Sensitivity Labels (e.g., "OFFEN") directly into COM objects to bypass blocking corporate IT popup dialogs during automated saves[cite: 14, 15, 23].
* **Intelligent DocxSplitter:** Safely slices massive 3GPP TS/TR specifications into individual Word documents based on Heading 1 or Heading 2 boundaries, perfectly preserving styles, images, and Visio objects[cite: 9, 14, 15, 23].
* **Background Word-to-PDF Converter:** A headless Word automation thread that silently converts generated files to PDFs or XPS without interrupting your workflow[cite: 9, 14, 15, 23].
* **Native Visio Extractor:** Parses the raw XML (`document.xml`) of a `.docx` file, identifies embedded `OLEObject` bins, and extracts raw `.vsdx` Visio diagrams straight out of the Word document to your local disk[cite: 9, 14, 15, 23].

---

### 🎨 Visio Tools (PlantUML & PowerPoint Converter)
* **Live Preview IDE:** A PlantUML code editor featuring syntax highlighting, line numbering, and a 500ms debounced live-rendering engine[cite: 14, 15, 23].
* **Batch Conversion Engine:** Drag and drop hundreds of `.puml`, `.txt`, or `.pptx` files to queue them for multi-threaded background conversion[cite: 14, 15, 23].
* **PowerPoint to Visio Pipeline:** Seamlessly convert entire PowerPoint presentations into multi-page Visio documents (`.vsdx`)[cite: 14, 15, 23]. Uses Enhanced Metafile (EMF) bridging to perfectly preserve editable native Office shapes, automatically aggressively ungroup them, and shrink wrap their text boundaries[cite: 14, 15, 23].
* **Custom Visio Stencil Engine:** Converts standard PlantUML shapes into grouped Visio shapes (`.vsdx`) mapped directly to custom 3GPP node stencils[cite: 14, 15, 23].

---

## <a id="architecture"></a>🏗️ Architecture & Data Flow

This application strictly adheres to the **Model-View-Controller (MVC)** and **Event-Driven Architecture (EDA)** paradigms using `PyQt5`[cite: 14, 15, 23]. 

1. **The UI Layer (`src/modules/*/ui/`):** Contains Qt Widgets and `QAbstractTableModel` implementations[cite: 3, 14, 15, 17, 23]. It never blocks the main GUI thread[cite: 14, 15, 23].
2. **The Core Layer (`src/modules/*/core/`):** Contains domain logic[cite: 14, 15, 17, 23]. All database transactions (`sqlite3` with FTS5 trigrams), FTP network scraping (`requests`), COM automation (`win32com` & `pythoncom`), headless LibreOffice conversions, and direct XML manipulation (`lxml` & `python-docx`) are isolated here[cite: 9, 10, 14, 15, 23].
3. **The Threading Bridge:** Worker tasks inherit from `QThread` (e.g., `GeneralEmailSyncThread`, `WordAgendaImporterThread`, `TDocsDownloaderThread`, `LLMExporterThread`)[cite: 4, 11, 12, 14, 15, 20, 23]. The UI dispatches tasks to the thread, and worker threads emit `pyqtSignals` back to the UI to update progress indicators, models, and logs asynchronously[cite: 4, 11, 12, 14, 15, 20, 23].
4. **The Singleton Managers:** Network configuration (proxies), Word configuration (Sensitivity Labels), database maintenance handlers, and Comparison Cart states are managed by thread-safe singletons and dynamic JSON config loaders to ensure cross-tab synchronization[cite: 4, 11, 14, 15, 23].

---

## <a id="prerequisites"></a>⚙️ Prerequisites

To run this application natively or build it from source, you must have the following installed on your Windows machine[cite: 14, 15, 23]:

1. **Python 3.10+**[cite: 14, 15, 23]
2. **Microsoft Word (Desktop App)** (Required for native COM Automation Splitter, Converter, and Diff Engine)[cite: 14, 15, 23]
3. **Microsoft Outlook (Desktop App)** (Required for the eMeeting and General Email Managers)[cite: 14, 15, 23]
4. **Java Runtime Environment (JRE) 11+** (Required for the local PlantUML generation engine)[cite: 14, 15, 23]
5. *(Optional but Recommended)* **LibreOffice (Installed or Portable)** (Required for safe, macro-free conversion of legacy Word 97–2003 `.doc` files, including SA2 Chairman's Notes and older specifications[cite: 10, 14, 15, 23]. If using portable LibreOffice, link `LibreOfficePortable.exe` using the **📂 Locate Executable** button in the Word Tools tab[cite: 9, 14, 15, 23].)
6. *(Optional)* **Microsoft Visio** (To view and edit generated `.vsdx` files)[cite: 14, 15, 23]
7. *(Optional)* **Microsoft PowerPoint** (For `.pptx` to `.vsdx` conversions)[cite: 14, 15, 23]

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
*Note: This installs `PyQt5`, `requests`, `python-docx`, `beautifulsoup4`, `openpyxl`, `pandas`, `plotly`, `networkx`, `lxml`, and `pywin32`.*[cite: 14, 15, 23]

### 3. Launch the Application
```bash
python src/main_tools.py
```
*Upon first launch, the app will automatically download the latest `plantuml.jar` from GitHub if it is not present in your assets folder.*[cite: 14, 15, 23]

---

## <a id="usage"></a>📖 How to Use the GUI

### 🔎 3GPP Specification Full-Text & Evolution Search
1. Navigate to the **🔎 Spec Search** tab[cite: 14, 15, 23].
2. **Importing Specifications:**
   * Click **📥 Import from Specs DB** to open the universal specification browser[cite: 14, 15, 23]. Select any 3GPP document (Series 01–55) or filter by Working Group[cite: 14, 15, 23]. Missing archives download and extract from the 3GPP FTP server automatically[cite: 14, 15, 23].
   * Use **`⚡ Select All Unindexed`** or **`⭐ Select Latest per Release`** to batch-select versions with checkboxes[cite: 14, 15, 23].
   * Click **📁 Import Local .docx** to ingest single or multi-part split documents (`_s00_s04.docx`, `_s05_s08.docx`) directly from your drive[cite: 14, 15, 23].
3. **Executing Substring Searches:**
   * Type any exact phrase or keyword into the search bar (e.g., `"slice replacement"`, `"ATSSS"`, `"emergency"`)[cite: 14, 15, 23]. Search queries with 3 or more characters automatically execute across the FTS5 trigram index[cite: 14, 15, 23].
   * Optionally enter a clause number in the **Filter clause** field (e.g., `5.2`, `4.3.2`) to focus on specific sections[cite: 14, 15, 23].
4. **Date Cutoff & "First Added" Text Analysis:**
   * Review the **Release Evolution Matrix** displayed in per-specification tabs (e.g., `TS 23.501 (32)`, `TS 23.502 (20)`)[cite: 14, 15, 23].
   * Toggle **🎯 Date Cutoff** and select a cutoff date[cite: 14, 15, 23]. Text introduced after that date will be highlighted with ⚡ **`⚡ Post-Cutoff Added`**[cite: 14, 15, 23].
   * Check **Show Only Post-Cutoff Additions** to filter out older prior art and show only clauses containing post-cutoff date modifications[cite: 14, 15, 23].
5. **Inspecting Matching Clause Content:**
   * Click any cell in the matrix to load the clause into the **Clause Content Inspector**[cite: 14, 15, 23].
   * The **💡 Key Match Excerpt** callout at the top highlights the matching paragraph with surrounding sentence context[cite: 14, 15, 23].
   * Use **`[ ◀ Prev ]`** and **`[ Next ▶ ]`** to cycle between match occurrences in long clauses[cite: 14, 15, 23].
   * Click **`[ 📋 Copy Citation ]`** to copy the formatted text with 3GPP document, version, and release date metadata directly to your clipboard[cite: 14, 15, 23].

---

### 🔬 3GPP Protocols Evolution Matrix (NAS & ASN.1 / RRC / NGAP)
1. Navigate to the **🔬 Protocols** (or **🔬 NAS**) tab[cite: 14, 15, 23].
2. **Importing Specifications:**
   * Click **📥 Import from Specs DB** to select specification releases across **TS 38.331 (NR RRC)**, **TS 36.331 (LTE RRC)**, **TS 38.413 (NGAP)**, **TS 24.501 (5GS NAS)**, or **TS 24.301 (EPS NAS)**[cite: 2, 14, 15, 23]. Missing versions download and convert automatically from the 3GPP FTP archive[cite: 2, 14, 15, 23].
   * Click **📁 Import Local .docx** to ingest local single-file or multi-part split specification documents directly[cite: 14, 15, 23].
3. **Selecting Releases & Messages:**
   * Use the **Specification Versions & Releases** tree to activate, deactivate, or right-click to delete specific releases or entire specification series[cite: 14, 15, 23].
   * Select a Message, SIB, or PDU from the list (e.g., `RRCReconfiguration`, `SIB1`, `REGISTRATION REQUEST`)[cite: 14, 15, 23]. The **Evolution Matrix** pivots all Information Elements and unrolls nested ASN.1 sequence fields (e.g. `└─ radioBearerConfig`), color-coding additions (🟢), removals (🔴), and modifications (🟡)[cite: 14, 15, 23].
4. **Filtering Fields and Descriptions:**
   * Use **Filter message/SIB name** to search message titles[cite: 14, 15, 23].
   * Use **Filter by IE / Field** to isolate specific parameters across the matrix[cite: 14, 15, 23].
   * Click the **`📖 Desc`** button to toggle extended description search, matching keywords located deep inside Clause 9 IE definitions and ASN.1 field description tables[cite: 14, 15, 23].
5. **Inspecting Structure & Reverse Lookup:**
   * Click any row in the matrix to render its Clause 9 coding diagram or ASN.1 syntax block and Field Descriptions table in the bottom **Inspector**[cite: 14, 15, 23].
   * Click the **Used in: N messages ▾** badge in the inspector header (or right-click any row in the matrix) to find all other messages referencing that parameter across active releases[cite: 14, 15, 23].

---

### 🗄️ Database Maintenance & Compaction
1. Click the **🗄️ Database** button located in the bottom system bar next to Task Manager and Proxy[cite: 14, 15, 23].
2. The dialog displays all SQLite database files (`3gpp_data.db`, `3gpp_protocol_data.db`, `3gpp_spec_search.db`), their current on-disk sizes, and Write-Ahead Log (`-wal`) statuses[cite: 14, 15, 23].
3. Click **Compact** on an individual database or **🧹 Compact All Databases** to flush WAL logs, execute SQLite `VACUUM`, optimize indices, and instantly reclaim free disk space[cite: 14, 15, 23].

---

### 📊 3GPP Meetings & Specifications
1. Navigate to the **Meetings** tab[cite: 14, 15, 23].
2. Click **Sync All Meetings** to trigger the 3-Phase scraper[cite: 14, 15, 23]. You can also use **Open Last Meeting** to instantly resume your previous working group session[cite: 14, 15, 23].
3. Use the **Global TDoc Search** input to instantly find a specific document[cite: 14, 15, 23]. Type a valid TDoc number (e.g., `S2-2605740`), and press **Enter** (or click **📄 Doc**) to fetch and open it immediately, or click **🗓️ Mtg** to launch its parent meeting table[cite: 14, 15, 23].
4. Right-click any meeting to access its FTP folders, view its info, or open its cached **TDocs List**[cite: 14, 15, 23].
5. In the TDocs Window, use the **Search** bar or dropdown filters to find specific documents[cite: 4, 11, 14, 15, 23]. Double-click any cell to open the Notes editor and assign a color-coded status to a document[cite: 4, 11, 14, 15, 23].
6. **Importing SA2 Session Documents & Chairman's Notes:**
   * **Drag & Drop:** Drag any `.docx`, `.doc`, or `.htm` session document anywhere onto the TDocs window[cite: 11, 23]. A visual frosted drop overlay will highlight the window[cite: 11, 23].
   * **Menu Import:** Alternatively, click the **🔄 Refresh** menu and select **📝 Import Word Document (.docx / .doc)...**[cite: 11, 23].
   * The file is automatically copied to `{meeting_dir}/Agenda/`, converted in the background via LibreOffice (if `.doc`), parsed, and merged into the table without freezing the UI[cite: 10, 11, 12, 23].
7. Click the Action column to automatically download, unzip, and open documents, or use the **⚖️ Add to Comparison Cart** submenu to select base versions or revisions for diffing[cite: 2, 4, 11, 14, 15, 23].
8. Under the Specifications tab, use **🎯 Quick Fetch** to surgically inject single specifications or series into the database without a full sync[cite: 14, 15, 23].

---

### 📧 Tracking Related Emails for TDocs (Universal Meeting Support)
1. **Configuring Folders & Tag Colors:**
   * In any open TDocs window, click the **📧 Emails ▾** header menu and select **⚙️ Configure Outlook Folders...**[cite: 18].
   * Click **➕ Add Folder via Outlook...** to browse and map your Working Group distribution list folders (e.g., `SA2_WG`, `SA2_DISC`, `RAN2_List`)[cite: 21].
   * Enter a short Tag (e.g., `WG`, `Disc`, `Offline`) and click the color button to assign a distinct visual badge color using the color picker[cite: 21]. Configurations are saved globally per Working Group[cite: 21].
2. **Syncing Outlook Emails:**
   * Click **📧 Emails ▾ $\rightarrow$ 🔄 Sync Related Emails...**[cite: 18].
   * Confirm the date range (defaults to meeting start/end dates $\pm 3$ days buffer) and click **🚀 Start Sync**[cite: 20, 21].
   * The background engine indexes all mentions of TDocs in both Subject lines and Message bodies without downloading physical `.msg` files[cite: 19, 20].
3. **Inspecting TDoc Conversation Threads:**
   * Review the **Emails** column in the main TDocs table[cite: 18]. Cells display total family counts and blue unread badges (e.g., `✉️ 4 (🔵 2)`)[cite: 18].
   * Double-click any cell in the **Emails** column (or right-click a row and select **📧 View Related Emails...**) to open the modeless inspection dialog[cite: 18].
   * **Family Breadcrumbs:** The top card displays the complete document revision lineage (e.g., `S2-2601000 ➔ S2-2601234 ➔ S2-2601555`)[cite: 21].
   * **Quotation Filter:** Uncheck **Include Quoted Matches** to filter out reply chains that only mentioned the TDoc in historical quoted text[cite: 21].
4. **Navigating & Reading Emails:**
   * Drag the interactive **vertical splitter** to expand the reading pane[cite: 21].
   * Click **⧉ Pop Out View** to detach the reading pane into an independent viewer window (`StandaloneEmailReaderWindow`), ideal for laptop screens or secondary monitors[cite: 21].
   * **Interactive TDoc Links:** Every 3GPP document number cited in the Subject line, Match Excerpt banner, or Body is rendered as an interactive link[cite: 21]. Click any link (e.g., `🔗 S2-2608457`) to open that document's related emails immediately[cite: 18, 21].
   * Click **🚀 Open in Outlook** to view the original message live in native Microsoft Outlook[cite: 21].
5. **Managing Read & Ignored Statuses:**
   * Selecting an email automatically marks it as read[cite: 21].
   * Select multiple rows using `Ctrl` or `Shift` to batch **Mark Read**, **Mark Unread**, **Ignore**, or **Delete**[cite: 21].
   * **`🚫 Ignore`:** Suppresses high-volume mailing list announcements or bulk compilation emails from badge counts across all referenced TDocs without deleting them from the database[cite: 19, 21].
   * Right-click any row in the main TDocs table to mark all emails for that document family as read or unread in one click[cite: 18].
   * To reset generic meeting email records, click **📧 Emails ▾ $\rightarrow$ 🗑️ Wipe Generic Emails Database...**[cite: 18].

---

### 📋 3GPP Work Items (WIs)
1. Navigate to the **3GPP Work Items** tab[cite: 14, 15, 23].
2. Click the **🔄 Sync 3GPP WIs** button (hover over it for tooltip details) to trigger the parallel multi-threaded scraper across all 19 Technical Specification Groups and Working Groups[cite: 14, 15, 23].
3. Monitor the real-time progress bar and status messages as records are fetched and bulk upserted into the shared database[cite: 14, 15, 23].
4. Use the **Local Search** bar and multi-select **Checkable Dropdowns** to debounce-filter the table by Acronym, Name, Code, Release, or Working Group[cite: 14, 15, 23]. Your selected filters are automatically saved and restored between application sessions[cite: 14, 15, 23].
5. **Interactive Columns:** Click any blue **Latest WID** hyperlink to download the document via the global search engine (or fall back to the 3GPP Web Portal)[cite: 14, 15, 23]. Click the interactive **💬 Remarks** button to view a chronologically sorted history of secretary remarks for that specific work item[cite: 14, 15, 23].

---

### 📧 eMeeting Email Manager (SA2 Electronic Sessions)
1. Open a specific electronic meeting from the database, click the **📧 Emails ▾** menu, and choose **📊 Open eMeeting Email Manager (Dashboard)**[cite: 18].
2. Click **⚙️ Folders** to browse your Outlook directory and safely map your Source (Inbox) and Target (Archive) folders[cite: 14, 15, 23].
3. Click **🔄 Sync Source** to download and index all eMeeting emails[cite: 14, 15, 23].
4. Select a TDoc thread from the **Left Panel** to view its chronological email history in the **Right Panel**[cite: 14, 15, 23].
5. Use the **⭐ Star** and **👀 Follow** buttons in the reading pane to track specific documents or entire topics[cite: 14, 15, 23]. Use the left-side filters to isolate these threads, and the right-side dropdowns to filter by Company or Sender strictly within a thread[cite: 14, 15, 23].
6. Select rows and click **➡️ Move Selected** (or **⏭️ Move All**) to organize emails into dynamic Agenda Item subfolders inside your Outlook archive[cite: 14, 15, 23].
7. Click **📊 Statistics** to generate and open an interactive visual analytics dashboard of the meeting's email traffic[cite: 14, 15, 23].

---

### 📝 Slicing & Comparing Word Documents
1. In the **Comparison Cart** at the bottom of the Meetings Tab, sequentially select documents[cite: 14, 15, 23]. The round-robin queue will automatically populate Slot A and Slot B with local files or fetched 3GPP Revisions[cite: 14, 15, 23].
2. Click **Compare in Word**[cite: 14, 15, 23]. The tool will spawn a background process, temporarily remove file locks, and present a native Word redline comparison[cite: 14, 15, 23].
3. For large specs, navigate to the **Spec Splitter** tab, drag a `.docx` file, choose a Heading depth (e.g., "Level 2" for clauses like `6.1`, `6.2`), and click Split[cite: 9, 14, 15, 23].

---

### 🎨 Visio Tools
1. **PlantUML Editor:** Type standard PlantUML code into the left pane[cite: 14, 15, 23]. The Live Preview will automatically update the image on the right[cite: 14, 15, 23].
2. Click **Export Diagram ▼** and select **To Visio (.vsdx)** to generate a native Visio file, or use other options like PowerPoint, SVG, or ASCII[cite: 14, 15, 23].
3. **Batch Process & PowerPoint Conversion:** Navigate to the **📂 Visio Tools** tab and drag-and-drop `.puml`, `.txt`, or `.pptx` (PowerPoint) files into the drop zone[cite: 14, 15, 23]. The system will detect the file type and process it into an editable Visio file in the background[cite: 14, 15, 23].

---

### ⚙️ Configuring Corporate Proxies & Networking
If you are behind a corporate firewall:
1. Glance at the **bottom right status bar** to see your active network status (Public Internet vs. 3GPP Local Network)[cite: 4, 11, 14, 15, 23].
2. Click the **Network Config** button in the Console Panel[cite: 14, 15, 23].
3. Enter your HTTP/HTTPS proxies into the global session without restarting the app[cite: 14, 15, 23].
4. Adjust the **Humanness Delays** to throttle network requests (to mimic human behavior) or set them to 0.0 for maximum download speed[cite: 14, 15, 23].

---

## <a id="troubleshooting"></a>🛠️ Known Quirks / Troubleshooting

* **Corporate IT "Aktion blockiert" on Drag & Drop:**
  * If Windows Defender Attack Surface Reduction (ASR) blocks dragging downloaded `.doc` files directly from your `Downloads` folder, either:
    1. Use the **🔄 Refresh $\rightarrow$ 📝 Import Word Document...** file picker menu[cite: 11, 23].
    2. Unblock the file via Right Click $\rightarrow$ Properties $\rightarrow$ **Zulassen (Unblock)**[cite: 23].
* **Legacy Word 97–2003 Macro Permissions:**
  * Legacy `.doc` files containing VBA macros (like SA2 Chairman's Notes) are blocked by Word COM security settings[cite: 8, 9, 23]. Ensure LibreOffice is installed or point the app to portable LibreOffice (`LibreOfficePortable.exe`) in the Word tab to enable automated, macro-free conversion[cite: 9, 10, 14, 15, 23].
* **Sensitivity Label Dialogs (Microsoft Purview / Azure Information Protection):**
  * If automated Word conversions or comparisons trigger corporate classification popups, configure your default sensitivity label string (e.g., `OFFEN` or `INTERNAL`) in `word_config.json` to allow silent headless saves[cite: 23].