# 📊 3GPP Meeting Tools & Diagram Converter

An advanced, component-based desktop IDE designed to bridge the gap between text-based diagramming (`PlantUML`) and corporate enterprise environments (`Microsoft Visio` and `PowerPoint`)[cite: 31]. 

Built specifically with telecommunications and 3GPP standards workflows in mind, this tool allows you to write highly efficient PlantUML sequence, activity, and network diagrams, instantly export them as fully editable native Office shapes, rapidly slice massive specification documents into manageable chapters, track NAS, ASN.1 (RRC / NGAP), and PFCP (TS 29.244) protocol message evolutions, search arbitrary substrings across specification releases using FTS5 trigram indexing with "First Added" and cutoff date detection, manage local SQLite databases with built-in compaction tools, track emails across working groups linked to specific TDocs and their revision families, and seamlessly navigate, filter, and synchronize the vast 3GPP meeting, specification, and work item archives locally[cite: 29, 31].

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
  * **Arbitrary Substring Matching:** Powered by an embedded SQLite Full-Text Search (FTS5) engine configured with a 3-character Trigram tokenizer (`tokenize="trigram"`)[cite: 31]. Enables near-instantaneous search for exact phrases, field substrings, acronyms, or protocol constants across millions of words without full-table scan delays[cite: 31].
  * **Targeted Release & Clause Filtering:** Filter queries by specific clause patterns (e.g., `5.2`, `8.1.4`, `Annex A`) or execute cross-specification queries across all active releases simultaneously[cite: 31].
* **Release Evolution Matrix & "First Added" Text Tracking:**
  * **Per-Specification Tabbed Matrix Visualization:** Automatically isolates search results into dedicated per-specification tabs (e.g., `TS 23.501 (32)`, `TS 23.502 (20)`)[cite: 31]. This prevents sparse empty matrices, eliminates colliding clause numbers, and preserves clean chronological column ordering per document[cite: 31].
  * **"First Added" Identification:** Automatically determines the exact earliest release where matching text was introduced, rendering clear visual indicators[cite: 31]:
    * 🟢 **`🟢 Added`**: Highlighted in soft green to indicate the exact version where text first appeared in that clause[cite: 31].
    * ⚪ **`✓ Present`**: Retained and present in subsequent releases[cite: 31].
    * 🔴 **`✗ Removed`**: Highlighted in soft red when text present in a previous release was deleted in that version[cite: 31].
    * ➖ **`-`**: Clause not matching or not present in that release[cite: 31].
* **Date Cutoff Analysis:**
  * **Official Release Date Storage:** Tracks official 3GPP portal upload and publication dates across all indexed specification releases[cite: 31].
  * **Post-Cutoff Date Additions Filter:** Toggle the **🎯 Date Cutoff** selector to highlight text introduced after a target date[cite: 31]:
    * ⚡ **`⚡ Post-Cutoff Added`**: Highlighted in soft amber/yellow to clearly identify text additions introduced after cutoff dates[cite: 31].
    * **Exclusive Filter Mode:** Check **Show Only Post-Cutoff Additions** to hide clauses where matching text was already present prior to the selected priority date (filtering out prior art)[cite: 31].
* **Universal Specification Ingestion Dialog:**
  * **Unrestricted Catalog Access:** Master-detail browser spanning all ~1,500+ specifications across Series 01 through 55 and all Working Groups (RAN1–4, SA1–6, CT1–4)[cite: 31].
  * **Live Search & Presets:** Filter by keyword, topic, or specification number with built-in quick presets for core 3GPP specifications[cite: 31].
  * **Explicit Checkbox Selection:** Dedicated checkbox column for unambiguous selection tracking with dynamic count badges (`Selected: N version(s)`)[cite: 31].
  * **Smart Batch Selectors:** One-click helpers including **`⚡ Select All Unindexed`**, **`⭐ Select Latest per Release`** (supporting both decimal and 3-digit lettered versions like `i40` / `g30`), **`☑️ Select All`**, and **`◻️ Clear`**[cite: 31].
  * **Revision-Mark Filtering:** Automatically discards 3GPP Word change-mark files (`-rm` / `_rm`) during unzipping and local imports, ensuring only clean (`-cl`) specification text is indexed[cite: 31].
* **Multi-Part Split Document Parsing:**
  * **Split Document Sequencing:** Automatically detects, sequences, and unifies modern multi-part specification archives (e.g., `_s00_s04.docx`, `_s05_s08.docx`, `_s09_s14.docx`) into a single consolidated release model in SQLite[cite: 31].
  * **High-Performance XML Extraction:** Direct `lxml` parsing extracts document structure directly from OpenXML without Microsoft Word COM runtime overhead[cite: 31].
  * **Intelligent Heading & TOC Sanitization:** Distinguishes genuine 3GPP headings from numbered procedure call flow steps and strips Table of Contents (TOC) dot leaders and stub entries[cite: 31].
* **Rich Clause Content Inspector:**
  * **💡 Key Match Excerpt:** Dedicated callout banner at the top of the inspector displaying the exact matching paragraph and surrounding sentence context[cite: 31].
  * **Interactive Term Navigation:** **`[ ◀ Prev ]`** and **`[ Next ▶ ]`** buttons with an active match counter (`🎯 N Match(es)`) and auto-scroll to match locations[cite: 31].
  * **One-Click Citation Copy:** **`[ 📋 Copy Citation ]`** button formats and copies complete clause text with official 3GPP document headers, versions, and release dates[cite: 31].
* **Persistent Search Configuration:**
  * Active search queries, clause filters, cutoff dates, and exact checked specification versions are automatically persisted to `spec_search_config.json` and restored across application sessions[cite: 31].
* **Non-Blocking Database Maintenance:**
  * Background `SpecSearchWipeWorker` thread enables fast, freeze-free database resets with automatic checkpointing and SQLite schema reconstruction[cite: 31].

---

### 🔬 3GPP Protocol Evolution Matrix & Inspector (NAS, ASN.1 & PFCP)
* **Comprehensive Multi-Protocol Ingestion:**
  * **NAS Protocols:** Complete support for **5GS NAS (TS 24.501)** and **EPS NAS (TS 24.301)**[cite: 31].
  * **ASN.1 Protocols:** Native support for **NR RRC (TS 38.331)**, **LTE RRC (TS 36.331)**, and **NGAP (TS 38.413)**[cite: 31].
  * **GTP-U Protocol:** Native support for **GTPv1-U (TS 29.281)** covering user plane tunnels across 5GS (`N3`, `N9`, `N19`, `F1-U`, `Xn-U`, `W1-U`), EPS (`S1-U`, `X2-U`, `S5/S8`), and legacy interfaces. Parses signalling messages (Echo Request/Response, Error Indication, End Marker, Tunnel Status) and synthesizes G-PDUs with Clause 5.2 Extension Headers (PDU Session Container, NR RAN Container, PDU Set Information Container) with per-interface filtering.
  * **PFCP Protocol:** Native support for **PFCP (TS 29.244)** spanning both 5GC (`N4`, `N4mb`) and EPC (`Sxa`, `Sxb`, `Sxc`) reference points[cite: 15, 25]. Parses top-level Node-Related (Clause 7.4) and Session-Related (Clause 7.5) PDU messages alongside the master Information Element Type registry (Table 8.1.2-1)[cite: 29].
  * **Hierarchical Grouped IE Unrolling:** Recursively traverses and unrolls nested PFCP Grouped IEs (e.g., `Create PDR └─ PDI └─ SDF Filter`, `Create FAR`, `Create URR`, `Usage Report`) into the Evolution Matrix with tree indentation and depth tracking[cite: 26, 29].
  * **Interface Applicability Metadata & Filtering:** Automatically indexes per-IE interface applicability tags (`Sxa`, `Sxb`, `Sxc`, `N4`, `N4mb`) and provides a dynamic UI filter dropdown (`[All Interfaces | N4 | N4mb | Sxa | Sxb | Sxc]`) that appears whenever a PFCP message is selected[cite: 27, 28].
  * **Release 20+ Multi-Part Document Ingestion:** Automatically detects, sequences, and parses modern split 3GPP specifications by aggregating all clause sub-documents into a single unified release model[cite: 31].
  * **Automated Legacy `.doc` Conversion:** Automatically converts older binary Word 97–2003 `.doc` specifications to `.docx` via headless COM automation or LibreOffice with Protected View bypass and NTFS Zone Identifier unblocking before parsing[cite: 31].
  * **High-Performance XML Parsing:** Direct `lxml` extraction parses message definition tables, ASN.1 syntax blocks, PFCP Grouped IE tables, and field description tables directly from `.docx` archives without requiring Word runtime overhead[cite: 29, 31].
* **Evolution Matrix & Visual Diffing:**
  * **Hierarchical Sequence & Group Unrolling:** Recursively unrolls nested ASN.1 sequence/choice fields and PFCP grouped structures, allowing you to track high-level and deep parameter changes across releases simultaneously[cite: 26, 31].
  * **Visual Release Diffing:** Color-coded matrix cells immediately highlight field additions (🟢 Green), removals (🔴 Red), and format/type modifications (🟡 Yellow) between chronological 3GPP releases[cite: 31].
  * **Hierarchical Specification Tree:** Interactive tree view (`QTreeWidget`) grouping releases under collapsible specification parents (`TS 38.331`, `TS 29.244`, `TS 24.501`, `TS 24.301`, etc.) with master toggles, specification-wide selection, and right-click deletion context menus[cite: 25, 31].
  * **Persistent Filter Configuration:** Tree expansion states, active releases, message selections, and search terms are automatically persisted in `nas_config.json` across sessions[cite: 31].
* **Dual-Layer & Extended Description Search:**
  * **Debounced Filtering:** Dedicated 250ms debounced search bars for Message Names and Information Elements / Fields[cite: 31].
  * **Extended Description Search (`📖 Desc`):** Toggle deep text search across underlying Clause 8/9 IE descriptions and ASN.1 field description tables (e.g., searching `"emergency"` or `"slicing"` highlights matching messages even if the keyword is not in the field name)[cite: 31].
* **Structure & Field Descriptions Inspector:**
  * **Bit-Level Structure Rendering:** Renders bit-level octet diagrams (Figure 8.x/9.x) and value coding tables (Table 8.x/9.x) with full OpenXML `gridSpan` (colspan) and `vMerge` (rowspan) support[cite: 29, 31].
  * **ASN.1 Syntax & Descriptions View:** Renders syntax-highlighted ASN.1 definition blocks accompanied by formatted 3GPP Field Description tables[cite: 31].
  * **Reverse Field/IE Lookup:** Interactive header badge (`Used in: N messages ▾`) and right-click matrix context menu to trace and jump to all messages referencing a given IE or ASN.1 type across active releases[cite: 31].

---

### 🗄️ Database Maintenance & Compaction
* **System-Level Database Manager:** Integrated `🗄️ Database` tool accessible directly from the bottom system bar[cite: 31].
* **Automatic Freelist & WAL Compaction:** Inspects on-disk database sizes and active Write-Ahead Logs (`-wal`), executing `PRAGMA wal_checkpoint(TRUNCATE)` and `VACUUM` to defragment pages and reclaim megabytes of disk space after heavy scraping or database wiping[cite: 31].
* **Batch Maintenance:** One-click **Compact All Databases** to optimize `3gpp_data.db`, `3gpp_protocol_data.db`, `3gpp_spec_search.db`, and auxiliary cache databases concurrently[cite: 31].

---

### 📡 3GPP Meeting, Specification & Work Items Database
* **Asynchronous Three-Phase Syncing Engine:** 
  * **Phase 1 (FTP Directory Mapping):** Scrapes the 3GPP FTP archives in parallel to instantly populate your database with all available meeting numbers, gracefully handling hidden RAN Ad-Hoc (`TSGR_AHs`) subdirectories[cite: 31].
  * **Phase 2 (Deep Document Scrape):** Crawls the `Docs/` folder of every meeting[cite: 31]. Uses smart regex stripping to ignore file extensions and revisions, mathematically sorting the files to determine the first and last TDocs of the meeting[cite: 31].
  * **Phase 3 (DynaReport Upserting):** Injects metadata (Location, Start/End Dates, Ad-Hoc/Electronic status) by fetching the legacy 3GPP Portal HTML tables[cite: 31].
* **Targeted Quick Fetch:** Instantly sync individual specifications (e.g., `23.801-01`) or entire specification series (e.g., `23`) directly from the FTP server without needing to run a lengthy full database sync[cite: 31].

* **3GPP Work Items (WIs) Synchronizer:**
  * **Parallel Multi-WG Scraper:** Concurrently scrapes active Work Items across all 19 Technical Specification Groups and Working Groups (SA, SA1-6, RAN, RAN1-6, CT, CT1-6) from official 3GPP dynamic report pages using multi-threaded execution (5 workers)[cite: 31].
  * **High-Performance Bulk Upsert:** Utilizes atomic SQLite bulk transactions (`executemany` with `ON CONFLICT DO UPDATE`) to instantly sync thousands of work items and map them to their respective working groups via relational sidecar tables (`work_items`, `wi_group_map`, `wi_remarks`)[cite: 31].
  * **Interactive UI Tab:** Features a dedicated tab with a real-time progress bar, status feedback, and helpful button tooltips[cite: 31]. Includes debounced multi-select CheckableComboBox filters (Release, WG) with persistent state-saving, chronologically sorted historical remarks via a custom interactive UI bubble, and clickable WID hyperlinks that automatically route through the global TDoc fetcher or 3GPP Portal[cite: 31].

* **3GPP Work Items (WIs) & Specification Linkage:**
  * **Relational Mapping (`spec_wi_map`):** Bi-directionally maps 3GPP Specifications to Work Items during Pass 2 DynaReport scraping without requiring rigid locks on un-synced WIs[cite: 31].
  * **Specification Inspector Chips:** Details dialogs display interactive primary (⭐) and secondary Work Item chips with direct 3GPP portal navigation[cite: 31].
  * **Work Items Table & Local Specs Inspector:** The Work Items tab features dedicated **WG** and **Linked Specs** columns, local specification inspectors (`LinkedSpecsDialog`), and one-click citation copy actions[cite: 31].

* **Intelligent TDocs Manager:**
  * **Smart Global TDoc Search:** Instantly locate and download any document across the entire database[cite: 31]. Just type a TDoc number (e.g., `S2-2605740r11`) and the UI will dynamically reveal minimalist quick-actions to download the specific file or open its parent meeting context—all without leaving the main dashboard[cite: 31].
  * **Persistent Personal Notes & Status (Sidecar Database):** Keep a private, local SQLite database that "overlays" your data onto the 3GPP list[cite: 31]. Double-click any TDoc to assign a color-coded status (🟢 Support, 🔴 Object, 🟡 Monitor) and save personal notes[cite: 31]. Your data survives perfectly even when downloading fresh 3GPP Excel updates[cite: 31].
  * **Smart Revision Inheritance:** When a TDoc gets a new revision during a meeting, the new child document automatically inherits a "Ghost" version of the personal notes and status you assigned to the base document[cite: 31]!
  * **Interactive Secretary Remarks:** TDocs mentioned in the Secretary Remarks are automatically identified and converted into hyperlinks[cite: 31]. Left-click a link to instantly jump to that row (intelligently wiping active filters if necessary), or right-click to download it or add it to your Comparison Cart[cite: 31].
  * **Natural Sorting & Smart Filtering:** Bulletproof multi-select dropdowns and natural numerical sorting for complex multi-level Agenda Items (e.g., AI 20.6.2 sorts correctly before 20.6.11)[cite: 31].
  * **Comprehensive Analytics Dashboards:** Generate interactive offline HTML Plotly reports detailing TDoc outcomes, top contributing companies, and complex strategic alliance network graphs (co-signing clusters) using Louvain community detection algorithms[cite: 31].
  * **SA2 Electronic Revisions & Agenda Parsing:** Automatically parses `TdocsByAgenda.htm` to extract comments, inject on-the-fly revisions directly into your table, and provides a "No Comments Only" filter[cite: 31]. For eMeetings, it automatically scrapes the `INBOX/Revisions/` FTP folder[cite: 31].
  * **SA2 Chairman's Notes & Session List Ingestion (`.doc` / `.docx`):**
    * **Frosted Drop Overlay:** Drag and drop `.doc`, `.docx`, `.htm`, or `.html` session documents onto the TDocs window; a visual frosted-blue drop overlay appears with dashed borders and instant drop targets[cite: 31].
    * **Non-Blocking Background Worker (`WordAgendaImporterThread`):** Copies imported files to `{meeting_dir}/Agenda/`, unblocks NTFS Zone Identifiers, converts legacy macro-bearing `.doc` files via headless LibreOffice, and parses table data in the background without freezing the UI[cite: 31].
  * **Multi-Action Resources Menu:** Instantly jump to local cache directories, fetched HTML Agenda files, Main FTP folders, Docs/ folders, or Revisions/ folders directly from the UI[cite: 31].
  * **Quick Launch History:** Remembers your active working group session, allowing you to bypass the database table and jump back into your last opened meeting with a single click[cite: 31].

* **Smart Network Detection:** Automatically detects when you are connected to the official "3GPPWIFI" network during live meetings[cite: 31]. It runs a lightweight background thread to ping the internal local server (e.g., `10.10.10.10`) and displays a persistent visual indicator in the status bar[cite: 31]. This enables dynamic features like bypassing public internet firewalls and routing downloads directly through the high-speed local meeting network[cite: 31].

* **3GPP FTP Session Manager:** Automatically injects randomized User-Agents and HTTP Keep-Alive headers[cite: 31]. Features a configurable **Humanness Delay** engine to bypass aggressive 3GPP server throttling and "Too Many Requests" blocks, which can be dialed down to 0.0 for maximum scraping speed[cite: 31].

---

### 📧 Universal TDoc Email Tracker & Inspection Dialog
* **Working Group-Agnostic Ingestion:** Indexes emails across any 3GPP Working Group (SA2, RAN2, CT1, etc.) directly from your Outlook folders without moving emails or touching server-side folders[cite: 31]. Operates independently of the dedicated eMeeting logic to prevent regressions[cite: 31].
* **WG-Dependent Multi-Folder Profiles & Custom Tag Colors:**
  * Configure specific Outlook folders per Working Group (saved globally in `emails_config.json`)[cite: 31].
  * Assign custom tags (e.g., `[WG]`, `[Disc]`, `[Offline]`, `[Inbox]`) and pick personalized badge colors using an interactive `QColorDialog`[cite: 31]. Tags render in the conversation stream with custom contrasting colors[cite: 31].
* **Smart Quotation Boundary & Direct Message Detection:**
  * Differentiates whether a TDoc was cited in the **Subject**, the **Direct Body** of the message, or an inherited historical reply chain (**Quoted**)[cite: 31].
  * Eliminates false-positive cascades where casual replies (`"ok, danke"`, `"+1"`) cite TDocs buried in older email footers[cite: 31].
  * Toggle **`☑️ Include Quoted Matches`** to hide or reveal conversational thread citations on demand[cite: 31].
* **Exchange Internal Senders & DMARC Resolution:**
  * Automatically resolves listserv rewrites (`LIST.ETSI.ORG`) and internal Exchange X.500 addresses (`/o=...` / `EX`) to primary SMTP addresses to ensure company sanitization recognizes internal colleagues[cite: 31].
* **Modeless, Multi-Window Architecture:**
  * The inspection dialog operates as an independent, modeless top-level window (`Qt.Window`)[cite: 31]. It never freezes or blocks the main TDocs list or background downloads, allowing you to snap windows side-by-side[cite: 31].
  * Multiple TDocs can be inspected concurrently without duplicate window spawning[cite: 31].
* **Interactive TDoc Linkifier:**
  * Automatically converts every detected 3GPP TDoc number in the Subject line, Match Excerpt banner, and Body text into a clickable link[cite: 31].
  * Current document family numbers are highlighted in amber (`#FFF176`), while cross-referenced TDocs appear with interactive links (e.g., `🔗 S2-2608457`)[cite: 31].
  * Clicking any referenced TDoc instantly launches an inspection window for that document[cite: 31].
* **Reading Pane Controls & Standalone Viewer:**
  * **Interactive Vertical Splitter:** Drag the splitter bar between the email list and the reading pane to adjust viewing proportions[cite: 31].
  * **Unicode Whitespace Compression:** Automatically strips invisible non-breaking spaces (`\xa0`) and collapses excessive blank lines from Word/Outlook formatting into clean, readable text[cite: 31].
  * **💡 Match Found Callout:** Displays an excerpt banner directly above the message showing the exact surrounding sentence context where the TDoc was found[cite: 31].
  * **`⧉ Pop Out View`:** Detaches the message preview into an independent, fully resizable viewer (`StandaloneEmailReaderWindow`) with live selection synchronization across multi-monitor or laptop setups[cite: 31].
* **Read / Unread Lifecycle & Ignore Engine:**
  * Track local read states in SQLite (`general_emails.is_read`)[cite: 31].
  * Selecting an email marks it as read after an 800ms debounce[cite: 31].
  * Multi-select rows with `Ctrl` or `Shift` to batch Mark Read, Mark Unread, Ignore, or Delete[cite: 31].
  * **`🚫 Ignore` Action:** Suppresses high-volume distribution list announcements or rapporteur compilation emails from all document counts without deleting them[cite: 31]. Ignored flags are preserved across re-syncs[cite: 31]. Toggle **`Show Ignored`** to review or un-ignore them[cite: 31].
* **TDocs Window Integration:**
  * **`Emails` Column:** Displays aggregate family email counts with unread badges (e.g., `✉️ 5 (🔵 2)`)[cite: 31].
  * **Context Menu:** Right-click any row to view related emails or toggle all emails for that TDoc's revision family between read and unread[cite: 31].
  * **`📧 Emails ▾` Header Menu:** One-click menu to sync related emails, configure folders, mark all as read, or execute a high-speed wipe of the generic emails database[cite: 31].

---

### 📧 eMeeting Email Manager (Dedicated SA2 eMeeting Dashboard)
* **High-Performance Sync Engine:** Connects directly to your local Microsoft Outlook via COM automation[cite: 31]. Pulls, parses, and indexes thousands of eMeeting mailing list emails in milliseconds using SQLite chunked batching (`executemany`) with zero memory spikes[cite: 31].
* **Master-Detail Thread Architecture:** Bypasses broken Outlook reply chains by logically grouping emails purely by parsed TDoc numbers[cite: 31]. The UI features a split-screen design: a Left Panel displaying active TDoc threads and a Right Panel displaying the isolated, chronological conversation for the selected topic[cite: 31].
* **Intelligent 3GPP Parser:** Uses smart regex to extract TDoc numbers (6-8 digits), Agenda Items, Revisions, and free text directly from standard 3GPP bracketed subject lines and email bodies[cite: 31].
* **DMARC Listserv Bypass:** Automatically detects when 3GPP mailing lists rewrite the sender address to `LIST.ETSI.ORG`[cite: 31]. It parses the actual sender's name and email address from the email body and maps them to known telecommunication companies[cite: 31].
* **Advanced Dual-Layer Filtering:** 
  * **Macro-Filters (Thread Level):** Use Star (⭐) and Follow (👀) buttons, or the global search bar, to instantly filter the left-hand thread list down to specific topics or Agenda Items of interest[cite: 31].
  * **Micro-Filters (Conversation Level):** Once a thread is selected, use the Company dropdown, Sender dropdown, or Text search boxes to isolate specific replies strictly within that single conversation[cite: 31].
* **Interactive Email Analytics:** Click the **Statistics** button to instantly generate an interactive, offline HTML Plotly dashboard visualizing Agenda Item volumes, company activity rankings, timeline histograms, and top delegate leaderboards[cite: 31].
* **Automated Archiving:** Safely extracts physical `.msg` files to your hard drive and dynamically builds a clean target folder hierarchy in Outlook (e.g., `Archive/SA2_175/9.1.1/`) to permanently organize your inbox[cite: 31].

---

### 📝 Word Document Manipulation & AI Integration
* **🤖 AI/LLM Corpus Exporter:**
  * **Smart Automation:** Automatically downloads missing TDocs from the 3GPP FTP and extracts the underlying Word documents in the background[cite: 31].
  * **Intelligent Parsing:** Uses a custom Regex State Machine to handle complex 3GPP formatting, including extracting Track Changes and parsing tricky "all new text" placeholder clauses (e.g., `6.4.5.X`)[cite: 31].
  * **Mega-File Compilation:** Compiles and groups the extracted text into clean, Agenda Item-specific Markdown files tailored specifically for LLM context windows (Gemini, Claude, GPT)[cite: 31].
* **Global Comparison Cart:** A persistent, round-robin state dashboard that bridges multiple meeting windows[cite: 31]. Intelligently push any Base TDoc or specific Revision into alternating slots, then launch a native Word comparison instantly[cite: 31].
* **Isolated Word Diff Engine:** Uses COM `DispatchEx` to spawn an invisible, isolated instance of Microsoft Word[cite: 31]. It safely opens files as Read-Only, auto-accepts tracked changes purely in RAM, assigns proper document names for the comparison pane, and generates a visual diff without freezing your active Word sessions or locking local files[cite: 31].
* **LibreOffice Integration Engine:**
  * **Macro-Free & Sandboxed Conversion:** Built-in adapter leveraging headless LibreOffice with isolated user profiles (`-env:UserInstallation`) to suppress network printer hangs and bypass macro security restrictions[cite: 31].
  * **Installed & Portable Support:** Seamless auto-detection of system-installed LibreOffice and single-click integration for portable distributions (`LibreOfficePortable.exe`)[cite: 31].
* **Corporate IT Bypass (Sensitivity Labels):** Automatically injects configurable Microsoft Purview Sensitivity Labels (e.g., "OFFEN") directly into COM objects to bypass blocking corporate IT popup dialogs during automated saves[cite: 31].
* **Intelligent DocxSplitter:** Safely slices massive 3GPP TS/TR specifications into individual Word documents based on Heading 1 or Heading 2 boundaries, perfectly preserving styles, images, and Visio objects[cite: 31].
* **Background Word-to-PDF Converter:** A headless Word automation thread that silently converts generated files to PDFs or XPS without interrupting your workflow[cite: 31].
* **Native Visio Extractor:** Parses the raw XML (`document.xml`) of a `.docx` file, identifies embedded `OLEObject` bins, and extracts raw `.vsdx` Visio diagrams straight out of the Word document to your local disk[cite: 31].

---

### 🎨 Visio Tools (PlantUML & PowerPoint Converter)
* **Live Preview IDE:** A PlantUML code editor featuring syntax highlighting, line numbering, and a 500ms debounced live-rendering engine[cite: 31].
* **Batch Conversion Engine:** Drag and drop hundreds of `.puml`, `.txt`, or `.pptx` files to queue them for multi-threaded background conversion[cite: 31].
* **PowerPoint to Visio Pipeline:** Seamlessly convert entire PowerPoint presentations into multi-page Visio documents (`.vsdx`)[cite: 31]. Uses Enhanced Metafile (EMF) bridging to perfectly preserve editable native Office shapes, automatically aggressively ungroup them, and shrink wrap their text boundaries[cite: 31].
* **Custom Visio Stencil Engine:** Converts standard PlantUML shapes into grouped Visio shapes (`.vsdx`) mapped directly to custom 3GPP node stencils[cite: 31].

---

## <a id="architecture"></a>🏗️ Architecture & Data Flow

This application strictly adheres to the **Model-View-Controller (MVC)** and **Event-Driven Architecture (EDA)** paradigms using `PyQt5`[cite: 31]. 

1. **The UI Layer (`src/modules/*/ui/`):** Contains Qt Widgets and `QAbstractTableModel` implementations[cite: 31]. It never blocks the main GUI thread[cite: 31].
2. **The Core Layer (`src/modules/*/core/`):** Contains domain logic[cite: 31]. All database transactions (`sqlite3` with FTS5 trigrams), FTP network scraping (`requests`), COM automation (`win32com` & `pythoncom`), headless LibreOffice conversions, and direct XML manipulation (`lxml` & `python-docx`) are isolated here[cite: 31].
3. **The Threading Bridge:** Worker tasks inherit from `QThread` (e.g., `GeneralEmailSyncThread`, `WordAgendaImporterThread`, `TDocsDownloaderThread`, `LLMExporterThread`)[cite: 31]. The UI dispatches tasks to the thread, and worker threads emit `pyqtSignals` back to the UI to update progress indicators, models, and logs asynchronously[cite: 31].
4. **The Singleton Managers:** Network configuration (proxies), Word configuration (Sensitivity Labels), database maintenance handlers, and Comparison Cart states are managed by thread-safe singletons and dynamic JSON config loaders to ensure cross-tab synchronization[cite: 31].

---

## <a id="prerequisites"></a>⚙️ Prerequisites

To run this application natively or build it from source, you must have the following installed on your Windows machine[cite: 31]:

1. **Python 3.10+**[cite: 31]
2. **Microsoft Word (Desktop App)** (Required for native COM Automation Splitter, Converter, and Diff Engine)[cite: 31]
3. **Microsoft Outlook (Desktop App)** (Required for the eMeeting and General Email Managers)[cite: 31]
4. **Java Runtime Environment (JRE) 11+** (Required for the local PlantUML generation engine)[cite: 31]
5. *(Optional but Recommended)* **LibreOffice (Installed or Portable)** (Required for safe, macro-free conversion of legacy Word 97–2003 `.doc` files, including SA2 Chairman's Notes and older specifications[cite: 31]. If using portable LibreOffice, link `LibreOfficePortable.exe` using the **📂 Locate Executable** button in the Word Tools tab[cite: 31].)
6. *(Optional)* **Microsoft Visio** (To view and edit generated `.vsdx` files)[cite: 31]
7. *(Optional)* **Microsoft PowerPoint** (For `.pptx` to `.vsdx` conversions)[cite: 31]

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
*Note: This installs `PyQt5`, `requests`, `python-docx`, `beautifulsoup4`, `openpyxl`, `pandas`, `plotly`, `networkx`, `lxml`, and `pywin32`.*[cite: 31]

### 3. Launch the Application
```bash
python src/main_tools.py
```
*Upon first launch, the app will automatically download the latest `plantuml.jar` from GitHub if it is not present in your assets folder.*[cite: 31]

---

## <a id="usage"></a>📖 How to Use the GUI

### 🔎 3GPP Specification Full-Text & Evolution Search
1. Navigate to the **🔎 Spec Search** tab[cite: 31].
2. **Importing Specifications:**
   * Click **📥 Import from Specs DB** to open the universal specification browser[cite: 31]. Select any 3GPP document (Series 01–55) or filter by Working Group[cite: 31]. Missing archives download and extract from the 3GPP FTP server automatically[cite: 31].
   * Use **`⚡ Select All Unindexed`** or **`⭐ Select Latest per Release`** to batch-select versions with checkboxes[cite: 31].
   * Click **📁 Import Local .docx** to ingest single or multi-part split documents (`_s00_s04.docx`, `_s05_s08.docx`) directly from your drive[cite: 31].
3. **Executing Substring Searches:**
   * Type any exact phrase or keyword into the search bar (e.g., `"slice replacement"`, `"ATSSS"`, `"emergency"`)[cite: 31]. Search queries with 3 or more characters automatically execute across the FTS5 trigram index[cite: 31].
   * Optionally enter a clause number in the **Filter clause** field (e.g., `5.2`, `4.3.2`) to focus on specific sections[cite: 31].
4. **Date Cutoff & "First Added" Text Analysis:**
   * Review the **Release Evolution Matrix** displayed in per-specification tabs (e.g., `TS 23.501 (32)`, `TS 23.502 (20)`)[cite: 31].
   * Toggle **🎯 Date Cutoff** and select a cutoff date[cite: 31]. Text introduced after that date will be highlighted with ⚡ **`⚡ Post-Cutoff Added`**[cite: 31].
   * Check **Show Only Post-Cutoff Additions** to filter out older prior art and show only clauses containing post-cutoff date modifications[cite: 31].
5. **Inspecting Matching Clause Content:**
   * Click any cell in the matrix to load the clause into the **Clause Content Inspector**[cite: 31].
   * The **💡 Key Match Excerpt** callout at the top highlights the matching paragraph with surrounding sentence context[cite: 31].
   * Use **`[ ◀ Prev ]`** and **`[ Next ▶ ]`** to cycle between match occurrences in long clauses[cite: 31].
   * Click **`[ 📋 Copy Citation ]`** to copy the formatted text with 3GPP document, version, and release date metadata directly to your clipboard[cite: 31].

---

### 🔬 3GPP Protocols Evolution Matrix (NAS, ASN.1 & PFCP)
1. Navigate to the **🔬 Protocols** (or **🔬 NAS**) tab[cite: 31].
2. **Importing Specifications:**
   * Click **📥 Import from Specs DB** to select specification releases across **TS 38.331 (NR RRC)**, **TS 36.331 (LTE RRC)**, **TS 38.413 (NGAP)**, **TS 29.244 (PFCP)**, **TS 24.501 (5GS NAS)**, or **TS 24.301 (EPS NAS)**[cite: 25, 31]. Missing versions download and convert automatically from the 3GPP FTP archive[cite: 31].
   * Click **📁 Import Local .docx** to ingest local single-file or multi-part split specification documents directly[cite: 31].
3. **Selecting Releases & Messages:**
   * Use the **Specification Versions & Releases** tree to activate, deactivate, or right-click to delete specific releases or entire specification series[cite: 31].
   * Select a Message, SIB, or PDU from the list (e.g., `PFCP Session Establishment Request`, `RRCReconfiguration`, `SIB1`, `REGISTRATION REQUEST`)[cite: 29, 31]. The **Evolution Matrix** pivots all Information Elements and unrolls nested ASN.1 sequence fields or PFCP Grouped IEs (e.g., `Create PDR └─ PDI └─ SDF Filter`), color-coding additions (🟢), removals (🔴), and modifications (🟡)[cite: 26, 29, 31].
4. **Filtering Fields, Descriptions & Interfaces:**
   * Use **Filter message/SIB name** to search message titles[cite: 31].
   * Use **Filter by IE / Field** to isolate specific parameters across the matrix[cite: 31].
   * Click the **`📖 Desc`** button to toggle extended description search, matching keywords located deep inside Clause 8/9 IE definitions and ASN.1 field description tables[cite: 31].
   * For PFCP messages, use the **Interface Selector Dropdown** (`All Interfaces`, `N4`, `N4mb`, `Sxa`, `Sxb`, `Sxc`) positioned above the matrix table to instantly filter parameters by target reference point[cite: 27].
5. **Inspecting Structure & Reverse Lookup:**
   * Click any row in the matrix to render its Clause 8/9 coding diagram (bit-level octet diagram) or ASN.1 syntax block and Field Descriptions table in the bottom **Inspector**[cite: 29, 31].
   * Click the **Used in: N messages ▾** badge in the inspector header (or right-click any row in the matrix) to find all other messages referencing that parameter across active releases[cite: 31].

---

### 🗄️ Database Maintenance & Compaction
1. Click the **🗄️ Database** button located in the bottom system bar next to Task Manager and Proxy[cite: 31].
2. The dialog displays all SQLite database files (`3gpp_data.db`, `3gpp_protocol_data.db`, `3gpp_spec_search.db`), their current on-disk sizes, and Write-Ahead Log (`-wal`) statuses[cite: 31].
3. Click **Compact** on an individual database or **🧹 Compact All Databases** to flush WAL logs, execute SQLite `VACUUM`, optimize indices, and instantly reclaim free disk space[cite: 31].

---

### 📊 3GPP Meetings & Specifications
1. Navigate to the **Meetings** tab[cite: 31].
2. Click **Sync All Meetings** to trigger the 3-Phase scraper[cite: 31]. You can also use **Open Last Meeting** to instantly resume your previous working group session[cite: 31].
3. Use the **Global TDoc Search** input to instantly find a specific document[cite: 31]. Type a valid TDoc number (e.g., `S2-2605740`), and press **Enter** (or click **📄 Doc**) to fetch and open it immediately, or click **🗓️ Mtg** to launch its parent meeting table[cite: 31].
4. Right-click any meeting to access its FTP folders, view its info, or open its cached **TDocs List**[cite: 31].
5. In the TDocs Window, use the **Search** bar or dropdown filters to find specific documents[cite: 31]. Double-click any cell to open the Notes editor and assign a color-coded status to a document[cite: 31].
6. **Importing SA2 Session Documents & Chairman's Notes:**
   * **Drag & Drop:** Drag any `.docx`, `.doc`, or `.htm` session document anywhere onto the TDocs window[cite: 31]. A visual frosted drop overlay will highlight the window[cite: 31].
   * **Menu Import:** Alternatively, click the **🔄 Refresh** menu and select **📝 Import Word Document (.docx / .doc)...**[cite: 31].
   * The file is automatically copied to `{meeting_dir}/Agenda/`, converted in the background via LibreOffice (if `.doc`), parsed, and merged into the table without freezing the UI[cite: 31].
7. Click the Action column to automatically download, unzip, and open documents, or use the **⚖️ Add to Comparison Cart** submenu to select base versions or revisions for diffing[cite: 31].
8. Under the Specifications tab, use **🎯 Quick Fetch** to surgically inject single specifications or series into the database without a full sync[cite: 31].

---

### 📧 Tracking Related Emails for TDocs (Universal Meeting Support)
1. **Configuring Folders & Tag Colors:**
   * In any open TDocs window, click the **📧 Emails ▾** header menu and select **⚙️ Configure Outlook Folders...**[cite: 31].
   * Click **➕ Add Folder via Outlook...** to browse and map your Working Group distribution list folders (e.g., `SA2_WG`, `SA2_DISC`, `RAN2_List`)[cite: 31].
   * Enter a short Tag (e.g., `WG`, `Disc`, `Offline`) and click the color button to assign a distinct visual badge color using the color picker[cite: 31]. Configurations are saved globally per Working Group[cite: 31].
2. **Syncing Outlook Emails:**
   * Click **📧 Emails ▾ $\rightarrow$ 🔄 Sync Related Emails...**[cite: 31].
   * Confirm the date range (defaults to meeting start/end dates $\pm 3$ days buffer) and click **🚀 Start Sync**[cite: 31].
   * The background engine indexes all mentions of TDocs in both Subject lines and Message bodies without downloading physical `.msg` files[cite: 31].
3. **Inspecting TDoc Conversation Threads:**
   * Review the **Emails** column in the main TDocs table[cite: 31]. Cells display total family counts and blue unread badges (e.g., `✉️ 4 (🔵 2)`)[cite: 31].
   * Double-click any cell in the **Emails** column (or right-click a row and select **📧 View Related Emails...**) to open the modeless inspection dialog[cite: 31].
   * **Family Breadcrumbs:** The top card displays the complete document revision lineage (e.g., `S2-2601000 ➔ S2-2601234 ➔ S2-2601555`)[cite: 31].
   * **Quotation Filter:** Uncheck **Include Quoted Matches** to filter out reply chains that only mentioned the TDoc in historical quoted text[cite: 31].
4. **Navigating & Reading Emails:**
   * Drag the interactive **vertical splitter** to expand the reading pane[cite: 31].
   * Click **⧉ Pop Out View** to detach the reading pane into an independent viewer window (`StandaloneEmailReaderWindow`), ideal for laptop screens or secondary monitors[cite: 31].
   * **Interactive TDoc Links:** Every 3GPP document number cited in the Subject line, Match Excerpt banner, or Body is rendered as an interactive link[cite: 31]. Click any link (e.g., `🔗 S2-2608457`) to open that document's related emails immediately[cite: 31].
   * Click **🚀 Open in Outlook** to view the original message live in native Microsoft Outlook[cite: 31].
5. **Managing Read & Ignored Statuses:**
   * Selecting an email automatically marks it as read[cite: 31].
   * Select multiple rows using `Ctrl` or `Shift` to batch **Mark Read**, **Mark Unread**, **Ignore**, or **Delete**[cite: 31].
   * **`🚫 Ignore`:** Suppresses high-volume mailing list announcements or bulk compilation emails from badge counts across all referenced TDocs without deleting them from the database[cite: 31].
   * Right-click any row in the main TDocs table to mark all emails for that document family as read or unread in one click[cite: 31].
   * To reset generic meeting email records, click **📧 Emails ▾ $\rightarrow$ 🗑️ Wipe Generic Emails Database...**[cite: 31].

---

### 📋 3GPP Work Items (WIs)
1. Navigate to the **3GPP Work Items** tab[cite: 31].
2. Click the **🔄 Sync 3GPP WIs** button (hover over it for tooltip details) to trigger the parallel multi-threaded scraper across all 19 Technical Specification Groups and Working Groups[cite: 31].
3. Monitor the real-time progress bar and status messages as records are fetched and bulk upserted into the shared database[cite: 31].
4. Use the **Local Search** bar and multi-select **Checkable Dropdowns** to debounce-filter the table by Acronym, Name, Code, Release, or Working Group[cite: 31]. Your selected filters are automatically saved and restored between application sessions[cite: 31].
5. **Interactive Columns:** Click any blue **Latest WID** hyperlink to download the document via the global search engine (or fall back to the 3GPP Web Portal)[cite: 31]. Click the interactive **💬 Remarks** button to view a chronologically sorted history of secretary remarks for that specific work item[cite: 31].

---

### 📧 eMeeting Email Manager (SA2 Electronic Sessions)
1. Open a specific electronic meeting from the database, click the **📧 Emails ▾** menu, and choose **📊 Open eMeeting Email Manager (Dashboard)**[cite: 31].
2. Click **⚙️ Folders** to browse your Outlook directory and safely map your Source (Inbox) and Target (Archive) folders[cite: 31].
3. Click **🔄 Sync Source** to download and index all eMeeting emails[cite: 31].
4. Select a TDoc thread from the **Left Panel** to view its chronological email history in the **Right Panel**[cite: 31].
5. Use the **⭐ Star** and **👀 Follow** buttons in the reading pane to track specific documents or entire topics[cite: 31]. Use the left-side filters to isolate these threads, and the right-side dropdowns to filter by Company or Sender strictly within a thread[cite: 31].
6. Select rows and click **➡️ Move Selected** (or **⏭️ Move All**) to organize emails into dynamic Agenda Item subfolders inside your Outlook archive[cite: 31].
7. Click **📊 Statistics** to generate and open an interactive visual analytics dashboard of the meeting's email traffic[cite: 31].

---

### 📝 Slicing & Comparing Word Documents
1. In the **Comparison Cart** at the bottom of the Meetings Tab, sequentially select documents[cite: 31]. The round-robin queue will automatically populate Slot A and Slot B with local files or fetched 3GPP Revisions[cite: 31].
2. Click **Compare in Word**[cite: 31]. The tool will spawn a background process, temporarily remove file locks, and present a native Word redline comparison[cite: 31].
3. For large specs, navigate to the **Spec Splitter** tab, drag a `.docx` file, choose a Heading depth (e.g., "Level 2" for clauses like `6.1`, `6.2`), and click Split[cite: 31].

---

### 🎨 Visio Tools
1. **PlantUML Editor:** Type standard PlantUML code into the left pane[cite: 31]. The Live Preview will automatically update the image on the right[cite: 31].
2. Click **Export Diagram ▼** and select **To Visio (.vsdx)** to generate a native Visio file, or use other options like PowerPoint, SVG, or ASCII[cite: 31].
3. **Batch Process & PowerPoint Conversion:** Navigate to the **📂 Visio Tools** tab and drag-and-drop `.puml`, `.txt`, or `.pptx` (PowerPoint) files into the drop zone[cite: 31]. The system will detect the file type and process it into an editable Visio file in the background[cite: 31].

---

### ⚙️ Configuring Corporate Proxies & Networking
If you are behind a corporate firewall:
1. Glance at the **bottom right status bar** to see your active network status (Public Internet vs. 3GPP Local Network)[cite: 31].
2. Click the **Network Config** button in the Console Panel[cite: 31].
3. Enter your HTTP/HTTPS proxies into the global session without restarting the app[cite: 31].
4. Adjust the **Humanness Delays** to throttle network requests (to mimic human behavior) or set them to 0.0 for maximum download speed[cite: 31].

---

## <a id="troubleshooting"></a>🛠️ Known Quirks / Troubleshooting

* **Corporate IT "Aktion blockiert" on Drag & Drop:**
  * If Windows Defender Attack Surface Reduction (ASR) blocks dragging downloaded `.doc` files directly from your `Downloads` folder, either:
    1. Use the **🔄 Refresh $\rightarrow$ 📝 Import Word Document...** file picker menu[cite: 31].
    2. Unblock the file via Right Click $\rightarrow$ Properties $\rightarrow$ **Zulassen (Unblock)**[cite: 31].
* **Legacy Word 97–2003 Macro Permissions:**
  * Legacy `.doc` files containing VBA macros (like SA2 Chairman's Notes) are blocked by Word COM security settings[cite: 31]. Ensure LibreOffice is installed or point the app to portable LibreOffice (`LibreOfficePortable.exe`) in the Word tab to enable automated, macro-free conversion[cite: 31].
* **Sensitivity Label Dialogs (Microsoft Purview / Azure Information Protection):**
  * If automated Word conversions or comparisons trigger corporate classification popups, configure your default sensitivity label string (e.g., `OFFEN` or `INTERNAL`) in `word_config.json` to allow silent headless saves[cite: 31].