# 📊 3GPP Meeting Tools & Diagram Converter

An advanced, component-based desktop IDE designed to bridge the gap between text-based diagramming (`PlantUML`) and corporate enterprise environments (`Microsoft Visio` and `PowerPoint`). 

Built specifically with telecommunications and 3GPP standards workflows in mind, this tool allows you to write highly efficient PlantUML sequence, activity, and network diagrams, instantly export them as fully editable native Office shapes, rapidly slice massive specification documents into manageable chapters, track NAS and ASN.1 (RRC / NGAP) protocol message evolutions, search arbitrary substrings across specification releases using FTS5 trigram indexing with "First Added" detection, manage local SQLite databases with built-in compaction tools, and seamlessly navigate, filter, and synchronize the vast 3GPP meeting, specification, and work item archives locally.

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
* **FTS5 Trigram Substring Search:**
  * **Arbitrary Substring Matching:** Powered by an embedded SQLite Full-Text Search (FTS5) engine configured with a 3-character Trigram tokenizer (`tokenize="trigram"`). Enables near-instantaneous search for exact phrases, field substrings, acronyms, or protocol constants across millions of words without full-table scan delays.
  * **Targeted Release & Clause Filtering:** Filter queries by specific clause patterns (e.g., `5.2`, `8.1.4`, `Annex A`) or execute cross-specification queries across all active releases simultaneously.
* **Release Evolution Matrix & "First Added" Text Tracking:**
  * **Chronological Release Diffing:** Pivots search hits across releases to immediately visualize text presence and modifications over time.
  * **"First Added" Identification:** Automatically determines the exact earliest release where matching text was introduced, rendering clear visual indicators:
    * 🟢 **`🟢 Added`**: Highlighted in soft green to indicate the exact version where text first appeared in that clause.
    * ⚪ **`✓ Present`**: Retained and present in subsequent releases.
    * 🔴 **`✗ Removed`**: Highlighted in soft red when text present in a previous release was deleted in that version.
    * ➖ **`-`**: Clause not matching or not present in that release.
* **High-Performance Direct XML Clause Parsing:**
  * **Direct `lxml` Ingestion:** Parses Word `.docx` specification archives directly at the XML layer (`document.xml`) without requiring Microsoft Word COM automation runtime overhead.
  * **Automatic Clause Segmentation:** Slices documents by numbered headings (`Heading 1` through `Heading 6`) and annexes, associating all nested paragraphs and tables with their respective clause numbers.
  * **Multi-Part & Legacy Support:** Seamlessly unzips split specification archives (e.g., `_s00_s04`, `_s05_s08`) and automatically converts legacy binary `.doc` files to OpenXML `.docx` prior to indexing.
* **Interactive Clause Content Inspector:**
  * **Live Match Highlighting:** Selecting any cell in the evolution matrix renders the full clause content with all matched search terms highlighted in bright yellow.
  * **Complete Context Inspection:** Inspect surrounding clause text, tables, and notes directly inside the application without opening external Word viewers.

### 🔬 3GPP Protocol Evolution Matrix & Inspector (NAS & ASN.1 / RRC / NGAP)
* **Comprehensive Multi-Protocol Ingestion:**
  * **NAS Protocols:** Complete support for **5GS NAS (TS 24.501)** and **EPS NAS (TS 24.301)**[cite: 2].
  * **ASN.1 Protocols:** Native support for **NR RRC (TS 38.331)**, **LTE RRC (TS 36.331)**, and **NGAP (TS 38.413)**.
  * **Release 20+ Multi-Part Document Ingestion:** Automatically detects, sequences, and parses modern split 3GPP specifications (e.g., `24501-k00.zip` split across `_s00_s04`, `_s06_s08`, `_s09_s10`, etc.) by aggregating all clause sub-documents into a single unified release model.
  * **Automated Legacy `.doc` Conversion:** Automatically converts older binary Word 97–2003 `.doc` specifications to `.docx` via headless COM automation with Protected View bypass and NTFS Zone Identifier unblocking before parsing.
  * **High-Performance XML Parsing:** Direct `lxml` extraction parses message definition tables, ASN.1 syntax blocks, and field description tables directly from `.docx` archives without requiring Word runtime overhead.
* **Evolution Matrix & Visual Diffing:**
  * **Hierarchical ASN.1 Sequence & CHOICE Unrolling:** Recursively unrolls nested sequence parameters and critical extensions (e.g., `└─ radioBearerConfig`, `└─ masterCellGroup`), allowing you to track high-level and deep field changes across releases simultaneously.
  * **Visual Release Diffing:** Color-coded matrix cells immediately highlight field additions (🟢 Green), removals (🔴 Red), and format/type modifications (🟡 Yellow) between chronological 3GPP releases.
  * **Hierarchical Specification Tree:** Interactive tree view (`QTreeWidget`) grouping releases under collapsible specification parents (`TS 38.331`, `TS 24.501`, `TS 24.301`, etc.) with master toggles, specification-wide selection, and right-click deletion context menus.
  * **Persistent Filter Configuration:** Tree expansion states, active releases, message selections, and search terms are automatically persisted in `nas_config.json` across sessions.
* **Dual-Layer & Extended Description Search:**
  * **Debounced Filtering:** Dedicated 250ms debounced search bars for Message Names and Information Elements / Fields.
  * **Extended Description Search (`📖 Desc`):** Toggle deep text search across underlying Clause 9 IE descriptions and ASN.1 field description tables (e.g., searching `"emergency"` or `"slicing"` highlights matching messages even if the keyword is not in the field name).
* **Structure & Field Descriptions Inspector:**
  * **NAS Bit-Level Structure Rendering:** Renders bit-level octet diagrams (Figure 9.x) and value coding tables (Table 9.x) with full OpenXML `gridSpan` (colspan) and `vMerge` (rowspan) support.
  * **ASN.1 Syntax & Descriptions View:** Renders syntax-highlighted ASN.1 definition blocks accompanied by formatted 3GPP Field Description tables.
  * **Reverse Field/IE Lookup:** Interactive header badge (`Used in: N messages ▾`) and right-click matrix context menu to trace and jump to all messages referencing a given IE or ASN.1 type across active releases.

### 🗄️ Database Maintenance & Compaction
* **System-Level Database Manager:** Integrated `🗄️ Database` tool accessible directly from the bottom system bar.
* **Automatic Freelist & WAL Compaction:** Inspects on-disk database sizes and active Write-Ahead Logs (`-wal`), executing `PRAGMA wal_checkpoint(TRUNCATE)` and `VACUUM` to defragment pages and reclaim megabytes of disk space after heavy scraping or database wiping.
* **Batch Maintenance:** One-click **Compact All Databases** to optimize `3gpp_data.db`, `3gpp_protocol_data.db`, `3gpp_spec_search.db`, and auxiliary cache databases concurrently.

### 📡 3GPP Meeting, Specification & Work Items Database
* **Asynchronous Three-Phase Syncing Engine:** 
  * **Phase 1 (FTP Directory Mapping):** Scrapes the 3GPP FTP archives in parallel to instantly populate your database with all available meeting numbers, gracefully handling hidden RAN Ad-Hoc (`TSGR_AHs`) subdirectories.
  * **Phase 2 (Deep Document Scrape):** Crawls the `Docs/` folder of every meeting. Uses smart Regex stripping to ignore file extensions and revisions, mathematically sorting the files to determine the first and last TDocs of the meeting.
  * **Phase 3 (DynaReport Upserting):** Injects metadata (Location, Start/End Dates, Ad-Hoc/Electronic status) by fetching the legacy 3GPP Portal HTML tables.
* **Targeted Quick Fetch:** Instantly sync individual specifications (e.g., `23.801-01`) or entire specification series (e.g., `23`) directly from the FTP server without needing to run a lengthy full database sync.

* **3GPP Work Items (WIs) Synchronizer:**
  * **Parallel Multi-WG Scraper:** Concurrently scrapes active Work Items across all 19 Technical Specification Groups and Working Groups (SA, SA1-6, RAN, RAN1-6, CT, CT1-6) from official 3GPP dynamic report pages using multi-threaded execution (5 workers).
  * **High-Performance Bulk Upsert:** Utilizes atomic SQLite bulk transactions (`executemany` with `ON CONFLICT DO UPDATE`) to instantly sync thousands of work items and map them to their respective working groups via relational sidecar tables (`work_items`, `wi_group_map`, `wi_remarks`).
  * **Interactive UI Tab:** Features a dedicated tab with a real-time progress bar, status feedback, and helpful button tooltips. Includes debounced multi-select CheckableComboBox filters (Release, WG) with persistent state-saving, chronologically sorted historical remarks via a custom interactive UI bubble, and clickable WID hyperlinks that automatically route through the global TDoc fetcher or 3GPP Portal.

* **Intelligent TDocs Manager:**
  * **Smart Global TDoc Search:** Instantly locate and download any document across the entire database. Just type a TDoc number (e.g., `S2-2605740r11`) and the UI will dynamically reveal minimalist quick-actions to download the specific file or open its parent meeting context—all without leaving the main dashboard.
  * **Persistent Personal Notes & Status (Sidecar Database):** Keep a private, local SQLite database that "overlays" your data onto the 3GPP list. Double-click any TDoc to assign a color-coded status (🟢 Support, 🔴 Object, 🟡 Monitor) and save personal notes. Your data survives perfectly even when downloading fresh 3GPP Excel updates.
  * **Smart Revision Inheritance:** When a TDoc gets a new revision during a meeting, the new child document automatically inherits a "Ghost" version of the personal notes and status you assigned to the base document!
  * **Interactive Secretary Remarks:** TDocs mentioned in the Secretary Remarks are automatically identified and converted into hyperlinks. Left-click a link to instantly jump to that row (intelligently wiping your active filters if necessary), or right-click to instantly download it or add it to your Comparison Cart.
  * **Natural Sorting & Smart Filtering:** Bulletproof multi-select dropdowns and natural numerical sorting for complex multi-level Agenda Items (e.g., AI 20.6.2 sorts correctly before 20.6.11).
  * **Comprehensive Analytics Dashboards:** Generate interactive offline HTML Plotly reports detailing TDoc outcomes, top contributing companies, and complex strategic alliance network graphs (co-signing clusters) using Louvain community detection algorithms.
  * **SA2 Electronic Revisions & Agenda Parsing:** Automatically parses messy Word-exported `TdocsByAgenda.htm` files to extract comments, inject on-the-fly revisions directly into your table, and provides a "No Comments Only" filter. For eMeetings, it automatically scrapes the `INBOX/Revisions/` FTP folder.
  * **Multi-Action Resources Menu:** Instantly jump to local cache directories, fetched HTML Agenda files, Main FTP folders, Docs/ folders, or Revisions/ folders directly from the UI.
  * **Quick Launch History:** Remembers your exact active working group session, allowing you to bypass the database table and instantly jump back into your last opened meeting with a single click.

* **Smart Network Detection:** Automatically detects when you are connected to the official "3GPPWIFI" network during live meetings. It runs a lightweight background thread to ping the internal local server (e.g., `10.10.10.10`) and displays a persistent visual indicator in the status bar. This enables dynamic features like bypassing public internet firewalls and routing downloads directly through the high-speed local meeting network.

* **3GPP FTP Session Manager:** Automatically injects randomized User-Agents and HTTP Keep-Alive headers. Features a configurable **Humanness Delay** engine to bypass aggressive 3GPP server throttling and "Too Many Requests" blocks, which can be dialed down to 0.0 for maximum scraping speed.

### 📧 eMeeting Email Manager (Native Outlook Integration)
* **High-Performance Sync Engine:** Connects directly to your local Microsoft Outlook via COM automation. Pulls, parses, and indexes thousands of eMeeting mailing list emails in milliseconds using SQLite chunked batching (`executemany`) with zero memory spikes.
* **Master-Detail Thread Architecture:** Bypasses broken Outlook reply chains by logically grouping emails purely by parsed TDoc numbers. The UI features a split-screen design: a Left Panel displaying active TDoc threads and a Right Panel displaying the isolated, chronological conversation for the selected topic.
* **Intelligent 3GPP Parser:** Uses smart regex to extract TDoc numbers (6-8 digits), Agenda Items, Revisions, and free text directly from standard 3GPP bracketed subject lines and email bodies.
* **DMARC Listserv Bypass:** Automatically detects when 3GPP mailing lists rewrite the sender address to `LIST.ETSI.ORG`. It parses the actual sender's name and email address from the email body and maps them to known telecommunication companies.
* **Advanced Dual-Layer Filtering:** 
  * **Macro-Filters (Thread Level):** Use Star (⭐) and Follow (👀) buttons, or the global search bar, to instantly filter the left-hand thread list down to specific topics or Agenda Items of interest.
  * **Micro-Filters (Conversation Level):** Once a thread is selected, use the Company dropdown, Sender dropdown, or Text search boxes to isolate specific replies strictly within that single conversation.
* **Interactive Email Analytics:** Click the **Statistics** button to instantly generate an interactive, offline HTML Plotly dashboard visualizing Agenda Item volumes, company activity rankings, timeline histograms, and top delegate leaderboards.
* **Automated Archiving:** Safely extracts physical `.msg` files to your hard drive and dynamically builds a clean target folder hierarchy in Outlook (e.g., `Archive/SA2_175/9.1.1/`) to permanently organize your inbox.

### 📝 Word Document Manipulation & AI Integration
* **🤖 AI/LLM Corpus Exporter:**
  * **Smart Automation:** Automatically downloads missing TDocs from the 3GPP FTP and extracts the underlying Word documents in the background.
  * **Intelligent Parsing:** Uses a custom Regex State Machine to handle complex 3GPP formatting, including extracting Track Changes and parsing tricky "all new text" placeholder clauses (e.g., `6.4.5.X`).
  * **Mega-File Compilation:** Compiles and groups the extracted text into clean, Agenda Item-specific Markdown files tailored specifically for LLM context windows (Gemini, Claude, GPT).
* **Global Comparison Cart:** A persistent, round-robin state dashboard that bridges multiple meeting windows. Intelligently push any Base TDoc or specific Revision into alternating slots, then launch a native Word comparison instantly.
* **Isolated Word Diff Engine:** Uses COM `DispatchEx` to spawn an invisible, isolated instance of Microsoft Word. It safely opens files as Read-Only, auto-accepts tracked changes purely in RAM, assigns proper document names for the comparison pane, and generates a visual diff without freezing your active Word sessions or locking local files.
* **Corporate IT Bypass (Sensitivity Labels):** Automatically injects configurable Microsoft Purview Sensitivity Labels (e.g., "OFFEN") directly into COM objects to bypass blocking corporate IT popup dialogs during automated saves.
* **Intelligent DocxSplitter:** Safely slices massive 3GPP TS/TR specifications (often hundreds of pages long) into individual Word documents based on Heading 1 or Heading 2 boundaries, perfectly preserving styles, images, and Visio objects.
* **Background Word-to-PDF Converter:** A headless Word automation thread that silently converts generated files to PDFs or XPS without interrupting your workflow.
* **Native Visio Extractor:** Parses the raw XML (`document.xml`) of a `.docx` file, identifies embedded `OLEObject` bins, and extracts raw `.vsdx` Visio diagrams straight out of the Word document to your local disk.

### 🎨 Visio Tools (PlantUML & PowerPoint Converter)
* **Live Preview IDE:** A PlantUML code editor featuring syntax highlighting, line numbering, and a 500ms debounced live-rendering engine.
* **Batch Conversion Engine:** Drag and drop hundreds of `.puml`, `.txt`, or `.pptx` files to queue them for multi-threaded background conversion.
* **PowerPoint to Visio Pipeline:** Seamlessly convert entire PowerPoint presentations into multi-page Visio documents (`.vsdx`). Uses Enhanced Metafile (EMF) bridging to perfectly preserve editable native Office shapes, automatically aggressively ungroup them, and shrink wrap their text boundaries.
* **Custom Visio Stencil Engine:** Converts standard PlantUML shapes into grouped Visio shapes (`.vsdx`) mapped directly to custom 3GPP node stencils.

---

## <a id="architecture"></a>🏗️ Architecture & Data Flow

This application strictly adheres to the **Model-View-Controller (MVC)** and **Event-Driven Architecture (EDA)** paradigms using `PyQt5`. 

1. **The UI Layer (`modules/*/ui/`):** Contains Qt Widgets and `QAbstractTableModel` implementations. It never blocks the main thread.
2. **The Core Layer (`modules/*/core/`):** Contains domain logic. All database transactions (`sqlite3` with FTS5 trigrams), FTP network scraping (`requests`), COM automation (`win32com` & `pythoncom`), and direct XML manipulation (`lxml` & `python-docx`) are isolated here.
3. **The Threading Bridge:** Worker tasks inherit from `QThread`. The UI dispatches tasks to the Thread, and the Thread emits `pyqtSignals` back to the UI to update progress bars, models, and logs asynchronously.
4. **The Singleton Managers:** Network configuration (proxies), Word configuration (Sensitivity Labels), database maintenance handlers, and Comparison Cart states are managed by thread-safe singletons and dynamic JSON config loaders to ensure cross-tab synchronization.

---

## <a id="prerequisites"></a>⚙️ Prerequisites

To run this application natively or build it from source, you must have the following installed on your Windows machine:

1. **Python 3.10+**
2. **Microsoft Word (Desktop App)** (Required for the COM Automation Splitter, Converter, and Diff Engine)
3. **Microsoft Outlook (Desktop App)** (Required for the eMeeting Email Manager)
4. **Java Runtime Environment (JRE) 11+** (Required for the local PlantUML generation engine)
5. *(Optional but Recommended)* **Microsoft Visio** (To view generated `.vsdx` files)
6. *(Optional but Recommended)* **Microsoft PowerPoint** (For `.pptx` conversions)

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
*Note: This installs `PyQt5`, `requests`, `python-docx`, `beautifulsoup4`, `openpyxl`, `pandas`, `plotly`, `networkx`, `lxml`, and `pywin32`.*

### 3. Launch the Application
```bash
python src/main_tools.py
```
*Upon first launch, the app will automatically download the latest `plantuml.jar` from GitHub if it is not present in your assets folder.*

---

## <a id="usage"></a>📖 How to Use the GUI

### 🔎 3GPP Specification Full-Text & Evolution Search
1. Navigate to the **🔎 Spec Search** tab.
2. **Importing Specifications:**
   * Click **📥 Import from Specs DB** to select specification releases from `3gpp_data.db` (e.g., `24.501`, `38.331`, `23.501`, `33.501`). Missing files will download and unzip from the 3GPP FTP archive automatically.
   * Click **📁 Import Local .docx** to ingest single or split `.docx` files directly from your disk.
3. **Executing Substring Searches:**
   * Type any exact phrase, substring, or keyword into the main search bar (e.g., `"initial registration"`, `"PDU session"`, `"disaster roaming"`). Search queries with 3 or more characters automatically leverage the $O(1)$ FTS5 trigram index.
   * Optionally enter a clause number in the **Filter clause** field (e.g., `5.2`, `8.2.4`) to narrow results to specific sections.
4. **Tracking Text Introduction & Evolution:**
   * Review the **Release Evolution Matrix**. Each row represents a matching clause, with columns arranged in chronological release order.
   * Look for 🟢 **`🟢 Added`** markers to identify the exact specification version where that text was first introduced into the standard.
   * Look for 🔴 **`✗ Removed`** markers to identify where text present in previous releases was deprecated or deleted.
5. **Inspecting Matching Clause Content:**
   * Click any row or cell in the matrix to load the complete clause text into the **Clause Content Inspector**.
   * All occurrences of your search query will be highlighted in yellow within the full context of the clause.

### 🔬 3GPP Protocols Evolution Matrix (NAS & ASN.1 / RRC / NGAP)
1. Navigate to the **🔬 Protocols** (or **🔬 NAS**) tab.
2. **Importing Specifications:**
   * Click **📥 Import from Specs DB** to select specification releases across **TS 38.331 (NR RRC)**, **TS 36.331 (LTE RRC)**, **TS 38.413 (NGAP)**, **TS 24.501 (5GS NAS)**, or **TS 24.301 (EPS NAS)**. Missing versions download and convert automatically from the 3GPP FTP archive[cite: 2].
   * Click **📁 Import Local .docx** to ingest local single-file or multi-part split specification documents directly.
3. **Selecting Releases & Messages:**
   * Use the **Specification Versions & Releases** tree to activate, deactivate, or right-click to delete specific releases or entire specification series.
   * Select a Message, SIB, or PDU from the list (e.g., `RRCReconfiguration`, `SIB1`, `REGISTRATION REQUEST`). The **Evolution Matrix** pivots all Information Elements and unrolls nested ASN.1 sequence fields (e.g. `└─ radioBearerConfig`), color-coding additions (🟢), removals (🔴), and modifications (🟡).
4. **Filtering Fields and Descriptions:**
   * Use **Filter message/SIB name** to search message titles.
   * Use **Filter by IE / Field** to isolate specific parameters across the matrix.
   * Click the **`📖 Desc`** button to toggle extended description search, matching keywords located deep inside Clause 9 IE definitions and ASN.1 field description tables.
5. **Inspecting Structure & Reverse Lookup:**
   * Click any row in the matrix to render its Clause 9 coding diagram or ASN.1 syntax block and Field Descriptions table in the bottom **Inspector**.
   * Click the **Used in: N messages ▾** badge in the inspector header (or right-click any row in the matrix) to find all other messages referencing that parameter across active releases.

### 🗄️ Database Maintenance & Compaction
1. Click the **🗄️ Database** button located in the bottom system bar next to Task Manager and Proxy.
2. The dialog displays all SQLite database files (`3gpp_data.db`, `3gpp_protocol_data.db`, `3gpp_spec_search.db`), their current on-disk sizes, and Write-Ahead Log (`-wal`) statuses.
3. Click **Compact** on an individual database or **🧹 Compact All Databases** to flush WAL logs, execute SQLite `VACUUM`, optimize indices, and instantly reclaim free disk space.

### 📊 3GPP Meetings & Specifications
1. Navigate to the **Meetings** tab.
2. Click **Sync All Meetings** to trigger the 3-Phase scraper. You can also use **Open Last Meeting** to instantly resume your previous working group session.
3. Use the **Global TDoc Search** input to instantly find a specific document. Type a valid TDoc number (e.g., `S2-2605740`), and press **Enter** (or click **📄 Doc**) to fetch and open it immediately, or click **🗓️ Mtg** to launch its parent meeting table.
4. Right-click any meeting to access its FTP folders, view its info, or open its cached **TDocs List**.
5. In the TDocs Window, use the **Search** bar or dropdown filters to find specific documents. Double-click any cell to open the Notes editor and assign a color-coded status to a document.
6. For SA2 meetings, use the **Refresh** menu to import `TdocsByAgenda.htm` and automatically merge secretary remarks and on-the-fly revisions into your list.
7. Click the Action column to automatically download, unzip, and open the `.doc` files, or use the **⚖️ Add to Comparison Cart** submenu to select base versions or revisions for diffing.
8. Under the Specifications tab, use **🎯 Quick Fetch** to surgically inject single specifications or series into the database without a full sync.

### 📋 3GPP Work Items (WIs)
1. Navigate to the **3GPP Work Items** tab.
2. Click the **🔄 Sync 3GPP WIs** button (hover over it for tooltip details) to trigger the parallel multi-threaded scraper across all 19 Technical Specification Groups and Working Groups.
3. Monitor the real-time progress bar and status messages as records are fetched and bulk upserted into the shared database.
4. Use the **Local Search** bar and multi-select **Checkable Dropdowns** to debounce-filter the table by Acronym, Name, Code, Release, or Working Group. Your selected filters are automatically saved and restored between application sessions.
5. **Interactive Columns:** Click any blue **Latest WID** hyperlink to download the document via the global search engine (or fall back to the 3GPP Web Portal). Click the interactive **💬 Remarks** button to view a chronologically sorted history of secretary remarks for that specific work item.

### 📧 eMeeting Email Manager
1. Open a specific meeting from the main database and click the yellow **📧 Emails** button.
2. Click **⚙️ Folders** to browse your Outlook directory and safely map your Source (Inbox) and Target (Archive) folders.
3. Click **🔄 Sync Source** to download and index all emails for this meeting.
4. Select a TDoc thread from the **Left Panel** to view its chronological email history in the **Right Panel**.
5. Use the **⭐ Star** and **👀 Follow** buttons in the reading pane to track specific documents or entire topics. Use the left-side filters to isolate these threads, and the right-side dropdowns to filter by Company or Sender strictly within a thread.
6. Select rows and click **➡️ Move Selected** (or **⏭️ Move All**) to organize emails into dynamic Agenda Item subfolders inside your Outlook archive.
7. Click **📊 Statistics** to generate and open an interactive visual analytics dashboard of the meeting's email traffic.
8. Click any blue Sender Name in the grid to open an email window directly to them, or click a blue Revision number to download and open that document in Word.

### 📝 Slicing & Comparing Word Documents
1. In the **Comparison Cart** at the bottom of the Meetings Tab, sequentially select documents. The round-robin queue will automatically populate Slot A and Slot B with local files or fetched 3GPP Revisions.
2. Click **Compare in Word**. The tool will spawn a background process, temporarily remove file locks, and present a native Word redline comparison.
3. For large specs, navigate to the **Spec Splitter** tab, drag a `.docx` file, choose a Heading depth (e.g., "Level 2" for clauses like `6.1`, `6.2`), and click Split.

### 🎨 Visio Tools
1. **PlantUML Editor:** Type standard PlantUML code into the left pane. The Live Preview will automatically update the image on the right.
2. Click **Export Diagram ▼** and select **To Visio (.vsdx)** to generate a native Visio file, or use other options like PowerPoint, SVG, or ASCII.
3. **Batch Process & PowerPoint Conversion:** Navigate to the **📂 Visio Tools** tab and drag-and-drop `.puml`, `.txt`, or `.pptx` (PowerPoint) files into the drop zone. The system will detect the file type and process it into an editable Visio file in the background.

### ⚙️ Configuring Corporate Proxies & Networking
If you are behind a corporate firewall:
1. Glance at the **bottom right status bar** to see your active network status (Public Internet vs. 3GPP Local Network).
2. Click the **Network Config** button in the Console Panel.
3. Enter your HTTP/HTTPS proxies into the global session without restarting the app.
4. Adjust the **Humanness Delays** to throttle network requests (to mimic human behavior) or set them to 0.0 for maximum download speed.