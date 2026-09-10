# 📊 3GPP Meeting Tools & Diagram Converter

An advanced, component-based desktop IDE designed to bridge the gap between text-based diagramming (`PlantUML`) and corporate enterprise environments (`Microsoft Visio` and `PowerPoint`). 

Built specifically with telecommunications and 3GPP standards workflows in mind, this tool allows you to write highly efficient PlantUML sequence, activity, and network diagrams, instantly export them as fully editable native Office shapes, rapidly slice massive specification documents into manageable chapters, track NAS, ASN.1 (RRC / NGAP), GTP-U, and PFCP (TS 29.244) protocol message evolutions, search arbitrary substrings across specification releases using FTS5 trigram indexing with "First Added" and cutoff date detection, conduct multi-meeting company contribution audits with interactive KPI analytics and co-signer tracking, manage local SQLite databases with built-in compaction tools, track emails across working groups linked to specific TDocs and their revision families, connect seamlessly to local LLMs via Ollama for zero-cost offline intelligence and automated call-flow diagram synthesis, and navigate, filter, and synchronize the vast 3GPP meeting, specification, and work item archives locally[cite: 5, 6, 7].

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

### 🦙 Local AI & LLM Integration (Ollama Core Service)
* **100% Offline & Private Intelligence:**
  * **Confidentiality & Zero Cost:** Connects directly to a locally hosted Ollama daemon (`http://127.0.0.1:11434`)[cite: 1, 7]. Sensitive company draft positions, unreleased Change Requests (CRs), and meeting notes never leave your machine or cross external APIs.
  * **Airplane & Meeting-Room Ready:** Operates completely offline during flights and face-to-face 3GPP sessions without requiring external internet connectivity[cite: 7].
* **3GPP Call Flow Sequence Generation:**
  * Directly transforms textual procedure steps (e.g., from TS 23.502 or TS 38.300) into standard PlantUML sequence diagrams[cite: 5, 7].
  * Integrates real-time token streaming and prompt hot-reloading within Visio Tools[cite: 1, 5, 6].
* **Persistent Status Bar Pill & Health Monitor:**
  * **Interactive Status Widget:** Permanent status button integrated directly into the bottom status bar displaying active connectivity and model selection (🟢 `🦙 <model_name>` when connected, 🔴 `🦙 Offline` when disconnected)[cite: 1, 7].
  * **Instant Socket Pre-Flight:** Worker thread performs sub-millisecond TCP socket pre-checks to `127.0.0.1:11434` prior to issuing HTTP requests, preventing UI stalls and timeout freezes when the Ollama server is stopped[cite: 1, 7].
  * **Log Transition Memoization:** Heartbeat monitoring runs silently in the background, only emitting Qt signals and writing informational log records when connection states or installed model inventories actually change[cite: 1, 7].
* **Dedicated AI Network Session & Corporate Proxy Bypassing:**
  * **Bypass Corporate Proxy Traps:** Routes through a dedicated AI session factory (`get_ai_session()`) equipped with RFC 1918 loopback detection (`is_local_address()`), guaranteeing local prompts never get routed into or blocked by enterprise proxies (e.g., Zscaler, BlueCoat)[cite: 1, 7].
  * **Zero Scraper Delays:** AI sessions are completely decoupled from 3GPP FTP scraper sessions, bypassing humanness sleep delays to stream inference tokens with maximum throughput[cite: 1, 7].
* **Centralized Configuration & Dynamic Model Switching:**
  * One-click configuration dialog (`OllamaConfigDialog`) to update host URLs, switch between proxy modes (*Direct/Bypass*, *Auto-Detect*, *App Proxy*), test connectivity, and hot-swap default models (e.g., `qwen2.5:14b`, `llama3.1:8b`, `deepseek-r1:14b`) with settings persisted across sessions in `ollama_config.json`[cite: 1, 7].

---

### 🎨 Visio Tools (PlantUML & PowerPoint Converter)
* **Live Preview IDE:** A PlantUML code editor featuring syntax highlighting, line numbering, and a 500ms debounced live-rendering engine[cite: 2, 7].
* **🤖 AI Call Flow to PlantUML Sequence Generator:**
  * **Modeless, Multi-Window Experience:** Runs as an independent, non-blocking window (`Qt.Window`) with its own taskbar presence[cite: 5, 6]. Delegates can browse specifications, search meetings, and edit diagrams simultaneously while generation takes place[cite: 5, 6, 7].
  * **Real-Time Token Streaming & Cancellation:** Leverages chunked HTTP streaming (`stream_chat`) from local Ollama instances to stream code into the preview editor as it is synthesized, with an instant cancel button to abort runaway outputs[cite: 1, 5].
  * **3GPP Template & Skin Harmony:** Automatically strips markdown wrapper fences and enforces project-standard 3GPP lifeline styling, monochrome skin parameters, and box padding from `plantuml_templates.py`[cite: 4, 5].
  * **Dynamic Hot-Reload Prompts:** Prompt instructions are decoupled into external configuration files (`config/prompts/callflow_system.txt` and `config/prompts/callflow_user.txt`)[cite: 5, 8]. Modifications made in external text editors are detected via file timestamps (`mtime`) and reloaded instantly on the next inference click without restarting the application[cite: 5].
  * **One-Click Editor Integration:** Provides dedicated buttons to **📥 Replace Editor** or **➕ Append to Editor** to insert generated sequence code directly into your active workspace[cite: 5, 6].
* **Batch Conversion Engine:** Drag and drop hundreds of `.puml`, `.txt`, or `.pptx` files to queue them for multi-threaded background conversion[cite: 2, 7].
* **PowerPoint to Visio Pipeline:** Seamlessly convert entire PowerPoint presentations into multi-page Visio documents (`.vsdx`)[cite: 2, 7]. Uses Enhanced Metafile (EMF) bridging to perfectly preserve editable native Office shapes, automatically aggressively ungroup them, and shrink wrap their text boundaries[cite: 7].
* **Custom Visio Stencil Engine:** Converts standard PlantUML shapes into grouped Visio shapes (`.vsdx`) mapped directly to custom 3GPP node stencils[cite: 7].

---

### 🔎 3GPP Specification Full-Text & Substring Search Engine
* **FTS5 Trigram Substring Search & Chronological Tracking:**
  * **Arbitrary Substring Matching:** Powered by an embedded SQLite Full-Text Search (FTS5) engine configured with a 3-character Trigram tokenizer (`tokenize="trigram"`)[cite: 7]. Enables near-instantaneous search for exact phrases, field substrings, acronyms, or protocol constants across millions of words without full-table scan delays[cite: 7].
  * **Targeted Release & Clause Filtering:** Filter queries by specific clause patterns (e.g., `5.2`, `8.1.4`, `Annex A`) or execute cross-specification queries across all active releases simultaneously[cite: 7].
* **Release Evolution Matrix & "First Added" Text Tracking:**
  * **Per-Specification Tabbed Matrix Visualization:** Automatically isolates search results into dedicated per-specification tabs (e.g., `TS 23.501 (32)`, `TS 23.502 (20)`)[cite: 7]. This prevents sparse empty matrices, eliminates colliding clause numbers, and preserves clean chronological column ordering per document[cite: 7].
  * **"First Added" Identification:** Automatically determines the exact earliest release where matching text was introduced, rendering clear visual indicators[cite: 7]:
    * 🟢 **`🟢 Added`**: Highlighted in soft green to indicate the exact version where text first appeared in that clause[cite: 7].
    * ⚪ **`✓ Present`**: Retained and present in subsequent releases[cite: 7].
    * 🔴 **`✗ Removed`**: Highlighted in soft red when text present in a previous release was deleted in that version[cite: 7].
    * ➖ **`-`**: Clause not matching or not present in that release[cite: 7].
* **Date Cutoff Analysis:**
  * **Official Release Date Storage:** Tracks official 3GPP portal upload and publication dates across all indexed specification releases[cite: 7].
  * **Post-Cutoff Date Additions Filter:** Toggle the **🎯 Date Cutoff** selector to highlight text introduced after a target date[cite: 7]:
    * ⚡ **`⚡ Post-Cutoff Added`**: Highlighted in soft amber/yellow to clearly identify text additions introduced after cutoff dates[cite: 7].
    * **Exclusive Filter Mode:** Check **Show Only Post-Cutoff Additions** to hide clauses where matching text was already present prior to the selected priority date (filtering out prior art)[cite: 7].
* **Universal Specification Ingestion Dialog:**
  * **Unrestricted Catalog Access:** Master-detail browser spanning all ~1,500+ specifications across Series 01 through 55 and all Working Groups (RAN1–4, SA1–6, CT1–4)[cite: 7].
  * **Live Search & Presets:** Filter by keyword, topic, or specification number with built-in quick presets for core 3GPP specifications[cite: 7].
  * **Explicit Checkbox Selection:** Dedicated checkbox column for unambiguous selection tracking with dynamic count badges (`Selected: N version(s)`)[cite: 7].
  * **Smart Batch Selectors:** One-click helpers including **`⚡ Select All Unindexed`**, **`⭐ Select Latest per Release`** (supporting both decimal and 3-digit lettered versions like `i40` / `g30`), **`☑️ Select All`**, and **`◻️ Clear`**[cite: 7].
  * **Revision-Mark Filtering:** Automatically discards 3GPP Word change-mark files (`-rm` / `_rm`) during unzipping and local imports, ensuring only clean (`-cl`) specification text is indexed[cite: 7].
* **Multi-Part Split Document Parsing:**
  * **Split Document Sequencing:** Automatically detects, sequences, and unifies modern multi-part specification archives (e.g., `_s00_s04.docx`, `_s05_s08.docx`, `_s09_s14.docx`) into a single consolidated release model in SQLite[cite: 7].
  * **High-Performance XML Extraction:** Direct `lxml` parsing extracts document structure directly from OpenXML without Microsoft Word COM runtime overhead[cite: 7].
  * **Intelligent Heading & TOC Sanitization:** Distinguishes genuine 3GPP headings from numbered procedure call flow steps and strips Table of Contents (TOC) dot leaders and stub entries[cite: 7].
* **Rich Clause Content Inspector:**
  * **💡 Key Match Excerpt:** Dedicated callout banner at the top of the inspector displaying the exact matching paragraph and surrounding sentence context[cite: 7].
  * **Interactive Term Navigation:** **`[ ◀ Prev ]`** and **`[ Next ▶ ]`** buttons with an active match counter (`🎯 N Match(es)`) and auto-scroll to match locations[cite: 7].
  * **One-Click Citation Copy:** **`[ 📋 Copy Citation ]`** button formats and copies complete clause text with official 3GPP document headers, versions, and release dates[cite: 7].
* **Persistent Search Configuration:**
  * Active search queries, clause filters, cutoff dates, and exact checked specification versions are automatically persisted to `spec_search_config.json` and restored across application sessions[cite: 7].
* **Non-Blocking Database Maintenance:**
  * Background `SpecSearchWipeWorker` thread enables fast, freeze-free database resets with automatic checkpointing and SQLite schema reconstruction[cite: 7].

---

### 🔬 3GPP Protocol Evolution Matrix & Inspector (NAS, ASN.1 & PFCP)
* **Comprehensive Multi-Protocol Ingestion:**
  * **NAS Protocols:** Complete support for **5GS NAS (TS 24.501)** and **EPS NAS (TS 24.301)**[cite: 7].
  * **ASN.1 Protocols:** Native support for **NR RRC (TS 38.331)**, **LTE RRC (TS 36.331)**, and **NGAP (TS 38.413)**[cite: 7].
  * **GTP-U Protocol:** Native support for **GTPv1-U (TS 29.281)** covering user plane tunnels across 5GS (`N3`, `N9`, `N19`, `F1-U`, `Xn-U`, `W1-U`), EPS (`S1-U`, `X2-U`, `S5/S8`), and legacy interfaces[cite: 7]. Parses signalling messages (Echo Request/Response, Error Indication, End Marker, Tunnel Status) and synthesizes G-PDUs with Clause 5.2 Extension Headers (PDU Session Container, NR RAN Container, PDU Set Information Container) with per-interface filtering[cite: 7].
  * **PDU Session User Plane Protocol:** Native support for **TS 38.415** (PDU Session & PDU Set Information User Plane Protocols) over `NG-U`, `Xn-U`, `F1-U`, and `N9` interfaces[cite: 7]. Parses DL/UL PDU Session Information and DL PDU Set Information frames into fine-grained bit-level fields, tracking delay measurement flags, QoS Monitoring timestamps, and PDU Set sequence numbering across releases[cite: 7].
  * **PFCP Protocol:** Native support for **PFCP (TS 29.244)** spanning both 5GC (`N4`, `N4mb`) and EPC (`Sxa`, `Sxb`, `Sxc`) reference points[cite: 7]. Parses top-level Node-Related (Clause 7.4) and Session-Related (Clause 7.5) PDU messages alongside the master Information Element Type registry (Table 8.1.2-1)[cite: 7].
  * **Hierarchical Grouped IE Unrolling:** Recursively traverses and unrolls nested PFCP Grouped IEs (e.g., `Create PDR └─ PDI └─ SDF Filter`, `Create FAR`, `Create URR`, `Usage Report`) into the Evolution Matrix with tree indentation and depth tracking[cite: 7].
  * **Interface Applicability Metadata & Filtering:** Automatically indexes per-IE interface applicability tags (`Sxa`, `Sxb`, `Sxc`, `N4`, `N4mb`) and provides a dynamic UI filter dropdown (`[All Interfaces | N4 | N4mb | Sxa | Sxb | Sxc]`) that appears whenever a PFCP message is selected[cite: 7].
  * **Release 20+ Multi-Part Document Ingestion:** Automatically detects, sequences, and parses modern split 3GPP specifications by aggregating all clause sub-documents into a single unified release model[cite: 7].
  * **Automated Legacy `.doc` Conversion:** Automatically converts older binary Word 97–2003 `.doc` specifications to `.docx` via headless COM automation or LibreOffice with Protected View bypass and NTFS Zone Identifier unblocking before parsing[cite: 7].
  * **High-Performance XML Parsing:** Direct `lxml` extraction parses message definition tables, ASN.1 syntax blocks, PFCP Grouped IE tables, and field description tables directly from `.docx` archives without requiring Word runtime overhead[cite: 7].
* **Evolution Matrix & Visual Diffing:**
  * **Hierarchical Sequence & Group Unrolling:** Recursively unrolls nested ASN.1 sequence/choice fields and PFCP grouped structures, allowing you to track high-level and deep parameter changes across releases simultaneously[cite: 7].
  * **Visual Release Diffing:** Color-coded matrix cells immediately highlight field additions (🟢 Green), removals (🔴 Red), and format/type modifications (🟡 Yellow) between chronological 3GPP releases[cite: 7].
  * **Hierarchical Specification Tree:** Interactive tree view (`QTreeWidget`) grouping releases under collapsible specification parents (`TS 38.331`, `TS 29.244`, `TS 24.501`, `TS 24.301`, etc.) with master toggles, specification-wide selection, and right-click deletion context menus[cite: 7].
  * **Persistent Filter Configuration:** Tree expansion states, active releases, message selections, and search terms are automatically persisted in `nas_config.json` across sessions[cite: 7].
* **Dual-Layer & Extended Description Search:**
  * **Debounced Filtering:** Dedicated 250ms debounced search bars for Message Names and Information Elements / Fields[cite: 7].
  * **Extended Description Search (`📖 Desc`):** Toggle deep text search across underlying Clause 8/9 IE descriptions and ASN.1 field description tables (e.g., searching `"emergency"` or `"slicing"` highlights matching messages even if the keyword is not in the field name)[cite: 7].
* **Structure & Field Descriptions Inspector:**
  * **Bit-Level Structure Rendering:** Renders bit-level octet diagrams (Figure 8.x/9.x) and value coding tables (Table 8.x/9.x) with full OpenXML `gridSpan` (colspan) and `vMerge` (rowspan) support[cite: 7].
  * **ASN.1 Syntax & Descriptions View:** Renders syntax-highlighted ASN.1 definition blocks accompanied by formatted 3GPP Field Description tables[cite: 7].
  * **Reverse Field/IE Lookup:** Interactive header badge (`Used in: N messages ▾`) and right-click matrix context menu to trace and jump to all messages referencing a given IE or ASN.1 type across active releases[cite: 7].

---

### 🗄️ Database Maintenance & Compaction
* **System-Level Database Manager:** Integrated `🗄️ Database` tool accessible directly from the bottom system bar[cite: 7].
* **Automatic Freelist & WAL Compaction:** Inspects on-disk database sizes and active Write-Ahead Logs (`-wal`), executing `PRAGMA wal_checkpoint(TRUNCATE)` and `VACUUM` to defragment pages and reclaim megabytes of disk space after heavy scraping or database wiping[cite: 7].
* **Batch Maintenance:** One-click **Compact All Databases** to optimize `3gpp_data.db`, `3gpp_protocol_data.db`, `3gpp_spec_search.db`, and auxiliary cache databases concurrently[cite: 7].

---

### 📡 3GPP Meeting, Specification & Work Items Database
* **Asynchronous Three-Phase Syncing Engine:** 
  * **Phase 1 (FTP Directory Mapping):** Scrapes the 3GPP FTP archives in parallel to instantly populate your database with all available meeting numbers, gracefully handling hidden RAN Ad-Hoc (`TSGR_AHs`) subdirectories[cite: 7].
  * **Phase 2 (Deep Document Scrape):** Crawls the `Docs/` folder of every meeting[cite: 7]. Uses smart regex stripping to ignore file extensions and revisions, mathematically sorting the files to determine the first and last TDocs of the meeting[cite: 7].
  * **Phase 3 (DynaReport Upserting):** Injects metadata (Location, Start/End Dates, Ad-Hoc/Electronic status) by fetching the legacy 3GPP Portal HTML tables[cite: 7].
* **Targeted Quick Fetch:** Instantly sync individual specifications (e.g., `23.801-01`) or entire specification series (e.g., `23`) directly from the FTP server without needing to run a lengthy full database sync[cite: 7].

* **3GPP Work Items (WIs) Synchronizer:**
  * **Parallel Multi-WG Scraper:** Concurrently scrapes active Work Items across all 19 Technical Specification Groups and Working Groups (SA, SA1-6, RAN, RAN1-6, CT, CT1-6) from official 3GPP dynamic report pages using multi-threaded execution (5 workers)[cite: 7].
  * **High-Performance Bulk Upsert:** Utilizes atomic SQLite bulk transactions (`executemany` with `ON CONFLICT DO UPDATE`) to instantly sync thousands of work items and map them to their respective working groups via relational sidecar tables (`work_items`, `wi_group_map`, `wi_remarks`)[cite: 7].
  * **Interactive UI Tab:** Features a dedicated tab with a real-time progress bar, status feedback, and helpful button tooltips[cite: 7]. Includes debounced multi-select CheckableComboBox filters (Release, WG) with persistent state-saving, chronologically sorted historical remarks via a custom interactive UI bubble, and clickable WID hyperlinks that automatically route through the global TDoc fetcher or 3GPP Portal[cite: 7].

* **3GPP Work Items (WIs) & Specification Linkage:**
  * **Relational Mapping (`spec_wi_map`):** Bi-directionally maps 3GPP Specifications to Work Items during Pass 2 DynaReport scraping without requiring rigid locks on un-synced WIs[cite: 7].
  * **Specification Inspector Chips:** Details dialogs display interactive primary (⭐) and secondary Work Item chips with direct 3GPP portal navigation[cite: 7].
  * **Work Items Table & Local Specs Inspector:** The Work Items tab features dedicated **WG** and **Linked Specs** columns, local specification inspectors (`LinkedSpecsDialog`), and one-click citation copy actions[cite: 7].

* **Intelligent TDocs Manager:**
  * **Smart Global TDoc Search:** Instantly locate and download any document across the entire database[cite: 7]. Just type a TDoc number (e.g., `S2-2605740r11`) and the UI will dynamically reveal minimalist quick-actions to download the specific file or open its parent meeting context—all without leaving the main dashboard[cite: 7].
  * **Persistent Personal Notes & Status (Sidecar Database):** Keep a private, local SQLite database that "overlays" your data onto the 3GPP list[cite: 7]. Double-click any TDoc to assign a color-coded status (🟢 Support, 🔴 Object, 🟡 Monitor) and save personal notes[cite: 7]. Your data survives perfectly even when downloading fresh 3GPP Excel updates[cite: 7].
  * **Smart Revision Inheritance:** When a TDoc gets a new revision during a meeting, the new child document automatically inherits a "Ghost" version of the personal notes and status you assigned to the base document[cite: 7]!
  * **Interactive Secretary Remarks:** TDocs mentioned in the Secretary Remarks are automatically identified and converted into hyperlinks[cite: 7]. Left-click a link to instantly jump to that row (intelligently wiping active filters if necessary), or right-click to download it or add it to your Comparison Cart[cite: 7].
  * **Natural Sorting & Smart Filtering:** Bulletproof multi-select dropdowns and natural numerical sorting for complex multi-level Agenda Items (e.g., AI 20.6.2 sorts correctly before 20.6.11)[cite: 7].
  * **Comprehensive Analytics Dashboards:** Generate interactive offline HTML Plotly reports detailing TDoc outcomes, top contributing companies, and complex strategic alliance network graphs (co-signing clusters) using Louvain community detection algorithms[cite: 7].
  * **SA2 Electronic Revisions & Agenda Parsing:** Automatically parses `TdocsByAgenda.htm` to extract comments, inject on-the-fly revisions directly into your table, and provides a "No Comments Only" filter[cite: 7]. For eMeetings, it automatically scrapes the `INBOX/Revisions/` FTP folder[cite: 7].
  * **SA2 Chairman's Notes & Session List Ingestion (`.doc` / `.docx`):**
    * **Frosted Drop Overlay:** Drag and drop `.doc`, `.docx`, `.htm`, or `.html` session documents onto the TDocs window; a visual frosted-blue drop overlay appears with dashed borders and instant drop targets[cite: 7].
    * **Non-Blocking Background Worker (`WordAgendaImporterThread`):** Copies imported files to `{meeting_dir}/Agenda/`, unblocks NTFS Zone Identifiers, converts legacy macro-bearing `.doc` files via headless LibreOffice, and parses table data in the background without freezing the UI[cite: 7].
  * **Multi-Action Resources Menu:** Instantly jump to local cache directories, fetched HTML Agenda files, Main FTP folders, Docs/ folders, or Revisions/ folders directly from the UI[cite: 7].
  * **Quick Launch History:** Remembers your active working group session, allowing you to bypass the database table and jump back into your last opened meeting with a single click[cite: 7].

* **📊 Cross-Meeting Company Contribution Audit & Reporting:**
  * **Multi-Meeting Filtering Engine:** Filter and audit contributions across arbitrary date ranges, multiple Working Groups (e.g., SA2, SA1, RAN2), specific 3GPP Work Items, and canonical company entities[cite: 7].
  * **Canonical Company Sanitization & Co-Author Rules:** Integrates directly with `CompanySanitizer` to normalize diverse corporate subsidiaries and brand spellings[cite: 7]. Easily toggle between **"Any co-author (joint contribution)"** and **"Primary source only (first listed)"**[cite: 7].
  * **Work Item (WI) Auto-Completing Token Box:** Search and add multiple Work Item acronyms/codes using a tag/chip input backed by `WorkItemsDatabase` autocompletion[cite: 7].
  * **High-Performance Bounded Concurrency & O(1) Cache:** Operates a laptop-safe `ThreadPoolExecutor` capped at 3 worker threads to prevent CPU throttling, memory bloat, and 3GPP server rate-limiting[cite: 7]. Bypasses expensive full-table regex scanning by resolving `Related WIs` directly from 3GPP document sheets and agenda descriptions[cite: 7]. Employs `TDocsParser` sidecar JSON caches (`.xlsx.json`) for sub-second cache hits on previously opened meetings, with a one-click toggle to **"Bypass local cache (re-fetch from 3GPP)"**[cite: 7].
  * **Modeless, Multi-Window Experience:** Operates as an independent, non-blocking top-level window (`Qt.Window`), enabling full simultaneous access to the main meetings tab, individual meeting tables, and email inspectors[cite: 7].
  * **Clean Application Aesthetic:** Renders in neutral monochrome text without distracting colored cells or heavy badge backgrounds, matching the application's clean design theme[cite: 7].
  * **Dedicated Working Group (WG) & Abstract Columns:** Dedicated **WG** column at the start of the table for instant committee attribution, and an **Abstract** column with multi-line tooltip wrapping and popup reading support[cite: 7].
  * **Direct Row Navigation & Context Menu:**
    * **Smart Double-Click:** Double-clicking `WG` or `Meeting` launches that meeting's full table in `TDocsWindow`[cite: 7]; double-clicking `Abstract` opens the clean pop-up reader[cite: 7]; double-clicking `TDoc` (or any other cell) downloads (if missing), extracts, and opens the actual document in Word via `TDocActionThread`[cite: 7].
    * **Right-Click Command Menu:** Context actions to **🗓️ Open Meeting Table**[cite: 7], **📂 Open Docs Folder** (opens the meeting's public FTP `Docs/` directory in your browser)[cite: 7], **📄 Open TDoc Document** (downloads, extracts, and launches the Word document directly)[cite: 7], **🌐 View on 3GU Portal**[cite: 7], **🔗 Copy URL** (copies direct ZIP download link to clipboard)[cite: 7], **📋 Copy TDoc / Title**[cite: 7], and **📝 View Abstract...**[cite: 7].
  * **In-App KPI Summary Bar:** Real-time summary metric cards displaying **Total Contributions** (with WG coverage count), **Agreed / Approved** (total count and percentage of total submissions)[cite: 7], **Collaboration** (% joint vs. solo)[cite: 7], **Top Co-signer** (most frequent partner organization with joint count)[cite: 7], and **Top Work Item** (accounting for multi-entry WIs while filtering out placeholders like `DUMMY` and `Unspecified`)[cite: 7].
  * **Interactive HTML Analytics Dashboard (`📊 Statistics Dashboard...`):** Generates an offline Plotly HTML report with native integer formatting featuring[cite: 7]:
    * **Monthly Contribution Trend:** Stacked bar chart of outcomes grouped chronologically by Meeting End Date / Month (`YYYY-MM`)[cite: 7].
    * **Overall Contribution Outcomes:** Donut chart with accurate counts and percentage tooltips using semantic 3GPP colors (Agreed/Approved in green, Not Treated in grey, Revised in indigo, Noted in amber, etc.)[cite: 7].
    * **Top Co-signers:** Horizontal bar chart displaying partner organizations segmented by Working Group in stacked colors[cite: 7].
    * **Working Group Activity:** Stacked bar chart showing total contributions per Working Group categorized by outcome status[cite: 7].
    * **Work Item & Study Item Allocation:** Ranking of top 20 active Work Items, splitting multi-entry rows and filtering placeholders[cite: 7].
  * **Corporate Deutsche Telekom Magenta Excel Export:** Exports the aggregated, filtered view into a styled `.xlsx` workbook featuring zebra striping, centered alignments, and active download hyperlinks pointing directly to each TDoc's specific meeting Docs folder[cite: 7].

* **Fast Native Wi-Fi & 3GPP Network Detection:**
  * Automatically detects when you are connected to the official "3GPPWIFI" network during live standardization meetings[cite: 7].
  * **Zero-Subprocess In-Memory Detection:** Queries Windows Network List Manager (NLM) via native COM (`netprofm.dll`) and `netsh wlan`, resolving active SSIDs in ~2ms with zero shell overhead or AMSI security locks.
  * **Instant TCP Server Reachability:** Performs non-blocking TCP socket probes directly to the internal meeting server (`10.10.10.10:80/443`), bypassing slow `ping.exe` subprocesses.
  * **Decoupled Status Display:** Persistent status indicator in the bottom system bar displaying `🟢 <SSID> (Local Server Active)` or `🌐 <SSID>`[cite: 7]. Enables dynamic local caching, proxy bypassing, and direct FTP routing through the high-speed local meeting network[cite: 7].

* **3GPP FTP Session Manager:** Automatically injects randomized User-Agents and HTTP Keep-Alive headers[cite: 7]. Features a configurable **Humanness Delay** engine to bypass aggressive 3GPP server throttling and "Too Many Requests" blocks, which can be dialed down to 0.0 for maximum scraping speed[cite: 7].

* **🔒 Secure Corporate Proxy Profiles & DPAPI Credential Store:**
  * **Full Profile Encryption:** Protects corporate network topology and credentials by encrypting all proxy parameters (HTTP/HTTPS hosts, ports, Active Directory usernames, and passwords) into an isolated payload blob using Windows DPAPI (`core.utils.dpapi.SecureCredentialStore`)[cite: 7].
  * **Zero-Plaintext On-Disk Invariant:** Plaintext hostnames, IP addresses, usernames, and passwords are permanently banned from `network_config.json`[cite: 7]. If encryption fails or DPAPI is unavailable, the engine refuses to persist unencrypted credentials to disk[cite: 7].
  * **Authenticode & Windows Security Catalog Integrity:** Dynamically verifies that `crypt32.dll` and `win32crypt` are genuine Microsoft system binaries using `WinVerifyTrust` with Windows Security Catalog (`CatRoot`) fallback, 64-bit module address checks in `System32`, and workspace anti-shadowing guards[cite: 7].
  * **Application-Specific Entropy Salt:** Applies a secondary cryptographic salt so that other scripts running under the same Windows account cannot by chance decrypt the profile blob without applying the exact application salt[cite: 7].
  * **Instant Profile Toggle:** One-click activation/deactivation enables seamless switching between corporate proxy routing and direct public internet without discarding saved credentials[cite: 7].

---

### 📧 Universal TDoc Email Tracker & Inspection Dialog
* **Working Group-Agnostic Ingestion:** Indexes emails across any 3GPP Working Group (SA2, RAN2, CT1, etc.) directly from your Outlook folders without moving emails or touching server-side folders[cite: 7]. Operates independently of the dedicated eMeeting logic to prevent regressions[cite: 7].
* **WG-Dependent Multi-Folder Profiles & Custom Tag Colors:**
  * Configure specific Outlook folders per Working Group (saved globally in `emails_config.json`)[cite: 7].
  * Assign custom tags (e.g., `[WG]`, `[Disc]`, `[Offline]`, `[Inbox]`) and pick personalized badge colors using an interactive `QColorDialog`[cite: 7]. Tags render in the conversation stream with custom contrasting colors[cite: 7].
* **Smart Quotation Boundary & Direct Message Detection:**
  * Differentiates whether a TDoc was cited in the **Subject**, the **Direct Body** of the message, or an inherited historical reply chain (**Quoted**)[cite: 7].
  * Eliminates false-positive cascades where casual replies (`"ok, danke"`, `"+1"`) cite TDocs buried in older email footers[cite: 7].
  * Toggle **`☑️ Include Quoted Matches`** to hide or reveal conversational thread citations on demand[cite: 7].
* **Exchange Internal Senders & DMARC Resolution:**
  * Automatically resolves listserv rewrites (`LIST.ETSI.ORG`) and internal Exchange X.500 addresses (`/o=...` / `EX`) to primary SMTP addresses to ensure company sanitization recognizes internal colleagues[cite: 7].
* **Modeless, Multi-Window Architecture:**
  * The inspection dialog operates as an independent, modeless top-level window (`Qt.Window`)[cite: 7]. It never freezes or blocks the main TDocs list or background downloads, allowing you to snap windows side-by-side[cite: 7].
  * Multiple TDocs can be inspected concurrently without duplicate window spawning[cite: 7].
* **Interactive TDoc Linkifier:**
  * Automatically converts every detected 3GPP TDoc number in the Subject line, Match Excerpt banner, and Body text into a clickable link[cite: 7].
  * Current document family numbers are highlighted in amber (`#FFF176`), while cross-referenced TDocs appear with interactive links (e.g., `🔗 S2-2608457`)[cite: 7].
  * Clicking any referenced TDoc instantly launches an inspection window for that document[cite: 7].
* **Reading Pane Controls & Standalone Viewer:**
  * **Interactive Vertical Splitter:** Drag the splitter bar between the email list and the reading pane to adjust viewing proportions[cite: 7].
  * **Unicode Whitespace Compression:** Automatically strips invisible non-breaking spaces (`\xa0`) and collapses excessive blank lines from Word/Outlook formatting into clean, readable text[cite: 7].
  * **💡 Match Found Callout:** Displays an excerpt banner directly above the message showing the exact surrounding sentence context where the TDoc was found[cite: 7].
  * **`⧉ Pop Out View`:** Detaches the message preview into an independent, fully resizable viewer (`StandaloneEmailReaderWindow`), ideal for laptop screens or secondary monitors[cite: 7].
* **Read / Unread Lifecycle & Ignore Engine:**
  * Track local read states in SQLite (`general_emails.is_read`)[cite: 7].
  * Selecting an email marks it as read after an 800ms debounce[cite: 7].
  * Multi-select rows with `Ctrl` or `Shift` to batch Mark Read, Mark Unread, Ignore, or Delete[cite: 7].
  * **`🚫 Ignore` Action:** Suppresses high-volume distribution list announcements or rapporteur compilation emails from all document counts without deleting them[cite: 7]. Ignored flags are preserved across re-syncs[cite: 7]. Toggle **`Show Ignored`** to review or un-ignore them[cite: 7].
* **TDocs Window Integration:**
  * **`Emails` Column:** Displays aggregate family email counts with unread badges (e.g., `✉️ 5 (🔵 2)`)[cite: 7].
  * **Context Menu:** Right-click any row to view related emails or toggle all emails for that TDoc's revision family between read and unread[cite: 7].
  * **`📧 Emails ▾` Header Menu:** One-click menu to sync related emails, configure folders, mark all as read, or execute a high-speed wipe of the generic emails database[cite: 7].

---

### 📧 eMeeting Email Manager (Dedicated SA2 eMeeting Dashboard)
* **High-Performance Sync Engine:** Connects directly to your local Microsoft Outlook via COM automation[cite: 7]. Pulls, parses, and indexes thousands of eMeeting mailing list emails in milliseconds using SQLite chunked batching (`executemany`) with zero memory spikes[cite: 7].
* **Master-Detail Thread Architecture:** Bypasses broken Outlook reply chains by logically grouping emails purely by parsed TDoc numbers[cite: 7]. The UI features a split-screen design: a Left Panel displaying active TDoc threads and a Right Panel displaying the isolated, chronological conversation for the selected topic[cite: 7].
* **Intelligent 3GPP Parser:** Uses smart regex to extract TDoc numbers (6-8 digits), Agenda Items, Revisions, and free text directly from standard 3GPP bracketed subject lines and email bodies[cite: 7].
* **DMARC Listserv Bypass:** Automatically detects when 3GPP mailing lists rewrite the sender address to `LIST.ETSI.ORG`[cite: 7]. It parses the actual sender's name and email address from the email body and maps them to known telecommunication companies[cite: 7].
* **Advanced Dual-Layer Filtering:** 
  * **Macro-Filters (Thread Level):** Use Star (⭐) and Follow (👀) buttons, or the global search bar, to instantly filter the left-hand thread list down to specific topics or Agenda Items of interest[cite: 7].
  * **Micro-Filters (Conversation Level):** Once a thread is selected, use the Company dropdown, Sender dropdown, or Text search boxes to isolate specific replies strictly within that single conversation[cite: 7].
* **Interactive Email Analytics:** Click the **Statistics** button to instantly generate an interactive, offline HTML Plotly dashboard visualizing Agenda Item volumes, company activity rankings, timeline histograms, and top delegate leaderboards[cite: 7].
* **Automated Archiving:** Safely extracts physical `.msg` files to your hard drive and dynamically builds a clean target folder hierarchy in Outlook (e.g., `Archive/SA2_175/9.1.1/`) to permanently organize your inbox[cite: 7].

---

### 📝 Word Document Manipulation & AI Integration
* **🤖 AI/LLM Corpus Exporter:**
  * **Smart Automation:** Automatically downloads missing TDocs from the 3GPP FTP and extracts the underlying Word documents in the background[cite: 7].
  * **Intelligent Parsing:** Uses a custom Regex State Machine to handle complex 3GPP formatting, including extracting Track Changes and parsing tricky "all new text" placeholder clauses (e.g., `6.4.5.X`)[cite: 7].
  * **Mega-File Compilation:** Compiles and groups the extracted text into clean, Agenda Item-specific Markdown files tailored specifically for LLM context windows (Gemini, Claude, GPT, Ollama)[cite: 7].
* **Global Comparison Cart:** A persistent, round-robin state dashboard that bridges multiple meeting windows[cite: 7]. Intelligently push any Base TDoc or specific Revision into alternating slots, then launch a native Word comparison instantly[cite: 7].
* **Isolated Word Diff Engine:** Uses COM `DispatchEx` to spawn an invisible, isolated instance of Microsoft Word[cite: 7]. It safely opens files as Read-Only, auto-accepts tracked changes purely in RAM, assigns proper document names for the comparison pane, and generates a visual diff without freezing your active Word sessions or locking local files[cite: 7].
* **LibreOffice Integration Engine:**
  * **Macro-Free & Sandboxed Conversion:** Built-in adapter leveraging headless LibreOffice with isolated user profiles (`-env:UserInstallation`) to suppress network printer hangs and bypass macro security restrictions[cite: 7].
  * **Installed & Portable Support:** Seamless auto-detection of system-installed LibreOffice and single-click integration for portable distributions (`LibreOfficePortable.exe`)[cite: 7].
* **Corporate IT Bypass (Sensitivity Labels):** Automatically injects configurable Microsoft Purview Sensitivity Labels (e.g., "OFFEN") directly into COM objects to bypass blocking corporate IT popup dialogs during automated saves[cite: 7].
* **Intelligent DocxSplitter:** Safely slices massive 3GPP TS/TR specifications into individual Word documents based on Heading 1 or Heading 2 boundaries, perfectly preserving styles, images, and Visio objects[cite: 7].
* **Background Word-to-PDF Converter:** A headless Word automation thread that silently converts generated files to PDFs or XPS without interrupting your workflow[cite: 7].
* **Native Visio Extractor:** Parses the raw XML (`document.xml`) of a `.docx` file, identifies embedded `OLEObject` bins, and extracts raw `.vsdx` Visio diagrams straight out of the Word document to your local disk[cite: 7].

---

## <a id="architecture"></a>🏗️ Architecture & Data Flow

This application strictly adheres to the **Model-View-Controller (MVC)** and **Event-Driven Architecture (EDA)** paradigms using `PyQt5`[cite: 7]. 

1. **The UI Layer (`src/modules/*/ui/` & `src/main_window.py`):** Contains Qt Widgets, `QAbstractTableModel` implementations, modeless top-level inspection dialogs (`CallFlowDialog`), and status bar permanent action widgets[cite: 5, 7]. The UI never performs synchronous network I/O or blocks the main event loop[cite: 7].
2. **The Core Layer (`src/core/` & `src/modules/*/core/`):** Contains domain logic[cite: 7]. All database transactions (`sqlite3` with FTS5 trigrams), REST AI communications (`core/ai/ollama_client.py`)[cite: 1], prompt management with timestamp hot-reloading (`modules/puml2visio/core/prompt_manager.py`)[cite: 5], FTP network scraping (`requests`), COM automation (`win32com` & `pythoncom`), headless LibreOffice conversions, and direct XML manipulation (`lxml` & `python-docx`) are isolated here[cite: 7].
3. **The Threading Bridge:** Independent worker tasks inherit from `QThread` (e.g., `OllamaMonitorThread`, `CallFlowGeneratorThread`, `WifiMonitorThread`, `GeneralEmailSyncThread`, `WordAgendaImporterThread`, `ContributionSearchWorker`, `TDocsDownloaderThread`, `LLMExporterThread`)[cite: 1, 5, 7]. Workers operate completely decoupled with their own error boundaries and communicate with the main thread strictly through thread-safe `pyqtSignals`[cite: 1, 5, 7].
4. **The Singleton Managers & Security Utilities:** Global network state (`NetworkState`), specialized AI HTTP session creation (`session.get_ai_session()`), cryptographic credential protection (`core.utils.dpapi.SecureCredentialStore`), and Comparison Cart states are managed by thread-safe singletons and dynamic JSON config loaders[cite: 1, 7].

---

## <a id="prerequisites"></a>⚙️ Prerequisites

To run this application natively or build it from source, you must have the following installed on your Windows machine[cite: 7]:

1. **Python 3.10+**[cite: 7]
2. **Microsoft Word (Desktop App)** (Required for native COM Automation Splitter, Converter, and Diff Engine)[cite: 7]
3. **Microsoft Outlook (Desktop App)** (Required for the eMeeting and General Email Managers)[cite: 7]
4. **Java Runtime Environment (JRE) 11+** (Required for the local PlantUML generation engine)[cite: 7]
5. *(Optional but Recommended for Local AI)* **Ollama** (Install from [ollama.ai](https://ollama.ai) to enable private, offline document summaries, call-flow to sequence conversion, and AI assistants)[cite: 5, 7].
6. *(Optional but Recommended)* **LibreOffice (Installed or Portable)** (Required for safe, macro-free conversion of legacy Word 97–2003 `.doc` files, including SA2 Chairman's Notes and older specifications[cite: 7]. If using portable LibreOffice, link `LibreOfficePortable.exe` using the **📂 Locate Executable** button in the Word Tools tab[cite: 7].)
7. *(Optional)* **Microsoft Visio** (To view and edit generated `.vsdx` files)[cite: 7]
8. *(Optional)* **Microsoft PowerPoint** (For `.pptx` to `.vsdx` conversions)[cite: 7]

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
*Note: This installs `PyQt5`, `requests`, `python-docx`, `beautifulsoup4`, `openpyxl`, `pandas`, `plotly`, `networkx`, `lxml`, and `pywin32`.*[cite: 7]

### 3. Launch the Application
```bash
python src/main_tools.py
```
*Upon first launch, the app will automatically download the latest `plantuml.jar` from GitHub if it is not present in your assets folder.*[cite: 7]

---

## <a id="usage"></a>📖 How to Use the GUI

### 🦙 Configuring Ollama & Local LLMs
1. **Status Bar Quick Glance:** Observe the **bottom right status bar**[cite: 7]. The pill will display `🦙 <selected_model>` in green if Ollama is running, or `🦙 Offline` in soft red if the server is stopped[cite: 1, 7].
2. **Launching Settings:** Click the `🦙` button in the status bar to open the **Ollama LLM Configuration** dialog[cite: 7].
3. **Connection & Testing:**
   * Enter your Ollama server URL (defaults to `http://127.0.0.1:11434`)[cite: 1, 7].
   * Select your **Proxy Routing Mode**:
     * **Direct (Bypass Proxy - Recommended):** Ensures loopback calls never get routed into corporate firewalls or proxy walls[cite: 1, 7].
     * **Auto-Detect:** Automatically uses direct connections for local/private IPs and the application proxy for external domains[cite: 7].
     * **Route via App Proxy:** Enforces routing through your configured corporate proxy profile[cite: 7].
   * Click **🔌 Test** to verify connectivity and refresh your list of locally installed models[cite: 7].
4. **Selecting Default Model:** Choose your preferred model from the **Active Model** dropdown (e.g., `qwen2.5:7b`, `llama3.1:8b`) and click **Save**[cite: 1, 7]. The status bar pill updates immediately[cite: 1, 7].

---

### 🎨 Visio Tools (PlantUML & PowerPoint Converter)
1. **PlantUML Editor:** Type standard PlantUML code into the left pane[cite: 2, 7]. The Live Preview will automatically update the image on the right[cite: 7].
2. **Generating Sequence Diagrams from Call Flow Text (AI):**
   * On the editor toolbar, click **🤖 Generate from Call Flow...**[cite: 6].
   * The generator dialog opens as an independent, modeless window[cite: 5]. You can continue working in other tabs or searching specifications while it remains open[cite: 5, 7].
   * Paste 3GPP procedural steps (e.g., from TS 23.502 or TS 38.300) into the left text box[cite: 5].
   * Select your installed Ollama model and target diagram type (default: *Sequence*)[cite: 5].
   * Click **🚀 Generate Diagram**[cite: 5]. The model will stream PlantUML syntax token-by-token directly into the preview area[cite: 1, 5].
   * Once finished, click **📥 Replace Editor** to overwrite the editor contents or **➕ Append to Editor** to insert the code into your existing diagram[cite: 5, 6].
   * *(Prompt Engineering)* You can customize generation prompts without restarting the app by editing `config/prompts/callflow_system.txt` or `config/prompts/callflow_user.txt` in any text editor[cite: 5, 8]. The tool reloads updated prompt files automatically upon clicking generate[cite: 5].
3. **Exporting Diagrams:** Click **Export Diagram ▼** and select **To Visio (.vsdx)** to generate a native Visio file, or use other options like PowerPoint, SVG, or ASCII[cite: 2, 7].
4. **Batch Process & PowerPoint Conversion:** Navigate to the **📂 Visio Tools** tab and drag-and-drop `.puml`, `.txt`, or `.pptx` (PowerPoint) files into the drop zone[cite: 2, 7]. The system will detect the file type and process it into an editable Visio file in the background[cite: 2, 7].

---

### 🔎 3GPP Specification Full-Text & Evolution Search
1. Navigate to the **🔎 Spec Search** tab[cite: 7].
2. **Importing Specifications:**
   * Click **📥 Import from Specs DB** to open the universal specification browser[cite: 7]. Select any 3GPP document (Series 01–55) or filter by Working Group[cite: 7]. Missing archives download and extract from the 3GPP FTP server automatically[cite: 7].
   * Use **`⚡ Select All Unindexed`** or **`⭐ Select Latest per Release`** to batch-select versions with checkboxes[cite: 7].
   * Click **📁 Import Local .docx** to ingest single or multi-part split documents (`_s00_s04.docx`, `_s05_s08.docx`) directly from your drive[cite: 7].
3. **Executing Substring Searches:**
   * Type any exact phrase or keyword into the search bar (e.g., `"slice replacement"`, `"ATSSS"`, `"emergency"`)[cite: 7]. Search queries with 3 or more characters automatically execute across the FTS5 trigram index[cite: 7].
   * Optionally enter a clause number in the **Filter clause** field (e.g., `5.2`, `4.3.2`) to focus on specific sections[cite: 7].
4. **Date Cutoff & "First Added" Text Analysis:**
   * Review the **Release Evolution Matrix** displayed in per-specification tabs (e.g., `TS 23.501 (32)`, `TS 23.502 (20)`)[cite: 7].
   * Toggle **🎯 Date Cutoff** and select a cutoff date[cite: 7]. Text introduced after that date will be highlighted with ⚡ **`⚡ Post-Cutoff Added`**[cite: 7].
   * Check **Show Only Post-Cutoff Additions** to filter out older prior art and show only clauses containing post-cutoff date modifications[cite: 7].
5. **Inspecting Matching Clause Content:**
   * Click any cell in the matrix to load the clause into the **Clause Content Inspector**[cite: 7].
   * The **💡 Key Match Excerpt** callout at the top highlights the matching paragraph with surrounding sentence context[cite: 7].
   * Use **`[ ◀ Prev ]`** and **`[ Next ▶ ]`** to cycle between match occurrences in long clauses[cite: 7].
   * Click **`[ 📋 Copy Citation ]`** to copy the formatted text with 3GPP document, version, and release date metadata directly to your clipboard[cite: 7].

---

### 🔬 3GPP Protocols Evolution Matrix (NAS, ASN.1 & PFCP)
1. Navigate to the **🔬 Protocols** (or **🔬 NAS**) tab[cite: 7].
2. **Importing Specifications:**
   * Click **📥 Import from Specs DB** to select specification releases across **TS 38.331 (NR RRC)**, **TS 36.331 (LTE RRC)**, **TS 38.413 (NGAP)**, **TS 29.244 (PFCP)**, **TS 24.501 (5GS NAS)**, or **TS 24.301 (EPS NAS)**[cite: 7]. Missing versions download and convert automatically from the 3GPP FTP archive[cite: 7].
   * Click **📁 Import Local .docx** to ingest local single-file or multi-part split specification documents directly[cite: 7].
3. **Selecting Releases & Messages:**
   * Use the **Specification Versions & Releases** tree to activate, deactivate, or right-click to delete specific releases or entire specification series[cite: 7].
   * Select a Message, SIB, or PDU from the list (e.g., `PFCP Session Establishment Request`, `RRCReconfiguration`, `SIB1`, `REGISTRATION REQUEST`)[cite: 7]. The **Evolution Matrix** pivots all Information Elements and unrolls nested ASN.1 sequence fields or PFCP Grouped IEs (e.g., `Create PDR └─ PDI └─ SDF Filter`), color-coding additions (🟢), removals (🔴), and modifications (🟡)[cite: 7].
4. **Filtering Fields, Descriptions & Interfaces:**
   * Use **Filter message/SIB name** to search message titles[cite: 7].
   * Use **Filter by IE / Field** to isolate specific parameters across the matrix[cite: 7].
   * Click the **`📖 Desc`** button to toggle extended description search, matching keywords located deep inside Clause 8/9 IE definitions and ASN.1 field description tables[cite: 7].
   * For PFCP messages, use the **Interface Selector Dropdown** (`All Interfaces`, `N4`, `N4mb`, `Sxa`, `Sxb`, `Sxc`) positioned above the matrix table to instantly filter parameters by target reference point[cite: 7].
5. **Inspecting Structure & Reverse Lookup:**
   * Click any row in the matrix to render its Clause 8/9 coding diagram (bit-level octet diagram) or ASN.1 syntax block and Field Descriptions table in the bottom **Inspector**[cite: 7].
   * Click the **Used in: N messages ▾** badge in the inspector header (or right-click any row in the matrix) to find all other messages referencing that parameter across active releases[cite: 7].

---

### 🗄️ Database Maintenance & Compaction
1. Click the **🗄️ Database** button located in the bottom system bar next to Task Manager and Proxy[cite: 7].
2. The dialog displays all SQLite database files (`3gpp_data.db`, `3gpp_protocol_data.db`, `3gpp_spec_search.db`), their current on-disk sizes, and Write-Ahead Log (`-wal`) statuses[cite: 7].
3. Click **Compact** on an individual database or **🧹 Compact All Databases** to flush WAL logs, execute SQLite `VACUUM`, optimize indices, and instantly reclaim free disk space[cite: 7].

---

### 📊 3GPP Meetings & Specifications
1. Navigate to the **Meetings** tab[cite: 7].
2. Click **Sync All Meetings** to trigger the 3-Phase scraper[cite: 7]. You can also use **Open Last Meeting** to instantly resume your previous working group session[cite: 7].
3. Use the **Global TDoc Search** input to instantly find a specific document[cite: 7]. Type a valid TDoc number (e.g., `S2-2605740`), and press **Enter** (or click **📄 Doc**) to fetch and open it immediately, or click **🗓️ Mtg** to launch its parent meeting table[cite: 7].
4. Right-click any meeting to access its FTP folders, view its info, or open its cached **TDocs List**[cite: 7].
5. In the TDocs Window, use the **Search** bar or dropdown filters to find specific documents[cite: 7]. Double-click any cell to open the Notes editor and assign a color-coded status to a document[cite: 7].
6. **Importing SA2 Session Documents & Chairman's Notes:**
   * **Drag & Drop:** Drag any `.docx`, `.doc`, or `.htm` session document anywhere onto the TDocs window[cite: 7]. A visual frosted drop overlay will highlight the window[cite: 7].
   * **Menu Import:** Alternatively, click the **🔄 Refresh** menu and select **📝 Import Word Document (.docx / .doc)...**[cite: 7].
   * The file is automatically copied to `{meeting_dir}/Agenda/`, converted in the background via LibreOffice (if `.doc`), parsed, and merged into the table without freezing the UI[cite: 7].
7. Click the Action column to automatically download, unzip, and open documents, or use the **⚖️ Add to Comparison Cart** submenu to select base versions or revisions for diffing[cite: 7].
8. Under the Specifications tab, use **🎯 Quick Fetch** to surgically inject single specifications or series into the database without a full sync[cite: 7].

---

### 📊 Cross-Meeting Company Contribution Audit & Reporting
1. Navigate to the **Meetings** tab and click **📊 Contributions Report...**[cite: 7].
2. The audit dialog opens as a modeless, non-blocking window, allowing you to use other app tabs and meeting windows concurrently[cite: 7].
3. **Configuring Search Filters:**
   * **Working Groups:** Select one or multiple Working Groups via the checkable dropdown (e.g., `SA2`, `SA1`, `RAN2`)[cite: 7].
   * **Date Range:** Define the meeting start and end date boundaries[cite: 7].
   * **Work Items:** Type a WI acronym or code (e.g., `FS_6G_ARC`, `AmbientIoT_Ph2-ARC`) into the token box and press Enter or click **Add**[cite: 7]. Autocompletion guides your input using the `WorkItemsDatabase`[cite: 7]. Multiple tokens can be added as removable chips[cite: 7].
   * **Target Companies:** Filter and check companies from the list populated from `CompanySanitizer`[cite: 7].
   * **Author Match Rule:** Select **Any co-author** to capture joint contributions, or **Primary source only** to match documents where the company is the first listed submitter[cite: 7].
   * **Bypass Cache:** Check **Bypass local cache (re-fetch from 3GPP)** to force downloading fresh TDoc lists from the 3GPP portal[cite: 7].
4. Click **🔍 Search Contributions**[cite: 7]. The background engine searches matching meetings concurrently (max 3 worker threads), leveraging O(1) sidecar JSON caches to display results without freezing your PC[cite: 7].
5. **Reviewing Results & In-App KPIs:**
   * Review the real-time KPI strip: **Total Contributions** (with WG coverage count), **Agreed / Approved** (total count and percentage of total submissions)[cite: 7], **Collaboration** (% joint vs. solo)[cite: 7], **Top Co-signer**, and **Top Work Item**[cite: 7].
   * Inspect the neutral, uncolored results table containing `WG`, `Meeting`, `TDoc`, `Title`, `Source`, `Matched Company`, `Type`, `For`, `Agenda Item`, `TDoc Status`, `Related WIs`, and `Abstract`[cite: 7].
6. **Row Actions & Direct Navigation:**
   * **Open Meeting Table:** Double-click the **WG** or **Meeting** cell (or right-click and select **🗓️ Open Meeting Table**) to launch that meeting's full table in `TDocsWindow`[cite: 7].
   * **Open TDoc Document:** Double-click the **TDoc** cell (or right-click and select **📄 Open TDoc Document**) to automatically download, extract, and open the document directly in Microsoft Word via `TDocActionThread`[cite: 7].
   * **Open Docs Folder:** Right-click and choose **📂 Open Docs Folder** to open the meeting's public FTP documents directory in your browser[cite: 7].
   * **View on 3GU Portal:** Right-click and select **🌐 View on 3GU Portal** to open the official portal page for that TDoc[cite: 7].
   * **Copy URL:** Right-click and choose **🔗 Copy URL** to copy the direct ZIP download link to your clipboard[cite: 7].
   * **Read Abstract:** Double-click the **Abstract** cell (or right-click and select **📝 View Abstract...**) to open the popup text reader[cite: 7].
7. **Exporting Data & Analytics:**
   * Click **📥 Export to Excel...** to generate a styled Deutsche Telekom magenta `.xlsx` spreadsheet with active download hyperlinks pointing directly to each TDoc's specific meeting Docs folder[cite: 7].
   * Click **📊 Statistics Dashboard...** to compile and open an interactive offline HTML Plotly report displaying monthly contribution trends, overall decision outcomes, top co-signers segmented by Working Group, Working Group activity breakdowns, and Work Item allocations[cite: 7].

---

### 📧 Tracking Related Emails for TDocs (Universal Meeting Support)
1. **Configuring Folders & Tag Colors:**
   * In any open TDocs window, click the **📧 Emails ▾** header menu and select **⚙️ Configure Outlook Folders...**[cite: 7].
   * Click **➕ Add Folder via Outlook...** to browse and map your Working Group distribution list folders (e.g., `SA2_WG`, `SA2_DISC`, `RAN2_List`)[cite: 7].
   * Enter a short Tag (e.g., `WG`, `Disc`, `Offline`) and click the color button to assign a distinct visual badge color using the color picker[cite: 7]. Configurations are saved globally per Working Group[cite: 7].
2. **Syncing Outlook Emails:**
   * Click **📧 Emails ▾ $\rightarrow$ 🔄 Sync Related Emails...**[cite: 7].
   * Confirm the date range (defaults to meeting start/end dates $\pm 3$ days buffer) and click **🚀 Start Sync**[cite: 7].
   * The background engine indexes all mentions of TDocs in both Subject lines and Message bodies without downloading physical `.msg` files[cite: 7].
3. **Inspecting TDoc Conversation Threads:**
   * Review the **Emails** column in the main TDocs table[cite: 7]. Cells display total family counts and blue unread badges (e.g., `✉️ 4 (🔵 2)`)[cite: 7].
   * Double-click any cell in the **Emails** column (or right-click a row and select **📧 View Related Emails...**) to open the modeless inspection dialog[cite: 7].
   * **Family Breadcrumbs:** The top card displays the complete document revision lineage (e.g., `S2-2601000 ➔ S2-2601234 ➔ S2-2601555`)[cite: 7].
   * **Quotation Filter:** Uncheck **Include Quoted Matches** to filter out reply chains that only mentioned the TDoc in historical quoted text[cite: 7].
4. **Navigating & Reading Emails:**
   * Drag the interactive **vertical splitter** to expand the reading pane[cite: 7].
   * Click **⧉ Pop Out View** to detach the reading pane into an independent viewer window (`StandaloneEmailReaderWindow`), ideal for laptop screens or secondary monitors[cite: 7].
   * **Interactive TDoc Links:** Every 3GPP document number cited in the Subject line, Match Excerpt banner, or Body is rendered as an interactive link[cite: 7]. Click any link (e.g., `🔗 S2-2608457`) to open that document's related emails immediately[cite: 7].
   * Click **🚀 Open in Outlook** to view the original message live in native Microsoft Outlook[cite: 7].
5. **Managing Read & Ignored Statuses:**
   * Selecting an email automatically marks it as read[cite: 7].
   * Select multiple rows using `Ctrl` or `Shift` to batch **Mark Read**, **Mark Unread**, **Ignore**, or **Delete**[cite: 7].
   * **`🚫 Ignore`:** Suppresses high-volume distribution list announcements or rapporteur compilation emails from badge counts across all referenced TDocs without deleting them from the database[cite: 7].
   * Right-click any row in the main TDocs table to mark all emails for that document family as read or unread in one click[cite: 7].
   * To reset generic meeting email records, click **📧 Emails ▾ $\rightarrow$ 🗑️ Wipe Generic Emails Database...**[cite: 7].

---

### 📋 3GPP Work Items (WIs)
1. Navigate to the **3GPP Work Items** tab[cite: 7].
2. Click the **🔄 Sync 3GPP WIs** button (hover over it for tooltip details) to trigger the parallel multi-threaded scraper across all 19 Technical Specification Groups and Working Groups[cite: 7].
3. Monitor the real-time progress bar and status messages as records are fetched and bulk upserted into the shared database[cite: 7].
4. Use the **Local Search** bar and multi-select **Checkable Dropdowns** to debounce-filter the table by Acronym, Name, Code, Release, or Working Group[cite: 7]. Your selected filters are automatically saved and restored between application sessions[cite: 7].
5. **Interactive Columns:** Click any blue **Latest WID** hyperlink to download the document via the global search engine (or fall back to the 3GPP Web Portal)[cite: 7]. Click the interactive **💬 Remarks** button to view a chronologically sorted history of secretary remarks for that specific work item[cite: 7].

---

### 📧 eMeeting Email Manager (SA2 Electronic Sessions)
1. Open a specific electronic meeting from the database, click the **📧 Emails ▾** menu, and choose **📊 Open eMeeting Email Manager (Dashboard)**[cite: 7].
2. Click **⚙️ Folders** to browse your Outlook directory and safely map your Source (Inbox) and Target (Archive) folders[cite: 7].
3. Click **🔄 Sync Source** to download and index all eMeeting emails[cite: 7].
4. Select a TDoc thread from the **Left Panel** to view its chronological email history in the **Right Panel**[cite: 7].
5. Use the **⭐ Star** and **👀 Follow** buttons in the reading pane to track specific documents or entire topics[cite: 7]. Use the left-side filters to isolate these threads, and the right-side dropdowns to filter by Company or Sender strictly within a thread[cite: 7].
6. Select rows and click **➡️ Move Selected** (or **⏭️ Move All**) to organize emails into dynamic Agenda Item subfolders inside your Outlook archive[cite: 7].
7. Click **📊 Statistics** to generate and open an interactive visual analytics dashboard of the meeting's email traffic[cite: 7].

---

### 📝 Slicing & Comparing Word Documents
1. In the **Comparison Cart** at the bottom of the Meetings Tab, sequentially select documents[cite: 7]. The round-robin queue will automatically populate Slot A and Slot B with local files or fetched 3GPP Revisions[cite: 7].
2. Click **Compare in Word**[cite: 7]. The tool will spawn a background process, temporarily remove file locks, and present a native Word redline comparison[cite: 7].
3. For large specs, navigate to the **Spec Splitter** tab, drag a `.docx` file, choose a Heading depth (e.g., "Level 2" for clauses like `6.1`, `6.2`), and click Split[cite: 7].

---

### ⚙️ Configuring Corporate Proxies & Networking
If you are behind a corporate firewall:
1. **Network Status Indicator:** Check the **bottom right status bar** to observe active network detection (Public Internet vs. 3GPP Local Network)[cite: 7].
2. **Configuring Secure Proxy Profiles:**
   * Click the **📡 Proxy** button in the Console Panel[cite: 7].
   * Check **Activate this proxy profile** to route outgoing HTTP/HTTPS traffic through the proxy[cite: 7].
   * Enter your HTTP and HTTPS proxy addresses (e.g., `proxy.company.com:8080`)[cite: 7]. Use the **Use the same proxy address for HTTPS** checkbox to sync them automatically[cite: 7].
   * For corporate networks requiring authentication (HTTP 407), enter your username (`Domain\Username` or standard `Username`) and password[cite: 7].
   * Ensure **Save proxy profile securely (Windows DPAPI)** is checked[cite: 7]. All parameters—including hosts, ports, username, and password—are encrypted via `core.utils.dpapi.SecureCredentialStore` into an unreadable binary blob before saving to `network_config.json`[cite: 7]. No plaintext proxy topology or credentials ever touch the disk[cite: 7].
   * Click **⚡ Test Connection** to verify end-to-end connectivity to `www.3gpp.org` via a non-blocking background thread (`ProxyTestWorker`)[cite: 7]. Sensitive credentials are automatically redacted from error popups and logs[cite: 7].
   * Click **Save & Apply** to apply the configuration immediately to the active `NetworkSession`[cite: 7].
3. **Deactivating Corporate Proxy (Direct Mode):**
   * When disconnecting from the corporate network, click **📡 Proxy** and select **Deactivate (Direct)**[cite: 7]. The session immediately routes directly to the public internet while retaining your encrypted profile for easy reactivation[cite: 7].
4. **Scraper Humanness & Request Delays:**
   * Click the **⚙️ Network** button in the Console Panel[cite: 7].
   * Adjust **Min Delay** and **Max Delay** to throttle network requests to mimic human browsing behavior, or set them to `0.0` for maximum download speed[cite: 7].
   * Toggle **Rotate Modern User-Agents** to cycle browser headers across requests[cite: 7].

---

## <a id="troubleshooting"></a>🛠️ Known Quirks / Troubleshooting

* **Corporate IT "Aktion blockiert" on Drag & Drop:**
  * If Windows Defender Attack Surface Reduction (ASR) blocks dragging downloaded `.doc` files directly from your `Downloads` folder, either:
    1. Use the **🔄 Refresh $\rightarrow$ 📝 Import Word Document...** file picker menu[cite: 7].
    2. Unblock the file via Right Click $\rightarrow$ Properties $\rightarrow$ **Zulassen (Unblock)**[cite: 7].
* **Legacy Word 97–2003 Macro Permissions:**
  * Legacy `.doc` files containing VBA macros (like SA2 Chairman's Notes) are blocked by Word COM security settings[cite: 7]. Ensure LibreOffice is installed or point the app to portable LibreOffice (`LibreOfficePortable.exe`) in the Word tab to enable automated, macro-free conversion[cite: 7].
* **Sensitivity Label Dialogs (Microsoft Purview / Azure Information Protection):**
  * If automated Word conversions or comparisons trigger corporate classification popups, configure your default sensitivity label string (e.g., `OFFEN` or `INTERNAL`) in `word_config.json` to allow silent headless saves[cite: 7].
* **Ollama Connection Refused or Showing Offline:**
  * Verify that the local daemon is active by running `ollama list` in PowerShell or Windows Terminal.
  * If running Ollama on another machine on your local network (e.g., a GPU workstation), ensure `OLLAMA_HOST=0.0.0.0` is set in the daemon environment and configure the target IP in the **🦙 Ollama Configuration** dialog[cite: 1, 7].
* **Tuning AI Diagram Output:**
  * If generated PlantUML diagrams omit expected participants or lifelines, tune the prompt templates in `config/prompts/callflow_system.txt`[cite: 5, 8]. The prompt engine automatically reloads changes without needing an application restart[cite: 5].