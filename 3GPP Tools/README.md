# 📊 3GPP Meeting Tools & Diagram Converter

An advanced, component-based desktop IDE designed to bridge the gap between text-based diagramming (`PlantUML`) and corporate enterprise environments (`Microsoft Visio` and `PowerPoint`). 

Built specifically with telecommunications and 3GPP standards workflows in mind, this tool allows you to write highly efficient PlantUML sequence, activity, and network diagrams, instantly export them as fully editable native Office shapes, rapidly slice massive specification documents into manageable chapters, and seamlessly navigate, filter, and synchronize the vast 3GPP meeting and specification databases locally.

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

### 📡 3GPP Meeting & Specifications Database
* **Asynchronous Three-Phase Syncing Engine:** * **Phase 1 (FTP Directory Mapping):** Scrapes the 3GPP FTP archives in parallel to instantly populate your database with all available meeting numbers, gracefully handling hidden RAN Ad-Hoc (`TSGR_AHs`) subdirectories.
  * **Phase 2 (Deep Document Scrape):** Crawls the `Docs/` folder of every meeting. Uses smart Regex stripping to ignore file extensions and revisions, mathematically sorting the files to determine the first and last TDocs of the meeting.
  * **Phase 3 (DynaReport Upserting):** Injects metadata (Location, Start/End Dates, Ad-Hoc/Electronic status) by fetching the legacy 3GPP Portal HTML tables.
* **Targeted Quick Fetch:** Instantly sync individual specifications (e.g., `23.801-01`) or entire specification series (e.g., `23`) directly from the FTP server without needing to run a lengthy full database sync.
* **Intelligent TDocs Manager:**
  * **Smart Global TDoc Search:** Instantly locate and download any document across the entire database. Just type a TDoc number (e.g., `S2-2605740r11`) and the UI will dynamically reveal minimalist quick-actions to download the specific file or open its parent meeting context—all without leaving the main dashboard.
  * **Natural Sorting & Smart Filtering:** Bulletproof multi-select dropdowns and natural numerical sorting for complex multi-level Agenda Items (e.g., AI 20.6.2 sorts correctly before 20.6.11).
  * **SA2 Electronic Revisions & Agenda Parsing:** Automatically scrapes `INBOX/Revisions/` for electronic meetings. Parses messy Word-exported `TdocsByAgenda.htm` files to extract comments, inject on-the-fly revisions directly into your table, and provides a "No Comments Only" filter.
  * **Multi-Action Resources Menu:** Instantly jump to local cache directories, fetched HTML Agenda files, Main FTP folders, Docs/ folders, or Revisions/ folders directly from the UI.
  * **Quick Launch History:** Remembers your exact active working group session, allowing you to bypass the database table and instantly jump back into your last opened meeting with a single click.
* **3GPP FTP Session Manager:** Automatically injects randomized User-Agents and HTTP Keep-Alive headers. Features a configurable **Humanness Delay** engine to bypass aggressive 3GPP server throttling and "Too Many Requests" blocks, which can be dialed down to 0.0 for maximum scraping speed.

### 📧 eMeeting Email Manager (Native Outlook Integration)
* **High-Performance Sync Engine:** Connects directly to your local Microsoft Outlook via COM automation. Pulls, parses, and indexes thousands of eMeeting mailing list emails in milliseconds using SQLite chunked batching (`executemany`) with zero memory spikes.
* **Intelligent 3GPP Parser:** Bypasses broken Outlook email threads by using smart regex to extract TDoc numbers (6-8 digits), Agenda Items, Revisions, and free text directly from standard 3GPP bracketed subject lines and email bodies.
* **DMARC Listserv Bypass:** Automatically detects when 3GPP mailing lists rewrite the sender address to `LIST.ETSI.ORG`. It parses the actual sender's name and email address from the email body and maps them to known telecommunication companies.
* **Smart Tracking & Focus Filters:** * **Star TDocs (⭐):** Highlight a specific TDoc to instantly group and track all past and future emails discussing it, neutralizing chaotic and broken reply chains.
  * **Follow AIs (👀):** Monitor entire Agenda Items (e.g., `9.1.1`) to easily filter topics of interest.
  * **Date Fencing:** Automatically restricts email scanning to the precise start and end dates of the meeting, jumping out of background loops early to drastically optimize sync speeds and prevent inbox pollution.
* **Automated Archiving:** Safely extracts physical `.msg` files to your hard drive and dynamically builds a clean target folder hierarchy in Outlook (e.g., `Archive/SA2_175/9.1.1/`) to permanently organize your inbox.

### 📝 Word Document Manipulation
* **Global Comparison Cart:** A persistent, round-robin state dashboard that bridges multiple meeting windows. Intelligently push any Base TDoc or specific Revision into alternating slots, then launch a native Word comparison instantly.
* **Isolated Word Diff Engine:** Uses COM `DispatchEx` to spawn an invisible, isolated instance of Microsoft Word. It safely opens files as Read-Only, auto-accepts tracked changes purely in RAM, assigns proper document names for the comparison pane, and generates a visual diff without freezing your active Word sessions or locking local files.
* **Corporate IT Bypass (Sensitivity Labels):** Automatically injects configurable Microsoft Purview Sensitivity Labels (e.g., "OFFEN") directly into COM objects to bypass blocking corporate IT popup dialogs during automated saves.
* **Intelligent DocxSplitter:** Safely slices massive 3GPP TS/TR specifications (often hundreds of pages long) into individual Word documents based on Heading 1 or Heading 2 boundaries, perfectly preserving styles, images, and Visio objects.
* **Background Word-to-PDF Converter:** A headless Word automation thread that silently converts generated files to PDFs or XPS without interrupting your workflow.
* **Native Visio Extractor:** Parses the raw XML (`document.xml`) of a `.docx` file, identifies embedded `OLEObject` bins, and extracts raw `.vsdx` Visio diagrams straight out of the Word document to your local disk.

### 🎨 PlantUML to Visio Converter
* **Live Preview IDE:** A code editor featuring syntax highlighting, line numbering, and a 500ms debounced live-rendering engine.
* **Batch Conversion Engine:** Drag and drop hundreds of `.puml` or `.txt` files to queue them for multi-threaded background conversion.
* **Custom Visio Stencil Engine:** Converts standard PlantUML shapes into grouped Visio shapes (`.vsdx`) mapped directly to custom 3GPP node stencils.

---

## <a id="architecture"></a>🏗️ Architecture & Data Flow

This application strictly adheres to the **Model-View-Controller (MVC)** and **Event-Driven Architecture (EDA)** paradigms using `PyQt5`. 

1. **The UI Layer (`modules/*/ui/`):** Contains only dumb Qt Widgets and standard `QAbstractTableModel` proxies. It never blocks the main thread.
2. **The Core Layer (`modules/*/core/`):** Contains the heavy lifting. All database transactions (`sqlite3`), FTP network scraping (`requests`), COM object automation (`win32com` & `pythoncom`), and XML manipulation (`python-docx`) are isolated here.
3. **The Threading Bridge:** Every Core module inherits from `QThread`. The UI sends data to the Thread, and the Thread emits `pyqtSignals` back to the UI to update progress bars or logs.
4. **The Singleton Managers:** The Network Configuration (proxies), Word Configuration (Sensitivity Labels), and Comparison Cart states are managed by robust Singletons and dynamic JSON config loaders to ensure cross-tab synchronization.

---

## <a id="prerequisites"></a>⚙️ Prerequisites

To run this application natively or build it from source, you must have the following installed on your Windows machine:

1. **Python 3.10+**
2. **Microsoft Word (Desktop App)** (Required for the COM Automation Splitter, Converter, and Diff Engine)
3. **Microsoft Outlook (Desktop App)** (Required for the eMeeting Email Manager)
4. **Java Runtime Environment (JRE) 11+** (Required for the local PlantUML generation engine)
5. *(Optional but Recommended)* **Microsoft Visio** (To view the generated outputs)

---

## <a id="installation"></a>🚀 Installation

### 1. Clone the Repository
```bash
git clone [https://github.com/your-repo/3GPP-Delegate-Helper.git](https://github.com/your-repo/3GPP-Delegate-Helper.git)
cd 3GPP-Delegate-Helper
```

### 2. Install Python Dependencies
```bash
pip install -r requirements.txt
```
*Note: This includes `PyQt5`, `requests`, `python-docx`, `beautifulsoup4`, `openpyxl`, and `pywin32`.*

### 3. Launch the Application
```bash
python src/main_puml2visio.py
```
*Upon first launch, the app will automatically attempt to download the latest `plantuml.jar` from GitHub if it is not present in your assets folder.*

---

## <a id="usage"></a>📖 How to Use the GUI

### 📊 3GPP Meetings & Specifications
1. Navigate to the **Meetings** tab.
2. Click **Sync All Meetings** to trigger the 3-Phase scraper. You can also use **Open Last Meeting** to instantly resume your previous working group session.
3. Use the **Global TDoc Search** input to instantly find a specific document. Type a valid TDoc number (e.g., `S2-2605740`), and press **Enter** (or click **📄 Doc**) to fetch and open it immediately, or click **🗓️ Mtg** to launch its parent meeting table.
4. Right-click any meeting to access its FTP folders, view its info, or open its cached **TDocs List**.
5. In the TDocs Window, use the **Search** bar or dropdown filters to find specific documents. 
6. For SA2 electronic meetings, use the **Refresh** menu to import `TdocsByAgenda.htm` and automatically merge secretary remarks and on-the-fly revisions into your list.
7. Click the Action column to automatically download, unzip, and open the `.doc` files, or use the **⚖️ Add to Comparison Cart** submenu to select base versions or revisions for diffing.
8. Under the Specifications tab, use **🎯 Quick Fetch** to surgically inject single specifications or series into the database without a full sync.

### 📧 eMeeting Email Manager
1. Open a specific meeting from the main database and click the yellow **📧 Emails** button.
2. Click **⚙️ Folders** to browse your Outlook directory and safely map your Source (Inbox) and Target (Archive) folders.
3. Click **🔄 Sync Source** to download and index all emails for this meeting.
4. Select rows and click **➡️ Move Selected** (or **⏭️ Move All**) to permanently organize the emails into dynamic Agenda Item subfolders inside your Outlook archive.
5. Use the **⭐ Star** and **👀 Follow** buttons in the reading pane to surgically track specific documents or entire topics, and use the top filter bar to instantly isolate them during chaotic sessions.
6. Click any blue Sender Name in the grid to open a new email window directly to them, or click a blue Revision number to automatically download and open that document in Word.

### 📝 Slicing & Comparing Word Documents
1. In the **Comparison Cart** at the bottom of the Meetings Tab, sequentially select documents. The round-robin queue will automatically populate Slot A and Slot B with local files or fetched 3GPP Revisions.
2. Click **Compare in Word**. The tool will spawn a background process, temporarily remove file locks and OS restrictions, and present you with a native Word redline document.
3. For large specs, navigate to the **Spec Splitter** tab, drag a `.docx` file, choose a Heading depth (e.g., "Level 2" for clauses like `6.1`, `6.2`), and click Split.

### 🎨 PlantUML Editor
1. Type standard PlantUML code into the left pane.
2. The Live Preview will automatically update the image on the right.
3. Click **Export Visio** to generate a native `.vsdx` file, or **Copy to Clipboard** to paste the image directly into PowerPoint.

### ⚙️ Configuring Corporate Proxies & Networking
If you are behind a corporate firewall:
1. Click the **Network Config** button.
2. Enter your HTTP/HTTPS proxies into the global session without restarting the app.
3. Adjust the **Humanness Delays** to throttle network requests (to mimic human behavior) or set them to 0.0 for maximum download speed.

---

## <a id="troubleshooting"></a>🛠️ Known Quirks / Troubleshooting

* **Missing / Ghost Meetings in DB:** If 3GPP deletes an old folder from their FTP server, Phase 1 cannot find it. However, the Phase 3 Upsert Engine will find the historical record on the DynaReport page and *force* it into your database anyway. These meetings will have a valid 3GPP Portal link, but their FTP links will lead to a 404.
* **The "3GPP Empty String" Crash:** The scraper uses advanced unicode normalization to handle 3GPP's notorious formatting errors (like using non-breaking hyphens in dates or completely omitting table columns). If a meeting appears with blank dates, try clicking "Sync this Meeting" via the right-click menu to run the heuristic parser again.
* **COM Errors & File Locks:** If Visio, Outlook, or PowerPoint crash in the background, invisible instances of the programs might get stuck in your system's memory and lock your files. Click the **🖥️ Task Manager** button in the app console and click **Kill Ghosts** to instantly clear them out without losing your active work.
* **The Table Resizing "Amnesia":** If a dropdown menu is open while filtering TDocs, the row heights might appear squished. Qt inherently suspends geometry calculations while popups are active to save memory. 
* **Word Splitter Memory Throttling:** Slicing a heavy `.docx` file unzips a massive XML tree into your RAM. To prevent memory crashes and disk thrashing, the parallel processing is hard-capped at 3 maximum threads. You will see the chapters output in batches of 3 in the console.