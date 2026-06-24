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
7. [📜 License](#license)

---

## <a id="features"></a>✨ Features

### 📡 3GPP Meeting & Specifications Database
* **Asynchronous Three-Phase Syncing Engine:** * **Phase 1 (FTP Directory Mapping):** Scrapes the 3GPP FTP archives in parallel to instantly populate your database with all available meeting numbers, gracefully handling hidden RAN Ad-Hoc (`TSGR_AHs`) subdirectories.
  * **Phase 2 (Deep Document Scrape):** Crawls the `Docs/` folder of every meeting. Uses smart Regex stripping to ignore file extensions and revisions, mathematically sorting the files to determine the first and last TDocs of the meeting.
  * **Phase 3 (DynaReport Upserting):** Injects metadata (Location, Start/End Dates, Ad-Hoc/Electronic status) by fetching the legacy 3GPP Portal HTML tables.
* **Intelligent TDocs Manager:**
  * **Natural Sorting & Smart Filtering:** Bulletproof multi-select dropdowns and natural numerical sorting for complex multi-level Agenda Items (e.g., AI 20.6.2 sorts correctly before 20.6.11).
  * **SA2 Electronic Revisions:** Automatically scrapes `INBOX/Revisions/` for electronic meetings, seamlessly mapping `rXX` versions to their base TDocs via an intuitive cascading context menu.
  * **Multi-Action Folder Integration:** Instantly jump to local cache directories, Main FTP folders, Docs/ folders, or Revisions/ folders directly from the UI.
* **3GPP FTP Session Manager:** Automatically injects randomized User-Agents and HTTP Keep-Alive headers to bypass aggressive 3GPP server throttling and "Too Many Requests" blocks.

### 📝 Word Document Manipulation
* **Global Comparison Cart:** A persistent state dashboard that bridges multiple meeting windows. Push any Base TDoc or specific Revision into "Slot A" and "Slot B", then launch a native Word comparison instantly.
* **Isolated Word Diff Engine:** Uses COM `DispatchEx` to spawn an invisible, isolated instance of Microsoft Word. It safely opens files as Read-Only, auto-accepts tracked changes purely in RAM, and generates a visual comparison without freezing your active Word sessions or locking local files.
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
2. **The Core Layer (`modules/*/core/`):** Contains the heavy lifting. All database transactions (`sqlite3`), FTP network scraping (`requests`), COM object automation (`win32com`), and XML manipulation (`python-docx`) are isolated here.
3. **The Threading Bridge:** Every Core module inherits from `QThread`. The UI sends data to the Thread, and the Thread emits `pyqtSignals` back to the UI to update progress bars or logs.
4. **The Singleton Managers:** The Network Configuration (proxies) and Comparison Cart states are managed by robust Singletons to ensure cross-tab synchronization.

---

## <a id="prerequisites"></a>⚙️ Prerequisites

To run this application natively or build it from source, you must have the following installed on your Windows machine:

1. **Python 3.10+**
2. **Microsoft Word (Desktop App)** (Required for the COM Automation Splitter, Converter, and Diff Engine)
3. **Java Runtime Environment (JRE) 11+** (Required for the local PlantUML generation engine)
4. *(Optional but Recommended)* **Microsoft Visio** (To view the generated outputs)

---

## <a id="installation"></a>🚀 Installation

### 1. Clone the Repository
```bash
git clone https://github.com/your-repo/3GPP-Delegate-Helper.git
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
2. Click **Sync All Meetings** to trigger the 3-Phase scraper. 
3. Right-click any meeting to access its FTP folders, view its info, or open its cached **TDocs List**.
4. In the TDocs Window, use the **Search** bar or dropdown filters to find specific documents. Click the Action column to automatically download, unzip, and open the `.doc` files, or use the **⚖️ Add to Comparison Cart** submenu to select base versions or revisions for diffing.

### 📝 Slicing & Comparing Word Documents
1. In the **Comparison Cart** at the bottom of the Meetings Tab, populate Slot A and Slot B with local files or fetched 3GPP Revisions.
2. Click **Compare in Word**. The tool will spawn a background process and present you with a native Word redline document.
3. For large specs, navigate to the **Spec Splitter** tab, drag a `.docx` file, choose a Heading depth (e.g., "Level 2" for clauses like `6.1`, `6.2`), and click Split.

### 🎨 PlantUML Editor
1. Type standard PlantUML code into the left pane.
2. The Live Preview will automatically update the image on the right.
3. Click **Export Visio** to generate a native `.vsdx` file, or **Copy to Clipboard** to paste the image directly into PowerPoint.

### ⚙️ Configuring Corporate Proxies
If you are behind a corporate firewall:
1. Click the **Network Config (Proxy)** button in the top right.
2. Enter your HTTP/HTTPS proxies into the global session without restarting the app.

---

## <a id="troubleshooting"></a>🛠️ Known Quirks / Troubleshooting

* **Missing / Ghost Meetings in DB:** If 3GPP deletes an old folder from their FTP server, Phase 1 cannot find it. However, the Phase 3 Upsert Engine will find the historical record on the DynaReport page and *force* it into your database anyway. These meetings will have a valid 3GPP Portal link, but their FTP links will lead to a 404.
* **The "3GPP Empty String" Crash:** The scraper uses advanced unicode normalization to handle 3GPP's notorious formatting errors (like using non-breaking hyphens in dates or completely omitting table columns). If a meeting appears with blank dates, try clicking "Sync this Meeting" via the right-click menu to run the heuristic parser again.
* **COM Errors & File Locks:** If Visio or PowerPoint crash in the background, invisible instances of the programs might get stuck in your system's memory and lock your files. Click the **🖥️ Task Manager** button in the app console and click **Kill Ghosts** to instantly clear them out without losing your active work.
* **The Table Resizing "Amnesia":** If a dropdown menu is open while filtering TDocs, the row heights might appear squished. Qt inherently suspends geometry calculations while popups are active to save memory. 
* **Word Splitter Memory Throttling:** Slicing a heavy `.docx` file unzips a massive XML tree into your RAM. To prevent memory crashes and disk thrashing, the parallel processing is hard-capped at 3 maximum threads. You will see the chapters output in batches of 3 in the console.