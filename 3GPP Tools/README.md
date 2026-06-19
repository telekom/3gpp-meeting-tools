# 📊 3GPP Meeting Tools & Diagram Converter

An advanced, component-based desktop IDE designed to bridge the gap between text-based diagramming (`PlantUML`) and corporate enterprise environments (`Microsoft Visio` and `PowerPoint`). 

Built specifically with telecommunications and 3GPP standards workflows in mind, this tool allows you to write highly efficient PlantUML sequence, activity, and network diagrams, instantly export them as fully editable native Office shapes, rapidly slice massive specification documents into manageable chapters, and seamlessly navigate, filter, and synchronize the vast 3GPP specification database locally.

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

### 📡 3GPP Specifications Database
* **Asynchronous Two-Pass Syncing:** * **Pass 1 (Lightning Fast):** Scrapes the 3GPP FTP archive in parallel to instantly populate your database with all available specification numbers and versions, unblocking the UI in seconds.
  * **Pass 2 (Deep Metadata):** Silently connects to the 3GPP ASP.NET DynaReport pages in the background to fetch rich metadata (Working Groups, Radio Technologies, Initial Release dates) without freezing your workflow.
* **1-Click Document Conversion (Word, PDF, HTML):** Replaces clunky "Download" buttons with instant action buttons (**📝 Word**, **📕 PDF**, **🌐 HTML**). Automatically downloads the `.zip`, flattens the internal folder structure, safely extracts the `.docx` files, and leverages native COM automation to convert the specification to your desired format.
* **Smart Visual Caching:** The UI actively monitors your hard drive. If a specification has already been downloaded or converted to a PDF/HTML file, the action buttons instantly illuminate with a **Green ✅** and bold styling so you know the file is available offline.
* **Data-Driven Precision Filtering:** All search filters (Series, Working Group, Radio Tech, Type) are strict, read-only dropdowns dynamically populated directly from your local database to eliminate typos and manual guessing.
* **Self-Healing Architecture:** Runs a silent SQLite cleaner on startup to automatically purge orphaned Working Groups or Series. Built on SQLite WAL mode, allowing the background scraper to save metadata while you simultaneously search the frontend GUI. Type-safe semantic sorting prevents crashes on malformed 3GPP versions.

### 🎨 Diagramming & Document Tools
* **Smart Code Editor:** A professional IDE experience featuring dynamic line numbering, active-line highlighting, native Undo/Redo history, and a background Auto-Save Cache that restores your session if the app is closed or crashes.
* **Intelligent Live Preview:** A debounced background rendering engine that automatically pipes your PlantUML code to a live browser tab as you type. It intercepts Java crashes and dynamically paints a red Syntax Error overlay directly in your browser.
* **Visio Export (.vsdx):** Perfect alignment via 2D SVG gap-measuring. Converts your PlantUML code into fully editable, ungroupable Microsoft Visio shapes.
* **PowerPoint Export (.pptx):** Automatically injects the generated shapes directly into a blank PowerPoint slide for immediate copy-pasting.
* **Bulk Document Splitter:** Extracts massive `.docx` specification documents and slices them chapter-by-chapter into standalone Word files using parallel processing.

---

## <a id="architecture"></a>🏗️ Architecture & Data Flow

* **PyQt5 GUI:** A multi-tabbed, highly responsive user interface with built-in dark/light contextual elements and custom global stylesheets.
* **Centralized Network Session:** All modules share a robust HTTP connection pool with automatic User-Agent spoofing and aggressive retry strategies (Backoff factor for 403, 429, 502 errors) to survive strict corporate firewalls.
* **Asynchronous Task Queue & Emergency Abort:** Long-running network or COM tasks are isolated in background `QThreads`, ensuring the GUI remains buttery smooth. Includes an emergency **🛑 Abort** button in the Queue Panel to forcefully halt active jobs.
* **Native Visio COM:** Interacts directly with the Windows Component Object Model (COM) via `pywin32` to manipulate the Visio application natively.

---

## <a id="prerequisites"></a>⚙️ Prerequisites

* **Python:** Version 3.8 or higher.
* **OS:** Windows 10 / 11 (Required for COM automation).
* **Microsoft Office:** Installed and activated (Visio, PowerPoint, Word) for document interaction.
* **Java (Optional but Recommended):** Java JRE or JDK must be installed for local PlantUML generation.

---

## <a id="installation"></a>🚀 Installation

1. **Install required dependencies:**
   ```bash
   pip install PyQt5 requests beautifulsoup4 pywin32 python-docx
   ```
   *(Note: Ensure you have your corporate proxy configured in your terminal if you are behind a firewall: `set HTTP_PROXY=http://...`)*

2. **Run the Application:**
   ```bash
   python main_window.py
   ```

---

## <a id="usage"></a>📖 How to Use the GUI

### 📂 3GPP Specifications Database
1. **Set your Download Directory:** Use the `📂 Browse` button to set where specifications should be saved. Click `↗️ Open` to view this folder in Windows Explorer at any time.
2. **Run a Full Sync:** Click `🔄 Full Sync` to map the 3GPP FTP server. The progress bar will disappear quickly, and an orange `⏳ Fetching deep metadata...` warning will appear while it safely downloads Working Group data in the background.
3. **Filter and Search:** Type a specification number (e.g., `23.501`) into the search bar, or use the `⚙️ Table Filters` to isolate specifications by specific Working Groups (e.g., `SA2`).
4. **Read Documents:** Select a version from the dropdown menu. Click **📝 Word** to open the raw document, or click **📕 PDF** / **🌐 HTML** to automatically convert and open it.
5. **Advanced Sync:** Click `⚙️ Filtered Sync` to forcefully re-download metadata only for a specific subset of documents (e.g., only update metadata for `Series 23`).
6. **Targeted Refresh:** Select specific rows in the table, right-click, and select `🔄 Update selected` to instantly fetch the latest files and metadata for those specific documents.

### 🖱️ Drag & Drop Interface
You can drag and drop `.puml`, `.txt`, `.svg`, or `.docx` files directly into the Batch Convert or Word Splitter tabs. The UI will visually indicate accepted files and automatically queue them for processing.

### ⚙️ Background Processing
Keep an eye on the **Queue panel** on the right side of the screen. You can queue dozens of files to be exported to PDF, HTML, XPS, RTF, or pure Text. Runs silently in the background via detached COM instances to keep your active work safe from freezing. If you need to stop a task immediately, click the **🛑 Abort** button.

### 🛠 System Toolbar (Console Header)
* **🖥️ Task Manager:** Open the COM Process Manager to identify and surgically kill headless/ghost instances of Visio or PowerPoint.
* **📡 Proxy:** Instantly test and inject HTTP/HTTPS proxies into the global session without restarting the app.
* **🔄 Update JAR:** Ping GitHub for newer versions of PlantUML.

---

## <a id="troubleshooting"></a>🛠️ Known Quirks / Troubleshooting

* **3GPP Database Sync Returns 0 Results:** If you type in the search bar and see no results, your local database is empty. Ensure your corporate proxy is configured correctly via the **Proxy** toolbar button, then run a **Full Sync**.
* **COM Errors & File Locks:** If Visio or PowerPoint crash in the background, invisible instances of the programs might get stuck in your system's memory and lock your files. Click the **🖥️ Task Manager** button in the app console and click **Kill Ghosts** to instantly clear them out without losing your active work.
* **Word Splitter Memory Throttling:** Slicing a heavy `.docx` file unzips a massive XML tree into your RAM. To prevent memory crashes and disk thrashing, the parallel processing is hard-capped at 3 maximum threads. You will see the chapters output in batches of 3 in the console.
* **PowerPoint "Leave Open" Behavior:** Unlike Visio exports (which save silently to your disk), clicking `Export PPTX` intentionally leaves the generated PowerPoint presentation open and unsaved so you can immediately copy the slide.
* **Missing Visio Source Code Alignment:** Modifying the PlantUML `textLength` attributes manually might cause Visio text boxes to behave erratically during export.