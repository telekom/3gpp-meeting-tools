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

### 📡 3GPP Database Synchronizer
* **Massive Network Parallelism:** Utilizes a Producer-Consumer architecture with 15 concurrent HTTP workers and a single safe SQLite writer thread to rapidly sync the 3GPP FTP archive without disk-locking bottlenecks.
* **Smart Metadata Scraper:** A dual-strategy HTML parser seamlessly extracts metadata (Title, Release, Working Group, Radio Tech) from both the legacy 3GPP HTML tables and the modern ASP.NET WebForms portal.
* **Advanced Filtered Sync:** Don't want to sync 4,000+ files? Use the Advanced Sync dialog to target specific cross-sections (e.g., *only `TS` specifications from the `23` series where `SA2` is responsible*).
* **Targeted Row Updates:** Shift-click multiple specifications in the UI and right-click to fast-track their synchronization instantly.

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
* **Initialize Database:** Navigate to the **Specifications** tab. Click **Full Sync** on your first run to map the 3GPP archive to your local database.
* **Filter and Download:** Use the dynamic search bars (e.g., Type `23.501`) to instantly filter the local database. Select the version you need from the dropdown and click the download button to open the zip file.
* **Filtered Sync:** Click **⚙️ Filtered Sync** to update only a subset of specifications (e.g., 5G, SA2) without scanning the entire 3GPP server.
* **Targeted Refresh:** Select specific rows in the table, right-click, and select **🔄 Update selected** to instantly fetch the latest files and metadata for those specific documents.

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
