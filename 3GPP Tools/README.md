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
  * **Phase 2 (Deep Document Scrape):** Crawls the `Docs/` folder of every meeting. Uses smart Regex stripping to ignore file extensions and revisions, mathematically sorting and saving the exact **First TDoc** and **Last TDoc** for chronological indexing.
  * **Phase 3 (DynaReport Metadata):** Connects to the 3GPP ASP.NET Portal to fetch rich metadata (Locations, Dates, Meeting Types). Built with a bulletproof **NLP Heuristic Engine** that universally normalizes unicode hyphens, bi-directionally searches for dates, and uses an **Upsert Engine** to add historical meetings even if their FTP folders have been deleted.
* **Advanced Database Filtering:** Instantly filter thousands of meetings by Working Group (SA2, RAN3, CT1, etc.), Date Range, Location, **In-Person vs. Electronic**, or **Regular vs. Ad-Hoc / BIS**.
* **Direct Portal Integration:** Right-click any meeting to instantly open its parent FTP folder, its `Docs/` subfolder, or launch the official 3GPP Web Portal directly via its extracted `MtgId`.

### 1-Click Document Slicer
* **Massive File Handling:** Upload a monolithic 3GPP `.docx` specification (often hundreds of pages long) and automatically slice it into individual, bite-sized Word documents based on Heading 1 chapters.
* **Parallel Processing:** Utilizes `concurrent.futures` to slice heavy XML structures across multiple CPU threads, finishing in seconds what would take minutes manually.

### 🎨 Native PlantUML to Visio / PPTX
* **Text-to-Vector:** Write standard PlantUML code and click "Export." The tool compiles the code to an SVG, parses the internal XML, and translates it entirely into native Microsoft COM Objects. 
* **Fully Editable Shapes:** No flat images! Every lifeline, arrow, and text box becomes a real, grouped Visio shape or PowerPoint element that you can drag, recolor, and edit natively.
* **Smart Alignment:** Calculates bounding boxes and font metrics dynamically to ensure your Visio exports look exactly like the generated PlantUML preview.

---

## <a id="architecture"></a>🏗️ Architecture & Data Flow

This application is built with a highly decoupled, modular **PyQt5** architecture.

1. **`core/`:** Contains network sessions with automatic proxy injection, multi-threaded worker classes, database managers (`SQLite3`), and the 3GPP HTML/Regex scraping engines.
2. **`ui/`:** Component-based UI elements. Features a dark-mode Syntax Highlighter (`QsciScintilla`), SVG preview panes, responsive splitters, and separated Data Models (`QAbstractTableModel`).
3. **`exporters/`:** The COM interop layer. Uses `win32com.client` to boot invisible background instances of Visio and PowerPoint, injecting XML path data directly into the Microsoft Office API.

---

## <a id="prerequisites"></a>⚙️ Prerequisites

* **Windows OS** (Required for the `win32com` Office exporting functionality).
* **Python 3.9+**
* **Microsoft Visio & PowerPoint** (Must be locally installed and activated. Web-only versions are not supported).
* **Java (JRE)** (Required by the local PlantUML `.jar` compiler).
* Graphviz / Dot (Optional, but recommended for advanced PlantUML activity diagrams).

---

## <a id="installation"></a>🚀 Installation

1. **Clone the repository:**
   ```bash
   git clone [https://github.com/yourusername/3gpp-meeting-tools.git](https://github.com/yourusername/3gpp-meeting-tools.git)
   cd 3gpp-meeting-tools
   ```
2. **Install Python Dependencies:**
   ```bash
   pip install -r requirements.txt
   ```
   *(Core dependencies include: `PyQt5`, `QScintilla`, `pywin32`, `beautifulsoup4`, `requests`, `python-docx`)*
3. **Run the Application:**
   ```bash
   python main.py
   ```

---

## <a id="usage"></a>📖 How to Use the GUI

### Tab 1: Meetings & Sync
* **Syncing:** On the right-side panel, ensure all three Scrape Configuration checkboxes are checked, then click **"🔄 Sync All Meetings"**. Watch the console as it maps directories, extracts TDocs, and back-fills metadata.
* **Filtering:** Use the dropdowns to quickly find "SA2" meetings that were "Electronic" and "Ad-Hoc".
* **Right-Click Menu:** Right-click any row to view full meeting info, jump to the FTP folder, or open the 3GPP Portal.

### Tab 2: Document Slicer
* Click **"Browse..."** to select a heavy `.docx` specification file.
* Select an output folder.
* Click **"Slice Document"**. The progress bar will track the asynchronous extraction of chapters.

### Tab 3: Diagram Converter
* **Editor:** Write your code in the left pane. Syntax is automatically highlighted. 
* **Preview:** Click `Render Preview` (or press `Ctrl+Enter` / `F5`) to generate an SVG preview on the right.
* **Export Visio:** Ensure Visio is closed or idle. Click `Export Visio`. The app will launch Visio in the background, draw the shapes, and prompt you to save the `.vsdx` file.
* **Export PPTX:** Click `Export PPTX`. PowerPoint will open, generate a new slide with your editable shapes, and leave the application open for you to review.

### 🌐 Top Toolbar Controls
* **🖥️ Task Manager:** Instantly lists all background `VISIO.EXE` and `POWERPNT.EXE` processes. Use "Kill Ghosts" if an export crashes and leaves a locked process behind.
* **🗑️ Clear Cache:** Deletes temporary SVG files and downloaded `.jar` files from your AppData temp point.
* **📡 Proxy:** Instantly test and inject HTTP/HTTPS proxies into the global session without restarting the app.
* **🔄 Update JAR:** Ping GitHub for newer versions of PlantUML.

---

## <a id="troubleshooting"></a>🛠️ Known Quirks / Troubleshooting

* **Missing / Ghost Meetings in DB:** If 3GPP deletes an old folder from their FTP server, Phase 1 cannot find it. However, the Phase 3 Upsert Engine will find the historical record on the DynaReport page and *force* it into your database anyway. These meetings will have a valid 3GPP Portal link, but their FTP links will lead to a 404.
* **The "3GPP Empty String" Crash:** The scraper uses advanced unicode normalization to handle 3GPP's notorious formatting errors (like using non-breaking hyphens in dates or completely omitting table columns). If a meeting appears with blank dates, try clicking "Sync this Meeting" via the right-click menu to run the heuristic parser again.
* **COM Errors & File Locks:** If Visio or PowerPoint crash in the background, invisible instances of the programs might get stuck in your system's memory and lock your files. Click the **🖥️ Task Manager** button in the app console and click **Kill Ghosts** to instantly clear them out without losing your active work.
* **Word Splitter Memory Throttling:** Slicing a heavy `.docx` file unzips a massive XML tree into your RAM. To prevent memory crashes and disk thrashing, the parallel processing is hard-capped at 3 maximum threads. You will see the chapters output in batches of 3 in the console.

