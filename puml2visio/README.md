# 📊 PlantUML to Visio/PowerPoint Converter (3GPP Tools)

An advanced, component-based desktop IDE designed to bridge the gap between text-based diagramming (`PlantUML`) and corporate enterprise environments (`Microsoft Visio` and `PowerPoint`). 

Built specifically with telecommunications and 3GPP standards workflows in mind, this tool allows you to write highly efficient PlantUML sequence, activity, and network diagrams and instantly export them as fully editable native Office shapes, SVGs, or Unicode Text Art.

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

* **Smart Code Editor:** A professional IDE experience featuring dynamic line numbering, active-line highlighting, native Undo/Redo history, and a background Auto-Save Cache that restores your session if the app is closed or crashes.
* **Intelligent Live Preview:** A debounced background rendering engine that automatically pipes your PlantUML code to a live browser tab as you type. If you make a typo, it intercepts the Java crash and dynamically paints a red Syntax Error overlay (with the exact line number) directly in your browser.
* **Enterprise Diagram Boilerplates:** Includes 29 built-in diagram templates ranging from standard UML to specialized IT formats (`nwdiag`, `rackdiag`, `packetdiag`). All UML templates are automatically styled with a flat, enterprise-ready "monochrome" theme (Arial font, white backgrounds, black lines) optimized for corporate documents.
* **Rich Export Engine:** * **Visio (.vsdx):** Perfect alignment via 2D SVG gap-measuring.
  * **PowerPoint (.pptx):** Bypasses buggy SVG engines via an EMF pipeline for natively ungroupable objects.
  * **ASCII Text Art (.txt):** Uses PlantUML's `-tutxt` engine to generate clean Unicode text diagrams for markdown or RFC specs.
* **Built-in COM Process Manager:** A native "kill switch" dialog that safely identifies and terminates headless "ghost" instances of Visio, PowerPoint, or Word left hanging in memory by background crashes, preventing file locks and memory leaks.
* **Native Windows Integration:** Automatically applies a dynamic vector App Icon to the Windows Taskbar (bypassing the generic Python logo), features an "Open Folder" explorer hook, and dynamically copies generated file paths to your clipboard.
* **Word Document Extractor:** Extracts hidden, embedded Visio (`.vsdx`) files natively trapped inside Word Document (`.docx`) OLE wrappers.
* **Modular MVC Architecture:** Built on a decoupled UI standard, utilizing dedicated UI Tabs, UI Panels, and a centralized Python `QueueManager` to handle threading without locking the GUI.

---

## <a id="architecture"></a>🏗️ Architecture & Data Flow

```mermaid
graph TD
    subgraph UI (View)
        E[puml2visio.py<br>Entry Point] --> M[main_window.py<br>The Traffic Cop]
        M --> T[ui_tabs.py<br>Code Editor & Drop Zones]
        M --> P[ui_panels.py<br>Console & Queue]
    end

    subgraph Controller
        M --> Q[queue_manager.py<br>Thread Orchestration]
    end

    subgraph Core Engines (Model)
        Q --> J[utils.py<br>Java Registry Scanner]
        Q --> V[visio_converter.py<br>COM Automation]
        Q --> PPT[powerpoint_converter.py]
        T --> LP[live_preview.py<br>Error Interceptor]
        P --> PM[process_manager.py<br>OS Process Monitor]
    end

    subgraph Output
        OutV[Visio .vsdx]
        OutP[PowerPoint .pptx<br>Native Shapes]
        OutA[ASCII .txt]
        OutB[Browser Preview]
    end

    T -- "Type / Edit" --> LP
    LP -- "Render / Error Hook" --> OutB
    
    T -- "Export Requested" --> Q
    Q -- "1. Generate Clean SVG" --> J
    
    J -- "2a. Visio Pipeline" --> V
    V --> OutV

    J -- "2b. PowerPoint EMF Pipeline" --> PPT
    PPT --> OutP
    
    Q -- "2c. Unicode Render" --> OutA
    
    PM -- "Identify/Kill Ghosts" --> V
```

---

## <a id="prerequisites"></a>⚙️ Prerequisites

Because this application relies heavily on Microsoft's Component Object Model (COM) to natively manipulate diagrams, it requires a specific environment:

1. **Windows OS** (Required for COM automation).
2. **Microsoft Visio** and **Microsoft PowerPoint** installed locally.
3. **Java Runtime Environment (JRE)** (Java 11+ recommended to support the newest PlantUML features; Java 8 minimum. The tool will auto-detect the best version).
4. **Python 3.8+**

---

## <a id="installation"></a>🚀 Installation

1. **Clone the repository:**
   ```bash
   git clone [https://github.com/telekom/3gpp-meeting-tools.git](https://github.com/telekom/3gpp-meeting-tools.git)
   cd 3gpp-meeting-tools/puml2visio
   ```

2. **Install the required Python packages:**
   Create a virtual environment (optional but recommended) and install dependencies:
   ```bash
   pip install PyQt5 pywin32
   ```

3. **Run the application:**
   ```bash
   python puml2visio.py
   ```
   *Note: On first launch, the app will automatically attempt to download `plantuml.jar`. If you are behind a corporate firewall, a proxy configuration dialog will appear to assist. You can click "Test Connection" to verify your proxy credentials.*

---

## <a id="usage"></a>📖 How to Use the GUI

The application features three main workspaces navigated via tabs, with a fully resizable bottom terminal and queue viewer.

### 📝 Tab 1: Code Editor (Single Diagram Mode)
* **Templates & Docs:** Use the dropdown menu to select from 29 diagram types. Click `Insert` to populate the editor with an enterprise-styled boilerplate, or `Docs` to instantly open the official syntax guide.
* **Auto-Save & Undo:** The editor automatically saves your work every 2 seconds to a hidden cache file. If you accidentally clear the editor, click `↩️ Undo` to restore it. 
* **Live Preview:** Click `👁️ Live Preview`. The app will open a browser tab and automatically render your diagram as you type. If you make a syntax error, the browser will flash red and tell you exactly which line failed.
* **Exporting:** Click the `📤 Export Diagram ▼` dropdown to select your target format (.vsdx, .pptx, .svg, or .txt). The application will generate the file, change the `Copy Path` tooltip, and allow you to click `📂 Open Folder` to instantly view the result in Windows Explorer.
* **Round-Trip Extract:** Drag and drop a previously generated `.vsdx` file directly into the text box to instantly retrieve its original source code.

### 📂 Tab 2: Batch Convert
* Drag a selection of `.txt` or `.puml` files from your file explorer and drop them onto the dashed area. 
* The application will queue them up in the **Queue Viewer** at the bottom right. You can select items in the queue and click `Remove` to cancel them before they process.

### 📄 Tab 3: Word Extractor
* When collaborating on 3GPP standards, Visio files are often deeply embedded inside Word documents as OLE objects. 
* Drag and drop a `.docx` file onto this tab. The app will unzip the archive in milliseconds, extract the clean `.vsdx` files, and place them right next to your original Word document.

### 🛠 System Toolbar (Console Header)
* **🖥️ Task Manager:** Open the COM Process Manager to identify and kill background headless instances of Visio or PowerPoint that may have crashed.
* **📡 Proxy:** Update your network configuration on the fly and test the connection to GitHub.
* **🔄 Update JAR:** Force the application to ping GitHub and check if a newer version of PlantUML is available to download.

---

## <a id="troubleshooting"></a>🛠️ Known Quirks / Troubleshooting

* **PowerPoint "Leave Open" Behavior:** Unlike Visio exports (which save silently to your disk), clicking `Export PPTX` intentionally leaves the generated PowerPoint presentation open and unsaved on your screen. This allows you to immediately copy the generated slide and paste it directly into your master deck. 
* **COM Errors & File Locks:** If Visio or PowerPoint crash in the background, invisible instances of the programs might get stuck in your system's memory and lock your files. If you start receiving `COM Error` messages, click the **🖥️ Task Manager** button in the app console and click **Kill Ghosts** to instantly clear them out without losing your active work.
* **Missing Visio Source Code Alignment:** Modifying the PlantUML `textLength` attributes manually might cause Visio text boxes to behave erratically. The tool automatically cleans standard SVG artifacts, but highly customized `skinparam` settings may override this.

