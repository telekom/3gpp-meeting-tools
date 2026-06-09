# 📊 PlantUML to Visio/PowerPoint Converter (3GPP Tools)

An advanced desktop application designed to bridge the gap between text-based diagramming (`PlantUML`) and corporate enterprise environments (`Microsoft Visio` and `PowerPoint`). 

Built specifically with telecommunications and 3GPP standards workflows in mind, this tool allows you to write highly efficient PlantUML sequence diagrams and instantly export them as fully editable native Office shapes.

> **🤖 AI-Assisted Development:** > The architecture, UI polishing, and complex Microsoft COM automation in this project were heavily co-developed using advanced Large Language Models (LLMs), allowing for rapid iteration and deep integration into native Windows APIs.

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

* **Visio Integration:** Generates native `.vsdx` files from PlantUML text. Intelligently strips structural SVG wrappers so shapes are easily editable.
* **Native PowerPoint Export:** Bypasses standard image embedding by coercing PowerPoint into converting SVG paths into **native, grouped Office Drawing objects**. 
* **Flawless Round-Tripping:** Embeds your original PlantUML source code directly into the Visio Page or PowerPoint Speaker Notes. Drop a generated file back into the app to retrieve your source code!
* **Word Document Extractor:** Extracts hidden, embedded Visio (`.vsdx`) files natively trapped inside Word Document (`.docx`) OLE wrappers.
* **Intelligent Auto-Setup:** Automatically detects your Java environment and downloads the correct `plantuml.jar` on first launch (with Corporate Proxy support).

---

## <a id="architecture"></a>🏗️ Architecture & Data Flow

```mermaid
graph TD
    subgraph Input
        A[PlantUML Text / .puml]
        W[Word Document .docx]
    end

    subgraph Core Engines
        J[Java: plantuml.jar]
        P[Python ZIP Extractor]
    end

    subgraph Windows COM Automation
        V[Microsoft Visio API]
        PPT[Microsoft PowerPoint API]
    end

    subgraph Output
        OutV[Visio .vsdx]
        OutP[PowerPoint .pptx<br>Native Shapes]
        OutS[Standard .svg]
    end

    A -->|1.