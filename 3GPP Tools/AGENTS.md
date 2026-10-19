# AGENTS.md — 3GPP Tools Developer & AI Agent Guide

This document is the authoritative technical reference for human developers and AI coding agents working on the **3GPP Meeting Tools & Diagram Converter** repository. It outlines project architecture, subsystem responsibilities, coding conventions, threading and execution rules, and security boundaries.

---

## 1. Project Overview

**3GPP Tools** is a comprehensive Python/PyQt5 desktop engineering environment tailored for telecommunications engineers and 3GPP standards delegates. It integrates:
- **PlantUML & Office Integration:** Live PlantUML editing with debounced preview and conversion to native editable Microsoft Visio (`.vsdx`) and PowerPoint shapes.
- **Specification Search & Evolution:** SQLite FTS5 trigram substring search across 3GPP specification releases with "First Added" chronological tracking, cutoff date detection, and OpenXML parsing of multi-part documents.
- **Protocol Inspection & Evolution:** Message and Information Element (IE) diffing across releases for 5GS/EPS NAS, NR/LTE RRC, NGAP, GTPv1-U, PDU Session UP (TS 38.415), and PFCP (TS 29.244).
- **Meeting & Work Item Management:** Offline synchronization of 3GPP meetings, TDocs, dynamic agendas, email discussion summarization, and company contribution audits.
- **Local AI Intelligence (Ollama):** Zero-cost, 100% offline, privacy-guaranteed AI integration for TDoc technical triage, call-flow sequence generation, email stance summarization, and an evidence-grounded 3GPP Research Assistant.

---

## 2. Environment & Execution

### Prerequisites
- **Operating System:** Windows (primary environment, using Win32 COM APIs and Windows shell hooks).
- **Python:** `>= 3.8` (typically executed with Python 3.10 to 3.14 64-bit).
- **Java JRE / JDK:** 64-bit Java 11+ recommended (or legacy Java 8) for PlantUML JAR rendering. Portable JREs are supported via zero-admin configuration.
- **Microsoft Office (Optional but required for Visio/Word features):** Microsoft Visio and Word desktop installations for COM automation. PlantUML editing and specification tools function without Office installed.
- **Local Ollama Daemon (Optional for AI features):** Running locally at `http://127.0.0.1:11434`.

### Installation & Launching
Dependencies are managed via `pyproject.toml` and `requirements.txt`:
```powershell
pip install -r requirements.txt
```

Launch the application:
```powershell
# Using the Windows batch launcher:
.\src\Launch.cmd

# Or directly via Python with proper PYTHONPATH:
$env:PYTHONPATH = ".;.\src;$env:PYTHONPATH"
python -m main_tools
```

---

## 3. Repository Layout

```text
3GPP Tools/
├── README.md                      # Comprehensive user & architectural documentation
├── pyproject.toml                 # Package definition & CLI entry points
├── requirements.txt               # Direct runtime dependencies
├── AGENTS.md                      # Root developer and AI agent guide (this file)
├── output/                        # Default export directory for converted diagrams/docs
├── src/
│   ├── Launch.cmd                 # Startup script with PYTHONPATH configuration
│   ├── main_tools.py              # Application entry point, logging, faulthandler setup
│   ├── main_window.py             # DragDropUI main window, tab instantiation, system bar
│   ├── assets/                    # PlantUML JAR, icons, and bundled runtime assets
│   ├── config/                    # JSON user preferences & externalized prompts
│   │   ├── prompts/               # Hot-reloadable prompt templates (.txt)
│   │   └── *.json                 # Subsystem configurations (Ollama, specs, meetings, etc.)
│   ├── core/                      # Shared cross-cutting infrastructure
│   │   ├── ai/                    # OllamaClient, health monitoring, OllamaConfigDialog
│   │   ├── config/                # Global application constants and paths
│   │   ├── network/               # NetworkSession singleton, RFC 1918 proxy bypassing
│   │   ├── ui/                    # Base components, global QSS style, console/queue panels
│   │   ├── utils/                 # Path resolution, Java runtime detection, helpers
│   │   └── queue_manager.py       # Multi-threaded background conversion & sync task queue
│   ├── db/                        # Persistent SQLite databases (3gpp_data.db, etc.)
│   └── modules/                   # Independent feature subsystems (tabs)
│       ├── assistant/             # Evidence-grounded local 3GPP Research Assistant
│       ├── emails/                # WG reflector email scraping, quote stripping, stance audit
│       ├── meetings/              # Meetings calendar, TDoc browser, AI triage, sync workers
│       ├── nas/                   # Protocol evolution matrix, bit-level rendering, parsers
│       ├── puml2visio/            # PlantUML editor, Live Preview, PPTX/Visio conversion, Java UI
│       ├── spec_search/           # FTS5 trigram search, release matrix, multipart doc parser
│       ├── specifications/        # 3GPP FTP spec archive synchronization & catalogue
│       ├── word_tools/            # Docx splitting by headings, Word COM comparison
│       └── work_items/            # 3GPP Work Item database, scraper, and TDoc correlation
```

---

## 4. Subsystems & Module Guide

### 4.1. Visio Tools (`src/modules/puml2visio`)
- **PlantUML IDE:** Debounced live SVG/PNG preview via `LivePreviewManager` utilizing `plantuml.jar`.
- **Java Maintenance:** `JavaMaintenanceDialog` inspects Java runtime versions (Legacy vs Modern) and auto-updates PlantUML JARs.
- **AI Call Flow Generator:** Modeless dialog streaming sequence diagram synthesis from textual 3GPP procedure descriptions into the editor via chunked Ollama HTTP streaming.
- **Conversion Pipeline:** PPTX to Visio conversion uses an Enhanced Metafile (EMF) bridge to preserve native editable vector shapes with text shrink-wrapping.

### 4.2. Word Tools (`src/modules/word_tools`)
- **Document Slicing:** Splits massive specification `.docx` files by heading level (`Heading 1`, `Heading 2`, etc.) using `python-docx`.
- **Document Comparison:** Automates Microsoft Word COM (`win32com.client`) to generate legal redline comparisons between two `.docx` or `.doc` revisions.

### 4.3. Specification Search (`src/modules/spec_search`)
- **FTS5 Trigram Substring Search:** Utilizes SQLite FTS5 configured with `tokenize="trigram"` in `3gpp_spec_search.db` for arbitrary substring, field, and acronym lookups.
- **Release Matrix & "First Added":** Visual matrix displaying text evolution (`🟢 Added`, `✓ Present`, `🔴 Removed`) across chronological releases.
- **Multi-Part Document Parser:** Automatically chains and sequences multi-part split documents (e.g., `_s00_s04.docx`, `_s05_s08.docx`) into a unified release model using direct `lxml` OpenXML parsing without Word COM overhead.

### 4.4. Protocol Matrix & Inspector (`src/modules/nas`)
- **Multi-Protocol Support:** Covers 5GS NAS (TS 24.501), EPS NAS (TS 24.301), NR RRC (TS 38.331), LTE RRC (TS 36.331), NGAP (TS 38.413), GTPv1-U (TS 29.281), PDU Session UP (TS 38.415), and PFCP (TS 29.244).
- **Parsing Architecture:** Unified under `ProtocolDocxDispatcher`. For in-depth parsing guidelines, refer to [`src/modules/nas/core/parsing/AGENTS.md`](file:///src/modules/nas/core/parsing/AGENTS.md).
- **Inspection UI:** Renders bit-level octet layouts, OpenXML tables with `gridSpan`/`vMerge`, and ASN.1 syntax blocks with field description tables.

### 4.5. Meetings & TDocs (`src/modules/meetings`)
- **Three-Phase Sync:** Asynchronously synchronizes meeting lists, TDoc metadata, and dynamic agenda reservations from 3GPP FTP servers.
- **AI TDoc Triage:** Modeless dialog extracting CR cover sheets and redlines to synthesize structured executive summaries (*Objective*, *Key Changes*, *Potential Impacts*).
- **Personal Notes & Stance:** Integrates local SQLite user notes (`user_tdocs.db`) with stance tagging (`⚪ Neutral`, `🔵 My TDoc`, `🟢 Support`, `🔴 Object`, `🟡 Monitor`).

### 4.6. Research Assistant (`src/modules/assistant`)
- **Architecture:** Bounded, evidence-grounded agent architecture connecting a local LLM to application knowledge services.
- **Authority Boundary:** Strictly **`KNOWLEDGE_READ`**. The model has no SQL access, shell execution, or database write permissions.
- **Semantic Tool Set:** Model interaction is restricted to predefined semantic tools (`find_specifications`, `get_specification`, `search_spec_text`, `get_spec_clause`, `list_protocols`, `find_protocol_message`, `find_protocol_ie`).
- **Detailed Design:** See [`src/modules/assistant/README.md`](file:///src/modules/assistant/README.md).

---

## 5. Architectural Rules & Constraints for AI Agents

When modifying or extending this codebase, adhere to the following architectural rules:

### 5.1. Threading & UI Responsiveness
- **Never block the Qt GUI thread:** Never execute network requests, database scraping, large document parsing, or LLM inference on the main thread.
- **Worker Conventions:** Use `QThread` subclasses or submit tasks to `QueueManager`.
- **Cooperative Cancellation:** Workers must periodically check `self.isInterruptionRequested()` or application `CancellationToken` instances. Never call `QThread.terminate()` as it can corrupt SQLite locks and C runtime states.
- **Safe Signal Delivery:** Pass data back to UI widgets via Qt signals (`pyqtSignal`) rather than directly manipulating widgets from background threads.

### 5.2. Win32 COM Automation Safety
- **Apartment Initialization:** Any background thread using `win32com.client` must call `pythoncom.CoInitialize()` before accessing COM objects and `pythoncom.CoUninitialize()` in a `finally` block.
- **Resource Cleanup:** Always wrap Word/Visio COM operations in `try...finally` blocks to explicitly call `app.Quit()` and delete references (`del app`), preventing orphaned background `WINWORD.EXE` or `VISIO.EXE` processes.
- **Legacy Conversion:** Prefer direct `lxml` extraction for reading DOCX files; only invoke COM when native Word layout/rendering is strictly necessary.

### 5.3. Local AI & Corporate Proxy Bypassing
- **Strictly Offline & Local:** Do not add cloud LLM fallbacks. All inference must target the local Ollama instance (`127.0.0.1:11434`).
- **Loopback Proxy Bypass:** Use `core.network.session.get_ai_session()` for AI HTTP calls. This factory automatically detects RFC 1918 loopback addresses (`is_local_address()`) and configures `trust_env=False` to prevent corporate intercepting proxies (e.g., Zscaler, BlueCoat) from hijacking local inference traffic.
- **Prompt Decoupling:** Decouple prompt text into `src/config/prompts/` to allow live hot-reloading based on file `mtime` without application restarts.

### 5.4. Database Integrity & Multi-Store Isolation
The project intentionally separates domain data into distinct SQLite databases located in `src/db/`:
1. `3gpp_data.db`: Meeting schedules, TDocs, specification catalogue metadata, work items.
2. `3gpp_protocol_data.db`: Structured protocol models, messages, and IE hierarchies.
3. `3gpp_spec_search.db`: Extracted clause text with FTS5 trigram indexes.
4. `user_tdocs.db`: User notes, personal review status, and triage archives.

- **Cross-Database Queries:** Do not couple queries across these databases using SQLite `ATTACH` unless strictly isolated inside a dedicated maintenance script.
- **WAL Compaction:** After massive writes or deletions, perform compaction via `PRAGMA wal_checkpoint(TRUNCATE)` and `VACUUM` (available via `DatabaseMaintenanceDialog`).

### 5.5. Logging & Diagnostics
- Use standard `logging.getLogger(__name__)`.
- Log records are mirrored to `3gpp_tools.log` (with timestamps/levels) and to the in-app `ConsolePanel` (clean message text via `GuiLogHandler`).
- Avoid dumping massive raw payloads (e.g., complete multi-megabyte specification clauses) into logs; log lengths and concise identifiers instead.

### 5.6. UI Consistency, Styling & Code Reuse
- **Reuse Common Styles:** Prefer shared design tokens and style constants in `src/core/ui/ui_components.py` (e.g., `BUTTON_STYLE_TOOLBAR_SECONDARY`, `CARD_FRAME_STYLE`, `BADGE_STYLE_*`, `TABLE_STYLE_CLEAN`) over ad-hoc inline stylesheets.
- **Support Interactive States:** Checkable buttons or toggleable controls must properly define interactive pseudo-states (such as `QPushButton:checked` and `QPushButton:checked:hover`) using the application's unified palette (`#EBF3FC` background, `#1E5C99` / `#395396` border and text).
- **Functionality & Component Reuse:** Always reuse existing core infrastructure, dialogs, network managers, and utility functions across feature modules rather than duplicating logic.

---

## 6. Development & Verification Workflow

When adding new features or fixing bugs:
1. **Verify Python Syntax & Imports:** Ensure imports resolve cleanly with `PYTHONPATH=".;src"`.
2. **Preserve Comments & Docstrings:** Do not strip existing architectural comments or docstrings from untouched code.
3. **Check Submodule Guides:** When touching protocol parsing, consult [`src/modules/nas/core/parsing/AGENTS.md`](file:///src/modules/nas/core/parsing/AGENTS.md). When touching the Assistant, consult [`src/modules/assistant/README.md`](file:///src/modules/assistant/README.md).
4. **Manual Verification:** Because the application is a graphical Qt desktop environment, test modified tabs or workers by running `python -m main_tools` and verifying that the Qt event loop starts and stops cleanly without console tracebacks or thread leaks.
