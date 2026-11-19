# Visio & PlantUML Studio Guide (`src/modules/puml2visio`)

Developer and AI agent guide for the PlantUML IDE, PowerPoint/Visio exporters, and Office format bridge.

---

## 1. Key Components

- **PlantUML IDE & Editor (`code_editor_tab.py`, `live_preview.py`):**
  - Debounced live SVG rendering via `LivePreviewManager` + `plantuml.jar`.
  - `CodeDropTextEdit` provides `PlantUMLHighlighter`, inline `FindReplaceBar`, line numbering, soft tab indentation, comment toggles (`Ctrl+/`), live error line pinpointing, and drag-and-drop across editor and preview pane for `.puml`, `.plantuml`, `.pu`, `.iuml`, `.wsd`, `.txt` (UTF-8, UTF-8 BOM, CP1252, Latin-1 fallback + single-step `Ctrl+Z` undo) and `.vsdx`/`.vsdm` embedded source extraction.
- **Snippet & Element Inserter Engine (`plantuml_snippets.py`, `builtin_presets.py`, `snippet_model.py`, `preset_dialog.py`):**
  - Declarative catalog with dynamic participant resolution (`extract_participants`), selection-aware wrapping, toolbar quick menu + collapsible `ElementsPaletteWidget`.
  - Externalized to `src/config/plantuml_presets.json` with Git-ignored `src/config/plantuml_presets_user.json`, `UserPresetDialog` for code-free editing, hot reload without restart, standalone "UE" presets (`participant "UE" as UE`, `actor "UE" as UE`, `Dual UEs`, `UE Subcomponents`), and individual 3GPP NFs (gNB, AMF, SMF, UPF, UDM, AUSF, PCF, NRF, NSSF, NEF, UDR, AF, SEPP, eNB, MME, SGW, PGW, HSS).
- **Dual Live Preview & Java Maintenance:**
  - Side-by-side embedded SVG viewer (`EmbeddedSvgPreview` with `QSvgRenderer`/`QGraphicsView`, wheel zoom, pan) and external browser live preview (`LivePreviewManager`).
  - `JavaMaintenanceDialog` inspects Legacy vs. Modern Java runtimes and auto-updates `plantuml.jar`.
- **AI Call Flow Generator:** Modeless dialog streaming sequence diagrams from textual 3GPP descriptions via chunked local Ollama HTTP streaming (`get_ai_session()`), with 3GPP style directives, TS 23.502/38.300 presets, live SVG preview tab, syntax validation with 1-click auto-fix, and multi-turn refinement.
- **PowerPoint Direct Presentation Append (`powerpoint_converter.py`, `select_presentation_dialog.py`, `ui_tabs.py`):**
  - Appends exported diagrams to an open PowerPoint deck via **"Add to Open PowerPoint Presentation..."** (`get_open_powerpoint_presentations` COM discovery).
  - `SelectPresentationDialog` offers active deck preselection, end vs. active slide placement, search filtering, `🔄 Refresh`, `⭐ Active Deck` / `🟠 Unsaved` badges, `+ New Presentation Instead`, editable grouped Office vector shapes, 16:9 / 4:3 scaling, and PlantUML source in Speaker Notes.
- **Visio & Office Conversion Studio (`batch_convert_tab.py`, `src/modules/puml2visio/ui/cards/`):**
  - Modular cards (`PumlBatchCard`, `OfficeBridgeCard`, `VsdxExtractorCard`, `StencilsCard`, `RecentExportsStrip`):
    1. **System Ribbon:** Live Visio/PowerPoint COM status, PlantUML/Java engine badge, `output/` folder controls.
    2. **Batch PlantUML Converter:** `.vsdx`, `.pptx`, `.svg`, `.txt` export with source embedding, auto-open, and diagram type detection (`Sequence`, `Class`, `Activity`, `State`, `Component`).
    3. **Office Format Bridge:** Bidirectional `.pptx` $\leftrightarrow$ `.vsdx` via EMF bridging and text shrink-wrapping.
    4. **Visio PlantUML Extractor:** Headless `<5ms` OpenXML extraction of embedded PlantUML from 3GPP-Tools-generated `.vsdx` files with `📝 Open in PlantUML IDE Tab` routing. (Note: extraction only works on `.vsdx` files created by 3GPP Tools.)
    5. **3GPP Stencils Reference:** Searchable 5GC/NG-RAN/EPS/IMS catalog with `⭐ User` presets and 1-click insertion.
    6. **Recent Exports Strip:** `👁️ Open`, `📂 Show in Folder`, `🔗 Copy Path`.

## 2. Backward-Compatibility & COM Invariants

- **`BatchConvertTab` API:** Preserve `set_state(state, text)` (accepting both status strings and boolean flags) and attribute aliases (`drop_label_puml`, `stencils_search_edit`, `puml_table`, `bridge_table`, `stencils_table`) used by `main_window.py` and tests.
- **COM & Threading:** All Visio/PowerPoint COM automation off the GUI thread must call `pythoncom.CoInitialize()` / `CoUninitialize()` and release COM references in `finally` blocks.
