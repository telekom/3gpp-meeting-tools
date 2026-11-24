# Visio & PlantUML Studio Guide (`src/modules/puml2visio`)

Developer and AI agent guide for the PlantUML IDE, PowerPoint/Visio exporters, and Office format bridge.

---

## 1. Directory Map

```text
src/modules/puml2visio/
├── AGENTS.md
├── core/
│   ├── ascii_converter.py / live_preview.py / prompt_manager.py / callflow_thread.py
│   ├── powerpoint_converter.py        # Native Office vector PPTX & open-deck append
│   ├── visio_converter.py             # PlantUML -> Visio (.vsdx)
│   └── pptx2visio_converter.py / visio2pptx_converter.py  # Bidirectional EMF bridge
├── templates/                         # builtin_presets.py, plantuml_snippets.py, snippet_model.py
└── ui/
    ├── editor/                        # Modular CodeDropTextEdit, LineNumberArea & FindReplaceBar (see editor/README.md)
    ├── ui_components.py               # Re-export shim -> ui/editor/ package
    ├── ui_tabs.py                     # PlantUML IDE tab (CodeDropTextEdit)
    ├── batch_convert_tab.py           # Visio & Office Conversion Studio (+ legacy aliases)
    ├── callflow/                      # Modular CallFlowDialog package (see callflow/README.md)
    ├── callflow_dialog.py             # Re-export shim -> ui/callflow/ package
    ├── java_dialog.py / preset_dialog.py / select_presentation_dialog.py
    ├── elements_palette.py / svg_preview_widget.py / syntax_highlighter.py
    └── cards/                         # puml_batch_card, office_bridge_card, vsdx_extractor_card, stencils_card, recent_exports_strip
```

---

## 2. Key Components & Invariants

- **PlantUML IDE & Editor (`ui_tabs.py`, `live_preview.py`, `syntax_highlighter.py`):**
  - Debounced live SVG rendering via `LivePreviewManager` + `plantuml.jar`.
  - `CodeDropTextEdit` provides `PlantUMLHighlighter`, inline `FindReplaceBar`, line numbering, soft tab indentation, comment toggles (`Ctrl+/`), live error line pinpointing, and drag-and-drop across editor and preview pane for `.puml`, `.plantuml`, `.pu`, `.iuml`, `.wsd`, `.txt` (UTF-8, UTF-8 BOM, CP1252, Latin-1 fallback + single-step `Ctrl+Z` undo) and `.vsdx`/`.vsdm` embedded source extraction.
- **Snippet & Element Inserter Engine (`plantuml_snippets.py`, `builtin_presets.py`, `snippet_model.py`, `preset_dialog.py`):**
  - Declarative catalog with dynamic participant resolution (`extract_participants`), selection-aware wrapping, toolbar quick menu + collapsible `ElementsPaletteWidget`.
  - Externalized to `src/config/plantuml_presets.json` with Git-ignored `src/config/plantuml_presets_user.json`, `UserPresetDialog` for code-free editing, hot reload without restart, standalone "UE" presets (`participant "UE" as UE`, `actor "UE" as UE`, `Dual UEs`, `UE Subcomponents`), and individual 3GPP NFs (gNB, AMF, SMF, UPF, UDM, AUSF, PCF, NRF, NSSF, NEF, UDR, AF, SEPP, eNB, MME, SGW, PGW, HSS).
- **Dual Live Preview & Java Maintenance:**
  - Side-by-side embedded SVG viewer (`EmbeddedSvgPreview` with `QSvgRenderer`/`QGraphicsView`, wheel zoom, pan) and external browser live preview (`LivePreviewManager`).
  - `JavaMaintenanceDialog` (`java_dialog.py`) inspects Legacy vs. Modern Java runtimes and auto-updates `plantuml.jar`.
- **AI Call Flow Generator (`callflow_dialog.py`, `callflow_thread.py`):** Modeless (`super().__init__(None)`, `Qt.Window`) dialog streaming sequence diagrams from textual 3GPP descriptions via chunked local Ollama HTTP streaming (`get_ai_session()`), with 3GPP style directives, TS 23.502/38.300 presets, live SVG preview tab, syntax validation with 1-click auto-fix, and multi-turn refinement.
- **PowerPoint Direct Presentation Append (`powerpoint_converter.py`, `select_presentation_dialog.py`, `ui_tabs.py`):**
  - Appends exported diagrams to an open PowerPoint deck via **"Add to Open PowerPoint Presentation..."** (`get_open_powerpoint_presentations` COM discovery).
  - `SelectPresentationDialog` offers active deck preselection, end vs. active slide placement, search filtering, `🔄 Refresh`, `⭐ Active Deck` / `🟠 Unsaved` badges, `+ New Presentation Instead`, editable grouped Office vector shapes, 16:9 / 4:3 scaling, and PlantUML source in Speaker Notes.
- **Visio & Office Conversion Studio (`batch_convert_tab.py`, `ui/cards/`):**
  1. **System Ribbon:** Live Visio/PowerPoint COM status, PlantUML/Java engine badge, `output/` folder controls.
  2. **Batch PlantUML Converter (`puml_batch_card.py`):** `.vsdx`, `.pptx`, `.svg`, `.txt` export with source embedding, auto-open, and diagram type detection (`Sequence`, `Class`, `Activity`, `State`, `Component`).
  3. **Office Format Bridge (`office_bridge_card.py`):** Bidirectional `.pptx` $\leftrightarrow$ `.vsdx` via EMF bridging and text shrink-wrapping.
  4. **Visio PlantUML Extractor (`vsdx_extractor_card.py`):** Headless `<5ms` OpenXML extraction of embedded PlantUML from **3GPP-Tools-generated `.vsdx` files only** with `📝 Open in PlantUML IDE Tab` routing.
  5. **3GPP Stencils Reference (`stencils_card.py`):** Searchable 5GC/NG-RAN/EPS/IMS catalog with `⭐ User` presets and 1-click insertion.
  6. **Recent Exports Strip (`recent_exports_strip.py`):** `👁️ Open`, `📂 Show in Folder`, `🔗 Copy Path`.
- **Backward-Compatibility & COM Invariants:**
  - Preserve `BatchConvertTab.set_state(state, text)` (accepting both status strings and boolean flags) and attribute aliases (`drop_label_puml`, `stencils_search_edit`, `puml_table`, `bridge_table`, `stencils_table`) used by `main_window.py` and tests.
  - All Visio/PowerPoint COM automation off the GUI thread must call `pythoncom.CoInitialize()` / `CoUninitialize()` and release COM references in `finally` blocks.
