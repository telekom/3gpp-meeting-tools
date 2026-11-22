# Specifications Catalogue Guide (`src/modules/specifications`)

Developer and AI agent guide for the 3GPP Specifications Catalogue, inline filtering, FTP sync, clause extraction, and tag annotations. For **Change History, Clause-to-TDoc Attribution & Origin CR Ranking** (`3gpp_spec_history.db`, `spec_history_*.py`, `spec_clause_attribution.py`), see [`../../../docs/SPEC_HISTORY_AND_ATTRIBUTION.md`](../../../docs/SPEC_HISTORY_AND_ATTRIBUTION.md).

---

## 1. Directory Map

```text
src/modules/specifications/
├── AGENTS.md
├── core/
│   ├── database.py / scraper.py               # 3gpp_data.db catalogue + 3GPP FTP sync
│   ├── spec_markdown_converter.py             # Mode 1 (unified) & Mode 2 (clause-split) Markdown export
│   └── spec_history_db.py / spec_history_extractor.py / spec_clause_attribution.py / spec_tdoc_opener.py / spec_tdoc_report.py
└── ui/
    ├── specs_tab/                             # Modular SpecificationsTab package (see specs_tab/README.md)
    ├── ui_tabs.py / components.py / spec_menu.py / spec_action_controller.py / threads.py
    ├── dialogs.py / spec_info_dialog.py / spec_annotations_dialog.py / spec_clause_dialog.py
    ├── spec_history/                          # Modular SpecVersionHistoryDialog package (see spec_history/README.md)
    ├── spec_history_dialog.py                 # Re-export shim -> spec_history/ package
    └── spec_tdoc_report_dialog.py
```

---

## 2. Key Capabilities & Entry Points

- **Unified Inline Filter Engine (`ui_tabs.py`, `tdocs_components.py`, `database.py`):**
  - Multi-selection `CheckableComboBox` for Series, Working Groups (WG), and Radio Technologies (`Tech`) backed by SQL `IN (...)` queries.
  - TSG-to-WG scoping (`SA`, `RAN`, `CT`) and dual-mode scope toggle (`Active WGs (18)` $\leftrightarrow$ `Full List (33)`): Active isolates the 18 active groups (SA1–6, RAN1–5, CT1/3/4/6, Plenaries); Full List orders active groups first (SA $\to$ RAN $\to$ CT) followed by disbanded/legacy groups (RAN6, GERAN1–3, CN1–5, T2–3, SMG, AHG, ITU).
  - Anchors `(Select All)` at top on open, provides `WG_DESCRIPTIONS` tooltips, `All`/`TS`/`TR` segmented buttons, and dismissible active filter pills.
- **Working Group Domain Badges:** Color-coded badges in the Specification column (SA indigo, RAN emerald, CT amber); clicking a badge filters immediately to that WG.
- **Document Resolution & Download (`spec_doc_resolver.py`, `SpecDownloadWorker`):** Discovers local `.docx`, `.doc`, `.pdf`, `.zip`, auto-unzips on the fly (skipping `-rm` revision-mark files), routes multi-part specs (`match_multipart_document_for_clause`), copies `Clause X.Y.Z` to clipboard for Word `Ctrl+F`, and downloads asynchronously.
- **Direct Version Indexing into Spec Search:** Row context menu, Column 0 `⋮` (`HoverMenuButton`), and Column 2 `Get Word ▾` expose `Index vX.Y.Z` / `Re-index vX.Y.Z` and batch `Index selected (N)`, using `is_version_indexed` and existing local Word files to avoid redundant FTP downloads.
- **Quick Copy & `HoverMenuButton` (`spec_menu.py`, `components.py`, `spec_info_dialog.py`):** `Ctrl+C` and menus export dual-format `QMimeData` (plain text + rich HTML hyperlinks for Outlook/Word/Teams). `HoverMenuButton` uses recursive menu inspection (`is_pos_inside`), diagonal corridor bounding hulls, grace-period debouncing, and click-to-pin so nested submenus never dismiss prematurely.
- **User Annotations & Tags (`spec_annotations_dialog.py`, `src/core/ui/flow_layout.py`):** Stored in `spec_user_annotations` inside `3gpp_data.db`. `EditSpecAnnotationsDialog` and `BatchTagDialog` use wrapping `FlowLayout` for `TagChip` pills and live-prioritized suggestions.
- **Specification Markdown Export (`spec_markdown_converter.py`):** Mode 1 (unified doc) and Mode 2 (clause-split dossier + `index.md` TOC) with version-isolated media, multi-part sequencing, and legacy `.doc` handling.
- **Interactive Clause Extraction (`spec_clause_dialog.py`, `spec_clause_exporter.py`):** Modeless floating window (`super().__init__(None)`, `Qt.Window`, `Qt.NonModal`). `WordExcerptThread` exports simultaneous `.docx` (`WordHeadingScanner` + `WordExcerptEngine` with `python-docx` fallback) and `.md` (`convert_document_to_markdown` + `{filename}_media/`).
- **Change History Performance:** `SpecificationsTab.refresh_table()` runs a single aggregate pre-fetch to badge version dropdowns (`📜 {cnt} CRs/p-CRs`) with zero $N+1$ query overhead.

## 3. Cross-Tab Deep Linking & DB Safety

- Preserve signals `navigate_to_specs_tab_requested`, `search_in_spec_search_requested`, `index_in_spec_search_requested`, `batch_index_in_spec_search_requested`, and deep-link helper `SpecificationsTab.select_and_show_spec`.
- Never execute destructive Git commands against `src/db/3gpp_data.db`.
