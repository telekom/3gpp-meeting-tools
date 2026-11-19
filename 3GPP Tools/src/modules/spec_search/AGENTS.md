# Specification Search & Release Evolution Guide (`src/modules/spec_search`)

Developer and AI agent guide for FTS5 trigram search, release evolution matrices, clause diffs, and docked Quick Spec Search. For **Change History, Clause-to-TDoc Attribution & Origin CR Ranking**, see [`../../../docs/SPEC_HISTORY_AND_ATTRIBUTION.md`](../../../docs/SPEC_HISTORY_AND_ATTRIBUTION.md).

---

## 1. Search & Parsing Core

- **FTS5 Trigram Search (`3gpp_spec_search.db`):** Uses SQLite FTS5 configured with `tokenize="trigram"` for arbitrary substring, field, and acronym lookups.
- **Multi-Part Document Parser:** Direct `lxml` OpenXML parsing (no Word COM overhead) chaining split specification parts (`_s00_s04.docx`, `_s05_s08.docx`) into unified chronological release models (`🟢 Added`, `✓ Present`, `🔴 Removed`, `⚡ Post-Cutoff`).
- **Clause-Level Comparison (`spec_clause_word_comparator.py`, `spec_clause_diff_dialog.py`):** Provides modeless (`super().__init__(None, Qt.Window)`) in-app Visual Diff (`+`/`-`), multi-tier LLM analysis prompts, Tab 3 `"🎯 Modifying TDoc"` attribution, and `SpecClauseWordCompareThread` Word COM `CompareDocuments` Track Changes.
- **Clause Extraction (`spec_clause_exporter.py`):** Extracts single clauses from the Evolution Matrix or `SpecClauseInspector` to `.docx` (`SpecClauseWordExtractorThread`), `.md` with media (`tdoc_markdown_converter`), or dual-format clipboard (`copy_clause_to_clipboard` HTML + plain text / GFM callout / citation).

## 2. UI Architecture (`spec_search_tabs.py`, `spec_search_components.py`, `spec_search_models.py`)

- **Consolidated Header & Concept B Release Pills:** No standalone top toolbar strip; `📥 Import ▾` and `⚙️ ▾` live inside the `Indexed Specs` card header with an inline 4px progress bar. Left panel renders `SpecCardWidget` cards with wrapping `FlowLayout` `ReleasePillWidget` pills (`[R18: 18.14.0]`), tri-state master checkboxes, `⋮` card menus, real-time `tree_filter_input` (matching both spec numbers and domain acronyms like `"UAS"` or `"RRC"`), and quick chips (`[All]`, `[None]`, `[⭐ Latest/Rel]`, `[Rel-18+]`).
- **Tree Backward-Compatibility Adapters:** Preserve `_TreeCompatSpecItem`, `_TreeCompatChildItem`, `_TreeCompatRootItem`, `select_and_search_spec`, and `focus_spec`.
- **Evolution Matrix & Two-Tier Inspector:**
  - Domain tab labels (`TS 23.256 · UAS`, `TS 23.501 · 5GS Architecture`).
  - Search console with fixed `125px` date cutoff input, KPI chips, and `🎯 Find Origin CR`.
  - `SpecEvolutionStatusDelegate` renders centered rounded pill badges and `#EBF3FC` / `#0F172A` soft sky-blue selection.
  - Inspector Tier 1: context title, first-added pill, `◀ N Match(es) ▶`, layout toggle (`⬒ Bottom Split` $\leftrightarrow$ `◫ Right Pane`), maximize (`⤢`/`⤡`), heading deduplication. Tier 2: `📋 Copy ▾`, `✂️ Extract ▾`, `📖 Full Doc ▾`, `🔍 Specs Tab`, `🔍 Diff Prev`, `🤖 LLM Diff`, `📜 Version Changes`, `🎯 Origin TDoc`.
- **Docked Quick Spec Search (`QuickSpecSearchPanel`, `src/modules/spec_search/ui/quick_search_panel.py`):** Right-dock mode (<kbd>Ctrl+Shift+F</kbd>) backed by `SpecSearchQueryWorker` (`250ms` debounce, monotonic `req_id` stale-response dropping, `Latest per Spec` default toggle, `<mark>` snippet delegate, collapsible clause preview, citation copy, and deep links to `SpecSearchTab`/`SpecificationsTab`).

## 3. Release Retention & Smart Import (`spec_search_db.py`, `spec_search_threads.py`, `spec_search_dialogs.py`)

- **Optional Auto-Supersede Policy:** Enabled by default (`spec_search_config.json`), automatically purges older minor releases within the same major 3GPP release on ingest (e.g. `15.17.0` replaces `15.16.0` in SQLite + FTS5). Toggleable in `⚙️` DB menu and `SpecSearchVersionSelectDialog`; batch cleanup via **`🧹 Prune Superseded Versions...`** (`SpecSearchPruneDialog`).
- **Smart Import (`SpecSmartImportDialog`):** Groups `SpecsDatabase.get_spec_files` by major release (`extract_major_release`), compares against `SpecSearchDatabase.get_versions_for_spec`, and badges `🟠 Update (vOld → vNew)`, `🔵 New Release`, or `🟢 Up to date (vX)` with quick-select chips (`[⚡ Updates & New Only]` default). Accessible from Spec Card `⋮`, `📥 Import ▾`, and Specifications tab menus.
