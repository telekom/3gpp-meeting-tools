# Specification Search & Release Evolution Guide (`src/modules/spec_search`)

Developer and AI agent guide for FTS5 trigram search, release evolution matrices, clause diffs, and docked Quick Spec Search. For **Change History, Clause-to-TDoc Attribution & Origin CR Ranking**, see [`../../../docs/SPEC_HISTORY_AND_ATTRIBUTION.md`](../../../docs/SPEC_HISTORY_AND_ATTRIBUTION.md).

---

## 1. Directory Map

```text
src/modules/spec_search/
├── AGENTS.md
├── core/
│   ├── spec_search_db.py              # 3gpp_spec_search.db + FTS5 trigram index + retention pruning
│   ├── spec_text_extractor.py         # Direct lxml multi-part OpenXML clause extractor
│   ├── spec_doc_resolver.py           # Shared spec archive discovery, unzip, multi-part routing
│   ├── spec_clause_diff.py / spec_clause_word_comparator.py / spec_clause_exporter.py
│   └── spec_search_threads.py
└── ui/
    ├── spec_search_tab/               # Modular SpecSearchTab package (see spec_search_tab/README.md)
    ├── spec_search_tabs.py            # Re-export shim -> spec_search_tab/ package
    ├── components/                    # Modular Version Cards, Delegate & Clause Inspector (see components/README.md)
    ├── spec_search_components.py      # Re-export shim -> components/ package
    ├── quick_search_panel.py          # Docked QuickSpecSearchPanel (Ctrl+Shift+F)
    ├── dialogs/                       # Modular Selection, Prune & Smart Import dialogs (see dialogs/README.md)
    ├── spec_search_dialogs.py         # Re-export shim -> dialogs/ package (+ SpecClauseDiffDialog alias)
    ├── clause_diff/                   # Modular SpecClauseDiffDialog package (see clause_diff/README.md)
    ├── spec_clause_diff_dialog.py     # Re-export shim -> clause_diff/ package
    └── spec_search_models.py
```

---

## 2. Search, Parsing & Comparison Core

- **FTS5 Trigram Search (`3gpp_spec_search.db`):** Uses SQLite FTS5 configured with `tokenize="trigram"` for arbitrary substring, field, and acronym lookups.
- **Multi-Part Document Parser (`spec_text_extractor.py`):** Direct `lxml` OpenXML parsing (no Word COM overhead) chaining split specification parts (`_s00_s04.docx`, `_s05_s08.docx`) into unified chronological release models (`🟢 Added`, `✓ Present`, `🔴 Removed`, `⚡ Post-Cutoff`), supporting `get_first_added_version` and `get_preceding_version`.
- **Clause-Level Comparison (`spec_clause_word_comparator.py`, `spec_clause_diff_dialog.py`):** Modeless (`super().__init__(None)`, `Qt.Window`, `Qt.NonModal`) in-app Visual Diff (`+`/`-`), multi-tier LLM analysis prompts, Tab 3 `"🎯 Modifying TDoc"` attribution, and `SpecClauseWordCompareThread` Word COM `CompareDocuments` Track Changes.
- **Clause Extraction (`spec_clause_exporter.py`):** Extracts single clauses from the Evolution Matrix or `SpecClauseInspector` to `.docx` (`SpecClauseWordExtractorThread`), `.md` with media (`tdoc_markdown_converter`), or dual-format clipboard (`copy_clause_to_clipboard` HTML + plain text / GFM callout / citation).

## 3. UI Architecture & Release Retention

- **Consolidated Header & Concept B Release Pills (`spec_search_tabs.py`, `spec_search_components.py`):** No standalone top toolbar strip; `📥 Import ▾` and `⚙️ ▾` live inside the `Indexed Specs` card header with an inline 4px progress bar. Left panel renders `SpecCardWidget` cards with wrapping `FlowLayout` `ReleasePillWidget` pills (`[R18: 18.14.0]`), tri-state master checkboxes, `⋮` card menus, real-time `tree_filter_input` (`filter_tree` matching both spec numbers and domain acronyms like `"UAS"` or `"RRC"`), and quick chips (`[All]`, `[None]`, `[⭐ Latest/Rel]`, `[Rel-18+]`).
- **Tree Backward-Compatibility Adapters:** Preserve `_TreeCompatSpecItem`, `_TreeCompatChildItem`, `_TreeCompatRootItem`, `select_and_search_spec`, and `focus_spec`.
- **Evolution Matrix & Two-Tier Inspector:** Domain tab labels (`TS 23.256 · UAS`, `TS 23.501 · 5GS Architecture`), search console (`125px` cutoff edit, KPI chips, `🎯 Find Origin CR`), `SpecEvolutionStatusDelegate` (`#EBF3FC` / `#0F172A`), and two-tier inspector (`⬒ Bottom Split` $\leftrightarrow$ `◫ Right Pane`, `⤢`/`⤡`, `📋 Copy ▾`, `✂️ Extract ▾`, `📖 Full Doc ▾`, `🔍 Specs Tab`, `🔍 Diff Prev`, `🤖 LLM Diff`, `📜 Version Changes`, `🎯 Origin TDoc`).
- **Docked Quick Spec Search (`QuickSpecSearchPanel`):** Right-dock mode (<kbd>Ctrl+Shift+F</kbd>) backed by `SpecSearchQueryWorker` (`250ms` debounce, monotonic `req_id` stale-response dropping, `Latest per Spec` default toggle, `<mark>` snippet delegate, collapsible clause preview, citation copy, and deep links to `SpecSearchTab`/`SpecificationsTab`).
- **Release Retention & Smart Import (`spec_search_config.json`, `SpecSearchPruneDialog`, `SpecSmartImportDialog`):** Optional auto-supersede policy (default enabled) purges older minor releases within the same major release on ingest (`15.17.0` replaces `15.16.0`); `SpecSmartImportDialog` badges `🟠 Update (vOld → vNew)`, `🔵 New Release`, or `🟢 Up to date (vX)` across major releases (`[⚡ Updates & New Only]` default).
