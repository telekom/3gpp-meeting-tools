# Specification Change History, Clause-to-TDoc Attribution & Origin CR Ranking

Deep architectural reference for change history ingestion, TDoc cross-referencing, OpenXML clause fingerprinting, and feature-level origin CR ranking (`src/modules/spec_search` & `src/modules/specifications`).

---

## 1. Specification Change History & Integrated TDocs
(`spec_history_db.py`, `spec_history_extractor.py`, `spec_tdoc_opener.py`, `spec_history_dialog.py`, `spec_tdoc_report.py`, `spec_tdoc_report_dialog.py`)

- **Isolated Storage (`3gpp_spec_history.db`):** Dedicated SQLite database decoupling historical CR/p-CR ledgers from core `3gpp_data.db`. Features master ledger `spec_change_history` and $O(1)$ cross-reference index `spec_tdoc_cross_ref`. Registered in `DatabaseMaintenanceDialog` for real-time size tracking and 1-click database wiping with zero risk to personal notes or specifications catalog.
- **High-Speed OpenXML Ingestion (<0.35s / spec):** Direct headless `lxml` streaming parser extracting full change history annexes across modern and legacy templates (TS 24.501, TS 38.331, TS 29.244, TS 31.102). Handles multi-part documents (`_sAnnexes.docx`), chained multi-table annexes (TS 24.007), paragraph-aware cell joining, `<w:gridSpan>` expansion, date carry-forward across blank rows, meeting repairs (`#112` -> `CT#112`), and draft version inference.
- **Working Group Pseudo-CR (p-CR) Tokenizer:** Audited regex `(?<![A-Za-z-])([A-Z][A-Z0-9]{1,4}-\d{5,7}(?:[rR]\d{1,2}(?!-))?)(?![0-9])` tolerant of concatenated rapporteur typos (e.g. `R3-261549R3-261553`), extracting 150+ p-CRs per draft row into first-class indexed entities with revision preservation.
- **Direct TDoc Document Opening & Meetings Cross-Linking:** 1-click opening of referenced TDoc documents (`.docx`, `.doc`, `.zip`) from `SpecVersionHistoryDialog`, `SpecClauseInspector`, and `SpecClauseDiffDialog`. Resolves against local meeting cache, automatically downloads missing archives from 3GPP servers via `MeetingsDatabase` (`docs_folder_url`) and `URLRouter`, provides 1-click **"🚀 Jump to TDoc in Meetings Tab"**, and formats plain text / rich HTML clipboard citations.
- **Interactive Notes & Inline Clickable TDoc Badges (`spec_history_dialog.py`):** Merges separate text fields and redundant button scroll areas into a unified, resizable `QTextBrowser`. Automatically parses rapporteur notes and renders all integrated TDocs, p-CRs, and revisions as clickable inline pill badges (`#E0F2FE`). Left-click instantly opens target Word/ZIP documents; right-click provides quick actions (open on 3GPP Portal 3GU, jump to Meetings tab, copy citation, copy identifier).
- **On-Demand TDoc Metadata Dossier & Report Generator (`spec_tdoc_report.py`, `spec_tdoc_report_dialog.py`):** Resolves full technical metadata (Title, Submitting Organizations/Companies, Agenda Item, Type, Approval Status, Technical Abstract / Proposed Solution, Work Item, and Secretary Remarks) directly from meeting agenda spreadsheets (`.xlsx`) and cached JSON without downloading the actual TDoc documents. Features an interactive modeless viewer dialog (`SpecTDocReportDialog`) with real-time multi-facet filtering (by company, type, status, keyword), technical solution drawer, Markdown dossier export, and CSV spreadsheet generation.
- **Modeless Multi-Window Usability:** `SpecVersionHistoryDialog` and `SpecTDocReportDialog` configure `Qt.Window` and `Qt.NonModal` window modality (`super().__init__(None)`).
- **Zero-Lag UI Performance & Logging:** All change history operations emit descriptive, emoji-badged log records to `ConsolePanel`. Single aggregate pre-fetch in `SpecificationsTab.refresh_table()` eliminates $N+1$ query overhead while badging version dropdown items (`📜 {cnt} CRs/p-CRs`). Accessible from Column 2 `Get Word ▾` menu, Column 0 context menu, `SpecInfoDialog`, `SpecSearchTab` Tier 2 toolbar, Evolution Matrix context menu, and clause diff dialog. All Qt actions escape literal ampersands (`&&`).

---

## 2. 3GPP Clause-to-TDoc Attribution Engine & OpenXML Fingerprinting
(`spec_clause_attribution.py`, `spec_clause_diff_dialog.py`, `spec_search_components.py`)

- **Workflow Grounding:** Pinpoints which specific Change Request (CR) or Working Group Pseudo-CR (p-CR) introduced or modified a clause across specification releases, authoring companies, and technical rationale.
- **Mode A vs Mode B Dichotomy:**
  - *Mode A (Normative TS Plenary CRs):* Individual CR rows in `spec_change_history` mapped to formal CR numbers, categories, and multi-file CR packs (`SP-xxxxxx.zip`, extracting only target `*CR<num>*.docx`).
  - *Mode B (Study TR Working Group p-CRs):* Resolves omnibus p-CR lists in study TRs where contributions originally used placeholder numbering (`6.x`, `Solution #X`) before final numbers were assigned by rapporteurs.
- **Ancestor Context Inheritance & The Key Issue Bridge:**
  - Generic leaf headings (`6.20.2.1 General`) automatically inherit full architectural titles by traversing `parent_clause` in `3gpp_spec_search.db` (`6.20 Solution #20: Protocol Architecture for Ambient IoT`).
  - The Key Issue Bridge links TR solution body declarations (`"addresses Key Issue #4"`) to candidate p-CRs via meeting agenda items (`AI 20.6.4`) and title/abstract citations (`KI #4`).
- **Multi-Stage Progressive Attribution Pipeline:**
  - *Stage 1 (Zero-Download Metadata & Structural Ranking, <50ms):* Scans rapporteur notes, matches Key Issue / Solution patterns, computes cleaned semantic title overlap, and scores salient diff keywords against persistent SQLite agenda metadata (`tdoc_metadata_cache` in `3gpp_spec_history.db`).
  - *Stage 2 (Track-Changes & Placeholder-Agnostic Content Fingerprinting):* Inspects candidate Word OpenXML documents (`word/document.xml`), excludes deleted `<w:del>` subtrees when matching added (`+`) diff lines, normalizes numbers (`6.x` / `6.20` -> `<NUM>`), and computes exact long-sentence and 6-word shingle containment for deterministic ground truth (`✅ Verified in DOCX (98% Text Match)`).
  - *Stage 3 (Local Ollama Grounding):* Feeds local Ollama with clause diffs and matched excerpts for multi-pCR attribution and synthesis.
- **UI Integration:**
  - *`SpecClauseDiffDialog` Tab 3 ("🎯 Modifying TDoc"):* Top candidate hero banner with 1-click `📄 Open TDoc`, `🚀 Jump to Meetings`, and `📋 Copy Citation`, ranked candidate table with confidence pills (`Verified`, `Very High`, `High`, `Moderate`), details drawer, and background `🔬 Deep Verify Top 3 (DOCX)` worker.
  - *`SpecClauseInspector` Tier 2 Action Toolbar:* Dedicated `🎯 Origin TDoc` action button opening the clause diff dialog directly focused on attribution.
  - *Modeless Usability:* `SpecClauseDiffDialog` runs modelessly (`super().__init__(None)`, `Qt.Window`).

---

## 3. Feature-Level Origin CR Ranking & Non-Exclusive Affected Clause Filter
(`spec_clause_attribution.py`, `spec_history_extractor.py`, `spec_history_dialog.py`, `spec_search_tabs.py`)

- **The Feature-Level Origin Challenge:** Discovers which foundational Change Request (CR) originally introduced an entire feature, technology, or architecture concept across all specification releases (e.g. `CR 2679 (Cat B)` in `SP-210083` introducing `NSAC`/`NSACF` into `TS 23.501` at `v17.0.0` / `SP#91E` citing `KI#1`).
- **Headless OpenXML Cover Sheet Extractor (`extract_cr_cover_sheet_metadata`):** High-speed stream extraction from `.docx` and `.zip` CR packages parsing `"Clauses affected"` tables, formal Category (`Cat B`, `Cat C`, `Cat F`), CR number, spec number, and work item codes.
- **Multi-Signal Origin CR Ranking Matrix (`rank_origin_crs_for_query`):**
  - *Signal 1 (Title / Subject Query Match & Origin Intent, Up to 45 pts):* Exact acronym and term matches, boosted by feature introduction keywords (`"definition"`, `"introduction"`, `"adding"`, `"KI#"`) and penalizing maintenance keywords (`"correction"`, `"clean-up"`, `"alignment"`).
  - *Signal 2 (3GPP Category Weighting, Up to 25 pts):* `Category B` (Addition of feature) heavily weighted (+25) over `Category C` (+15) and `Category F` (+5).
  - *Signal 3 (Additive, Non-Exclusive Affected Clauses Alignment, Up to 30 pts):* Rewards candidate CRs matching affected clauses from the search hit list (+25 exact, +15 prefix), while strictly ensuring that CRs with omitted cover sheet metadata are NEVER penalized or excluded.
  - *Signal 4 (Chronological Release Primacy, Up to 20 pts):* Rewards the earliest release baseline (`vX.0.0`, `vX.1.0`) and earliest plenary meeting in the release cycle.
  - *Role Assignment:* Categorizes candidates into `🎯 Foundational Origin CR`, `🏗️ Architecture Expansion`, and `🔧 Maintenance / Correction`.
- **UI Connections:**
  - *Spec Search (`🎯 Find Origin CR`):* 1-click launch button in the search console bar pre-setting the active search query, clause, and enabling origin ranking across all releases.
  - *`SpecVersionHistoryDialog` Enhancements:* Pre-populates search and clause filters, adds dual-mode clause filtering (`Rank by Clause (Additive)` vs `Strict Clause Filter`), dynamically switches to 9 columns (`Relevance` with role/score badges and tooltips), and renders an executive role & evidence banner in the interactive details drawer.
