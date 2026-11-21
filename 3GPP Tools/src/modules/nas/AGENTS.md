# Protocol Matrix & Inspector Guide (`src/modules/nas`)

Developer and AI agent guide for multi-protocol parsing, evolution matrix diffing, and field inspection (`3gpp_protocol_data.db`). For detailed OpenXML/ASN.1 parser engineering rules, also read [`core/parsing/AGENTS.md`](core/parsing/AGENTS.md).

---

## 1. Directory Map

```text
src/modules/nas/
├── AGENTS.md
├── core/
│   ├── nas_db.py / nas_exporter.py / nas_threads.py
│   └── parsing/                       # ProtocolDocxDispatcher + specialized parsers (see parsing/AGENTS.md)
│       ├── nas_parser.py / rrc_asn1_parser.py / ran3_asn1_parser.py / xnap_asn1_parser.py
│       ├── asn1_base_parser.py / asn1_parser.py / gtpu_parser.py / pdu_session_up_parser.py / pfcp_parser.py
│       └── protocol_parser_common.py / protocol_parser_constants.py / protocol_parser_utils.py
└── ui/
    ├── nas_tab/                       # Modular NASTab package (see nas_tab/README.md)
    ├── nas_tabs.py / nas_components.py / nas_models.py / nas_dialogs.py
    └── nas_ingest_controller.py / nas_matrix_controller.py / nas_workers.py
```

---

## 2. Supported Protocols & Parsing Rules

- **Covered Specifications:** 5GS NAS (TS 24.501), 5GS UE Policies (TS 24.526 URSP & ANDSP, TS 24.588 V2XP, TS 24.555 ProSeP, TS 24.514 RSLPP, TS 24.578 A2XP), EPS NAS (TS 24.301), NR RRC (TS 38.331), LTE RRC (TS 36.331), NGAP (TS 38.413), XnAP (TS 38.423), F1AP (TS 38.473), S1AP (TS 36.413), GTPv1-U (TS 29.281), PDU Session UP (TS 38.415), and PFCP (TS 29.244). Unified under `ProtocolDocxDispatcher`.
- **5GS NAS Annex D & UE Policies (`nas_parser.py`, `protocol_parser_utils.py`, `nas_matrix_controller.py`):**
  - *Strict Clause 8 / Annex D.5 Message Table Scoping:* `RE_CAPTION` restricts to `Table 8.x` and `Table D.5.x` with header validation in `_parse_clause_8_table`, preserving Annex D.6 IE value tables (`Table D.6.1.1`–`D.6.8.1`) inside `D.6.x` IEs.
  - *Half-Octet `<w:vMerge>` Split-Row Collapsing (`_collapse_word_half_octet_layout_rows`):* Normalizes split half-octet Word bit-diagrams (`Figure D.6.2.7`, `Figure D.6.2.3`) into single aligned octet rows.
  - *Multi-Line Coding Preservation:* Preserves `<w:br/>` and multi-paragraph boundaries inside description tables (e.g. TS 24.526 `Table 5.2.1`) as `<br/>`.
  - *UE Policy Container Unrolling (`D.6.2` / `Payload container` `9.11.3.39`):* Unrolls `0001` URSP $\to$ TS 24.526 Cl. 5.2, `0010` ANDSP $\to$ TS 24.526 Cl. 5.3.1, `0011` V2XP $\to$ TS 24.588 Cl. 5.2, `0100` ProSeP $\to$ TS 24.555 Cl. 5.2, `0101` A2XP $\to$ TS 24.578 Cl. 5.2, `0110` RSLPP $\to$ TS 24.514 Cl. 12.2, with 1-click Word/archive prompts.
- **RAN3 ASN.1 Engine (`ran3_asn1_parser.py`, `xnap_asn1_parser.py`, `asn1_base_parser.py`):**
  - *XnAP (`xnap_asn1_parser.py`):* Aligns with TS 38.423 Cl. 9.1, avoids misclassifying Cl. 9.2/9.3 IEs as messages, synthesizes `MessageType` (Cl. 9.2.3.1), maps `CPCCancel` $\to$ Cl. 9.1.2.31, parses `<w:sym>` Wingdings arrows (`node A → node B`).
  - *`SEQUENCE OF` List Unrolling & `ProtocolExtensionContainer` Resolution:* Recursively unrolls list/item fields and dereferences `-ExtIEs` object sets inline while suppressing empty `{ ... }` extension containers.
  - *Assigned Criticality & Elementary Procedures:* Maps ASN.1 criticality (`reject`/`ignore`/`notify`) into matrix cells (`M | IE | reject`), synthesizes `MessageType` from `NGAP-ELEMENTARY-PROCEDURE`, and backfills Clause 9.2 headings onto inline list/item types.
- **Message Description, Direction & RRC Disambiguation (`rrc_asn1_parser.py`, `nas_db.py`):**
  - `BaseAsn1DocxParser` appends `<Name> field descriptions` after `-- ASN1STOP` to pre-`-- ASN1START` prose (stripping `<pre>` and `RE_ASN1_BOX_CAPTION`). `RRCAsn1Parser._is_rrc_message()` excludes `*-IEs-rXX`/`-vXX`, `SetupRelease`, helpers, and lowercase inline fields while resolving versioned SIB PDUs (e.g. `SIB10-r16`).
  - *Umbrella Clause Disambiguation:* `NASDatabase.get_ie_definitions_by_clause` and `get_message_context_info` prioritize exact `ie_name` matches and reject shared umbrella-clause buckets (e.g. TS 38.331 `6.3.2`) when absent; `get_message_context_info` falls back on the fly to `ie_definitions` without requiring reimport; `📖 Desc` indexes descriptions and directions.

## 3. UI, Evolution Matrix, Document Opening & Reloading

- **Left Pane & Protocol Library (`nas_components.py`, `nas_tabs.py`, `nas_dialogs.py`):**
  - Card Frame architecture with KPI stats ribbon (`📦 Specs`, `🏷️ Versions`, `📨 Messages · IEs`) and `⚙️ Manage ▾`.
  - `NASVersionCardsWidget` (`NASSpecCardWidget` collapsible `▾`/`▸` pill drawer + `NASReleasePillWidget` `R20: 20.0.0`, preserving `NASVersionTreeWidget` test compatibility), quick chips (`All`, `None`, `⭐ Latest/Rel`, `Latest`, `Rel-18+`), WG badges, `PROTOCOL_SPEC_CATALOG` tooltips.
  - `NASMessageItemDelegate` (2-line delegate, `QFont.Normal` / `QFont.DemiBold` active, `ScrollBarAlwaysOff`, `Messages | All IEs / Types` bar).
- **Evolution Matrix & Exports (`nas_models.py`, `nas_exporter.py`):**
  - Checkable header diff filter pills (`diff_counts` + `diff_filter`), `NASMatrixTreeDelegate` guide lines, `depth_combo` (`max_depth`), auto-hidden empty `IEI / ID` columns (excluded from exports when hidden), `Rel-20\nv20.0.0` headers, unclipped `NASInspectorWidget` header (`TS XX.YYY · RNN (v...)`, hidden `usage_btn` when 0 refs), Message Context Banner (`⇄ Direction: ...`, `msg_meta_badge`: `SRB`, `RLC`, `Ch`).
  - **📋 Copy Table ▾** (Word/Outlook styled HTML `#E8F5E9`/`#FFF9C4`/`#FFEBEE`, Markdown, TSV) & **📋 Copy Definition ▾** (Word HTML, Markdown ````asn1 ````, Plain Text, Citation).
- **Spec Document Resolution & Reloading (`nas_ingest_controller.py`, `spec_doc_resolver.py`):**
  - Multi-surface `📖 Open Spec ▾` (Matrix header, Inspector, Matrix double-click, Releases tree, Protocol list) copies `Clause X.Y.Z` to clipboard and routes multi-part specs (`_s06_s08.docx`, `_s09_s10.docx`, `_sA_sHistory.docx`).
  - 1-click Version Reload, Batch Spec Group Reload, and Bulk Checked Reload use `build_spec_reload_task` + atomic `insert_parsed_spec`.
  - **QObject Collision Guard:** Always check `is_ingestion_running` before inspecting worker state (uninitialized `QWidget` instances resolve `tab.thread` to `QObject.thread()`, which lacks `.isRunning()`).
