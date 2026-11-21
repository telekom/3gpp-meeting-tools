# Meetings Module Developer & AI Agent Guide (`src/modules/meetings`)

Concise architectural guide, directory map, and critical invariants for the **Meetings & TDocs** subsystem. Detailed feature specifications live in `docs/` so this file stays fast to load:

- **Meetings Browser, `TDocsWindow`, Chair's Notes & Exporters:** [`../../../docs/MEETINGS_TDOCS_AND_EXPORTERS.md`](../../../docs/MEETINGS_TDOCS_AND_EXPORTERS.md)
- **SA2 Session Plan Studio, Parsers & Delegation Model:** [`../../../docs/MEETINGS_SESSION_PLAN.md`](../../../docs/MEETINGS_SESSION_PLAN.md)
- **Cross-WG Topic Dossiers, AI Discovery & Inputs Viewer:** [`../../../docs/MEETINGS_TOPIC_DOSSIERS.md`](../../../docs/MEETINGS_TOPIC_DOSSIERS.md)
- **Tracked Email Threads, Outlook Sync & Inbox Triage:** [`../emails/AGENTS.md`](../emails/AGENTS.md)

---

## 1. Directory Structure

```text
src/modules/meetings/
├── AGENTS.md                          # This concise module guide
├── core/
│   ├── exporters/                     # Multi-format export engines
│   │   ├── dossier_exporter.py        # Per-TDoc dossier exporter
│   │   ├── excel_exporter.py          # Formatted Excel (.xlsx) & Review Matrix export
│   │   ├── llm_exporter.py            # Batch/single Markdown extraction
│   │   ├── markdown_exporter.py       # Meeting summary reports
│   │   ├── post_meeting/              # Modular Post-Meeting Dossier package (see post_meeting/README.md)
│   │   ├── post_meeting_exporter.py   # Re-export shim -> exporters/post_meeting/ package
│   │   ├── pre_meeting/               # Modular Pre-Meeting Briefing & Dossier package (see pre_meeting/README.md)
│   │   ├── pre_meeting_exporter.py    # Re-export shim -> exporters/pre_meeting/ package
│   │   ├── tdoc2md.py                 # [PROTECTED] Standalone 3GPP TDoc -> Markdown core
│   │   ├── tdocs_package_exporter.py  # Packaged TDoc archive exporter
│   │   ├── topic_digest/              # Modular Topic Digest & Brief builder (see topic_digest/README.md)
│   │   ├── topic_digest_builder.py    # Re-export shim -> exporters/topic_digest/ package
│   │   ├── topic_dossier/             # Modular Cross-WG Topic Dossier package (see topic_dossier/README.md)
│   │   └── topic_dossier_exporter.py  # Re-export shim -> exporters/topic_dossier/ package
│   ├── session_plan/                  # SA2 Session Plan parsing, diffing, DB & HTML export
│   │   ├── db/                        # Modular SessionPlanDatabase package (see db/README.md)
│   │   ├── session_plan_db.py         # Re-export shim -> session_plan/db/ package
│   │   ├── session_plan_diff.py       # Slot-by-slot snapshot diff engine
│   │   ├── session_plan_exporter.py   # Re-export shim -> session_plan/exporter/ package (see exporter/README.md)
│   │   ├── session_plan_models.py     # Domain models & boundary-aware delegate resolvers
│   │   ├── session_plan_parser.py     # Fault-tolerant Word (.docx) & Markdown parser
│   │   └── session_plan_timezone.py   # Venue vs Vienna UTC timezone resolution
│   ├── tdocs/                         # TDoc & Chairman's Notes parsing and importer threads
│   │   ├── chair_notes/               # Modular RAN1/2/3 & SA2 agreements/WA parser (see chair_notes/README.md)
│   │   ├── chair_notes_agreements_parser.py # Re-export shim -> tdocs/chair_notes/ package
│   │   ├── tdocs_parser.py            # Excel/HTML/Word TDoc list & agenda parser
│   │   └── tdocs_threads.py           # Background threads for TDoc & Word agenda loading
│   ├── agreement_tracker_worker.py    # Cross-WG Study Item & Topic agreement tracker
│   ├── chair_notes_downloader.py      # Smart FTP downloader & ZIP extractor for Chair's Notes
│   ├── cohort_resolver.py             # Multi-WG meeting round & timeframe clustering
│   ├── cross_wg_ls_tracker.py         # Inter-group Liaison Statement dialogue tracker
│   ├── meeting_wi_enricher.py         # Async Work Item discovery & enrichment
│   ├── meetings_data_loader.py        # 3gpp_data.db queries and meeting list loaders
│   ├── meetings_network_sync.py       # FTP sync for meetings and TDoc lists
│   ├── tdoc_file_handler.py           # Downloading and caching TDoc ZIP/Word documents
│   ├── tdoc_url_resolver.py           # Official 3GPP server HTTPS download URL resolution
│   ├── tdoc_markdown_converter.py     # Application adapter & Office wrappers for tdoc2md
│   ├── tdoc_triage_worker.py          # Local LLM technical triage generation
│   ├── tdocs_db.py                    # SQLite interface for user_tdocs.db
│   ├── topic_decision.py              # Unified technical decision model
│   ├── topic_manager.py               # Persistent Meta-Topic manager & rule evaluation
│   └── stats/                         # Meeting statistics & HTML dashboard compilation
└── ui/
    ├── agreement_tracker_dialog.py    # Cross-Meeting Agreement Tracker dialog
    ├── contribution_dialog.py         # Multi-meeting company contribution audit dialog
    ├── delegate_owner_picker.py       # Multi-delegate checkbox dropdown picker
    ├── dialogs.py                     # MeetingInfoDialog, AddMeetingDialog, MeetingsConfigDialog
    ├── meetings_actions.py            # Async download, caching, and export actions
    ├── meetings_delegates.py          # DynamicHoverMenuDelegate, TDocsButtonDelegate, WGTagDelegate
    ├── meetings_menu.py               # Context menu & shared add_meeting_portal_actions()
    ├── models.py                      # MeetingsTableModel (7 consolidated columns)
    ├── dialogs/                       # Export & Topic Dossier dialogs
    │   ├── post_meeting_export_dialog.py
    │   ├── pre_meeting_export_dialog.py
    │   ├── tdocs_export_controller.py
    │   └── topic_dossier_dialog.py
    ├── search_controller.py           # GlobalSearchController (Quick TDoc Jump)
    ├── session_plan/dialogs/          # Modular Timezone, Roster, Dispatch, Import & Diff dialogs (see dialogs/README.md)
    ├── session_plan_dialogs.py        # Re-export shim (ui/session_plan/session_plan_dialogs.py -> dialogs/)
    ├── session_plan_window.py         # Re-export shim -> session_plan/window/ mixins
    ├── tdocs_components.py            # CheckableComboBox and styled UI widgets
    ├── tdocs_dialogs.py               # Statistics & export configuration dialogs
    ├── tdocs_menus.py                 # Right-click context menus & centralized Quick Copy
    ├── tdocs_models.py                # TDocsTableModel & TDocsFilterProxyModel
    ├── tdocs_window.py                # Re-export shim -> tdocs/window/ mixins
    ├── meetings_tab/                  # Modular MeetingsTab package (see meetings_tab/README.md)
    └── ui_tabs.py                     # Re-export shim -> meetings_tab/ package
```

---

## 2. Critical Architectural Rules & Invariants

1. **Protected Drop-In Core (`core/exporters/tdoc2md.py`):**
   - Upstream-maintained standalone stdlib-only engine (`3gpp-tdoc-markdown-exporter`). Never add third-party dependencies, Qt imports, or Office COM calls here; place application adapters in `tdoc_markdown_converter.py`.
   - Maintain backward compatibility for v2.2.0 **Incremental All-New pCR** detection (`ALL TEXT is NEW` + `<w:ins>`/`<w:del>` deltas without `[ADDED BLOCK]:` on baseline text) and **Subscript/Superscript/Font Style** preservation (`K<sub>NF</sub>` in Markdown, `K_NF` in plain text, `<m:oMath>`).
2. **Background Threading & COM Isolation:**
   - All FTP downloads, Excel/Word parsing, COM automations, and LLM triage run in `QThread` workers (`QueueManager` or dedicated threads) with `pythoncom.CoInitialize()` / `CoUninitialize()` in `finally` blocks.
   - Phase 1/Phase 2 exporters (`llm_exporter.py`, `post_meeting_exporter.py`, `pre_meeting_exporter.py`, `topic_dossier_exporter.py`) pre-download concurrently (`ThreadPoolExecutor`) to `.part` files and convert sequentially on a single thread to isolate Word COM.
3. **TDoc List Empty-Export Rule (`tdocs_parser.py`, `ui_tabs.py`):**
   - `TDocsParser.parse_tdocs_excel` must **never persist an empty `[]` `.xlsx.json` cache** and must ignore/remove existing empty caches so re-downloaded TDoc lists are re-parsed cleanly. Empty exports log `WARNING`; exceptions log `logger.exception`.
4. **Official 3GPP URLs & Portal Links:**
   - Use `tdoc_url_resolver.py` for FTP/HTTPS file URLs and `core.utils.portal_utils` (`contributionUid=...`, never `tdocId`) for 3GU portal URLs.
5. **Strict WG Scoping for Chairman's Notes Agreements (`is_ran = 'RAN' in wg_name`):**
   - Inline text agreements, working assumptions, `[⚖️ Agreements]`, and `[⚖️]` quick-filter chips are strictly scoped to RAN WGs (`RAN1`, `RAN2`, `RAN3`, etc.). In SA/CT WGs, hide agreement controls so chair triage notes never pollute official TDoc status.
   - Downloading Chairman's Notes (`ChairNotesPickerDialog` -> `Agenda/ChairNotes/`) is enabled for **all** WGs under a download-first workflow.
6. **Session Plan GUI $\leftrightarrow$ HTML Export Parity:**
   - Keep `SessionPlanWindow` (`session_plan/window/`) and `SessionPlanExporter.export_mobile_html` visually and semantically aligned (venue-only times, half-quarter `TU 1`/`TU 2` conflict precision, 3-state `Assigned`/`Unassigned`/`Not Followed` model, `meeting_locations.json`, solid `#RRGGBB` pastel tints — **never 8-digit `#RRGGBBAA` hex codes**). Verify with `tests/test_session_plan_ui_and_html.py`.
7. **Topic Dossier Non-Destructive LLM Delta Rule (`topic_dossier_exporter.py`, `topic_digest_builder.py`):**
   - Never overwrite existing `_company_positions_analysis.md` files and never create companion `.pre_meeting.md` / `_delta.md` files; inject `STAGE1_DELTA_START`..`STAGE1_DELTA_END` blocks in-place into `_meeting_brief.md` and support dated `stage2_<YYYY-MM-DD>_rN/` packages.
8. **`CompanySanitizer` Default Exclusion:**
   - Exclude tentative/carry-over sources (`?`, `(?)`, `[...]`) by default (`include_tentative=False`) from filters, statistics, and primary `- **Source:**` Markdown headers (isolating them on `- **Unconfirmed / Carry-over Sources:**`).
9. **`TDocsWindow` Layout & Scaling:**
   - Keep header/filter bars in `FlowLayout` (`minimumSizeHint().width() <= 1050 px`), maintain `Secretary Remarks` baseline floor (`scale_px(225)`), use scoped `.QWidget { background: transparent; border: none; }` selectors, and keep `CheckableComboBox` summary-based `sizeHint()`s.
10. **Documentation Rule:** Add detailed new Meetings feature specs to the relevant `docs/MEETINGS_*.md` file, keeping this `AGENTS.md` under ~15 KB.
