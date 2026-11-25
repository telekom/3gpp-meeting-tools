# Meetings Browser, TDocs Window, Chair's Notes & Exporters Reference

Deep architectural reference for the Meetings & TDocs workflows (`src/modules/meetings`). For Session Plans see [`MEETINGS_SESSION_PLAN.md`](MEETINGS_SESSION_PLAN.md); for Cross-WG Topic Dossiers see [`MEETINGS_TOPIC_DOSSIERS.md`](MEETINGS_TOPIC_DOSSIERS.md); for Tracked Email Threads & Inbox see [`../src/modules/emails/AGENTS.md`](../src/modules/emails/AGENTS.md).

---

## 1. Meetings Browser (`MeetingsTab` in `ui_tabs.py`, `models.py`, `meetings_delegates.py`)

- **Full-Width Table & 2-Row Top Ribbon:**
  - *Row 1:* `🚀 Last Meeting`, inline `⚡ TDoc:` jump box (`📄 Doc`, `🗓️ Mtg`, `⚖️ vs Word`, `📋 Copy ▾` via centralized `populate_quick_copy_menu`), right-aligned `🔄 Sync ▾` (*Quick Sync — All WGs*, dynamic *Sync \<Selected WGs\> Only*, *Custom / Deep Sync...* via `MeetingsSyncDialog`, *Add / Fetch Meeting...*, *Cache & Preferences*), `📊 Reports ▾`, and `⚙️ Settings`. An inline `sync_status_banner` (`MeetingsTabSyncMixin`) displays live step progress (`Step 1/3`..`Step 3/3`) and a cooperative **Stop** button (`QueueManager.abort_task_format`) above the table during syncs.
  - *Row 2 (Filter Engine):* `🔍 Search`, TSG toggles (`[All] [SA] [RAN] [CT]`), `Working Group` checkable combo, Nature (`Regular`/`Ad-Hoc`), Format (`In-Person`/`Electronic`), and Date Presets (`All Dates`, `Upcoming`, `This Year`, `Past 12 Months`, `Custom Range...` revealing calendar pickers, and `Unknown Dates` isolating undated folders so they never pollute dated/upcoming views).
- **Fast-First 3-Step Sync Pipeline (`MeetingsCrawlerThread`, `MeetingsDatabase`):**
  - *Step 1/3 (FTP Discovery):* Maps WG folders (with `MEETING_DENY_LIST` filtering out `TSGR*_AH*` container directories, hyphenated `-bis`/`-AH`/`-Ad_Hoc`/`-e` meeting-number normalization, and folder-date fallback e.g. `_2027-01` $\rightarrow$ `2027-01-01`), links any DynaReport-only placeholders by `(wg_id, meeting_number)`, and cleans up legacy container rows.
  - *Step 2/3 (DynaReport Metadata):* Runs immediately after Step 1 (~3s) so newly created meetings receive official dates/locations/`MtgId` right away, normalizes `/Invitation`/`/docs` suffixes off `url_key`, and inserts scheduled future meetings that do not yet have an FTP folder.
  - *Step 3/3 (Smart `Docs/` Scan):* By default (`docs_mode="smart"`), scans `Docs/` TDoc bounds only for recent (12 months), upcoming, newly discovered, or unindexed meetings; `docs_mode="all"` (full historical re-scan) is explicitly opt-in in `MeetingsSyncDialog`.
- **7-Column Table & Delegates:** `Action` (`TDocsButtonDelegate`: `✔ Open`, `⬇ Get`, `⏳ Fetching`), `WG` (`WGTagDelegate`: SA indigo, RAN emerald, CT amber via `get_wg_colors()`), `Location` (`📍 City, Country` vs `🌐 Online (eMeeting)`), `Dates` (`format_meeting_dates()`, e.g. `12 – 16 Oct 2026`, with temporal tooltip), `Allocated TDocs` (`C4-264009 → C4-264023`, `—` when unallocated). Double-clicking opens the meeting viewer.
- **Floating Comparison Cart (`ComparisonCartWidget`, `ComparisonManager`):** Shared lower-right status bar widget in `MeetingsTab` and `TDocsWindow`. Collapses to `⚖️ Stage Active Word ▾` when empty; expands to `[Slot A ▾] ⇄ [Slot B ▾] [🪟 Active Word] [Compare in Word] [Clear]`. Auto-re-snapshots modified open Word docs (`resolve_comparison_paths()`) before launching `WordComparatorThread`, and supports 1-click **Active Word vs. TDoc** comparison from Quick Jump, Instant Fetch, row menus, and Related TDoc hyperlinks.
- **Maintenance:** `Clear All Meetings...` lives in `MeetingsConfigDialog` (*⚠️ Database Maintenance (Danger Zone)*).

---

## 2. `TDocsWindow` Layout, Ordering, Caching & Ergonomics
(`tdocs/window/` package, `tdocs_window.py` shim, `tdocs_models.py`, `tdoc_delegates.py`, `tdocs_components.py`, `tdocs_cacher.py`)

- **Responsive Wrapping Header & Filter Bars (`FlowLayout`):** `_setup_header()` and `_setup_filters()` use `FlowLayout` (`QSizePolicy.Preferred`/`Maximum`, `setHeightForWidth(True)`) keeping `minimumSizeHint().width() <= 1050 px`. `count_lbl` and `last_mod_lbl` live in `_footer_status_widget` (left of `ComparisonCartWidget`).
- **Word-Wrapped Delegates & Balanced Column Budget (`_apply_scaled_table_columns`):**
  - `HtmlDelegate` (`Secretary Remarks`, `Related TDocs`) constrains `QTextDocument.setTextWidth()` cleanly in `paint()`, `sizeHint()`, and `editorEvent()`; `_adjust_table_row_heights()` (debounced via `_schedule_row_height_adjustment()`) adapts row heights between `scale_px(48)` and `scale_px(124)`.
  - `_apply_scaled_table_columns()` keeps non-icon columns in `Interactive` mode with calibrated widths (`Secretary Remarks` floor `scale_px(225)`), distributing horizontal surplus `55%` Remarks / `35%` Title / `10%` Source while respecting `_user_resized_cols`.
- **Scoped Filter Containers, `CheckableComboBox` & Per-Meeting Filter Persistence:** Filter containers use `.QWidget { background: transparent; border: none; }`; `CheckableComboBox` renders via `CheckableComboItemDelegate` on `#checkableComboPopupView` (flattening newlines in display labels while keeping raw `Qt.UserRole`), computes `sizeHint()` from summary strings, supports empty-selection/empty-string `(Empty)` restoration (`setCheckedItems(..., allow_empty=True)`), and `CheckableComboBox.updateItems` preserves selections on comment save (`super().setCurrentIndex(0)`). `sub_item_combo` includes `"(No Sub-Item)"` whenever rows lack a `Sub-Item` so unfiltered `Sub-Item: All` never hides ungrouped rows. `TDocsWindow` persists per-meeting filters (`search_input`, quick toggles, `No Comments`, all 8 `CheckableComboBox` facets, `order_combo`, `row_mode_combo`) in `QSettings("3GPP_Tools", "TDocsWindow")` (`_save_filter_settings` / `_load_filter_settings`, flushed in `closeEvent`), re-applying pending combo filters across async cache/enrichment refreshes (`_apply_pending_combo_filter`); `MeetingsTab.shutdown()` flushes any active `save_filters_timer` on exit.
- **Agenda & Sub-Item Ordering Engine (`tdocs_parser.py`, `AgendaItemDelegate`):**
  - 3-tier precedence: (1) Live `TdocsByAgenda.htm` / Word Chairman's Notes (`_agenda_order` / `_sub_item_order`); (2) Excel `"TDoc sort order within agenda item"`; (3) Revision anchor inheritance (`base_order + rev_offset`); fallback alphanumeric.
  - Precomputed $O(1)$ sort tuples (`_precompute_agenda_sort_tuples`) sort 1,800+ rows in `< 1ms`. `AgendaItemDelegate` renders `Sub-Item` as a rounded `#F1F5F9` pill below the AI number inside 48px rows; `Order By ▾` syncs bi-directionally with column headers.
- **Tracked Work Items Quick Filters & Auto-Enrichment (`meeting_wi_enricher.py`):**
  - `TDocsTableModel._precompute_row_wi_tokens` normalizes `Related WIs`, `agenda_map` acronyms, and description parentheticals (stripping WG suffixes `-ARC`/`-CT` and `FS_` prefixes). `TDocsFilterProxyModel` evaluates `btn_followed_wis` (`🎯`) and `btn_starred_wis` (`⭐`) with `include_unassigned_ls_in=True` (right-click toggleable, persisted in `meetings_config.json`).
  - `MeetingWIEnricherThread` runs bounded openpyxl scans for `workitemId=(\d+)`, skips existing `WorkItemsDatabase` codes, scrapes missing WIs, and caches `<excel_path>.wi_checked.json` for ~0ms repeat opens.
- **Personal Notes, Stance, Tags & Review Matrix (`user_tdocs.db`, `excel_export_dialog.py`):**
  - `InteractiveNotesDialog` shares `TagCloudWidget` personal tags (`🏷️`) and stances (`⚪ Neutral`, `🔵 My TDoc`, `🟢 Support`, `🔴 Object`, `🟡 Monitor`) uniformly across TDocs and Tracked Email Threads (`THR-...`).
  - `ExcelExportDialog` **📋 Review Matrix** preset exports `Status (<Company>)` / `Comments (<Company>)` with soft status tints and 42-char wrapped comment widths.
- **Parallel TDoc Caching (`TDocsCacherThread`), Quick Open & Email Send:**
  - `ThreadPoolExecutor(max_workers=5)` pre-fetches `Docs/` into `{meeting_dir}/{base_tdoc}/{filename}.zip` (`is_tdoc_cached` checks both direct and legacy `TDocs/` dirs).
  - `TDocsWindow._reveal_tdoc_in_table` scrolls/highlights opened TDocs (`TDocsTableModel.set_highlighted_tdoc`, clearing filters if hidden); `_send_tdoc_via_email` (`TDocFileHandler.find_local_tdoc_files`, `OutlookClient.create_email_draft`) opens Classic Outlook COM drafts with attachments (with `mailto:` fallback).
- **In-App Help (`TDocsHelpDialog`, `F1` / `❓ Help`) & `make_rich_tooltip`:** 4-tab modeless quick guide (`Toolbar && Tools`, `Table && Double-Clicks`, `Search && Filtering`, `Keyboard Shortcuts`).

---

## 3. Unconfirmed / Tentative Source Company Handling (`CompanySanitizer`)
(`company_sanitizer.py`, `src/config/company_synonyms.json`, `company_synonyms_user.json`, `tdoc_delegates.py`)

- **Two-Tier Catalog:** Factory definitions in `src/config/company_synonyms.json` (`pattern`, `full_name`, `category`, `description`) + Git-preserved `company_synonyms_user.json` overrides + zero-I/O `_BUILTIN_SYNONYMS_PATTERNS` fallback. `CompanySanitizer.reload_rules()` recompiles the C-accelerated 3-stage matcher and clears `@lru_cache(maxsize=65536)`.
- **Lossless Segment Tokenizer (`tokenize_source_segments`):** Splits raw `Source` strings into `(segment_text, is_tentative)`, treating `?`/`(?)`/`？` and `[...]` identically while preserving role brackets (`ROLE_BRACKET_RE`: `[pen holder]`, `[Rapporteur]`), legal suffixes (`CORP_SUFFIX_RE`: `, Inc?`), and missing-comma chains (`Huawei? Samsung?`).
- **Default Exclusion & Rendering:** `include_tentative=False` excludes carry-over companies from filters, single-meeting stats (`StatisticsExporterThread`), cross-meeting audits (`ContributionSearchWorker`, `ContributionStatsExporterThread`), and Markdown company reports (`MarkdownExporterThread`) — with opt-in toggles on `TDocsWindow.company_combo` (`MeetingsSettings.get_include_tentative_sources()`) and `ContributionReportDialog` (`CompanySelectorWidget` `chk_include_tentative`). `SourceCompanyDelegate` checks `CompanySanitizer.has_tentative_sources(source_str)` and renders tentative segments in muted italic slate (`#94A3B8`) via `CompanySanitizer.format_source_html`, with structured `CompanySanitizer.build_source_tooltip_html` hover cards.
- **LLM-Safe Two-Line Markdown Headers (Approach B):** `- **Source:**` lists confirmed contributors only (`CompanySanitizer.get_confirmed_source_text`); `- **Unconfirmed / Carry-over Sources:** <tentative> *(carry-over...)*` isolates carry-overs; table cells render `Confirmed *(carry-over: Tentative)*` via `CompanySanitizer.format_source_markdown_table_cell`; Ollama triage prompts warn local models via `CompanySanitizer.format_source_triage_context`; `dossier_exporter.py` uses `format_source_html` / `format_source_triage_context`.

---

## 4. TDoc Markdown Conversion, Parallel Pre-Download & Completeness
(`tdoc_markdown_converter.py`, `tdoc_url_resolver.py`, `llm_exporter.py`, `tdoc_file_handler.py`, `export_manifest.py`)

- **Revision Mode & Provenance:** `convert_document_to_markdown()` runs `resolve_revision_mode_for_type()` (`markup` for CR/pCR; `accept` for LS in/out correspondence; `is_conversion_stale()` upgrades older `markup` LS dossiers). `resolve_tdoc_download_url()` embeds official HTTPS URLs (`https://www.3gpp.org/ftp/...`, routing revisions to `/INBOX/Revisions/`).
- **Two-Phase Parallel Pipeline (`llm_exporter.py`, `post_meeting_exporter.py`, `pre_meeting_exporter.py`):**
  - *Phase 1:* `ThreadPoolExecutor(max_workers=5..6)` pre-downloads uncached binaries to `.zip.part` before atomic `os.replace`, validating via `zipfile.is_zipfile` with `revisions_url` $\leftrightarrow$ `docs_ftp_url` failover.
  - *Phase 2:* Sequential in-memory OpenXML conversion (isolating Word COM & vector metafile conversions).
- **Completeness Guard & Recovery (`export_manifest.py`):** `is_export_complete()` rejects `<120`-byte stubs, `"Document conversion was not available"`, corrupted `"__-________"` prefixes, presentation `.md` files lacking `"## Slide "`, or missing slide PNGs; `plan_incremental_session()` schedules incomplete files into `tdocs_to_update`; `_clean_for_json` converts `set`/`tuple`/`Path` before JSON serialization.

---

## 5. Pre-Meeting, Post-Meeting & Analytics Exporters

- **Pre-Meeting Briefing (`PreMeetingExportDialog`, `PreMeetingExporterThread`):**
  - Structure: `Export/PreMeeting_<WG>_<Meeting>_<timestamp>/Rel-XX/<WI>/{LS_IN,LS_OUT,pCR,CR,Discussion}/`.
  - Features `deduplicate_revisions=True` (`(rev of S2-...)`), 1-to-many candidate LS reply clustering (with `Related WIs` inheritance from parent `LS in`), contributor activity matrix, `⏳ Reserved` awareness, personal stance embedding (`> [!IMPORTANT] Delegate Pre-Meeting Stance`), and `.export_manifest.json` a-posteriori incremental updates with collapsible `<details>` LLM changelogs.
- **Post-Meeting Dossier (`PostMeetingExportDialog`, `PostMeetingExporterThread`):**
  - 2-column dialog with synchronized dual-mode tree (`🗂️ Work Items` $\leftrightarrow$ `📋 Agenda Items`, natural numeric sorting) and live KPI scope strip.
  - Exports received `LS in` and approved/agreed `pCR`/`CR`/`LS out` (plus `Disc`/`WID` when `include_disc`), replicates multi-WI docs with cross-WI callouts, preserves unassigned `LS in` under `General/LS_IN/`, guarantees bidirectional LS reply targets (`General/LS_OUT/` if out-of-scope), integrates RAN Chairman's Notes agreements (`include_chair_agreements=True`), and supports `.export_manifest.json` incremental extension.
- **Meeting Analytics (`StatisticsExporterThread`, `html_dashboard_compiler.py`, `contribution_stats.py`):**
  - Generates `Statistics_Report.html` (6-card KPI strip, `core/config/plot_styles.py` semantic outcome palette, Louvain alliance graphs, Plotly cross-filtering, dual Chart/Data tables) and 10-sheet `Statistics_Data.xlsx`, plus cross-meeting contribution audits.

---

## 6. Chairman's Notes & Cross-Meeting Agreement Tracker
(`chair_notes_downloader.py`, `chair_notes_picker_dialog.py`, `chair_notes_agreements_parser.py`, `agreement_tracker_worker.py`, `agreement_tracker_dialog.py`, `meeting_agreements_dialog.py`)

- **Strict WG Scoping (`is_ran = 'RAN' in wg_name`):**
  - **All WGs (`SA2`, `RAN1..3`, `CT`, etc.):** `🔄 Refresh ▾ → 📥 Download Chairman's Notes` opens `ChairNotesPickerDialog` (`ChairNotesClassifier` auto-picking newest file per chair/track) and stores files cleanly in `Agenda/ChairNotes/`, exposed in `🗂️ Resources`.
  - **RAN Only (`is_ran`):** Inline text agreements (`[⚖️ Agreements]`, `WA:`, `FFS:`) are strictly scoped to RAN WGs; SA/CT hide agreement buttons/chips/menu items so chair triage notes never pollute official TDoc status.
- **Dual-Mode Word Parser (`ChairNotesAgreementsParser`):**
  - *RAN1/RAN2 Document-Order Block Mode (`_parse_block_based_docx`):* Walks `iter_block_items()`, captures multi-paragraph `Agreements` blocks across empty lines (`consecutive_empty_lines`), 1-column boxed consensus/guidance tables (Pattern A, filtering antitrust/IPR boilerplate), and `style == 'Agreement'` paragraphs.
  - *RAN3/SA2 Color & Banner Table Mode (`_parse_table_based_docx`):* Scans rows 0–4 headers (including `Rel` and `Work Item`), tracks AI transitions on non-TDoc header rows so the first `Sub-Item` under each AI is preserved, inherits `WID`/`SID` from shaded AI banners, separates `From RAN3#<prev>:` recaps, and classifies OpenXML green (`#008000`/`#00B050`/`#538135`) as Agreements/WAs and blue (`#0000FF`/`#3333FF`) as FFS.
  - *Chair Result Normalization (`parse_chair_result` in `chair_notes_agreements_parser.py`, `TDocsTableModel.merge_agenda_data`):* Separates canonical 3GPP outcomes (`Approved`, `Agreed`, `Noted`, `Open`, `Postponed`, `Withdrawn`, `Revised`, `Merged`, `Replied to`, `Not Handled`) from embedded revision targets (`Revised to S2-...` $\rightarrow$ `Revised` + bidirectional `Revised to`/`Is revision of` links), merge pointers (`Merged into S2-...` $\rightarrow$ `Merged` + clickable remark), multi-line action notes (`OPEN\n...`), bulk-approval markers (`Block Approve?` $\rightarrow$ `Approved`, treated as approved in post-meeting/Markdown/topic exporters), and free-text remarks (`...not opened yet.`) so `TDoc Status` remains clean while preserving all remarks in `Secretary Remarks`.
  - *Cache Versioning (`CHAIR_NOTES_CACHE_VERSION = 2`):* Auto-invalidates v1 caches in `Agenda/chair_notes_cache.json`; populates `TDocsTableModel` rows with `[⚖️ AGREEMENT]` pills, emerald/amber/blue row tinting, `MeetingAgreementsDialog`, and AI-to-WI inheritance in `Agreements.md`.
- **Cross-Meeting Agreement Tracker (`AgreementTrackerDialog`, `AgreementTrackerWorker`):** Scans multiple WGs/dates in parallel, auto-downloads Final/EOM notes, filters by WI/SI/AI/keyword/decision type, and exports `.md`, `.html`, and `.xlsx`.

---

## 7. External LLM CLI Agent Integration in Meetings
(`src/core/ai/cli_detector.py`, `cli_launcher.py`, `src/core/ui/cli_context_dialog.py`, `tdocs_export_controller.py`)

- Discovers `codex`, `claude`, and `gemini` in `<2 ms` (`discover_cli_models()` reads `~/.codex/models_cache.json` + `~/.codex/config.toml` + `src/config/cli_agents_config.json`; `build_interactive_cli_args()` passes `-m`/`-c model_reasoning_effort=...`, `--model`/`--effort`, and `-m`), stages `_CLI_CONTEXT.md` with two-layer persistent notes (`src/config/cli_user_context.md` + per-workspace `_USER_NOTES.md`, `CliContextEditorDialog`), and launches `wt.exe` / `powershell.exe` via `-EncodedCommand`.
- `prepare_and_launch_cli_session(...)` runs single-TDoc actions (`TDocsWindow` context menu, `TDocTriageDialog` **`💬 Continue in CLI`**) directly inside canonical `{meeting_dir}/{base_tdoc}/` (`stage_copies=False`, checking `is_canonical_tdoc_markdown_ready`, zero `Export/` clutter); multi-TDoc/compare actions stage in `Export/CLI_Batch_{timestamp}/`; `TopicDossierDialog` & Pre/Post-Meeting export dialogs launch interactively with `bootstrap_prompt=""` (zero automatic multi-file consumption).
