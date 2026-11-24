# Cross-WG Topic Dossiers, AI Discovery & Inputs Viewer

Deep architectural reference for the **Cross-WG, Multi-Meeting Topic Dossier & Briefing Studio** (`meta_topics.json`, `topic_manager.py`, `cohort_resolver.py`, `topic_decision.py`, `cross_wg_ls_tracker.py`, `topic_dossier_exporter.py`, `topic_digest_builder.py`, `topic_dossier_dialog.py`, `ai_suggester.py`, `ai_suggestion_dialog.py`, `topic_inputs_dialog.py`, `topic_pills.py`).

---

## 1. Persistent Meta-Topics & Pill/Tag Editor
(`src/config/meta_topics.json`, `meta_topics_user.json`, `topic_manager.py`, `topic_pills.py`, `topic_dossier_dialog.py`)

- **Two-Tier Topic Storage:** Separates factory presets (`src/config/meta_topics.json`, Git-tracked) from user edits/deletions (`src/config/meta_topics_user.json`, Git-ignored, mirrored to `~/3GPP_Delegate_Helper/meta_topics_user.json`). Factory topics can be restored non-destructively in 1 click (`🔄 Restore Built-in Topics...`) without losing custom topics.
- **Correlation Rules (`MetaTopicDefinition`, `MetaTopicRules`):** Matches contributions via: (1) normalized WI/SI tokens (`ambientiot`, `fs_ambientiot`), (2) WG-specific AI wildcard patterns (`SA2: 5.14*`, `RAN2: 8.7*`), (3) title/abstract keywords, and (4) manual pinned/excluded TDoc IDs.
- **Pill & Tag UI (`topic_pills.py`):**
  - `WgTogglePill`: Checkable WG pills with group family colors (SA blue `#DBEAFE`/`#1E40AF`, RAN purple `#EDE9FE`/`#5B21B6`, CT teal `#CCFBF1`/`#0F766E`).
  - `TagCloudWidget`: Removable pill chips (`[ token ✕ ]`) with wide inline entry (`minimumWidth >= 190`), comma/Enter ingestion, copy-paste parsing, and backspace removal.
  - `AgendaItemPillsWidget`: Per-WG grouped AI patterns (`[ 20.6.21* ✕ ]`, `+ Add AI`), synchronous row detachment (`setParent(None)`), and top-aligned WG labels.
  - Scoped `#objectName` QSS frames (`#AiInferredBanner`, `#AiParamFrame`, `#AiProgFrame`, `#AiPreviewFrame`, `#TopicInputs*Frame`) with `background: transparent; border: none;` on child labels; `TopicDossierDialog` (`1020×680` fit, `2×2` options grid, `FlowLayout` actions, elided `outline: none` topic list) and scrollable `TopicEditorDialog` (`QScrollArea`).

---

## 2. Cross-WG Agenda Item (AI) Auto-Discovery Engine
(`ai_suggester.py`, `ai_suggestion_dialog.py`)

- **Single-AI Seeding (`infer_topic_from_known_ais`):** Entering an AI for a single WG (e.g. `SA2: 20.6.21`) queries cached agendas, resolves the official procedure/sub-topic description, extracts WI tokens (`fs_6g_arc`), and identifies recurring keywords (`data framework`, `data transfer`) to propagate across target WGs (`RAN2`, `RAN3`, `SA5`, `CT4`).
- **Scanning & Noise Suppression (`AgendaItemSuggester`):** Fast-scans `Agenda/*.xlsx.json` / Excel workbooks, suppresses broad umbrella WI noise (when keywords exist, candidate AIs must match keywords in AI description or TDoc titles), and protects administrative AIs (`GENERIC_AI_CATEGORIES`: `Corrections`, `Review of incoming LS`, `Drafting`, `Miscellaneous`, `AOB`) from `🟢 High` wildcard pollution (`7.2.2*`).
- **Confidence Classification:** `🟢 High` (description match or `title_matches >= 3` + WI correlation on non-generic AI), `🟡 Medium` (title keyword hits), `⚪ Low`.
- **`AiSuggestionDialog`:** Runs `AiSuggestionScanWorker` (`QThread`), displays multi-WG tabs, inferred context banner, `✅ Select High Confidence`, live rule preview, and non-destructive merge back into `TopicEditorDialog`.

---

## 3. Meeting Scope & Timeframe Resolution (`cohort_resolver.py`)

- Uses natural 3GPP standards terminology (**"Meeting Scope && Timeframe"**, **"Synchronized Meeting Rounds"**, **"Meeting Cycles"**, **"Custom Date Range"**) via `MeetingRoundResolver` (`MeetingCohortResolver`), clustering staggered multi-WG dates into unified cross-group rounds.
- **Folder-Name Date Fallback (`infer_meeting_date_from_folder` / `get_effective_meeting_dates`):** Meetings missing DB `start_date`/`end_date` (e.g. `TSGS2_178-AH-e_Electronic_2027-01`, `TSGS2_180_Stockholm_2027-04`) infer month bounds from `YYYY-MM` / `YYYY-MM-DD` folder tokens instead of being dropped.
- **Open-Ended `"since_date"` Mode & Corpus Auto-Anchoring:** Presets start with **`📅 Since Start Date → Present (+ Upcoming)`** (`since_date`, default 45-day upcoming lookahead with `Include all future` override), followed by `Past 4/6/8 Meeting Rounds`, `Release 20/19/18 Timeframe`, `📅 Custom Date Range...` (index 6), and `All Available Meetings`. When a topic has no explicit `since_date` rule yet and an existing `.export_manifest.json` is present, `TopicDossierDialog` auto-anchors `From:` to the earliest exported meeting `start_date`; `TopicEditorDialog` can also persist an explicit `Default Start Date` (`cycle_mode="since_date"`).
- **Per-WG Latest Summary & Full Tooltip:** `lbl_cohorts_summary` displays the latest resolved meeting per target WG (plus any upcoming meeting `→ #<num>`) and exposes the full round-by-round meeting list in its tooltip.

---

## 4. Considered Inputs & Decisions Visual Inspector
(`TopicInputsViewerDialog` in `topic_inputs_dialog.py`)

- **Modeless & Non-Blocking:** `super().__init__(None)`, `Qt.Window`, `Qt.NonModal`, backed by `InputsCollectorThread` (`QThread`).
- **Smart Live-Refresh Parity (`is_meeting_cache_stale_for_preview`):** Recent/concluded meetings within 45 days whose `.xlsx.json` cache predates meeting end (or is missing/stale within 14 days) trigger live refresh in `InputsCollectorThread`, and `_load_meeting_tdocs` never lets a stale `.xlsx.json` shadow a newer valid `.xlsx`. Matches use the same `abstract` + `ai_description` inputs as `TopicDossierExporterThread`.
- **Executive KPI Strip & 2-Row Wrapping Quick-Filter Pill Bar (`FilterPillButton`):** Row 1 (Working Groups) + Row 2 (Categories: Decisions, Agreements, Approved pCRs, Approved CRs, Liaisons, Discussions) inside `FlowLayout` containers (`Preferred` height-for-width sizing for 12+ WGs) with live counts, instant text search, and `↺ Reset`.
- **4 Specialized Tabs:** `⚖️ Technical Decisions && Agreements`, `📄 All Considered Contributions`, `📬 Inter-WG Liaison Statements`, `📊 WG && Round Distribution` (native `QFont` bold totals).
- **Decision Provenance Badges:** Differentiates `🟢 Text Agreement (Notes)` (RAN verbatim consensus; background paper Noted) from `🟢 Approved pCR (Doc)` / `🟢 Approved CR (Doc)` (full document approved).
- **Collapsible Detail Reader Drawer (`QSplitter` + `▲ Collapse`/`▼ Expand` + `📋 Copy Verbatim Text`):** Synchronizes across tab/row selection changes, rendering full verbatim agreement text (`Consolas`), abstracts, or LS reply chains.
- **Ultra-Fast Batch Population (<0.12s for 1,500+ items):** Suspends updates (`setUpdatesEnabled(False)`, `blockSignals(True)`), populates in `Interactive` mode, and runs `resizeColumnsToContents()` once at the end.

---

## 5. Unified Decisions, Standalone RAN Agreements, Exporter & Incremental LLM Feed
(`topic_decision.py`, `cross_wg_ls_tracker.py`, `topic_dossier_exporter.py`, `topic_digest_builder.py`, `export_manifest.py`)

- **Per-WG $\to$ Per-Meeting Topology:** Exports into `{WG}/{MeetingFolder}/` (`resolve_meeting_folder_id` disambiguating `-bis`/ad-hoc/electronic meetings like `RAN3_133` vs `RAN3_133-bis`, `SA2_175-AHE`, auto-migrating legacy `YYYY-QN/` exports).
- **Unified Technical Decisions (`topic_decision.py`):** Elevates RAN Chairman's Notes text agreements (strictly RAN-scoped so SA/CT notes never pollute `Text Agreement`), Working Assumptions, and Endorsed WFs alongside SA/CT approved pCRs/CRs in `agreements_timeline.md`.
- **Standalone RAN Agreement Markdown Exporter (`_write_agreement_markdown`):** Generates `{wg}/{meeting_folder_id}/Agreements/Agreement_{tdoc_or_ai}.md` with metadata frontmatter, authoritative 3GPP URL (`resolve_tdoc_download_url`), verbatim consensus blockquotes, topic traceability, and bidirectional link to `../Discussion/{tdoc}.md`, auto-including `decision_referenced_tids`.
- **Cross-WG LS Threading (`cross_wg_ls_tracker.py`):** Traces inter-group `LS OUT` $\to$ `LS IN` $\to$ reply chains into `ls_dialogue.md`.
- **Canonical Cache Contract & Parallel Pre-Download:** Checks `{cache_dir}/{folder_name}/{base_tdoc}/{tid}.md` (and legacy `_LLM_v{CONVERTER_VERSION}.md`) with `is_conversion_stale` *before* FTP download, converts cache misses directly into the canonical meeting TDoc folder, rewrites `media/` $\to$ `{tid}_media/`, honors `deduplicate_revisions` (`_prune_superseded_topic_docs` keeping global cache & decision-referenced/terminal revisions) and `include_discussions`. `allow_network_download=True` by default uses `ThreadPoolExecutor(max_workers=5)` with 12s failover timeouts, auto-downloads RAN Chairman's Notes into `Agenda/ChairNotes/`, and live-refreshes upcoming/empty/stub/recent ($\le 45$ days) meetings (`_is_meeting_live_for_refresh`).
- **3-State Meeting Lifecycle & Two-Stage E2E LLM Analysis (`topic_digest_builder.py`):**
  - Classifies each meeting deterministically (`classify_meeting_lifecycle` $\to$ `✅ Concluded`, `🧭 Upcoming / Pre-Meeting`, `⚪ No Activity (0 Inputs)`).
  - Generates `{WG}/{MeetingFolder}/_meeting_brief.md` (with `Brought vs. Decided Disposition`, `Decision / Outcome` column, normalized contributor matrix via `CompanySanitizer` excluding unconfirmed carry-over `[Company?]`) and company-grouped `_company_positions_input.md` (split along company boundaries at 140k chars).
  - **Non-Destructive Delta Merge:** Never overwrites `_company_positions_analysis.md` and never creates companion `.pre_meeting.md`/`_delta.md` files. Detects post-meeting changes via `compute_meeting_snapshot`/`diff_meeting_snapshots`, marks the meeting `🔄 Stale (Delta Ready)`, and injects an in-place `STAGE1_DELTA_START`..`STAGE1_DELTA_END` section into `_meeting_brief.md`.
  - **Dated Stage-2 Synthesis Packages:** `discover_stage2_packages` discovers `stage2_<YYYY-MM-DD>_rN/` (`COMPANY_VIEWS_AND_PROGRESS.md`, `COVERAGE.md`, `working/package_state.json`, `working/validation.json`) alongside legacy root `company_positions_e2e_synthesis.md`, tracking `pending`/`in_progress`/`stale`/`analyzed` states (`_evaluate_stage2_synthesis_state`), `changed_meetings`, and `--baseline` chaining in `company_positions_e2e_prompt.md`.
- **Offline Checkup & Manifest (`export_manifest.py`, `check_and_refresh_dossier`):** `.export_manifest.json` (`MANIFEST_VERSION = 2`) checkpoints atomically after each meeting and writes `deltas/delta_session_N.md`. **`🔍 Check Status & Refresh Index`** (`<0.5s`, zero network/conversion) scans `_company_positions_analysis.md` and `stage2_*` timestamps, backfills manifest metadata, promotes merged analyses, strips resolved `STAGE1_DELTA` blocks, and regenerates `index.md`, `company_positions_e2e_prompt.md`, and `_meeting_brief.md`.
