# Emails, Tracked Threads & Unified Inbox Guide (`src/modules/emails`)

Developer and AI agent guide for WG reflector scraping, Outlook COM integration, Tracked Email Threads (`THR-...`), the Unified Email Triage Cockpit (`MeetingEmailInboxDialog`), Safelinks/Drafts caching, and Option A Markdown rendering.

---

## 1. Directory Map

```text
src/modules/emails/
├── AGENTS.md
├── core/
│   ├── outlook_client.py              # Win32 Outlook COM + MAPI GetTable/GetConversation wrappers
│   ├── emeeting_signals.py            # <<START>>..<<END>> parser, revision hunter, objection detector & jump strip builder
│   ├── email_deduplicator.py          # Fast lossless fingerprint deduplication across emails.db & tracked_threads.db
│   ├── legacy_email_migrator.py       # Idempotent auto-migration from legacy `emails` table -> `general_emails`
│   ├── sync_watermark.py              # Persisted UTC sync watermark, MAPI DASL Table.Restrict filters & Live Auto-Sync state
│   ├── email_db.py / email_parser.py / email_threads.py / email_config.py / email_summary_worker.py
│   ├── general_email_db.py / general_email_sync.py    # Unpinned WG folder + Sent Items sync (`sync_metadata` + Table.Restrict cache)
│   ├── tracked_threads/               # Modular MAPI sync, scoped_sync, discovery & Drafts package (see tracked_threads/README.md)
│   ├── tracked_threads_db.py / tracked_threads_service.py (shim) / tracked_threads_workers.py
│   ├── ai_summary/                    # Incremental multi-scope `email_summaries.db` store & 2-stage Per-AI Map-Reduce (see ai_summary/README.md)
│   ├── markdown_export/               # Modular Option A cards, Safelinks, HTML tables & quote stripping package (see markdown_export/README.md)
│   ├── thread_markdown_exporter.py    # Re-export shim -> core/markdown_export/ package
│   └── stats/                         # Unified email analytics (dataset_builder, plot_agenda, plot_companies, plot_delegates, plot_timeline)
└── ui/
    ├── ai_summary/                    # Shared `InlineEmailSummaryCard`, unified `llm_runner` model picker & defaults dialog
    ├── briefings/                     # Modeless `EmailBriefingsDialog` for Per-AI cross-thinking & saved selections
    ├── manager_window/                # Modular EmailManagerWindow package (see manager_window/README.md)
    ├── email_window.py                # Re-export shim -> manager_window/ package
    ├── general_email_dialog.py / email_dialogs.py / email_summary_dialog.py
    ├── meeting_inbox/                 # Modular MeetingEmailInboxDialog package (see meeting_inbox/README.md)
    ├── meeting_email_inbox_dialog.py  # Re-export shim for MeetingEmailInboxDialog & shared QSS/widgets
    ├── merge_threads_dialog.py        # Searchable, similarity-ranked split-thread merge dialog
    ├── tracked_thread/                # Modular TrackedThreadDialog & PinFromOutlookDialog package (see tracked_thread/README.md)
    ├── tracked_thread_dialog.py       # Re-export shim -> tracked_thread/
    ├── tdoc_emails/                   # Modular TDocEmailsDialog & StandaloneEmailReaderWindow package (see tdoc_emails/README.md)
    ├── tdoc_emails_dialog.py          # Re-export shim -> tdoc_emails/
    └── email_models.py / email_toolbar.py / config_dialog.py / sync_progress_bar.py
```

---

## 2. Multi-Source Discovery & Tracked Threads

- **Untracked Thread Discovery (`PinFromOutlookDialog`, `classify_discovered_threads`):**
  - Accepts `.msg` files, Outlook drags (`FileGroupDescriptorW` / `RenPrivateMessages` $\to$ `pin_threads_from_seeds_batch` / `pin_thread_from_outlook_selection` / `pin_threads_from_msg_files`), and WG/`Sent Items` discovery indexing `tracked_threads.db` + `general_emails`.
  - Uses MAPI `folder.GetTable()` (up to 1,500 items/folder matching `normalize_subject`/`ConversationID`), `GetConversation().GetTable()`, recursive subfolders (`OutlookClient.iter_folder_tree`, `include_subfolders=True` default in `GeneralEmailFoldersDialog`), lossless SQLite deduplication (`email_deduplicator.py` preserving `MAX(is_read)`, `MAX(flag_status)`, attachments, TDoc matches), and DMARC sender resolution (`GeneralEmailSyncThread._resolve_sender`).
  - Classifies via `valid_tdocs` & `get_tdoc_match_locations_by_email_id()` into `📌 Pinned`, `🟢 Subject TDoc`, `🟡 TDoc in Body Only`, `🔴 No TDoc` (default `UNTRACKED_AND_BODY`), with exclusions in `<meeting_dir>/Agenda/email_discovery_config.json`.
- **Cascading Selectors & Inheritance (`TrackedThreadDialog`, `topic_pills.py`):**
  - `TagCloudWidget` (`strict_dropdown=True` searchable popups with `Clear`) scopes `Sub-Topics` to `AIs` (`ai_to_sub_items`, pruning orphans); `compact=True, prefix_label=..., max_visible_pills=...` renders borderless pills + `+N more`/`Show less` in `MeetingEmailInboxDialog` (collapsed default + live toggle summary) and `TrackedThreadDialog` (`links_panel` + live toggle preview).
  - `_expand_tdoc_family_with_merges()` tracks revisions/merges; `auto_suggest_links()` + `infer_wis_for_ais()` match `(?<![A-Za-z0-9])` TDocs (`S2-2612919_was_S2-2612519`), `AI#`/`KI#`, mixed-case WIs (`FS_enIMS_ARC`), attachments, only inherit unanimous body-only AIs/Sub-Topics, and infer WIs via `TDocsWindow._build_tdoc_meta_map()`.
  - **Split-Thread Merge (`MergeThreadsDialog`, `merge_threads`):** Ranked by `score_thread_similarity()` / `choose_default_primary_thread()` from `TrackedThreadDialog`, `TDocsWindow` (`tdocs_menu.py`), and `MeetingEmailInboxDialog` (auto-pinning unpinned items with `refresh_from_outlook=False`), combining folders, tags, notes, and status.
- **Meeting Table Integration:** Tracked threads appear in `TDocsTableModel` (`sync_tracked_thread_rows`, `📧 Email Thread`), share unread/flag counters (`📧 Emails (N new) ▾`, `🚩 N`), `🔵`/`🚩` filters, and sync via `TrackedThreadsRefreshWorker`, `SingleThreadRefreshWorker`, and `OutlookFolderScanWorker` with `AsyncSyncProgressBar`.

## 3. Unified Email & e-Meeting Triage Cockpit (`MeetingEmailInboxDialog`)

- **Deduplication, Zero-Data-Loss & Legacy e-Meeting Auto-Migration:** Combines and `EntryID`-deduplicates messages across `TrackedThreadsDatabase` and `GeneralEmailDatabase.get_all_emails_with_tdocs()`, persisting `QSettings` filters (`meetings/<meeting_key>/...`) immediately and auto-reloading on sync/edit signals (`TrackedThreadsRefreshWorker.thread_updated`/`finished_sync`, `GeneralEmailSyncThread`). `GeneralEmailDatabase._init_db()` automatically runs `migrate_legacy_emeeting_emails()` so historical e-Meeting `emails` rows are reconstructed (`<<START>>...<<END>>`) into `general_emails`, and all upserts preserve non-empty cached `body_text`/`clean_body` when Exchange purges old server items.
- **Unified Adaptive Async Sync (`InboxOutlookSyncWorker`, `sync_watermark.py`):**
  - *Unified engine across `TDocsWindow` (`📧 Emails ▾`) & `MeetingEmailInboxDialog`:* Both UIs share `InboxOutlookSyncWorker` (`incremental` / `full` / `scoped`) and persisted `🟢 Live Auto-Sync (Every 2m)` (`load_meeting_live_sync_enabled` / `save_meeting_live_sync_enabled`), skipping duplicate background ticks when the Inbox dialog is open and driving live sync.
  - *Incremental Watermark & MAPI `Table.Restrict()` (`⚡ Sync New` / `⚡ Sync New Emails Now`):* `GeneralEmailDatabase.sync_metadata` persists `last_sync_utc` (updated only on clean non-cancelled runs, cleared on `wipe_generic_emails()`). `resolve_sync_window()` uses `last_sync_utc - 5 min` (automatically covering app closes / laptop restarts), falling back on first sync to `meeting_start - 5d .. meeting_end + 5d`. Both `refresh_threads_batch()` and `GeneralEmailSyncThread.sync_folders()` apply locale-independent DASL `apply_mapi_since_restriction()` (`@SQL="urn:schemas:httpmail:datereceived" >= '...'`) with fallback early-break on `ReceivedTime`/`SentOn`, skipping 9,000+ historical items. `📅 Full Meeting Sync (Date Range)...` opens `GeneralEmailSyncDialog` for explicit windows.
  - *Filter-scoped quick sync (`⚡ Sync Filtered (K)`) & single-thread sync (`🔄 Sync Thread`):* When `AI`/`Company`/`Search`/`Source`/`Flagged`/`Priority` filters are active, `sync_scoped_inbox_items()` (`tracked_threads/scoped_sync.py`) syncs only filtered threads & unpinned conversations via `EntryID` + `GetConversation().GetTable()` without advancing `last_sync_utc`.
- **2-Row Responsive Header, 3-Way View Mode (`📄 By TDoc`), Priority Radar & `🛑 Objected` Smart Parking:**
  - *Row 1 (Navigation & Actions):* State pills (`🔥 Priority`, `🔵 Unread Only`, `🚩 Flagged`, `🕒 Last 24h`, `📋 All`), 3-way view toggle (`📨 By Message` / `🧵 By Thread` / `📄 By TDoc`), `🔄 Reload`, `⚡ Sync New ▾`, `🟢 Live (2m)`, `📊 Stats` (`build_unified_email_stats_dataset` + `EmailStatsExporterThread`), and `⬒ List on Top` / `◫ Side-by-Side` (`Ctrl+L`, `QSettings`-persisted).
  - *Row 2 (Search & Scope Filters):* `🔍 Search`, `👀 Followed AIs` (`followed_ais_helper.py`, auto-inherited from Followed/Starred WIs, Session Plan assignments, and `🔵 My TDoc` rows with per-meeting override & right-click reset), `🛑 Hide Objected`, and `AI`/`Company`/`Status`/`Source` filters.
  - *e-Meeting Signal Engine & Smart Parking (`emeeting_signals.py`, `tdoc_grouping.py`):* Extracts `<<START>>...<<END>>` blocks (`📋`), shorthand/full revision mentions (`🆕 r01/r02`), and direct-body 3GPP objection phrases (`⚠️`). Marking a Thread or TDoc `🛑 Objected` records `objected_at_rev` (`tracked_threads.db` & `user_tdocs.db`), parking it out of `🔥 Priority` unless a strictly newer revision (`r02` > `r01`) arrives (`🛑→🆕r02`).
  - *`📄 By TDoc` Mode & 2-Way `user_tdocs.db` Sync:* Groups all meeting emails by Base TDoc ID (`S2-26xxxx`) displaying `State / Signals | TDoc & Author | AI / WI | TDoc Title & Context | Email Contributors | Latest`, synchronizing `My Status`, `🛑 Objected`, and `📝 Note` live with `TDocsWindow`, alongside `📋 <<START>>` template copy, `🎯 Table` jump, `🤖 LLM && Export`, and clickable In-Thread Message Jump Strip (`jump://msg/<idx>`).
- **Incremental Multi-Scope AI Summaries & Per-AI Briefings (`core/ai_summary/`, `ui/ai_summary/`, `ui/briefings/`):**
  - *Unified Runner (`core.ai.llm_runner`):* Executes local Ollama and headless external CLIs (`codex exec`, `claude -p`, `gemini`) identically via `ModelSpec(provider, model, reasoning_effort)` with persisted `leaf` (fast) and `synthesis` (deep) defaults (`src/config/email_summary_config.json`).
  - *Watermark + Covered-ID Store (`email_summaries.db`):* Tracks `cutoff_utc` + `covered_message_ids` JSON set (catching backfilled/merged messages), `incremental_hops`, optional parsed `summary-meta`, and 5-deep undo history across `thread`, `tdoc`, `selection`, and `ai` scopes.
  - *2-Stage Per-AI Map-Reduce & Inline UI:* `EmailBriefingsDialog` reuses fresh leaf summaries and refreshes only stale leaves before cross-thinking synthesis (`Stage 2`), while `InlineEmailSummaryCard` embeds directly inside `MeetingEmailInboxDialog`, `TrackedThreadDialog`, and `TDocEmailsDialog` alongside Inbox freshness badges (`✨` / `✨+N`), contextual 1-AI strip, multi-select summarization, and flexible `SummaryReaderLayoutController` modes (`⬒`, `◫`, `⤢`, `▸`, and multi-monitor `⧉` `PopoutSummaryWindow`).
  - *Microsoft 365 Copilot Handoff (`m365_handoff.py`, `M365CopilotHandoffDialog`):* 2-step zero-IT-ticket Send & Paste-Back workflow across all readers and `EmailBriefingsDialog` that stages clipboard prompts / upload `.md` packs (for Web `copilot.cloud.microsoft` or Windows `ms-copilot:` app) and saves pasted replies cleanly into `email_summaries.db` with full watermark/ID tracking.

## 4. URL Safety, Drafts Caching, Quote Stripping & Viewer Ergonomics

- **Safelinks/Proofpoint Unwrapping & URL Shortening:** `unwrap_outlook_safelinks()` decodes Microsoft Safelinks, Proofpoint `urldefense`, and FireEye wrappers, normalizes `10.10.10.10/ftp/...` and `3gpp.org/ftp/...`, and collapses duplicate adjacent URLs. With `shorten_long_urls=True`, `linkify_markdown_urls()` shortens doc links to `[📄 <filename>](url)` and >72-char web links to `[🔗 <netloc>/…/<tail>](url)`, while `apply_email_reader_document_style()` connects `QTextBrowser.highlighted[str]` (`handle_reader_link_hover`) for full-URL `QToolTip` previews on `http(s)://` links only.
- **Thread & Proactive `/Drafts/` Caching:** Clicking linked docs caches into `<meeting_dir>/Agenda/ThreadAttachments/<thread_id>/` (`10.10.10.10` $\leftrightarrow$ `Meetings_3GPP_SYNC` failover, 1.5s connect timeout on `10.10.10.10`) and opens via `modules.word_tools.core.document_opener.open_document()` (applying MIP sensitivity labels). Additionally, `cache_drafts_for_messages()` proactively downloads **all** `/Drafts/` URLs (`is_drafts_document_url()`, excluding `/Revisions/` and `/Docs/`) referenced in newly synced/cached Pinned & `GeneralEmailDatabase` emails into `<meeting_dir>/Agenda/DraftsCache/`, reused automatically by `cache_and_open_thread_document_url()`.
- **Option A Message Cards, Localized Quotes & Shared Reader Behavior:**
  - `TrackedThreadDialog`, `TDocEmailsDialog`, `MeetingEmailInboxDialog`, `StandaloneEmailReaderWindow`, and `EmailManagerWindow` share `bind_reader_zoom_settings()` / `zoom_text_widget()` (`Ctrl+Wheel` / `Ctrl+0`, flushing `zoom_steps` to `QSettings` immediately on change), `apply_email_reader_document_style()`, `get_compact_action_bar_qss()`, `get_status_combo_qss()`, `_NotePlainTextEdit`, and discrete `? Help` pills (`create_help_tips_pill()`). `TDocEmailsDialog` receives `tdoc_info`, `open_tdoc_callback`, and `user_data_callback` from `email_compose._open_tdoc_emails()` to render a 2-line TDoc context header (Title, Source, AI/Sub-Item, WI, Status, clickable Family links), always-editable comment bar, compact scope pills (`🔗 Family`, `💬 Quoted`, `🚫 Ignored`), `📄 Open <TDoc>`, and live `My Status` selector alongside `🤖 LLM & Export` and `Launch LLM CLI` (`auto_execute=False`).
  - `export_thread_to_markdown(..., interactive_links=True)` renders Option A state-tinted `<table>` banners (`core.ui.theme.C` tokens, 4px accent column, `split_sender_company()`, `outlook://...` links) preserving `<h3 ...>` for deterministic `QTextDocument` block auto-scroll (`find_unread_heading_blocks()` / `scroll_browser_to_block()`, because Qt 5 MD4C strips `<a name>`), while `interactive_links=False` emits 100% pure Markdown for LLMs.
  - Localized trailing-quote detection (`find_trailing_quote_line_index()` / `strip_trailing_quotes()` / `split_direct_and_quoted_body()` across `From:`/`Sent:`, `发件人:`/`发送时间:`/`主题:`, `Von:`/`Gesendet:`, `De:`/`Envoyé:`) runs at sync and render time (`deduplicate_quotes=True`, or gray blockquotes via `_format_trailing_quoted_history()`), paired with `inject_html_tables_into_body()` GFM tables, `partition_email_attachments()` signature-image filtering, `reload_general_email_from_outlook()` / `reload_tdoc_emails_from_outlook()`, and `Launch LLM CLI` (`auto_execute=False`) staging. (Legacy `TDocEmailsDialog` 800ms auto-read timer is intentionally removed.)
