# Emails, Tracked Threads & Unified Inbox Guide (`src/modules/emails`)

Developer and AI agent guide for WG reflector scraping, Outlook COM integration, Tracked Email Threads (`THR-...`), the Unified Email Triage Cockpit (`MeetingEmailInboxDialog`), Safelinks/Drafts caching, and Option A Markdown rendering.

---

## 1. Directory Map

```text
src/modules/emails/
├── AGENTS.md
├── core/
│   ├── outlook_client.py              # Win32 Outlook COM + MAPI GetTable/GetConversation wrappers
│   ├── email_db.py / email_parser.py / email_threads.py / email_config.py / email_summary_worker.py
│   ├── general_email_db.py / general_email_sync.py    # Unpinned WG folder + Sent Items sync
│   ├── tracked_threads_db.py / tracked_threads_service.py / tracked_threads_workers.py
│   ├── thread_markdown_exporter.py    # Option A cards, Safelinks, quote stripping, Drafts caching
│   └── stats/                         # Email analytics (plot_agenda, plot_companies, plot_delegates, plot_timeline)
└── ui/
    ├── email_window.py / general_email_dialog.py / email_dialogs.py / email_summary_dialog.py
    ├── meeting_email_inbox_dialog.py  # Unified Unread & Flagged Email Triage Cockpit
    ├── merge_threads_dialog.py        # Searchable, similarity-ranked split-thread merge dialog
    ├── tracked_thread_dialog.py       # Tracked thread viewer & AI/WI/Sub-Topic editor
    ├── tdoc_emails_dialog.py          # Unified per-TDoc email viewer
    └── email_models.py / email_toolbar.py / config_dialog.py / sync_progress_bar.py
```

---

## 2. Multi-Source Discovery & Tracked Threads

- **Untracked Thread Discovery (`PinFromOutlookDialog`, `classify_discovered_threads`):**
  - Accepts local `.msg` files, direct Outlook drags (`FileGroupDescriptorW` / `RenPrivateMessages` $\to$ `pin_thread_from_outlook_selection`), and structured discovery across configured WG folders + `Sent Items` (`GetDefaultFolder(5)`).
  - Uses fast MAPI `folder.GetTable()` (up to 1,500 items/folder, matching normalized subject `normalize_subject` or `ConversationID`), `GetConversation().GetTable()`, and Listserv/DMARC sender resolution (`GeneralEmailSyncThread._resolve_sender`).
  - Classifies threads via `valid_tdocs` and `GeneralEmailDatabase.get_tdoc_match_locations_by_email_id()` into `📌 Pinned`, `🟢 Subject TDoc`, `🟡 TDoc in Body Only`, `🔴 No TDoc` (default `UNTRACKED_AND_BODY`), with per-meeting exclusions in `<meeting_dir>/Agenda/email_discovery_config.json`.
- **Cascading Selectors & Inheritance (`TrackedThreadDialog`, `topic_pills.py`):**
  - `TagCloudWidget` with `strict_dropdown=True` (scrollable searchable multi-token checklist popups) scopes `Sub-Topics` strictly to assigned `AIs` (`ai_to_sub_items`, pruning orphans when an AI is removed while preserving legacy links until `✕` removal).
  - `_expand_tdoc_family_with_merges()` tracks revisions/merges; `auto_suggest_links()` resolves `KI#...` / `AI ...` / `WI` tokens and inherits `Agenda Item`, `Sub-Item`, and `Related WIs` from `TDocsWindow._build_tdoc_meta_map()`.
  - **Split-Thread Search & Select-to-Merge (`MergeThreadsDialog`, `TrackedThreadsDatabase.merge_threads`):** Searchable, multi-select merge workflow ranked by `score_thread_similarity()` (subject tokens, shared TDocs/Sub-Topics/AIs/WIs) and smart primary target selection (`choose_default_primary_thread()`), launched from `TrackedThreadDialog` (`More ▾ → 🔗 Merge Split Threads...`), multi-row right-click in `TDocsWindow` (`tdocs_menu.py` → `🔗 Merge Selected Threads (N)...`), and `MeetingEmailInboxDialog` (`btn_merge_selected` / row context menu, auto-pinning unpinned items with `refresh_from_outlook=False`). `merge_threads()` preserves & combines `outlook_folders`, `user_tags`, `user_notes`, and non-Neutral `user_status`.
- **Meeting Table Integration:** Tracked threads appear inline in `TDocsTableModel` (`sync_tracked_thread_rows`, `📧 Email Thread`), share deduplicated unread counters (`📧 Emails (N new) ▾`, `🚩 N`), `🔵`/`🚩` quick filters, and run Outlook syncs asynchronously via `TrackedThreadsRefreshWorker`, `SingleThreadRefreshWorker`, and `OutlookFolderScanWorker` with `AsyncSyncProgressBar`.

## 3. Unified Email Triage Cockpit (`MeetingEmailInboxDialog`)

- **Deduplication & Persistence:** Combines and `EntryID`-deduplicates messages across `TrackedThreadsDatabase` and `GeneralEmailDatabase.get_all_emails_with_tdocs()`, persisting `QSettings` filters (`meetings/<meeting_key>/...`) immediately and auto-reloading on sync/edit signals (`TrackedThreadsRefreshWorker.thread_updated`/`finished_sync`, `GeneralEmailSyncThread`).
- **Adaptive Async Sync (`InboxOutlookSyncWorker`):**
  - *Full incremental sync (`📡 Sync Outlook`):* Single-pass MAPI `GetTable()` (`refresh_threads_batch()` / `refresh_meeting_threads()` + `GeneralEmailSyncThread.sync_folders()` using `UnRead` / `0x10900003` flag columns) with known-`EntryID` fast-skip (`get_existing_message_states()`, `get_cached_email_states()`, `update_message_flags_batch()`, `update_email_flags_batch()`) to update read/flag state without reading `.Body`/`.HTMLBody` or Exchange SMTP.
  - *Filter-scoped quick sync (`⚡ Sync Filtered (K)`) & single-thread sync (`🔄 Sync Thread`):* When `AI`/`Company`/`Search`/`Source`/`Flagged` filters are active, `sync_scoped_inbox_items()` syncs only filtered threads & unpinned conversations via `EntryID` + `GetConversation().GetTable()` with zero WG folder scans (`📡 Full Sync` stays one click away in `▾`). Clicking `🔄 Sync Thread` (`btn_sync_selected`), `▾ → 🔄 Sync Selected Thread`, or right-clicking a row (`_on_table_context_menu`) runs `InboxOutlookSyncWorker` scoped exclusively to the selected thread/conversation(s).
- **Triage, Status/Comments, Column Resizing & Bidirectional Sync:** State pills (`🔵 Unread Only`, `🚩 Flagged`, `🕒 Last 24h`, `📋 All`), `📨 By Message` / `🧵 By Thread`, `⬒ List on Top` / `◫ Side-by-Side` (`QSettings`-persisted), `Company`/`AI`/`Status` `CheckableComboBox` filters (`❓ Unassigned (No AI)` + `MY_STATUS_OPTIONS`, all persisted per meeting), interactively resizable table columns (`QHeaderView.Interactive` with proportional `compute_fitted_column_widths()` / `_fit_columns_to_viewport()` auto-fill on resize & `List on Top`/`Side-by-Side` toggle, per-meeting `table_header_state` persistence, and `_InboxStatusStripDelegate` 5px left-edge status color stripe + `📝` note indicator on Column 0), collapsible inline assignment bar (`🏷️ AI / WI ▴/▾`) with `ExtendedSelection` batch assign/auto/pin (`refresh_from_outlook=False`)/read/flag, always-editable inline comment strip (`note_frame` auto-visible whenever `user_notes` is non-empty, toggled via `📝 Note`, auto-saved to `TrackedThreadsDatabase.update_thread_annotations`), live `detail_status_combo` (`MY_STATUS_OPTIONS`) synchronized with `TDocsWindow`, compact reading-panel action bar (hiding single-message `↩️ Reply`/`📧 Outlook` in grouped thread mode), clickable `tdoc:` links (`clean_tdoc_id` stripping `2F`/`20` URL encoding artifacts, routing external TDocs via `global_action_requested("open_doc")`), `✔️ Read && Next`, and bidirectional Outlook read/flag sync (`OutlookClient.set_messages_read_status`, `set_messages_flag_status`, `mark_thread_read_with_outlook`, `set_thread_message_read_with_outlook`, `set_thread_message_flag_with_outlook`).

## 4. URL Safety, Drafts Caching, Quote Stripping & Viewer Ergonomics

- **Safelinks/Proofpoint Unwrapping & URL Shortening:** `unwrap_outlook_safelinks()` decodes Microsoft Safelinks, Proofpoint `urldefense`, and FireEye wrappers, normalizes `10.10.10.10/ftp/...` and `3gpp.org/ftp/...`, and collapses duplicate adjacent URLs. With `shorten_long_urls=True`, `linkify_markdown_urls()` shortens doc links to `[📄 <filename>](url)` and >72-char web links to `[🔗 <netloc>/…/<tail>](url)`, while `apply_email_reader_document_style()` connects `QTextBrowser.highlighted[str]` (`handle_reader_link_hover`) for full-URL `QToolTip` previews on `http(s)://` links only.
- **Thread & Proactive `/Drafts/` Caching:** Clicking linked docs caches into `<meeting_dir>/Agenda/ThreadAttachments/<thread_id>/` (`10.10.10.10` $\leftrightarrow$ `Meetings_3GPP_SYNC` failover, 1.5s connect timeout on `10.10.10.10`) and opens via `modules.word_tools.core.document_opener.open_document()` (applying MIP sensitivity labels). Additionally, `cache_drafts_for_messages()` proactively downloads **all** `/Drafts/` URLs (`is_drafts_document_url()`, excluding `/Revisions/` and `/Docs/`) referenced in newly synced/cached Pinned & `GeneralEmailDatabase` emails into `<meeting_dir>/Agenda/DraftsCache/`, reused automatically by `cache_and_open_thread_document_url()`.
- **Option A Message Cards, Localized Quotes & Shared Reader Behavior:**
  - `TrackedThreadDialog`, `TDocEmailsDialog`, `MeetingEmailInboxDialog`, `StandaloneEmailReaderWindow`, and `EmailManagerWindow` share `bind_reader_zoom_settings()` / `zoom_text_widget()` (`Ctrl+Wheel` / `Ctrl+0`, flushing `zoom_steps` to `QSettings` immediately on change), `apply_email_reader_document_style()`, `get_compact_action_bar_qss()`, `get_status_combo_qss()`, `_NotePlainTextEdit`, and discrete `? Help` pills (`create_help_tips_pill()`). `TDocEmailsDialog` receives `tdoc_info`, `open_tdoc_callback`, and `user_data_callback` from `email_compose._open_tdoc_emails()` to render a 2-line TDoc context header (Title, Source, AI/Sub-Item, WI, Status, clickable Family links), always-editable comment bar, compact scope pills (`🔗 Family`, `💬 Quoted`, `🚫 Ignored`), `📄 Open <TDoc>`, and live `My Status` selector alongside `🤖 LLM & Export` and `Launch LLM CLI` (`auto_execute=False`).
  - `export_thread_to_markdown(..., interactive_links=True)` renders Option A state-tinted `<table>` banners (`core.ui.theme.C` tokens, 4px accent column, `split_sender_company()`, `outlook://...` links) preserving `<h3 ...>` for deterministic `QTextDocument` block auto-scroll (`find_unread_heading_blocks()` / `scroll_browser_to_block()`, because Qt 5 MD4C strips `<a name>`), while `interactive_links=False` emits 100% pure Markdown for LLMs.
  - Localized trailing-quote detection (`find_trailing_quote_line_index()` / `strip_trailing_quotes()` / `split_direct_and_quoted_body()` across `From:`/`Sent:`, `发件人:`/`发送时间:`/`主题:`, `Von:`/`Gesendet:`, `De:`/`Envoyé:`) runs at sync and render time (`deduplicate_quotes=True`, or gray blockquotes via `_format_trailing_quoted_history()`), paired with `inject_html_tables_into_body()` GFM tables, `partition_email_attachments()` signature-image filtering, `reload_general_email_from_outlook()` / `reload_tdoc_emails_from_outlook()`, and `Launch LLM CLI` (`auto_execute=False`) staging. (Legacy `TDocEmailsDialog` 800ms auto-read timer is intentionally removed.)
