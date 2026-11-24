# Tracked Threads & Discovery Service Package (`tracked_threads/`)

Modular decomposition of the Outlook MAPI sync, discovery, Drafts caching, and topic auto-linking service previously housed in `tracked_threads_service.py`.

## Submodule Navigation Map

| Module | Responsibility | Key Functions / Symbols |
|---|---|---|
| [`parsing.py`](parsing.py) | Subject/token parsing, AI→WI inference, MailItem dict normalization, `.msg` reading, folder path extraction, shim config forwarding | `auto_suggest_links`, `infer_wis_for_ais`, `normalize_subject`, `_mail_item_to_dict`, `parse_msg_file`, `_extract_folder_paths`, `call_load_wg_email_config` |
| [`documents.py`](documents.py) | Proactive `/Drafts/` caching, TDoc URL failover download/open, LLM workspace prep | `cache_drafts_for_messages`, `cache_and_open_thread_document_url`, `cache_all_thread_documents_and_attachments`, `prepare_thread_llm_workspace` |
| [`mapi_collection.py`](mapi_collection.py) | MAPI `GetConversation()` and bulk `GetTable()` folder message collection | `_collect_conversation_items`, `_collect_matching_from_folders` |
| [`pinning.py`](pinning.py) | Batch multi-thread `.msg`/selection pinning & dual-DB (`general_email_db`) indexing | `_pin_from_seed_dict`, `pin_threads_from_seeds_batch`, `pin_threads_from_msg_files`, `pin_thread_from_msg`, `pin_thread_from_outlook_selection` |
| [`collection_and_pinning.py`](collection_and_pinning.py) | Thin re-export shim preserving legacy imports from `mapi_collection` + `pinning` | `_collect_conversation_items`, `pin_threads_from_seeds_batch`, `pin_thread_from_msg` |
| [`sync.py`](sync.py) | Single-thread, single-pass batch folder sync with MAPI `Table.Restrict` watermark filtering, shim refresh forwarding | `refresh_thread`, `refresh_threads_batch`, `refresh_meeting_threads`, `sync_scoped_inbox_items` (re-export), `call_refresh_thread` |
| [`scoped_sync.py`](scoped_sync.py) | Fast filter-scoped Inbox synchronization via Outlook `EntryID` + `GetConversation().GetTable()` | `sync_scoped_inbox_items` |
| [`discovery.py`](discovery.py) | Per-meeting exclusions, recent folder scanning, classification, subject pinning | `load_meeting_discovery_config`, `save_meeting_discovery_config`, `list_recent_threads_in_folders`, `classify_discovered_threads`, `pin_thread_by_subject_or_cid` |
| [`outlook_actions.py`](outlook_actions.py) | Outlook COM open, reload, attachments, reply, and read/flag mutations | `open_outlook_message`, `reload_outlook_message`, `open_outlook_attachment`, `reply_to_outlook_message`, `mark_thread_read_with_outlook`, `set_thread_message_read_with_outlook`, `set_thread_message_flag_with_outlook` |
| [`__init__.py`](__init__.py) | Package exports | Re-exports all public and internal symbols above |

## Backward Compatibility & Key Invariants

- [`../tracked_threads_service.py`](../tracked_threads_service.py) remains as a thin re-export shim exposing all symbols (including `load_wg_email_config` and `OutlookClient`).
- Submodules invoke `call_load_wg_email_config()` and `call_refresh_thread()`, checking `sys.modules["modules.emails.core.tracked_threads_service"]` so test monkeypatches on `tracked_threads_service.load_wg_email_config` and `tracked_threads_service.refresh_thread` apply seamlessly.
- **Regex Boundaries:** Never use `\b` around TDoc, `AI#`, or WI regexes because `_` is a `\w` character and breaks underscore-delimited names like `S2-2612919_was_S2-2612519`; always use `(?<![A-Za-z0-9])...(?![A-Za-z0-9])`.
- **Automatic Deduplication:** Both `GeneralEmailDatabase` and `TrackedThreadsDatabase` invoke `email_deduplicator.py` on init and write paths to collapse multi-folder / `.msg` duplicates losslessly.
