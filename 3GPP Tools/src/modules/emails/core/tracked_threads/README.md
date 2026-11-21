# Tracked Threads & Discovery Service Package (`tracked_threads/`)

Modular decomposition of the Outlook MAPI sync, discovery, Drafts caching, and topic auto-linking service previously housed in `tracked_threads_service.py`.

## Submodule Navigation Map

| Module | Responsibility | Key Functions / Symbols |
|---|---|---|
| [`parsing.py`](parsing.py) | Subject/token parsing, MailItem dict normalization, `.msg` reading, folder path extraction, shim config forwarding | `auto_suggest_links`, `normalize_subject`, `_mail_item_to_dict`, `parse_msg_file`, `_extract_folder_paths`, `call_load_wg_email_config` |
| [`documents.py`](documents.py) | Proactive `/Drafts/` caching, TDoc URL failover download/open, LLM workspace prep | `cache_drafts_for_messages`, `cache_and_open_thread_document_url`, `cache_all_thread_documents_and_attachments`, `prepare_thread_llm_workspace` |
| [`collection_and_pinning.py`](collection_and_pinning.py) | MAPI `GetConversation`/`GetTable` collection and `.msg`/selection pinning | `_collect_conversation_items`, `_collect_matching_from_folders`, `pin_thread_from_msg`, `pin_thread_from_outlook_selection` |
| [`sync.py`](sync.py) | Single-thread, single-pass batch folder sync, scoped inbox sync, shim refresh forwarding | `refresh_thread`, `refresh_threads_batch`, `refresh_meeting_threads`, `sync_scoped_inbox_items`, `call_refresh_thread` |
| [`discovery.py`](discovery.py) | Per-meeting exclusions, recent folder scanning, classification, subject pinning | `load_meeting_discovery_config`, `save_meeting_discovery_config`, `list_recent_threads_in_folders`, `classify_discovered_threads`, `pin_thread_by_subject_or_cid` |
| [`outlook_actions.py`](outlook_actions.py) | Outlook COM open, reload, attachments, reply, and read/flag mutations | `open_outlook_message`, `reload_outlook_message`, `open_outlook_attachment`, `reply_to_outlook_message`, `mark_thread_read_with_outlook`, `set_thread_message_read_with_outlook`, `set_thread_message_flag_with_outlook` |
| [`__init__.py`](__init__.py) | Package exports | Re-exports all public and internal symbols above |

## Backward Compatibility & Shim Monkeypatching

- [`../tracked_threads_service.py`](../tracked_threads_service.py) remains as a thin re-export shim exposing all symbols (including `load_wg_email_config` and `OutlookClient`).
- Submodules invoke `call_load_wg_email_config()` and `call_refresh_thread()`, checking `sys.modules["modules.emails.core.tracked_threads_service"]` so test monkeypatches on `tracked_threads_service.load_wg_email_config` and `tracked_threads_service.refresh_thread` apply seamlessly.
