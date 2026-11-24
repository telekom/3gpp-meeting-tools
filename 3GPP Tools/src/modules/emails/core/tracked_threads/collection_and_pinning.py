"""Backward-compatible re-export shim -> mapi_collection and pinning modules."""

from __future__ import annotations

from .mapi_collection import (
    _collect_conversation_items,
    _collect_matching_from_folders,
)
from .pinning import (
    _index_messages_into_general_db,
    _merge_suggested_links_into_thread,
    _pin_from_seed_dict,
    pin_thread_from_msg,
    pin_thread_from_outlook_selection,
    pin_threads_from_msg_files,
    pin_threads_from_seeds_batch,
)

__all__ = [
    "_collect_conversation_items",
    "_collect_matching_from_folders",
    "_merge_suggested_links_into_thread",
    "_index_messages_into_general_db",
    "_pin_from_seed_dict",
    "pin_threads_from_seeds_batch",
    "pin_threads_from_msg_files",
    "pin_thread_from_msg",
    "pin_thread_from_outlook_selection",
]
