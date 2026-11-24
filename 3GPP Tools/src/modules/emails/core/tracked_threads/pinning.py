"""Thread pinning from .msg files, seed dicts, and Outlook ActiveExplorer selections."""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from modules.emails.core.tracked_threads_db import TrackedThreadsDatabase
from .documents import _auto_cache_db_drafts
from .mapi_collection import _collect_conversation_items, _collect_matching_from_folders
from .parsing import (
    _extract_folder_paths,
    _mail_item_to_dict,
    auto_suggest_links,
    call_load_wg_email_config as load_wg_email_config,
    normalize_subject,
    parse_msg_file,
)

logger = logging.getLogger(__name__)


def _merge_suggested_links_into_thread(
    db: TrackedThreadsDatabase,
    thread_id: str,
    suggested: Dict[str, List[str]],
    *,
    linked_ais: Optional[List[str]] = None,
    linked_wis: Optional[List[str]] = None,
    linked_tdocs: Optional[List[str]] = None,
    linked_sub_items: Optional[List[str]] = None,
):
    """Unions new suggested AIs/Sub-Topics/WIs/TDocs onto an existing tracked thread."""
    existing = db.get_thread(thread_id)
    if not existing:
        return
    merged_ais = list(
        dict.fromkeys(
            (existing.get("linked_ais") or [])
            + (linked_ais if linked_ais is not None else suggested.get("ais") or [])
        )
    )
    merged_wis = list(
        dict.fromkeys(
            (existing.get("linked_wis") or [])
            + (linked_wis if linked_wis is not None else suggested.get("wis") or [])
        )
    )
    merged_tdocs = list(
        dict.fromkeys(
            (existing.get("linked_tdocs") or [])
            + (linked_tdocs if linked_tdocs is not None else suggested.get("tdocs") or [])
        )
    )
    merged_subs = list(
        dict.fromkeys(
            (existing.get("linked_sub_items") or [])
            + (linked_sub_items if linked_sub_items is not None else suggested.get("sub_items") or [])
        )
    )
    db.update_thread_links(
        thread_id,
        linked_ais=merged_ais,
        linked_wis=merged_wis,
        linked_tdocs=merged_tdocs,
        linked_sub_items=merged_subs,
    )


def _index_messages_into_general_db(
    general_email_db: Optional[Any],
    messages: List[Dict[str, Any]],
    *,
    folder_tag: str = "Imported",
    fallback_tdocs: Optional[List[str]] = None,
) -> int:
    """Indexes TDoc-matched messages into GeneralEmailDatabase (`emails.db`) so Related Emails dialogs show them."""
    if general_email_db is None or not messages or not hasattr(general_email_db, "save_emails_batch"):
        return 0
    from modules.emails.core.general_email_sync import GeneralEmailSyncThread

    email_records: List[Dict[str, Any]] = []
    match_records: List[Dict[str, Any]] = []
    for msg in messages:
        eid = str(msg.get("id") or "").strip()
        if not eid:
            continue
        subj = str(msg.get("subject") or "")
        body = str(msg.get("body_text") or msg.get("clean_body") or "")
        atts = list(msg.get("attachments") or [])
        matches = GeneralEmailSyncThread.extract_tdocs_from_parts(subj, body, attachments=atts)
        if not matches and fallback_tdocs:
            matches = [
                {"tdoc_id": str(t).strip().upper(), "rev_matched": "", "location": "Thread"}
                for t in fallback_tdocs
                if str(t or "").strip()
            ]
        if not matches:
            continue
        email_records.append(
            {
                "id": eid,
                "folder_tag": folder_tag,
                "subject": subj,
                "sender_name": str(msg.get("sender_name") or ""),
                "sender_email": str(msg.get("sender_email") or ""),
                "company": str(msg.get("company") or ""),
                "date_received": str(msg.get("date_sent") or msg.get("date_received") or ""),
                "body_text": str(msg.get("body_text") or "")[:12000],
                "clean_body": str(msg.get("clean_body") or "")[:12000],
                "attachments": atts,
                "is_read": int(msg.get("is_read", 0) or 0),
                "is_ignored": 0,
                "flag_status": 2 if int(msg.get("flag_status", 0) or 0) == 2 else 0,
            }
        )
        for m in matches:
            match_records.append(
                {
                    "email_id": eid,
                    "tdoc_id": m["tdoc_id"],
                    "rev_matched": m.get("rev_matched", ""),
                    "match_location": m.get("location", "Body"),
                }
            )
    if not email_records:
        return 0
    general_email_db.save_emails_batch(email_records, match_records, sync_unread=False)
    return len(email_records)


def _pin_from_seed_dict(
    db: TrackedThreadsDatabase,
    seed: Dict[str, Any],
    meeting_id: str,
    wg: str,
    linked_ais: Optional[List[str]] = None,
    linked_wis: Optional[List[str]] = None,
    linked_tdocs: Optional[List[str]] = None,
    linked_sub_items: Optional[List[str]] = None,
    seed_live_item: Any = None,
    agenda_map: Optional[Dict[str, Any]] = None,
    valid_wis: Optional[List[str]] = None,
    tdoc_meta_map: Optional[Dict[str, Dict[str, Any]]] = None,
    general_email_db: Optional[Any] = None,
) -> str:
    norm_subj = normalize_subject(seed["subject"])
    folders = _extract_folder_paths(load_wg_email_config(wg)) if wg else []
    suggested = auto_suggest_links(
        seed["subject"],
        seed.get("clean_body") or seed.get("body_text", ""),
        agenda_map=agenda_map,
        valid_wis=valid_wis,
        tdoc_meta_map=tdoc_meta_map,
        extra_tdocs=linked_tdocs,
        attachments=seed.get("attachments") or [],
    )
    thread_id = db.create_or_get_thread(
        title=norm_subj or seed["subject"],
        normalized_subject=norm_subj,
        conversation_id=seed.get("conversation_id", ""),
        outlook_folders=folders,
        meeting_id=meeting_id,
        wg=wg,
        linked_ais=linked_ais if linked_ais is not None else suggested["ais"],
        linked_wis=linked_wis if linked_wis is not None else suggested["wis"],
        linked_tdocs=linked_tdocs if linked_tdocs is not None else suggested["tdocs"],
        linked_sub_items=(
            linked_sub_items if linked_sub_items is not None else suggested.get("sub_items")
        ),
    )
    _merge_suggested_links_into_thread(
        db,
        thread_id,
        suggested,
        linked_ais=linked_ais,
        linked_wis=linked_wis,
        linked_tdocs=linked_tdocs,
        linked_sub_items=linked_sub_items,
    )

    conv_ids = [seed["conversation_id"]] if seed.get("conversation_id") else []
    patterns = [norm_subj] if norm_subj else []
    found = _collect_matching_from_folders(
        folders, conv_ids, patterns, seed_live_item=seed_live_item
    )
    if not any(
        m["id"] == seed["id"]
        or (
            seed.get("date_sent")
            and m.get("date_sent", "")[:16] == seed["date_sent"][:16]
            and m.get("sender_email", "").lower() == seed.get("sender_email", "").lower()
        )
        for m in found
    ):
        found.append(seed)

    db.upsert_messages(thread_id, found)
    discovered_cids = [m["conversation_id"] for m in found if m.get("conversation_id")]
    if discovered_cids:
        db.add_conversation_ids(thread_id, discovered_cids)
    _auto_cache_db_drafts(db, found)
    if general_email_db is not None:
        _index_messages_into_general_db(
            general_email_db,
            found,
            folder_tag="Imported",
            fallback_tdocs=suggested.get("tdocs"),
        )
    return thread_id


def pin_threads_from_seeds_batch(
    db: TrackedThreadsDatabase,
    seeds_with_items: List[Tuple[Dict[str, Any], Any]],
    *,
    meeting_id: str,
    wg: str,
    agenda_map: Optional[Dict[str, Any]] = None,
    valid_wis: Optional[List[str]] = None,
    tdoc_meta_map: Optional[Dict[str, Dict[str, Any]]] = None,
    general_email_db: Optional[Any] = None,
    refresh_outlook: bool = True,
) -> Dict[str, Any]:
    """
    Batch-pins a collection of dropped emails/threads `(seed_dict, live_outlook_item_or_None)`.
    Groups seeds belonging to the same conversation/subject so each distinct thread is created once,
    resolves each thread's own AIs/Sub-Topics/WIs/TDocs independently, runs a single shared Outlook
    batch scan when `refresh_outlook=True`, and indexes TDoc matches into `general_email_db`.
    """
    if not seeds_with_items:
        return {"thread_ids": [], "message_count": 0}

    folders = _extract_folder_paths(load_wg_email_config(wg)) if wg else []

    buckets: List[Dict[str, Any]] = []
    cid_to_bucket: Dict[str, int] = {}
    subj_to_bucket: Dict[str, int] = {}

    for seed, live_item in seeds_with_items:
        if not seed or not seed.get("subject"):
            continue
        cid = str(seed.get("conversation_id") or "").strip()
        norm_subj = normalize_subject(seed["subject"]) or seed["subject"]
        subj_key = norm_subj.lower()

        b_idx = None
        if cid and cid in cid_to_bucket:
            b_idx = cid_to_bucket[cid]
        elif subj_key in subj_to_bucket:
            b_idx = subj_to_bucket[subj_key]

        if b_idx is None:
            b_idx = len(buckets)
            buckets.append(
                {
                    "title": norm_subj,
                    "norm_subj": norm_subj,
                    "cids": [cid] if cid else [],
                    "items": [],
                }
            )
        bucket = buckets[b_idx]
        if cid:
            cid_to_bucket[cid] = b_idx
            if cid not in bucket["cids"]:
                bucket["cids"].append(cid)
        subj_to_bucket[subj_key] = b_idx
        bucket["items"].append((seed, live_item))

    pinned_ids: List[str] = []
    threads_for_refresh: List[Dict[str, Any]] = []
    pulled_cids: set = set()
    all_indexed_msgs: List[Dict[str, Any]] = []

    for bucket in buckets:
        seeds = [s for s, _ in bucket["items"]]
        combined_subj = " \n ".join(str(s.get("subject") or "") for s in seeds)
        combined_body = " \n ".join(
            str(s.get("clean_body") or s.get("body_text") or "")[:1500] for s in seeds
        )
        combined_atts = [a for s in seeds for a in (s.get("attachments") or [])]

        suggested = auto_suggest_links(
            combined_subj,
            combined_body,
            agenda_map=agenda_map,
            valid_wis=valid_wis,
            tdoc_meta_map=tdoc_meta_map,
            attachments=combined_atts,
        )
        primary_cid = bucket["cids"][0] if bucket["cids"] else ""
        tid = db.create_or_get_thread(
            title=bucket["title"],
            normalized_subject=bucket["norm_subj"],
            conversation_id=primary_cid,
            outlook_folders=folders,
            meeting_id=meeting_id,
            wg=wg,
            linked_ais=suggested["ais"],
            linked_wis=suggested["wis"],
            linked_tdocs=suggested["tdocs"],
            linked_sub_items=suggested.get("sub_items"),
        )
        _merge_suggested_links_into_thread(db, tid, suggested)
        if bucket["cids"]:
            db.add_conversation_ids(tid, bucket["cids"])

        bucket_msgs: List[Dict[str, Any]] = list(seeds)
        for seed, live_item in bucket["items"]:
            if live_item is not None:
                scid = str(seed.get("conversation_id") or "").strip()
                if not scid or scid not in pulled_cids:
                    if scid:
                        pulled_cids.add(scid)
                    for cm in _collect_conversation_items(live_item):
                        bucket_msgs.append(cm)

        db.upsert_messages(tid, bucket_msgs)
        cids_found = [m["conversation_id"] for m in bucket_msgs if m.get("conversation_id")]
        if cids_found:
            db.add_conversation_ids(tid, cids_found)

        if tid not in pinned_ids:
            pinned_ids.append(tid)
        th_rec = db.get_thread(tid)
        if th_rec:
            threads_for_refresh.append(th_rec)
        all_indexed_msgs.extend(bucket_msgs)

    if refresh_outlook and threads_for_refresh:
        try:
            from .sync import refresh_threads_batch

            refresh_threads_batch(db, threads_for_refresh)
        except Exception as exc:
            logger.debug(f"Batch refresh after drop skipped cleanly: {exc}")

    final_msgs: List[Dict[str, Any]] = []
    for tid in pinned_ids:
        final_msgs.extend(db.get_thread_messages(tid))
    if not final_msgs:
        final_msgs = all_indexed_msgs

    _auto_cache_db_drafts(db, final_msgs)
    if general_email_db is not None:
        _index_messages_into_general_db(general_email_db, final_msgs, folder_tag="Imported")

    return {"thread_ids": pinned_ids, "message_count": len(final_msgs)}


def pin_threads_from_msg_files(
    db: TrackedThreadsDatabase,
    msg_paths: List[Path],
    *,
    meeting_id: str,
    wg: str,
    agenda_map: Optional[Dict[str, Any]] = None,
    valid_wis: Optional[List[str]] = None,
    tdoc_meta_map: Optional[Dict[str, Dict[str, Any]]] = None,
    general_email_db: Optional[Any] = None,
    refresh_outlook: bool = True,
) -> Dict[str, Any]:
    """Batch-imports multiple `.msg` files, grouping same-thread emails and syncing once."""
    seeds_with_items: List[Tuple[Dict[str, Any], Any]] = []
    for mp in msg_paths or []:
        seed = parse_msg_file(Path(mp))
        if seed:
            seeds_with_items.append((seed, None))
    return pin_threads_from_seeds_batch(
        db,
        seeds_with_items,
        meeting_id=meeting_id,
        wg=wg,
        agenda_map=agenda_map,
        valid_wis=valid_wis,
        tdoc_meta_map=tdoc_meta_map,
        general_email_db=general_email_db,
        refresh_outlook=refresh_outlook,
    )


def pin_thread_from_msg(
    db: TrackedThreadsDatabase,
    msg_path: Path,
    meeting_id: str,
    wg: str,
    linked_ais: Optional[List[str]] = None,
    linked_wis: Optional[List[str]] = None,
    linked_tdocs: Optional[List[str]] = None,
    agenda_map: Optional[Dict[str, Any]] = None,
    valid_wis: Optional[List[str]] = None,
    tdoc_meta_map: Optional[Dict[str, Dict[str, Any]]] = None,
    general_email_db: Optional[Any] = None,
) -> Optional[str]:
    """Parses a dropped .msg, locates the full thread in configured Outlook folders, and pins it."""
    seed = parse_msg_file(msg_path)
    if not seed:
        return None
    return _pin_from_seed_dict(
        db,
        seed,
        meeting_id=meeting_id,
        wg=wg,
        linked_ais=linked_ais,
        linked_wis=linked_wis,
        linked_tdocs=linked_tdocs,
        agenda_map=agenda_map,
        valid_wis=valid_wis,
        tdoc_meta_map=tdoc_meta_map,
        general_email_db=general_email_db,
    )


def pin_thread_from_outlook_selection(
    db: TrackedThreadsDatabase,
    meeting_id: str,
    wg: str,
    agenda_map: Optional[Dict[str, Any]] = None,
    valid_wis: Optional[List[str]] = None,
    tdoc_meta_map: Optional[Dict[str, Dict[str, Any]]] = None,
    general_email_db: Optional[Any] = None,
) -> List[str]:
    """Pins thread(s) from the currently selected MailItem(s) in Microsoft Outlook's ActiveExplorer."""
    import win32com.client

    seeds_with_items: List[Tuple[Dict[str, Any], Any]] = []
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        explorer = outlook.ActiveExplorer()
        if not explorer:
            return []
        selection = explorer.Selection
        for i in range(1, selection.Count + 1):
            item = selection.Item(i)
            if getattr(item, "Class", 0) != 43:
                continue
            seed = _mail_item_to_dict(item)
            if seed:
                seeds_with_items.append((seed, item))
    except Exception as e:
        logger.error(f"Failed to read Outlook ActiveExplorer.Selection: {e}")
        return []

    if not seeds_with_items:
        return []
    res = pin_threads_from_seeds_batch(
        db,
        seeds_with_items,
        meeting_id=meeting_id,
        wg=wg,
        agenda_map=agenda_map,
        valid_wis=valid_wis,
        tdoc_meta_map=tdoc_meta_map,
        general_email_db=general_email_db,
        refresh_outlook=True,
    )
    return list(res.get("thread_ids") or [])
