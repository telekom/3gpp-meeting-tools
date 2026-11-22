"""MAPI conversation/folder collection and thread pinning from .msg or Outlook selection."""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Tuple

from modules.emails.core.outlook_client import (
    PR_FLAG_STATUS_PROP,
    OutlookClient,
    add_mapi_table_columns,
    extract_item_read_flag,
    extract_table_read_flag,
    safe_mapi_row_val,
)
from modules.emails.core.tracked_threads_db import TrackedThreadsDatabase
from .documents import _auto_cache_db_drafts
from .parsing import (
    _extract_folder_paths,
    _extract_folder_sub_flags,
    _mail_item_to_dict,
    auto_suggest_links,
    call_load_wg_email_config as load_wg_email_config,
    normalize_subject,
    parse_msg_file,
)

logger = logging.getLogger(__name__)


def _collect_conversation_items(
    seed_item: Any,
    max_items: int = 250,
    existing_states: Optional[Dict[str, Dict[str, Any]]] = None,
    flag_updates_out: Optional[List[tuple]] = None,
    extracted_cache: Optional[Dict[str, Dict[str, Any]]] = None,
) -> List[Dict[str, Any]]:
    """Uses Outlook's GetConversation() table to fetch new items in a thread while fast-updating known EntryIDs."""
    results: List[Dict[str, Any]] = []
    seen_ids = set()
    try:
        conv = seed_item.GetConversation()
        if not conv:
            return results
        tbl = conv.GetTable()
        add_mapi_table_columns(tbl, ("EntryID", "MessageClass", "UnRead", "FlagStatus", PR_FLAG_STATUS_PROP))
        ns = OutlookClient.get_namespace()
        if not ns:
            return results
        count = 0
        while not tbl.EndOfTable and count < max_items:
            count += 1
            try:
                row = tbl.GetNextRow()
                eid = row("EntryID")
                if not eid or eid in seen_ids:
                    continue
                seen_ids.add(eid)
                msg_cls = str(safe_mapi_row_val(row, "MessageClass", default="") or "")
                if msg_cls and not msg_cls.startswith("IPM.Note"):
                    continue

                cached = existing_states.get(eid) if existing_states else None
                if cached and cached.get("has_body"):
                    state = extract_table_read_flag(row)
                    if state is None:
                        state = extract_item_read_flag(ns, eid)
                    if state is not None and flag_updates_out is not None:
                        if state[0] != cached.get("is_read") or state[1] != cached.get("flag_status"):
                            flag_updates_out.append((eid, state[0], state[1]))
                            cached["is_read"], cached["flag_status"] = state
                    continue

                if extracted_cache is not None and eid in extracted_cache:
                    results.append(extracted_cache[eid])
                    continue

                item = ns.GetItemFromID(eid)
                if getattr(item, "Class", 43) != 43:
                    continue
                d = _mail_item_to_dict(item)
                if extracted_cache is not None:
                    extracted_cache[eid] = d
                results.append(d)
            except Exception:
                continue
    except Exception as e:
        logger.debug(f"GetConversation() unavailable for item: {e}")
    return results


def _collect_matching_from_folders(
    folder_paths: List[Any],
    conversation_ids: List[str],
    subject_patterns: List[str],
    max_scan_per_folder: int = 500,
    seed_live_item: Any = None,
    progress_cb: Optional[Callable[[int, int, str, List[str]], None]] = None,
    cancel_check: Optional[Callable[[], bool]] = None,
    existing_states: Optional[Dict[str, Dict[str, Any]]] = None,
    flag_updates_out: Optional[List[tuple]] = None,
) -> List[Dict[str, Any]]:
    """Scans configured Outlook folders via fast MAPI GetTable and Outlook GetConversation() for all thread messages."""
    matches: List[Dict[str, Any]] = []
    seen_ids = set()
    norm_patterns = {p.lower() for p in subject_patterns if p}
    conv_set = {c for c in conversation_ids if c}

    # 1. If we already have a live Outlook MailItem, pull its full cross-folder conversation first
    conv_pulled_for = set()
    if seed_live_item is not None:
        seed_cid = getattr(seed_live_item, "ConversationID", "") or ""
        if seed_cid:
            conv_pulled_for.add(seed_cid)
        for d in _collect_conversation_items(
            seed_live_item,
            existing_states=existing_states,
            flag_updates_out=flag_updates_out,
        ):
            if d["id"] not in seen_ids:
                seen_ids.add(d["id"])
                matches.append(d)
                if d.get("conversation_id"):
                    conv_set.add(d["conversation_id"])
                    conv_pulled_for.add(d["conversation_id"])

    # 2. Scan configured folders + Outlook Sent Items using bulk MAPI GetTable
    ns = OutlookClient.get_namespace()
    extracted_paths = _extract_folder_paths(folder_paths)
    sub_flags = _extract_folder_sub_flags(folder_paths)
    if ns:
        try:
            sent_folder = ns.GetDefaultFolder(5)  # olFolderSentMail
            sent_path = getattr(sent_folder, "FolderPath", "") or ""
            if sent_path and sent_path.lower() not in {p.lower() for p in extracted_paths}:
                extracted_paths.append(sent_path)
                sub_flags.setdefault(sent_path.lower(), False)
        except Exception:
            pass
    total_folders = len(extracted_paths)
    visited_folder_keys: set = set()
    expanded_folders: List[tuple] = []
    for idx, fpath in enumerate(extracted_paths):
        if cancel_check and cancel_check():
            break
        if progress_cb:
            progress_cb(idx, total_folders, fpath, extracted_paths[idx + 1 :])
        root_folder = OutlookClient.get_folder_by_path(fpath)
        if not root_folder or not ns:
            continue
        inc_sub = sub_flags.get(fpath.lower(), True)
        for sub_f in OutlookClient.iter_folder_tree(
            root_folder,
            include_subfolders=inc_sub,
            visited_keys=visited_folder_keys,
            fallback_path=fpath,
        ):
            expanded_folders.append((sub_f, fpath))

    for folder, fpath in expanded_folders:
        if cancel_check and cancel_check():
            break
        try:
            tbl = folder.GetTable()
            try:
                tbl.Sort("[ReceivedTime]", True)
            except Exception:
                try:
                    tbl.Sort("[SentOn]", True)
                except Exception:
                    pass
            add_mapi_table_columns(
                tbl,
                ("EntryID", "Subject", "MessageClass", "ConversationID", "UnRead", "FlagStatus", PR_FLAG_STATUS_PROP),
            )
            scanned = 0
            while not tbl.EndOfTable and scanned < max_scan_per_folder:
                if cancel_check and cancel_check():
                    break
                scanned += 1
                try:
                    row = tbl.GetNextRow()
                    msg_cls = str(safe_mapi_row_val(row, "MessageClass", default="") or "")
                    if msg_cls and not msg_cls.startswith("IPM.Note"):
                        continue
                    eid = row("EntryID")
                    subj = str(safe_mapi_row_val(row, "Subject", default="") or "")
                    row_cid = str(safe_mapi_row_val(row, "ConversationID", default="") or "")
                    norm = normalize_subject(subj).lower()
                    if (norm and norm in norm_patterns) or (row_cid and row_cid in conv_set):
                        if eid in seen_ids:
                            continue
                        cached = existing_states.get(eid) if existing_states else None
                        if cached and cached.get("has_body") and (not row_cid or row_cid in conv_pulled_for):
                            state = extract_table_read_flag(row)
                            if state is None:
                                state = extract_item_read_flag(ns, eid)
                            if state is not None and flag_updates_out is not None:
                                if state[0] != cached.get("is_read") or state[1] != cached.get("flag_status"):
                                    flag_updates_out.append((eid, state[0], state[1]))
                                    cached["is_read"], cached["flag_status"] = state
                            seen_ids.add(eid)
                            continue

                        item = ns.GetItemFromID(eid)
                        cid = getattr(item, "ConversationID", "") or ""
                        if cid and cid not in conv_pulled_for:
                            conv_pulled_for.add(cid)
                            conv_set.add(cid)
                            for cd in _collect_conversation_items(
                                item,
                                existing_states=existing_states,
                                flag_updates_out=flag_updates_out,
                            ):
                                if cd["id"] not in seen_ids:
                                    seen_ids.add(cd["id"])
                                    matches.append(cd)
                        if cached and cached.get("has_body"):
                            if flag_updates_out is not None:
                                st = extract_table_read_flag(row)
                                if st is None:
                                    is_r = 1 if getattr(item, "UnRead", False) is False else 0
                                    try:
                                        fl_s = 2 if int(getattr(item, "FlagStatus", 0) or 0) == 2 else 0
                                    except Exception:
                                        fl_s = 0
                                    st = (is_r, fl_s)
                                if st[0] != cached.get("is_read") or st[1] != cached.get("flag_status"):
                                    flag_updates_out.append((eid, st[0], st[1]))
                                    cached["is_read"], cached["flag_status"] = st
                            seen_ids.add(eid)
                        else:
                            d = _mail_item_to_dict(item)
                            if d["id"] not in seen_ids:
                                seen_ids.add(d["id"])
                                matches.append(d)
                except Exception:
                    continue
        except Exception as e:
            logger.warning(f"Error scanning Outlook folder '{fpath}': {e}")
    return matches


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
    merged_ais = list(dict.fromkeys((existing.get("linked_ais") or []) + (linked_ais if linked_ais is not None else suggested.get("ais") or [])))
    merged_wis = list(dict.fromkeys((existing.get("linked_wis") or []) + (linked_wis if linked_wis is not None else suggested.get("wis") or [])))
    merged_tdocs = list(dict.fromkeys((existing.get("linked_tdocs") or []) + (linked_tdocs if linked_tdocs is not None else suggested.get("tdocs") or [])))
    merged_subs = list(dict.fromkeys((existing.get("linked_sub_items") or []) + (linked_sub_items if linked_sub_items is not None else suggested.get("sub_items") or [])))
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
        email_records.append({
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
        })
        for m in matches:
            match_records.append({
                "email_id": eid,
                "tdoc_id": m["tdoc_id"],
                "rev_matched": m.get("rev_matched", ""),
                "match_location": m.get("location", "Body"),
            })
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
        linked_sub_items=linked_sub_items if linked_sub_items is not None else suggested.get("sub_items"),
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
    found = _collect_matching_from_folders(folders, conv_ids, patterns, seed_live_item=seed_live_item)
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

    # Group seeds into distinct threads by ConversationID or normalized subject
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
            buckets.append({
                "title": norm_subj,
                "norm_subj": norm_subj,
                "cids": [cid] if cid else [],
                "items": [],
            })
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
        combined_body = " \n ".join(str(s.get("clean_body") or s.get("body_text") or "")[:1500] for s in seeds)
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

    # Collect final stored messages across pinned threads for Drafts caching & GeneralEmailDB indexing
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
    for mp in (msg_paths or []):
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
        db, seed, meeting_id=meeting_id, wg=wg,
        linked_ais=linked_ais, linked_wis=linked_wis, linked_tdocs=linked_tdocs,
        agenda_map=agenda_map, valid_wis=valid_wis, tdoc_meta_map=tdoc_meta_map,
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
