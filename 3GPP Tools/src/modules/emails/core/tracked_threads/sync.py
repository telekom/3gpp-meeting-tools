"""Single-thread, single-pass batch folder sync, and scoped inbox synchronization."""

from __future__ import annotations

import logging
import sys
from typing import Any, Callable, Dict, List, Optional

from modules.emails.core.outlook_client import (
    PR_FLAG_STATUS_PROP,
    OutlookClient,
    add_mapi_table_columns,
    extract_item_read_flag,
    extract_table_read_flag,
    safe_mapi_row_val,
)
from modules.emails.core.tracked_threads_db import TrackedThreadsDatabase
from .collection_and_pinning import (
    _collect_conversation_items,
    _collect_matching_from_folders,
)
from .documents import _auto_cache_db_drafts
from .parsing import (
    _extract_folder_paths,
    _extract_folder_sub_flags,
    _mail_item_to_dict,
    call_load_wg_email_config as load_wg_email_config,
    normalize_subject,
)

logger = logging.getLogger(__name__)


def refresh_thread(
    db: TrackedThreadsDatabase,
    thread_id: str,
    progress_cb: Optional[Callable[[int, int, str, List[str]], None]] = None,
    cancel_check: Optional[Callable[[], bool]] = None,
) -> int:
    """Refreshes a single tracked thread from its configured Outlook folders. Returns new message count."""
    thread = db.get_thread(thread_id)
    if not thread:
        return 0
    folders = thread.get("outlook_folders") or (load_wg_email_config(thread.get("wg", "")) if thread.get("wg") else [])
    if not folders:
        return 0
    existing_states = db.get_existing_message_states(thread_id)
    flag_updates: List[tuple] = []
    found = _collect_matching_from_folders(
        folders,
        thread.get("conversation_ids") or [],
        thread.get("subject_patterns") or [thread.get("normalized_subject", "")],
        progress_cb=progress_cb,
        cancel_check=cancel_check,
        existing_states=existing_states,
        flag_updates_out=flag_updates,
    )
    discovered_cids = [m["conversation_id"] for m in found if m.get("conversation_id")]
    if discovered_cids:
        db.add_conversation_ids(thread_id, discovered_cids)
    new_cnt = db.upsert_messages(thread_id, found, sync_unread=True)
    if flag_updates:
        db.update_message_flags_batch(thread_id, flag_updates)
    _auto_cache_db_drafts(db, db.get_thread_messages(thread_id), cancel_check=cancel_check)
    return new_cnt


def refresh_threads_batch(
    db: TrackedThreadsDatabase,
    threads: List[Dict[str, Any]],
    progress_cb: Optional[Callable[[int, int, str, List[str]], None]] = None,
    cancel_check: Optional[Callable[[], bool]] = None,
    max_scan_per_folder: int = 500,
) -> Dict[str, int]:
    """Refreshes multiple tracked threads in a single MAPI GetTable pass per Outlook folder, skipping known EntryIDs."""
    if not threads:
        return {}

    ns = OutlookClient.get_namespace()
    if not ns:
        _auto_cache_db_drafts(
            db,
            [m for t in threads if t.get("thread_id") for m in db.get_thread_messages(t["thread_id"])],
            cancel_check=cancel_check,
        )
        return {}

    tid_map: Dict[str, Dict[str, Any]] = {}
    tid_existing: Dict[str, Dict[str, Dict[str, Any]]] = {}
    conv_to_tids: Dict[str, set] = {}
    subj_to_tids: Dict[str, set] = {}
    folder_to_tids: Dict[str, set] = {}
    folder_sub_flags: Dict[str, bool] = {}
    folder_order: List[str] = []
    all_tids: set = set()

    for t in threads:
        tid = str(t.get("thread_id") or "").strip()
        if not tid:
            continue
        raw_folders = t.get("outlook_folders") or (load_wg_email_config(t.get("wg", "")) if t.get("wg") else [])
        extracted = _extract_folder_paths(raw_folders)
        if not extracted:
            continue
        folder_sub_flags.update(_extract_folder_sub_flags(raw_folders))
        if t.get("wg"):
            folder_sub_flags.update(_extract_folder_sub_flags(load_wg_email_config(t.get("wg", ""))))
        tid_map[tid] = t
        all_tids.add(tid)
        tid_existing[tid] = db.get_existing_message_states(tid)
        for cid in (t.get("conversation_ids") or []):
            c_clean = str(cid or "").strip()
            if c_clean:
                conv_to_tids.setdefault(c_clean, set()).add(tid)
        patterns = t.get("subject_patterns") or [t.get("normalized_subject", "")]
        for pat in patterns:
            p_clean = str(pat or "").strip().lower()
            if p_clean:
                subj_to_tids.setdefault(p_clean, set()).add(tid)
        for fp in extracted:
            fp_key = fp.lower()
            if fp_key not in folder_to_tids:
                folder_to_tids[fp_key] = set()
                folder_order.append(fp)
            folder_to_tids[fp_key].add(tid)

    if not tid_map or not folder_order:
        _auto_cache_db_drafts(
            db,
            [m for t in threads if t.get("thread_id") for m in db.get_thread_messages(t["thread_id"])],
            cancel_check=cancel_check,
        )
        return {}

    try:
        sent_folder = ns.GetDefaultFolder(5)  # olFolderSentMail
        sent_path = getattr(sent_folder, "FolderPath", "") or ""
        if sent_path:
            sk = sent_path.lower()
            if sk not in folder_to_tids:
                folder_order.append(sent_path)
            folder_to_tids.setdefault(sk, set()).update(all_tids)
            folder_sub_flags.setdefault(sk, False)
    except Exception:
        pass

    new_msgs_by_tid: Dict[str, List[Dict[str, Any]]] = {tid: [] for tid in tid_map}
    seen_eids_by_tid: Dict[str, set] = {tid: set() for tid in tid_map}
    flag_updates_by_tid: Dict[str, List[tuple]] = {tid: [] for tid in tid_map}
    cids_by_tid: Dict[str, set] = {tid: set() for tid in tid_map}
    conv_pulled_for: set = set()
    extracted_cache: Dict[str, Dict[str, Any]] = {}

    total_folders = len(folder_order)
    visited_folder_keys: set = set()
    expanded_folder_jobs: List[tuple] = []
    for f_idx, fpath in enumerate(folder_order):
        if cancel_check and cancel_check():
            break
        if progress_cb:
            short_name = fpath.replace("\\", "/").rstrip("/").rsplit("/", 1)[-1] or fpath
            progress_cb(f_idx, total_folders, f"Folder: {short_name}", folder_order[f_idx + 1 :])
        root_folder = OutlookClient.get_folder_by_path(fpath)
        if not root_folder:
            continue
        allowed_tids = folder_to_tids.get(fpath.lower(), set())
        if not allowed_tids:
            continue
        inc_sub = folder_sub_flags.get(fpath.lower(), True)
        for sub_f in OutlookClient.iter_folder_tree(
            root_folder,
            include_subfolders=inc_sub,
            visited_keys=visited_folder_keys,
            fallback_path=fpath,
        ):
            expanded_folder_jobs.append((sub_f, fpath, allowed_tids))

    for folder, fpath, allowed_tids in expanded_folder_jobs:
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
                    eid = str(safe_mapi_row_val(row, "EntryID", default="") or "")
                    if not eid:
                        continue
                    subj = str(safe_mapi_row_val(row, "Subject", default="") or "")
                    row_cid = str(safe_mapi_row_val(row, "ConversationID", default="") or "")
                    norm = normalize_subject(subj).lower()

                    matched = set()
                    if row_cid and row_cid in conv_to_tids:
                        matched |= conv_to_tids[row_cid]
                    if norm and norm in subj_to_tids:
                        matched |= subj_to_tids[norm]
                    matched &= allowed_tids
                    if not matched:
                        continue

                    pending_tids = [t_id for t_id in matched if eid not in seen_eids_by_tid[t_id]]
                    if not pending_tids:
                        continue

                    needs_body_tids = []
                    cached_tids = []
                    for t_id in pending_tids:
                        c_info = tid_existing[t_id].get(eid)
                        if c_info and c_info.get("has_body"):
                            cached_tids.append(t_id)
                        else:
                            needs_body_tids.append(t_id)

                    need_conv_pull = bool(row_cid and row_cid not in conv_pulled_for)
                    if cached_tids and not needs_body_tids and not need_conv_pull:
                        st = extract_table_read_flag(row)
                        if st is None:
                            st = extract_item_read_flag(ns, eid)
                        if st is not None:
                            for t_id in cached_tids:
                                c_info = tid_existing[t_id][eid]
                                if st[0] != c_info.get("is_read") or st[1] != c_info.get("flag_status"):
                                    flag_updates_by_tid[t_id].append((eid, st[0], st[1]))
                                    c_info["is_read"], c_info["flag_status"] = st
                        for t_id in cached_tids:
                            seen_eids_by_tid[t_id].add(eid)
                        continue

                    item = None
                    if eid not in extracted_cache or (not row_cid or row_cid not in conv_pulled_for):
                        item = ns.GetItemFromID(eid)
                        cid = getattr(item, "ConversationID", "") or row_cid
                    else:
                        cid = row_cid

                    if cid and cid not in conv_pulled_for:
                        conv_pulled_for.add(cid)
                        shared_cached_eids = set.intersection(
                            *(
                                {ceid for ceid, cst in tid_existing[t_id].items() if cst.get("has_body")}
                                for t_id in pending_tids
                            )
                        )
                        shared_existing = (
                            {ceid: dict(tid_existing[pending_tids[0]][ceid]) for ceid in shared_cached_eids}
                            if shared_cached_eids
                            else {}
                        )
                        conv_flag_updates: List[tuple] = []
                        conv_siblings = _collect_conversation_items(
                            item,
                            existing_states=shared_existing or None,
                            flag_updates_out=conv_flag_updates,
                            extracted_cache=extracted_cache,
                        )
                        for sib_eid, sib_ir, sib_fs in conv_flag_updates:
                            for t_id in pending_tids:
                                seen_eids_by_tid[t_id].add(sib_eid)
                                c_info = tid_existing[t_id].get(sib_eid)
                                if c_info and (sib_ir != c_info.get("is_read") or sib_fs != c_info.get("flag_status")):
                                    flag_updates_by_tid[t_id].append((sib_eid, sib_ir, sib_fs))
                                    c_info["is_read"], c_info["flag_status"] = sib_ir, sib_fs
                        for cd in conv_siblings:
                            cd_id = cd["id"]
                            if cd.get("conversation_id"):
                                for t_id in pending_tids:
                                    cids_by_tid[t_id].add(cd["conversation_id"])
                            for t_id in pending_tids:
                                if cd_id in seen_eids_by_tid[t_id]:
                                    continue
                                seen_eids_by_tid[t_id].add(cd_id)
                                c_info = tid_existing[t_id].get(cd_id)
                                if c_info and c_info.get("has_body"):
                                    if cd["is_read"] != c_info.get("is_read") or cd["flag_status"] != c_info.get("flag_status"):
                                        flag_updates_by_tid[t_id].append((cd_id, cd["is_read"], cd["flag_status"]))
                                        c_info["is_read"], c_info["flag_status"] = cd["is_read"], cd["flag_status"]
                                else:
                                    new_msgs_by_tid[t_id].append(cd)

                    if needs_body_tids:
                        if eid in extracted_cache:
                            d = extracted_cache[eid]
                        else:
                            if item is None:
                                item = ns.GetItemFromID(eid)
                            d = _mail_item_to_dict(item)
                            extracted_cache[eid] = d
                        if d.get("conversation_id"):
                            for t_id in needs_body_tids:
                                cids_by_tid[t_id].add(d["conversation_id"])
                        for t_id in needs_body_tids:
                            if eid not in seen_eids_by_tid[t_id]:
                                seen_eids_by_tid[t_id].add(eid)
                                new_msgs_by_tid[t_id].append(d)

                    if cached_tids:
                        st = extract_table_read_flag(row)
                        if st is None and item is not None:
                            is_r = 1 if getattr(item, "UnRead", False) is False else 0
                            try:
                                fl_s = 2 if int(getattr(item, "FlagStatus", 0) or 0) == 2 else 0
                            except Exception:
                                fl_s = 0
                            st = (is_r, fl_s)
                        if st is not None:
                            for t_id in cached_tids:
                                c_info = tid_existing[t_id][eid]
                                if st[0] != c_info.get("is_read") or st[1] != c_info.get("flag_status"):
                                    flag_updates_by_tid[t_id].append((eid, st[0], st[1]))
                                    c_info["is_read"], c_info["flag_status"] = st
                        for t_id in cached_tids:
                            seen_eids_by_tid[t_id].add(eid)
                except Exception:
                    continue
        except Exception as e:
            logger.warning(f"Error scanning Outlook folder '{fpath}': {e}")

    results: Dict[str, int] = {}
    all_thread_msgs: List[Dict[str, Any]] = []
    for tid in tid_map:
        if cids_by_tid[tid]:
            db.add_conversation_ids(tid, list(cids_by_tid[tid]))
        new_count = db.upsert_messages(tid, new_msgs_by_tid[tid], sync_unread=True) if new_msgs_by_tid[tid] else 0
        if flag_updates_by_tid[tid]:
            db.update_message_flags_batch(tid, flag_updates_by_tid[tid])
        if new_count > 0:
            results[tid] = new_count
        all_thread_msgs.extend(db.get_thread_messages(tid))
    _auto_cache_db_drafts(db, all_thread_msgs, cancel_check=cancel_check)
    return results


def refresh_meeting_threads(
    db: TrackedThreadsDatabase,
    meeting_id: str,
    progress_cb: Optional[Callable[[int, int, str, List[str]], None]] = None,
    cancel_check: Optional[Callable[[], bool]] = None,
) -> Dict[str, int]:
    """Refreshes all tracked threads for a meeting using a single-pass MAPI folder scan."""
    threads = db.get_threads_for_meeting(meeting_id)
    return refresh_threads_batch(db, threads, progress_cb=progress_cb, cancel_check=cancel_check)


def sync_scoped_inbox_items(
    tracked_db: Optional[TrackedThreadsDatabase],
    general_db: Optional[Any],
    tracked_thread_ids: Optional[List[str]] = None,
    unpinned_email_ids: Optional[List[str]] = None,
    progress_cb: Optional[Callable[[str], None]] = None,
    cancel_check: Optional[Callable[[], bool]] = None,
) -> tuple:
    """
    Fast filter-scoped sync: refreshes ONLY the filtered tracked threads and unpinned conversations
    directly via Outlook EntryID + GetConversation() without scanning full WG folders.
    Returns (tracked_changes_count, general_changes_count).
    """
    ns = OutlookClient.get_namespace()
    if not ns:
        return 0, 0

    tracked_changes = 0
    general_changes = 0
    t_ids = [str(x).strip() for x in (tracked_thread_ids or []) if str(x or "").strip()]
    u_ids = [str(x).strip() for x in (unpinned_email_ids or []) if str(x or "").strip() and not str(x).startswith("MSG-")]

    # 1. Scoped Tracked Threads via cached EntryIDs + GetConversation()
    if tracked_db is not None and t_ids and not (cancel_check and cancel_check()):
        fallback_threads: List[Dict[str, Any]] = []
        extracted_cache: Dict[str, Dict[str, Any]] = {}
        for idx, tid in enumerate(t_ids):
            if cancel_check and cancel_check():
                break
            thread = tracked_db.get_thread(tid)
            if not thread:
                continue
            if progress_cb:
                progress_cb(f"Syncing thread ({idx + 1}/{len(t_ids)}): {str(thread.get('title') or tid)[:28]}")
            existing = tracked_db.get_existing_message_states(tid)
            outlook_eids = [eid for eid in existing if not eid.startswith("MSG-")]
            if not outlook_eids:
                fallback_threads.append(thread)
                continue

            flag_updates: List[tuple] = []
            new_msgs: List[Dict[str, Any]] = []
            seen_ids = set()
            seed_item = None
            for cand_eid in outlook_eids[-2:][::-1]:
                try:
                    seed_item = ns.GetItemFromID(cand_eid)
                    if seed_item is not None:
                        break
                except Exception:
                    continue

            if seed_item is not None:
                for d in _collect_conversation_items(
                    seed_item,
                    existing_states=existing,
                    flag_updates_out=flag_updates,
                    extracted_cache=extracted_cache,
                ):
                    seen_ids.add(d["id"])
                    if d["id"] not in existing or not existing[d["id"]].get("has_body"):
                        new_msgs.append(d)

            # Lightly check any cached EntryIDs not present in seed_item's conversation
            for eid in outlook_eids:
                if eid not in seen_ids:
                    st = extract_item_read_flag(ns, eid)
                    c_info = existing[eid]
                    if st is not None and (st[0] != c_info.get("is_read") or st[1] != c_info.get("flag_status")):
                        flag_updates.append((eid, st[0], st[1]))

            if new_msgs:
                cids = [m["conversation_id"] for m in new_msgs if m.get("conversation_id")]
                if cids:
                    tracked_db.add_conversation_ids(tid, cids)
                tracked_changes += tracked_db.upsert_messages(tid, new_msgs, sync_unread=True)
            if flag_updates:
                tracked_changes += tracked_db.update_message_flags_batch(tid, flag_updates)

        if fallback_threads and not (cancel_check and cancel_check()):
            upd = refresh_threads_batch(tracked_db, fallback_threads, cancel_check=cancel_check)
            tracked_changes += sum((upd or {}).values())

        if not (cancel_check and cancel_check()):
            scoped_msgs: List[Dict[str, Any]] = []
            for tid in t_ids:
                scoped_msgs.extend(tracked_db.get_thread_messages(tid))
            _auto_cache_db_drafts(tracked_db, scoped_msgs, progress_cb=progress_cb, cancel_check=cancel_check)

    # 2. Scoped Unpinned Conversations via direct EntryID + GetConversation()
    if general_db is not None and u_ids and not (cancel_check and cancel_check()):
        from modules.emails.core.general_email_sync import GeneralEmailSyncThread

        if progress_cb:
            progress_cb(f"Syncing {len(u_ids)} filtered conversation(s)...")
        cached_gen = general_db.get_cached_email_states()
        syncer = GeneralEmailSyncThread([], general_db.db_path)

        tdocs_by_eid: Dict[str, List[str]] = {}
        try:
            ph = ",".join(["?"] * len(u_ids))
            with general_db._get_connection() as conn:
                for eid_r, td_r in conn.execute(
                    f"SELECT email_id, tdoc_id FROM general_email_tdoc_matches WHERE email_id IN ({ph})",
                    u_ids,
                ).fetchall():
                    tdocs_by_eid.setdefault(eid_r, []).append(td_r)
        except Exception:
            pass

        flag_updates_gen: List[tuple] = []
        new_emails: List[dict] = []
        new_matches: List[dict] = []
        seen_gen = set()

        for seed_eid in u_ids:
            if cancel_check and cancel_check():
                break
            parent_tdocs = tdocs_by_eid.get(seed_eid, [])
            try:
                seed_item = ns.GetItemFromID(seed_eid)
            except Exception:
                seed_item = None
            if seed_item is None:
                continue

            if seed_eid not in seen_gen and seed_eid in cached_gen:
                seen_gen.add(seed_eid)
                is_r = 1 if getattr(seed_item, "UnRead", False) is False else 0
                try:
                    fl_s = 2 if int(getattr(seed_item, "FlagStatus", 0) or 0) == 2 else 0
                except Exception:
                    fl_s = 0
                c_info = cached_gen[seed_eid]
                if is_r != c_info.get("is_read") or fl_s != c_info.get("flag_status"):
                    flag_updates_gen.append((seed_eid, is_r, fl_s))
                    c_info["is_read"], c_info["flag_status"] = is_r, fl_s

            try:
                conv = seed_item.GetConversation()
                tbl = conv.GetTable() if conv else None
            except Exception:
                tbl = None
            if tbl is None:
                continue

            add_mapi_table_columns(tbl, ("EntryID", "MessageClass", "UnRead", "FlagStatus", PR_FLAG_STATUS_PROP))
            cnt = 0
            while not tbl.EndOfTable and cnt < 250:
                cnt += 1
                try:
                    row = tbl.GetNextRow()
                    eid = str(safe_mapi_row_val(row, "EntryID", default="") or "")
                    if not eid or eid in seen_gen:
                        continue
                    seen_gen.add(eid)
                    msg_cls = str(safe_mapi_row_val(row, "MessageClass", default="") or "")
                    if msg_cls and not msg_cls.startswith("IPM.Note"):
                        continue

                    c_info = cached_gen.get(eid)
                    if c_info and c_info.get("has_body"):
                        st = extract_table_read_flag(row)
                        if st is None:
                            st = extract_item_read_flag(ns, eid)
                        if st is not None and (st[0] != c_info.get("is_read") or st[1] != c_info.get("flag_status")):
                            flag_updates_gen.append((eid, st[0], st[1]))
                            c_info["is_read"], c_info["flag_status"] = st
                        continue

                    item = ns.GetItemFromID(eid)
                    if getattr(item, "Class", 43) != 43:
                        continue
                    rec, recs_m = syncer._build_email_record(
                        item,
                        folder_tag="Inbox",
                        entry_id=eid,
                        require_tdoc=bool(not parent_tdocs),
                        fallback_tdocs=parent_tdocs,
                    )
                    if rec is not None:
                        cached_gen[eid] = {
                            "is_read": rec["is_read"],
                            "flag_status": rec["flag_status"],
                            "has_body": bool(rec["body_text"]),
                        }
                        new_emails.append(rec)
                        new_matches.extend(recs_m)
                except Exception:
                    continue

        if new_emails:
            general_db.save_emails_batch(new_emails, new_matches, sync_unread=True)
            general_changes += len(new_emails)
        if flag_updates_gen:
            general_changes += general_db.update_email_flags_batch(flag_updates_gen)

        if not (cancel_check and cancel_check()):
            scoped_gen_msgs: List[Dict[str, Any]] = list(new_emails)
            try:
                all_eids = list({*u_ids, *(e["id"] for e in new_emails)})
                if all_eids:
                    ph = ",".join(["?"] * len(all_eids))
                    with general_db._get_connection() as conn:
                        for b_txt, c_txt in conn.execute(
                            f"SELECT body_text, clean_body FROM general_emails WHERE id IN ({ph})",
                            all_eids,
                        ).fetchall():
                            scoped_gen_msgs.append({"body_text": b_txt or "", "clean_body": c_txt or ""})
            except Exception:
                pass
            _auto_cache_db_drafts(general_db, scoped_gen_msgs, progress_cb=progress_cb, cancel_check=cancel_check)

    return tracked_changes, general_changes


_ORIG_REFRESH_THREAD = refresh_thread


def call_refresh_thread(*args: Any, **kwargs: Any) -> Any:
    """Calls refresh_thread, honoring monkeypatches on the legacy tracked_threads_service shim."""
    for mod_name in (
        "modules.emails.core.tracked_threads_service",
        "src.modules.emails.core.tracked_threads_service",
    ):
        shim = sys.modules.get(mod_name)
        if shim is not None:
            shim_fn = getattr(shim, "refresh_thread", None)
            if shim_fn is not None and shim_fn is not _ORIG_REFRESH_THREAD:
                return shim_fn(*args, **kwargs)
    return refresh_thread(*args, **kwargs)
