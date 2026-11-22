"""Fast filter-scoped Inbox synchronization via Outlook EntryID + GetConversation()."""

from __future__ import annotations

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
from .collection_and_pinning import _collect_conversation_items
from .documents import _auto_cache_db_drafts


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
            from .sync import refresh_threads_batch

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
