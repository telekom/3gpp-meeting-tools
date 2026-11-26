"""Per-AI leaf bucket builder combining TrackedThreadsDatabase and GeneralEmailDatabase."""

from __future__ import annotations

from typing import Any, Dict, Optional, Set

from modules.emails.core.ai_summary import sort_messages_chronologically
from modules.emails.core.emeeting_signals import extract_base_tdoc_id


def build_meeting_ai_buckets(
    *,
    tracked_db: Optional[Any] = None,
    general_email_db: Optional[Any] = None,
    meeting_id: str = "",
    tdoc_ai_lookup: Optional[Dict[str, str]] = None,
    tdoc_meta_map: Optional[Dict[str, Dict[str, Any]]] = None,
) -> Dict[str, Dict[str, Dict[str, Any]]]:
    """
    Build `{ai_code: {bucket_key: {'scope_type', 'scope_key', 'title', 'messages', 'tdoc_meta'}}}`
    deduplicated across `tracked_db` and `general_email_db`.
    """
    ai_lookup = {str(k).upper(): str(v).strip() for k, v in (tdoc_ai_lookup or {}).items() if str(v).strip()}
    meta_map = tdoc_meta_map or {}
    by_id: Dict[str, Dict[str, Any]] = {}

    if tracked_db is not None:
        for th in tracked_db.get_threads_for_meeting(meeting_id):
            tid = str(th.get("thread_id") or "").strip()
            th_ais = [str(a).strip() for a in (th.get("linked_ais") or []) if str(a).strip()]
            th_tdocs = list(dict.fromkeys(extract_base_tdoc_id(t) for t in (th.get("linked_tdocs") or []) if extract_base_tdoc_id(t)))
            for m in tracked_db.get_thread_messages(tid):
                eid = str(m.get("id") or "").strip()
                if not eid:
                    continue
                by_id[eid] = {
                    "id": eid,
                    "thread_id": tid,
                    "thread_title": str(th.get("title") or "").strip(),
                    "subject": str(m.get("subject") or th.get("title") or "").strip(),
                    "sender_name": str(m.get("sender_name") or "").strip(),
                    "sender_email": str(m.get("sender_email") or "").strip(),
                    "company": str(m.get("company") or "").strip(),
                    "date_sent": str(m.get("date_sent") or "").strip(),
                    "body_text": str(m.get("body_text") or ""),
                    "clean_body": str(m.get("clean_body") or ""),
                    "attachments": list(m.get("attachments") or []),
                    "linked_ais": list(th_ais),
                    "linked_tdocs": list(th_tdocs),
                    "is_objected": bool(th.get("is_objected")),
                    "objected_at_rev": str(th.get("objected_at_rev") or "").strip(),
                }

    if general_email_db is not None and hasattr(general_email_db, "get_all_emails_with_tdocs"):
        for row in general_email_db.get_all_emails_with_tdocs(show_ignored=False):
            eid = str(row.get("id") or "").strip()
            if not eid:
                continue
            raw_tds = row.get("linked_tdocs")
            if not raw_tds and row.get("matched_tdocs"):
                raw_tds = [x.strip() for x in str(row.get("matched_tdocs") or "").split(",") if x.strip()]
            linked_tdocs = list(
                dict.fromkeys(extract_base_tdoc_id(t) for t in (raw_tds or []) if extract_base_tdoc_id(t))
            )
            inferred_ais = list(dict.fromkeys(ai_lookup.get(t, "") for t in linked_tdocs if ai_lookup.get(t, "")))
            if eid in by_id:
                existing = by_id[eid]
                existing["linked_tdocs"] = list(dict.fromkeys(existing["linked_tdocs"] + linked_tdocs))
                existing["linked_ais"] = list(dict.fromkeys(existing["linked_ais"] + inferred_ais))
                if not existing.get("clean_body") and row.get("clean_body"):
                    existing["clean_body"] = str(row.get("clean_body") or "")
                continue
            by_id[eid] = {
                "id": eid,
                "thread_id": "",
                "thread_title": "",
                "subject": str(row.get("subject") or "").strip(),
                "sender_name": str(row.get("sender_name") or "").strip(),
                "sender_email": str(row.get("sender_email") or "").strip(),
                "company": str(row.get("company") or "").strip(),
                "date_sent": str(row.get("date_received") or "").strip(),
                "body_text": str(row.get("body_text") or ""),
                "clean_body": str(row.get("clean_body") or ""),
                "attachments": list(row.get("attachments") or []),
                "linked_ais": inferred_ais,
                "linked_tdocs": linked_tdocs,
            }

    ai_buckets: Dict[str, Dict[str, Dict[str, Any]]] = {}

    def _add_to_bucket(ai_code: str, bkey: str, *, scope_type: str, scope_key: str, title: str, msg: dict, tmeta: Optional[dict] = None) -> None:
        per_ai = ai_buckets.setdefault(ai_code, {})
        bucket = per_ai.get(bkey)
        if bucket is None:
            bucket = {
                "scope_type": scope_type,
                "scope_key": scope_key,
                "title": title,
                "messages": [],
                "_seen_ids": set(),
                "tdoc_meta": tmeta,
            }
            per_ai[bkey] = bucket
        mid = str(msg.get("id") or "")
        if mid and mid not in bucket["_seen_ids"]:
            bucket["_seen_ids"].add(mid)
            bucket["messages"].append(msg)

    for msg in by_id.values():
        tdocs = [extract_base_tdoc_id(t) for t in (msg.get("linked_tdocs") or []) if extract_base_tdoc_id(t)]
        msg_ais: Set[str] = {str(a).strip() for a in (msg.get("linked_ais") or []) if str(a).strip()}

        if tdocs:
            for btd in tdocs:
                tmeta = meta_map.get(btd) or {}
                ai_for_tdoc = str(tmeta.get("ai") or ai_lookup.get(btd) or "").strip()
                target_ais = {ai_for_tdoc} if ai_for_tdoc else msg_ais
                title = str(tmeta.get("title") or msg.get("subject") or btd).strip()
                for ai_c in target_ais:
                    if not ai_c:
                        continue
                    _add_to_bucket(
                        ai_c,
                        f"TDOC::{btd}",
                        scope_type="tdoc",
                        scope_key=btd,
                        title=title,
                        msg=msg,
                        tmeta=tmeta or None,
                    )
        elif msg_ais:
            tid = str(msg.get("thread_id") or "").strip()
            if tid:
                bkey = f"THR::{tid}"
                skey = tid
                title = str(msg.get("thread_title") or msg.get("subject") or tid).strip()
            else:
                subj = str(msg.get("subject") or msg.get("id") or "Thread").strip()
                bkey = f"UNPIN::{subj.lower()}"
                skey = bkey
                title = subj
            for ai_c in msg_ais:
                _add_to_bucket(
                    ai_c,
                    bkey,
                    scope_type="thread",
                    scope_key=skey,
                    title=title,
                    msg=msg,
                )

    for ai_code, per_ai in list(ai_buckets.items()):
        for bucket in per_ai.values():
            bucket.pop("_seen_ids", None)
            bucket["messages"] = sort_messages_chronologically(bucket["messages"])
        ai_buckets[ai_code] = _deduplicate_identical_message_buckets(per_ai)

    return ai_buckets


def _deduplicate_identical_message_buckets(
    per_ai: Dict[str, Dict[str, Any]],
) -> Dict[str, Dict[str, Any]]:
    """
    Collapse duplicate TDoc buckets within an AI that share the exact same email set
    (e.g. an omnibus Draft TR review thread) into one discussion carrying `alias_tdocs`.
    """
    grouped_keys: Dict[tuple, list] = {}
    passthrough: Dict[str, Dict[str, Any]] = {}
    for bkey, bucket in per_ai.items():
        msg_sig = tuple(sorted(str(m.get("id") or "") for m in bucket.get("messages") or [] if m.get("id")))
        if not msg_sig:
            passthrough[bkey] = bucket
            continue
        grouped_keys.setdefault(msg_sig, []).append((bkey, bucket))

    merged: Dict[str, Dict[str, Any]] = {}
    for _sig, items in grouped_keys.items():
        if len(items) == 1:
            bkey, bucket = items[0]
            merged[bkey] = bucket
            continue

        tdoc_keys = [
            str(b.get("scope_key") or "")
            for _, b in items
            if str(b.get("scope_type") or "") == "tdoc" and str(b.get("scope_key") or "")
        ]
        thread_item = next(((k, b) for k, b in items if str(b.get("scope_type") or "") == "thread"), None)
        if thread_item is not None:
            chosen_key, chosen = thread_item[0], dict(thread_item[1])
            chosen["alias_tdocs"] = list(dict.fromkeys(tdoc_keys))
            merged[chosen_key] = chosen
            continue

        sample_msgs = items[0][1].get("messages") or []
        common_tids = {str(m.get("thread_id") or "").strip() for m in sample_msgs if str(m.get("thread_id") or "").strip()}
        if len(common_tids) == 1:
            tid = next(iter(common_tids))
            th_title = next(
                (str(m.get("thread_title") or "").strip() for m in sample_msgs if str(m.get("thread_title") or "").strip()),
                items[0][1].get("title") or tid,
            )
            merged[f"THR::{tid}"] = {
                "scope_type": "thread",
                "scope_key": tid,
                "title": th_title,
                "messages": list(sample_msgs),
                "tdoc_meta": items[0][1].get("tdoc_meta"),
                "alias_tdocs": list(dict.fromkeys(tdoc_keys)),
            }
        else:
            primary_key, primary_bucket = items[0]
            chosen = dict(primary_bucket)
            chosen["alias_tdocs"] = list(dict.fromkeys(tdoc_keys))
            merged[primary_key] = chosen

    merged.update(passthrough)
    return merged
