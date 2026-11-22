# --- File: src/modules/emails/ui/meeting_inbox/tdoc_grouping.py ---
"""
Pure-Python grouping, signal rollup, and Priority Radar evaluator for `MeetingEmailInboxDialog`.
Supports 3 view modes (`MESSAGES`, `THREADS`, `TDOCS`) without bloating the dialog mixins.
"""

from __future__ import annotations

from typing import Any, Dict, List, Optional, Sequence, Set

from modules.emails.core.emeeting_signals import (
    analyze_message_signals,
    compute_group_signals,
    extract_base_tdoc_id,
    format_rev_tag,
    parse_rev_number,
)
from modules.emails.core.tracked_threads_service import normalize_subject


def enrich_message_item_signals(
    it: Dict[str, Any],
    tdoc_meta_map: Optional[Dict[str, Dict[str, Any]]] = None,
) -> Dict[str, Any]:
    """Attaches `_signals`, `is_objected`, and `objected_at_rev` to an individual email item."""
    meta_map = tdoc_meta_map or {}
    linked_tdocs = it.get("linked_tdocs") or []
    primary_tdoc = extract_base_tdoc_id(linked_tdocs[0]) if linked_tdocs else ""
    td_meta = meta_map.get(primary_tdoc) or {}

    if (not it.get("user_status") or it.get("user_status") == "⚪ Neutral") and td_meta.get("my_status"):
        it["user_status"] = str(td_meta.get("my_status") or "⚪ Neutral").strip() or "⚪ Neutral"
    if not it.get("user_notes") and td_meta.get("my_notes"):
        it["user_notes"] = str(td_meta.get("my_notes") or "").strip()

    is_obj = bool(it.get("is_objected") or td_meta.get("is_objected"))
    obj_rev = str(it.get("objected_at_rev") or td_meta.get("objected_at_rev") or "").strip()
    it["is_objected"] = is_obj
    it["objected_at_rev"] = obj_rev

    msg_sig = analyze_message_signals(it, primary_tdoc=primary_tdoc)
    obj_rev_num = parse_rev_number(obj_rev)
    unparked = bool(is_obj and msg_sig["max_rev_num"] > obj_rev_num)
    msg_sig["is_objected"] = is_obj
    msg_sig["objected_at_rev"] = format_rev_tag(obj_rev_num) if obj_rev_num > 0 else obj_rev
    msg_sig["unparked_by_new_rev"] = unparked
    msg_sig["effectively_parked"] = bool(is_obj and not unparked)
    it["_signals"] = msg_sig
    return it


def build_grouped_thread_items(
    all_items: Sequence[Dict[str, Any]],
    tdoc_meta_map: Optional[Dict[str, Dict[str, Any]]] = None,
) -> List[Dict[str, Any]]:
    """Groups messages by Tracked Thread ID or normalized subject (`THREADS` view mode)."""
    meta_map = tdoc_meta_map or {}
    buckets: Dict[str, List[dict]] = {}
    for it in all_items:
        if it.get("in_tracked_db") and it.get("thread_id"):
            key = f"THR::{it['thread_id']}"
        else:
            norm = normalize_subject(it.get("subject") or "").lower()
            key = f"UNPIN::{norm or it['id']}"
        buckets.setdefault(key, []).append(it)

    grouped: List[dict] = []
    for gkey, msgs in buckets.items():
        msgs_asc = sorted(msgs, key=lambda m: str(m.get("date_str") or ""))
        latest = msgs_asc[-1]
        in_tracked = any(m.get("in_tracked_db") for m in msgs_asc)
        tid = next((m.get("thread_id") for m in msgs_asc if m.get("thread_id")), "")
        th_title = next((m.get("thread_title") for m in msgs_asc if m.get("thread_title")), "")
        linked_ais = list(dict.fromkeys(ai for m in msgs_asc for ai in (m.get("linked_ais") or []) if ai))
        linked_subs = list(dict.fromkeys(s for m in msgs_asc for s in (m.get("linked_sub_items") or []) if s))
        linked_wis = list(dict.fromkeys(w for m in msgs_asc for w in (m.get("linked_wis") or []) if w))
        linked_tdocs = list(dict.fromkeys(td for m in msgs_asc for td in (m.get("linked_tdocs") or []) if td))
        primary_tdoc = extract_base_tdoc_id(linked_tdocs[0]) if linked_tdocs else ""
        td_meta = meta_map.get(primary_tdoc) or {}

        companies = list(dict.fromkeys(m.get("company") for m in msgs_asc if m.get("company")))
        unread_cnt = sum(1 for m in msgs_asc if not m["is_read"])
        flagged_cnt = sum(1 for m in msgs_asc if m["is_flagged"])
        user_status = next(
            (
                str(m.get("user_status") or "").strip()
                for m in msgs_asc
                if str(m.get("user_status") or "").strip() and m.get("user_status") != "⚪ Neutral"
            ),
            str(td_meta.get("my_status") or (msgs_asc[0].get("user_status") if msgs_asc else "") or "⚪ Neutral").strip() or "⚪ Neutral",
        )
        user_notes = next(
            (str(m.get("user_notes") or "").strip() for m in msgs_asc if str(m.get("user_notes") or "").strip()),
            str(td_meta.get("my_notes") or "").strip(),
        )
        is_obj = any(bool(m.get("is_objected")) for m in msgs_asc) or bool(td_meta.get("is_objected"))
        obj_rev = next(
            (str(m.get("objected_at_rev") or "").strip() for m in msgs_asc if str(m.get("objected_at_rev") or "").strip()),
            str(td_meta.get("objected_at_rev") or "").strip(),
        )
        sig = compute_group_signals(
            msgs_asc,
            primary_tdoc=primary_tdoc,
            is_objected=is_obj,
            objected_at_rev=obj_rev,
        )

        grouped.append({
            "id": gkey,
            "_is_group": True,
            "_group_kind": "THREAD",
            "_group_messages": msgs_asc,
            "_unread_count": unread_cnt,
            "_total_count": len(msgs_asc),
            "_flagged_count": flagged_cnt,
            "_companies": companies,
            "_signals": sig,
            "source_kind": "thread" if in_tracked else "tdoc",
            "thread_id": tid,
            "thread_title": th_title or normalize_subject(latest.get("subject") or ""),
            "subject": th_title or latest.get("subject") or "(No Subject)",
            "sender_name": latest.get("sender_name") or "",
            "sender_email": latest.get("sender_email") or "",
            "company": ", ".join(companies) if companies else (latest.get("company") or ""),
            "date_str": latest.get("date_str") or "",
            "body_text": latest.get("body_text") or "",
            "is_read": unread_cnt == 0,
            "is_flagged": flagged_cnt > 0,
            "attachments": latest.get("attachments") or [],
            "linked_ais": linked_ais,
            "linked_sub_items": linked_subs,
            "linked_wis": linked_wis,
            "linked_tdocs": linked_tdocs,
            "in_tracked_db": in_tracked,
            "in_general_db": any(m.get("in_general_db") for m in msgs_asc),
            "conversation_id": next((m.get("conversation_id") for m in msgs_asc if m.get("conversation_id")), ""),
            "user_status": user_status or "⚪ Neutral",
            "user_notes": user_notes,
            "is_objected": is_obj,
            "objected_at_rev": sig["objected_at_rev"],
        })
    return sorted(grouped, key=lambda x: x.get("date_str", ""), reverse=True)


def build_grouped_tdoc_items(
    all_items: Sequence[Dict[str, Any]],
    tdoc_meta_map: Optional[Dict[str, Dict[str, Any]]] = None,
    tdoc_ai_lookup: Optional[Dict[str, str]] = None,
) -> List[Dict[str, Any]]:
    """
    Groups messages by Base TDoc ID (`TDOCS` view mode for e-Meetings & TDoc-centric triage).
    Messages without any linked TDoc are grouped by thread/subject under a `📌 [Thread]` row
    so zero messages are lost.
    """
    meta_map = tdoc_meta_map or {}
    ai_lookup = tdoc_ai_lookup or {}
    buckets: Dict[str, List[dict]] = {}

    for it in all_items:
        raw_tdocs = [extract_base_tdoc_id(t) for t in (it.get("linked_tdocs") or []) if extract_base_tdoc_id(t)]
        base_tdocs = list(dict.fromkeys(raw_tdocs))
        if base_tdocs:
            for btd in base_tdocs:
                buckets.setdefault(f"TDOC::{btd}", []).append(it)
        else:
            if it.get("in_tracked_db") and it.get("thread_id"):
                buckets.setdefault(f"NOTDOC_THR::{it['thread_id']}", []).append(it)
            else:
                norm = normalize_subject(it.get("subject") or "").lower()
                buckets.setdefault(f"NOTDOC_UNPIN::{norm or it['id']}", []).append(it)

    grouped: List[dict] = []
    for gkey, msgs in buckets.items():
        # Deduplicate messages by id preserving chronological ascending order
        by_mid: Dict[str, dict] = {}
        for m in msgs:
            mid = str(m.get("id") or "")
            if mid:
                by_mid[mid] = m
        msgs_asc = sorted(by_mid.values(), key=lambda m: str(m.get("date_str") or ""))
        if not msgs_asc:
            continue
        latest = msgs_asc[-1]
        is_tdoc_bucket = gkey.startswith("TDOC::")
        base_tdoc = gkey.split("TDOC::", 1)[1] if is_tdoc_bucket else ""
        td_meta = meta_map.get(base_tdoc) or {}

        td_title = str(td_meta.get("title") or "").strip()
        td_source = str(td_meta.get("source") or "").strip()
        td_type = str(td_meta.get("type") or "").strip()
        td_official_status = str(td_meta.get("tdoc_status") or "").strip()

        in_tracked = any(m.get("in_tracked_db") for m in msgs_asc)
        tid = next((m.get("thread_id") for m in msgs_asc if m.get("thread_id")), "")
        th_title = next((m.get("thread_title") for m in msgs_asc if m.get("thread_title")), "")

        linked_ais: List[str] = []
        if base_tdoc:
            meta_ai = str(td_meta.get("ai") or ai_lookup.get(base_tdoc) or "").strip()
            if meta_ai:
                linked_ais.append(meta_ai)
        for m in msgs_asc:
            for ai in (m.get("linked_ais") or []):
                if ai and ai not in linked_ais:
                    linked_ais.append(ai)

        linked_subs: List[str] = []
        if td_meta.get("sub_item"):
            linked_subs.append(str(td_meta["sub_item"]).strip())
        for m in msgs_asc:
            for s in (m.get("linked_sub_items") or []):
                if s and s not in linked_subs:
                    linked_subs.append(s)

        linked_wis: List[str] = list(dict.fromkeys(w for w in (td_meta.get("wis") or []) if w))
        for m in msgs_asc:
            for w in (m.get("linked_wis") or []):
                if w and w not in linked_wis:
                    linked_wis.append(w)

        linked_tdocs = [base_tdoc] if base_tdoc else []
        for m in msgs_asc:
            for t in (m.get("linked_tdocs") or []):
                if t and t not in linked_tdocs:
                    linked_tdocs.append(t)

        comp_counts: Dict[str, int] = {}
        companies: List[str] = []
        for m in msgs_asc:
            c = str(m.get("company") or "").strip()
            if c:
                if c not in comp_counts:
                    companies.append(c)
                    comp_counts[c] = 0
                comp_counts[c] += 1
        latest_company = str(latest.get("company") or latest.get("sender_name") or "Unknown").strip()

        unread_cnt = sum(1 for m in msgs_asc if not m["is_read"])
        flagged_cnt = sum(1 for m in msgs_asc if m["is_flagged"])

        # Prefer TDoc user_status/notes from user_tdocs.db when available, fallback to thread status/notes
        td_my_status = str(td_meta.get("my_status") or "").strip()
        if td_my_status and td_my_status != "⚪ Neutral":
            user_status = td_my_status
        else:
            user_status = next(
                (
                    str(m.get("user_status") or "").strip()
                    for m in msgs_asc
                    if str(m.get("user_status") or "").strip() and m.get("user_status") != "⚪ Neutral"
                ),
                td_my_status or "⚪ Neutral",
            )
        user_notes = str(td_meta.get("my_notes") or "").strip() or next(
            (str(m.get("user_notes") or "").strip() for m in msgs_asc if str(m.get("user_notes") or "").strip()),
            "",
        )
        is_obj = bool(td_meta.get("is_objected")) or any(bool(m.get("is_objected")) for m in msgs_asc)
        obj_rev = str(td_meta.get("objected_at_rev") or "").strip() or next(
            (str(m.get("objected_at_rev") or "").strip() for m in msgs_asc if str(m.get("objected_at_rev") or "").strip()),
            "",
        )
        sig = compute_group_signals(
            msgs_asc,
            primary_tdoc=base_tdoc,
            is_objected=is_obj,
            objected_at_rev=obj_rev,
        )

        display_subject = td_title or th_title or str(latest.get("subject") or "(No Subject)")
        grouped.append({
            "id": gkey,
            "_is_group": True,
            "_group_kind": "TDOC" if is_tdoc_bucket else "THREAD",
            "_primary_tdoc": base_tdoc,
            "_tdoc_title": td_title,
            "_tdoc_source": td_source,
            "_tdoc_type": td_type,
            "_tdoc_status": td_official_status,
            "_group_messages": msgs_asc,
            "_unread_count": unread_cnt,
            "_total_count": len(msgs_asc),
            "_flagged_count": flagged_cnt,
            "_companies": companies,
            "_company_counts": comp_counts,
            "_latest_company": latest_company,
            "_signals": sig,
            "source_kind": "tdoc" if is_tdoc_bucket else ("thread" if in_tracked else "tdoc"),
            "thread_id": tid,
            "thread_title": f"{base_tdoc} — {td_title}" if (base_tdoc and td_title) else (base_tdoc or th_title or display_subject),
            "subject": display_subject,
            "sender_name": latest.get("sender_name") or "",
            "sender_email": latest.get("sender_email") or "",
            "company": ", ".join(companies) if companies else latest_company,
            "date_str": latest.get("date_str") or "",
            "body_text": latest.get("body_text") or "",
            "is_read": unread_cnt == 0,
            "is_flagged": flagged_cnt > 0,
            "attachments": latest.get("attachments") or [],
            "linked_ais": linked_ais,
            "linked_sub_items": linked_subs,
            "linked_wis": linked_wis,
            "linked_tdocs": linked_tdocs,
            "in_tracked_db": in_tracked,
            "in_general_db": any(m.get("in_general_db") for m in msgs_asc),
            "conversation_id": next((m.get("conversation_id") for m in msgs_asc if m.get("conversation_id")), ""),
            "user_status": user_status or "⚪ Neutral",
            "user_notes": user_notes,
            "is_objected": is_obj,
            "objected_at_rev": sig["objected_at_rev"],
        })
    return sorted(grouped, key=lambda x: x.get("date_str", ""), reverse=True)


def is_item_effectively_parked(it: Dict[str, Any]) -> bool:
    """Returns True if the item is marked `🛑 Objected` and has NOT been unparked by a newer revision."""
    sig = it.get("_signals") or {}
    if "effectively_parked" in sig:
        return bool(sig["effectively_parked"])
    return bool(it.get("is_objected"))


def is_priority_item(it: Dict[str, Any], followed_ais: Set[str]) -> bool:
    """
    Evaluates whether an item belongs in the `🔥 Priority` triage radar:
    - Excluded if effectively parked (`🛑 Objected` without a newer revision).
    - Included if unparked by a newer revision (`🛑→🆕rXX`), OR if it has unread activity AND:
      * non-Neutral `My Status` (`🔵 My TDoc`, `🟢 Support`, `🔴 Object`, `🟡 Monitor`), OR
      * is Flagged (`🚩`), OR
      * belongs to one of the user's `👀 Followed AIs`.
    """
    if is_item_effectively_parked(it):
        return False
    sig = it.get("_signals") or {}
    if sig.get("unparked_by_new_rev") and not it.get("is_read", True):
        return True
    if it.get("is_read", True):
        return False
    status = str(it.get("user_status") or "⚪ Neutral").strip() or "⚪ Neutral"
    if status != "⚪ Neutral":
        return True
    if it.get("is_flagged"):
        return True
    item_ais = {str(a).strip() for a in (it.get("linked_ais") or []) if str(a).strip()}
    if followed_ais and (item_ais & followed_ais):
        return True
    return False
