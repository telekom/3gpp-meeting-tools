# --- File: src/modules/emails/core/tracked_threads_service.py ---
import hashlib
import json
import logging
import re
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional

from modules.emails.core.email_config import load_wg_email_config
from modules.emails.core.outlook_client import (
    PR_FLAG_STATUS_PROP,
    OutlookClient,
    add_mapi_table_columns,
    extract_item_read_flag,
    extract_table_read_flag,
    safe_mapi_row_val,
)
from modules.emails.core.thread_markdown_exporter import (
    extract_drafts_document_urls,
    extract_thread_document_urls,
    inject_html_tables_into_body,
    is_drafts_document_url,
    strip_trailing_quotes,
    unwrap_outlook_safelinks,
)
from modules.emails.core.tracked_threads_db import TrackedThreadsDatabase

logger = logging.getLogger(__name__)

_PREFIX_RE = re.compile(r"^(?:\s*(?:re|fwd|fw|aw|wg|sv|antw|回复|答复|转发)\s*[:：]\s*)+", re.IGNORECASE)
_TAG_RE = re.compile(r"^\s*\[[^\]]{2,80}\]\s*")
_TDOC_RE = re.compile(r"\b(?:S[1-6]|R[1-5]|C[1-6]|SP|RP|CP)-\d{6,7}(?:(?:r|rev)\d{1,2}[a-z]?)?\b", re.IGNORECASE)
_KI_AI_RE = re.compile(r"\b(?:KI\s*#?\s*\d+(?:[._]\d+)?|AI\s*\d+(?:\.\d+)+)\b", re.IGNORECASE)
_WI_RE = re.compile(r"\b(?:FS|TEI\d+|e?[A-Z0-9]{2,})_[A-Z0-9_]{2,}\b")


def _resolve_ais_against_agenda(
    raw_ais: List[str],
    raw_wis: List[str],
    agenda_map: Optional[Dict[str, Any]],
    valid_ais: Optional[List[str]],
) -> List[str]:
    """Resolves AI/KI subject tokens against actual meeting agenda items (using WI tokens to disambiguate KIs)."""
    ai_descs: Dict[str, str] = {}
    for k, item in (agenda_map or {}).items():
        ai_key = str(k).strip()
        if ai_key:
            ai_descs[ai_key] = str(getattr(item, "description", "") or "").strip()
    for entry in (valid_ais or []):
        txt = str(entry or "").strip()
        if not txt:
            continue
        if " — " in txt:
            num, desc = txt.split(" — ", 1)
            ai_descs.setdefault(num.strip(), desc.strip())
        else:
            ai_descs.setdefault(txt, "")

    known_by_lower = {k.lower(): k for k in ai_descs}
    resolved: List[str] = []
    for tok in raw_ais:
        t_clean = tok.strip()
        # Direct AI number match (e.g. "AI20.6.1" -> "20.6.1")
        num_part = re.sub(r"^AI\s*", "", t_clean, flags=re.IGNORECASE).strip()
        if num_part.lower() in known_by_lower:
            cand = known_by_lower[num_part.lower()]
            if cand not in resolved:
                resolved.append(cand)
            continue

        # KI token match (e.g. "KI#4", "KI#1_1") against agenda item descriptions
        ki_m = re.match(r"^KI#?(\d+(?:[._]\d+)?)$", t_clean, re.IGNORECASE)
        if not ki_m:
            continue
        ki_num_re = re.escape(ki_m.group(1)).replace(r"\_", r"[._]").replace(r"\.", r"[._]")
        ki_pat = re.compile(rf"\bKI\s*#?\s*{ki_num_re}\b", re.IGNORECASE)
        matches = [ai_k for ai_k, desc in ai_descs.items() if desc and ki_pat.search(desc)]
        if len(matches) > 1 and raw_wis:
            wi_matches = [
                ai_k for ai_k in matches
                if any(w.lower() in ai_descs[ai_k].lower() for w in raw_wis)
            ]
            if wi_matches:
                matches = wi_matches
        if len(matches) == 1 and matches[0] not in resolved:
            resolved.append(matches[0])
    return sorted(resolved)


def auto_suggest_links(
    subject: str,
    body: str = "",
    *,
    agenda_map: Optional[Dict[str, Any]] = None,
    valid_ais: Optional[List[str]] = None,
    valid_wis: Optional[List[str]] = None,
    tdoc_meta_map: Optional[Dict[str, Dict[str, Any]]] = None,
    extra_tdocs: Optional[List[str]] = None,
) -> Dict[str, List[str]]:
    """
    Extracts candidate AIs/KIs, Sub-Topics, WIs, and TDocs from an email subject, body, and linked TDocs.
    When meeting metadata (`agenda_map`/`valid_ais`/`valid_wis`/`tdoc_meta_map`) is provided,
    tokens are resolved against real meeting Agenda Items, Sub-Items, and WIs.
    """
    text_subj = subject or ""
    text_all = f"{text_subj}\n{(body or '')[:2000]}"
    raw_ais = sorted({re.sub(r"\s+", "", m.upper()) for m in _KI_AI_RE.findall(text_subj)})
    raw_wis = sorted({m.strip() for m in _WI_RE.findall(text_subj)})
    tdocs = sorted(
        {m.upper() for m in _TDOC_RE.findall(text_all)}
        | {str(t).strip().upper() for t in (extra_tdocs or []) if str(t).strip()}
    )

    if agenda_map is not None or valid_ais is not None:
        ais = _resolve_ais_against_agenda(raw_ais, raw_wis, agenda_map, valid_ais)
    else:
        ais = raw_ais

    if valid_wis is not None:
        wi_by_lower = {str(w).strip().lower(): str(w).strip() for w in valid_wis if str(w).strip()}
        wis = sorted({wi_by_lower[w.lower()] for w in raw_wis if w.lower() in wi_by_lower})
    else:
        wis = raw_wis

    sub_items: List[str] = []
    if tdoc_meta_map:
        wi_by_lower = (
            {str(w).strip().lower(): str(w).strip() for w in valid_wis if str(w).strip()}
            if valid_wis is not None
            else None
        )
        for td in tdocs:
            base_td = re.sub(r"(?:R|REV)\d{1,2}[A-Z]?$", "", td, flags=re.IGNORECASE)
            meta = tdoc_meta_map.get(td) or tdoc_meta_map.get(base_td)
            if not isinstance(meta, dict):
                continue
            ai_val = str(meta.get("ai") or "").strip()
            if ai_val and ai_val not in ("N/A", "-", "None") and ai_val not in ais:
                ais.append(ai_val)
            sub_val = str(meta.get("sub_item") or "").strip()
            if sub_val and sub_val not in ("N/A", "-", "None") and sub_val not in sub_items:
                sub_items.append(sub_val)
            for w in (meta.get("wis") or []):
                wc = str(w).strip()
                if not wc or wc.lower() in ("(no wi)", "none", "nan", "-"):
                    continue
                if wi_by_lower is not None:
                    wc = wi_by_lower.get(wc.lower(), wc)
                if wc not in wis:
                    wis.append(wc)

    return {"ais": sorted(set(ais)), "sub_items": sub_items, "wis": sorted(set(wis)), "tdocs": tdocs}


def normalize_subject(subject: str) -> str:
    """Normalizes an email subject by stripping reply/forward prefixes and WG tags."""
    s = (subject or "").strip()
    changed = True
    while changed:
        changed = False
        new_s = _PREFIX_RE.sub("", s).strip()
        if new_s != s:
            s = new_s
            changed = True
        new_s = _TAG_RE.sub("", s).strip()
        if new_s != s:
            s = new_s
            changed = True
    return re.sub(r"\s+", " ", s).strip()


def _extract_folder_paths(raw_folders: Optional[List[Any]]) -> List[str]:
    """Extracts string folder paths from either string lists or emails_config dict entries."""
    paths: List[str] = []
    for entry in (raw_folders or []):
        if isinstance(entry, dict):
            fp = str(entry.get("folder_path", "") or "").strip()
        else:
            fp = str(entry or "").strip()
        if fp and fp not in paths:
            paths.append(fp)
    return paths


def _extract_company(sender_name: str, sender_email: str) -> str:
    try:
        from modules.emails.core.general_email_sync import GeneralEmailSyncThread
        comp = GeneralEmailSyncThread._resolve_company(sender_name, sender_email)
        if comp and comp != "Unknown":
            return comp
    except Exception:
        pass
    if "(" in sender_name and ")" in sender_name:
        m = re.search(r"\(([^)]+)\)", sender_name)
        if m:
            return m.group(1).strip()
    if "@" in (sender_email or ""):
        domain = sender_email.split("@", 1)[1].split(".")[0]
        if domain and domain.lower() not in ("gmail", "outlook", "hotmail", "yahoo", "icloud"):
            return domain.capitalize()
    return ""


def _strip_trailing_quotes(body_text: str) -> str:
    """Strips pure trailing quoted reply chains while preserving inline replies."""
    return strip_trailing_quotes(body_text)


def _mail_item_to_dict(item: Any) -> Dict[str, Any]:
    """Extracts standardized message fields from an Outlook COM MailItem."""
    subject = getattr(item, "Subject", "") or ""
    sender_name = getattr(item, "SenderName", "") or ""
    try:
        sender_email = getattr(item, "SenderEmailAddress", "") or ""
    except Exception:
        sender_email = ""
    raw_body = getattr(item, "Body", "") or ""
    try:
        html_body = getattr(item, "HTMLBody", "") or ""
    except Exception:
        html_body = ""
    if html_body and "<table" in html_body.lower():
        raw_body = inject_html_tables_into_body(raw_body, html_body)
    body_text = unwrap_outlook_safelinks(raw_body)
    conv_id = getattr(item, "ConversationID", "") or ""
    try:
        from modules.emails.core.general_email_sync import GeneralEmailSyncThread
        sender_name, sender_email = GeneralEmailSyncThread._resolve_sender(item, sender_name, sender_email, body_text)
    except Exception:
        pass
    dt = getattr(item, "ReceivedTime", None) or getattr(item, "SentOn", None)
    date_sent = dt.strftime("%Y-%m-%d %H:%M:%S") if hasattr(dt, "strftime") else str(dt or "")
    entry_id = getattr(item, "EntryID", "")
    if not entry_id:
        digest = hashlib.sha1(f"{subject}|{sender_name}|{date_sent}".encode("utf-8", errors="ignore")).hexdigest()[:16]
        entry_id = f"MSG-{digest}"

    attachments = []
    try:
        atts = getattr(item, "Attachments", None)
        if atts:
            for i in range(1, atts.Count + 1):
                attachments.append(atts.Item(i).FileName)
    except Exception:
        pass

    try:
        raw_flag = int(getattr(item, "FlagStatus", 0) or 0)
    except Exception:
        raw_flag = 0

    return {
        "id": entry_id,
        "subject": subject,
        "sender_name": sender_name,
        "sender_email": sender_email,
        "company": _extract_company(sender_name, sender_email),
        "date_sent": date_sent,
        "body_text": body_text,
        "clean_body": _strip_trailing_quotes(body_text) if body_text else "",
        "message_id_header": "",
        "in_reply_to": "",
        "conversation_id": conv_id,
        "is_read": 1 if getattr(item, "UnRead", False) is False else 0,
        "flag_status": 2 if raw_flag == 2 else 0,
        "attachments": attachments,
    }


def parse_msg_file(msg_path: Path) -> Optional[Dict[str, Any]]:
    """Parses a local .msg file via Outlook COM OpenSharedItem."""
    from modules.emails.core.outlook_client import OutlookClient

    namespace = OutlookClient.get_namespace()
    if not namespace:
        return None
    try:
        item = namespace.OpenSharedItem(str(Path(msg_path).resolve()))
        return _mail_item_to_dict(item)
    except Exception as e:
        logger.error(f"Failed to parse .msg file {msg_path}: {e}")
        return None


def _iter_folder_tree(folder: Any, max_depth: int = 3) -> List[Any]:
    """Returns [folder, ...subfolders] up to max_depth."""
    result = [folder]
    if max_depth <= 0:
        return result
    try:
        for sub in folder.Folders:
            result.extend(_iter_folder_tree(sub, max_depth - 1))
    except Exception:
        pass
    return result


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
    if ns:
        try:
            sent_folder = ns.GetDefaultFolder(5)  # olFolderSentMail
            sent_path = getattr(sent_folder, "FolderPath", "") or ""
            if sent_path and sent_path.lower() not in {p.lower() for p in extracted_paths}:
                extracted_paths.append(sent_path)
        except Exception:
            pass
    total_folders = len(extracted_paths)
    for idx, fpath in enumerate(extracted_paths):
        if cancel_check and cancel_check():
            break
        if progress_cb:
            progress_cb(idx, total_folders, fpath, extracted_paths[idx + 1 :])
        folder = OutlookClient.get_folder_by_path(fpath)
        if not folder or not ns:
            continue
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
    return thread_id


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
) -> Optional[str]:
    """Parses a dropped .msg, locates the full thread in configured Outlook folders, and pins it."""
    seed = parse_msg_file(msg_path)
    if not seed:
        return None
    return _pin_from_seed_dict(
        db, seed, meeting_id=meeting_id, wg=wg,
        linked_ais=linked_ais, linked_wis=linked_wis, linked_tdocs=linked_tdocs,
        agenda_map=agenda_map, valid_wis=valid_wis, tdoc_meta_map=tdoc_meta_map,
    )


def pin_thread_from_outlook_selection(
    db: TrackedThreadsDatabase,
    meeting_id: str,
    wg: str,
    agenda_map: Optional[Dict[str, Any]] = None,
    valid_wis: Optional[List[str]] = None,
    tdoc_meta_map: Optional[Dict[str, Dict[str, Any]]] = None,
) -> List[str]:
    """Pins thread(s) from the currently selected MailItem(s) in Microsoft Outlook's ActiveExplorer."""
    import win32com.client

    pinned: List[str] = []
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        explorer = outlook.ActiveExplorer()
        if not explorer:
            return pinned
        selection = explorer.Selection
        for i in range(1, selection.Count + 1):
            item = selection.Item(i)
            if getattr(item, "Class", 0) != 43:
                continue
            seed = _mail_item_to_dict(item)
            tid = _pin_from_seed_dict(
                db, seed, meeting_id=meeting_id, wg=wg, seed_live_item=item,
                agenda_map=agenda_map, valid_wis=valid_wis, tdoc_meta_map=tdoc_meta_map,
            )
            if tid and tid not in pinned:
                pinned.append(tid)
    except Exception as e:
        logger.error(f"Failed to read Outlook ActiveExplorer.Selection: {e}")
    return pinned


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
    except Exception:
        pass

    new_msgs_by_tid: Dict[str, List[Dict[str, Any]]] = {tid: [] for tid in tid_map}
    seen_eids_by_tid: Dict[str, set] = {tid: set() for tid in tid_map}
    flag_updates_by_tid: Dict[str, List[tuple]] = {tid: [] for tid in tid_map}
    cids_by_tid: Dict[str, set] = {tid: set() for tid in tid_map}
    conv_pulled_for: set = set()
    extracted_cache: Dict[str, Dict[str, Any]] = {}

    total_folders = len(folder_order)
    for f_idx, fpath in enumerate(folder_order):
        if cancel_check and cancel_check():
            break
        if progress_cb:
            short_name = fpath.replace("\\", "/").rstrip("/").rsplit("/", 1)[-1] or fpath
            progress_cb(f_idx, total_folders, f"Folder: {short_name}", folder_order[f_idx + 1 :])
        folder = OutlookClient.get_folder_by_path(fpath)
        if not folder:
            continue
        allowed_tids = folder_to_tids.get(fpath.lower(), set())
        if not allowed_tids:
            continue
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


def get_meeting_discovery_config_path(meeting_dir: Optional[Path]) -> Optional[Path]:
    if not meeting_dir:
        return None
    return Path(meeting_dir) / "Agenda" / "email_discovery_config.json"


def load_meeting_discovery_config(
    meeting_dir: Optional[Path],
    default_tag: str = "",
    default_start: str = "",
    default_end: str = "",
) -> Dict[str, Any]:
    """Loads per-meeting Outlook thread discovery preferences from `<meeting_dir>/Agenda/email_discovery_config.json`."""
    defaults: Dict[str, Any] = {
        "subject_tag": str(default_tag or "").strip(),
        "require_subject_tag": bool(str(default_tag or "").strip()),
        "use_date_filter": bool(default_start or default_end),
        "start_date": str(default_start or "").strip(),
        "end_date": str(default_end or "").strip(),
        "days_before": 14,
        "days_after": 7,
        "coverage_filter": "UNTRACKED_AND_BODY",
        "ignored_thread_keys": [],
    }
    cfg_path = get_meeting_discovery_config_path(meeting_dir)
    if cfg_path and cfg_path.is_file():
        try:
            loaded = json.loads(cfg_path.read_text(encoding="utf-8"))
            if isinstance(loaded, dict):
                defaults.update(loaded)
        except Exception as exc:
            logger.debug(f"Could not read email discovery config {cfg_path}: {exc}")
    defaults["ignored_thread_keys"] = sorted({
        str(k).strip().lower() for k in (defaults.get("ignored_thread_keys") or []) if str(k).strip()
    })
    return defaults


def save_meeting_discovery_config(meeting_dir: Optional[Path], config: Dict[str, Any]) -> Optional[Path]:
    """Persists per-meeting Outlook thread discovery preferences into `<meeting_dir>/Agenda/email_discovery_config.json`."""
    cfg_path = get_meeting_discovery_config_path(meeting_dir)
    if not cfg_path:
        return None
    cfg_path.parent.mkdir(parents=True, exist_ok=True)
    clean = dict(config or {})
    clean["ignored_thread_keys"] = sorted({
        str(k).strip().lower() for k in (clean.get("ignored_thread_keys") or []) if str(k).strip()
    })
    cfg_path.write_text(json.dumps(clean, indent=2), encoding="utf-8")
    return cfg_path


def _add_thread_row_to_bucket(
    threads_by_key: Dict[str, Dict[str, Any]],
    *,
    eid: str,
    subj: str,
    cid: str,
    date_str: str,
    sender: str,
    is_unread: bool,
    fpath: str,
):
    norm = normalize_subject(subj)
    if not norm:
        return
    key = norm.lower()
    existing = threads_by_key.get(key)
    if eid and existing and eid in existing["entry_ids"]:
        return
    tdoc_in_subj = sorted({m.upper() for m in _TDOC_RE.findall(subj or "")})
    if key not in threads_by_key:
        threads_by_key[key] = {
            "title": norm,
            "normalized_subject": norm,
            "norm_key": key,
            "raw_subject": subj,
            "conversation_id": cid,
            "conversation_ids": [cid] if cid else [],
            "entry_ids": [eid] if eid else [],
            "last_date": date_str,
            "first_date": date_str,
            "last_sender": sender,
            "senders": [sender] if sender else [],
            "count": 1,
            "unread_count": 1 if is_unread else 0,
            "folder": fpath,
            "subject_tdocs": tdoc_in_subj,
        }
    else:
        entry = threads_by_key[key]
        if eid:
            entry["entry_ids"].append(eid)
        if cid and cid not in entry["conversation_ids"]:
            entry["conversation_ids"].append(cid)
        if not entry["conversation_id"] and cid:
            entry["conversation_id"] = cid
        if sender and sender not in entry["senders"]:
            entry["senders"].append(sender)
        for td in tdoc_in_subj:
            if td not in entry["subject_tdocs"]:
                entry["subject_tdocs"].append(td)
        if date_str and (not entry["last_date"] or date_str > entry["last_date"]):
            entry["last_date"] = date_str
            if sender:
                entry["last_sender"] = sender
            entry["raw_subject"] = subj or entry["raw_subject"]
        if date_str and (not entry["first_date"] or date_str < entry["first_date"]):
            entry["first_date"] = date_str
        entry["count"] += 1
        if is_unread:
            entry["unread_count"] += 1


def list_recent_threads_in_folders(
    wg: str,
    max_items_per_folder: int = 1500,
    progress_cb: Optional[Callable[[int, int, str, List[str]], None]] = None,
    cancel_check: Optional[Callable[[], bool]] = None,
) -> List[Dict[str, Any]]:
    """Lists recent unique conversation threads across configured Outlook WG folders + Sent Items via MAPI GetTable()."""
    from modules.emails.core.outlook_client import OutlookClient

    folders = _extract_folder_paths(load_wg_email_config(wg)) if wg else []
    ns = OutlookClient.get_namespace()
    if ns:
        try:
            sent_path = getattr(ns.GetDefaultFolder(5), "FolderPath", "") or ""
            if sent_path and sent_path.lower() not in {p.lower() for p in folders}:
                folders.append(sent_path)
        except Exception:
            pass

    threads_by_key: Dict[str, Dict[str, Any]] = {}
    total_folders = len(folders)

    for idx, fpath in enumerate(folders):
        if cancel_check and cancel_check():
            break
        if progress_cb:
            progress_cb(idx, total_folders, fpath, folders[idx + 1 :])
        folder = OutlookClient.get_folder_by_path(fpath)
        if not folder:
            continue
        try:
            if hasattr(folder, "GetTable"):
                tbl = folder.GetTable()
                try:
                    tbl.Sort("[ReceivedTime]", True)
                except Exception:
                    try:
                        tbl.Sort("[SentOn]", True)
                    except Exception:
                        pass
                cols = tbl.Columns
                for col_name in ("EntryID", "Subject", "MessageClass", "ConversationID", "SenderName", "ReceivedTime", "UnRead"):
                    try:
                        cols.Add(col_name)
                    except Exception:
                        pass
                scanned = 0
                while not tbl.EndOfTable and scanned < max_items_per_folder:
                    if cancel_check and cancel_check():
                        break
                    scanned += 1
                    try:
                        row = tbl.GetNextRow()
                        msg_cls = str(row("MessageClass") or "")
                        if msg_cls and not msg_cls.startswith("IPM.Note"):
                            continue
                        subj = str(row("Subject") or "")
                        eid = str(row("EntryID") or "")
                        try:
                            cid = str(row("ConversationID") or "")
                        except Exception:
                            cid = ""
                        try:
                            sender = str(row("SenderName") or "")
                        except Exception:
                            sender = ""
                        try:
                            dt = row("ReceivedTime")
                            date_str = dt.strftime("%Y-%m-%d %H:%M") if hasattr(dt, "strftime") else str(dt or "")[:16]
                        except Exception:
                            date_str = ""
                        try:
                            is_unread = bool(row("UnRead"))
                        except Exception:
                            is_unread = False
                        _add_thread_row_to_bucket(
                            threads_by_key, eid=eid, subj=subj, cid=cid,
                            date_str=date_str, sender=sender, is_unread=is_unread, fpath=fpath,
                        )
                    except Exception:
                        continue
            else:
                items = folder.Items
                items.Sort("[ReceivedTime]", True)
                total = min(len(items), max_items_per_folder)
                for i in range(1, total + 1):
                    if cancel_check and cancel_check():
                        break
                    try:
                        item = items.Item(i)
                        if getattr(item, "Class", 43) != 43:
                            continue
                        subj = getattr(item, "Subject", "") or ""
                        eid = str(getattr(item, "EntryID", "") or "")
                        cid = str(getattr(item, "ConversationID", "") or "")
                        dt = getattr(item, "ReceivedTime", None) or getattr(item, "SentOn", None)
                        date_str = dt.strftime("%Y-%m-%d %H:%M") if hasattr(dt, "strftime") else str(dt or "")[:16]
                        sender = str(getattr(item, "SenderName", "") or "")
                        is_unread = bool(getattr(item, "UnRead", False))
                        _add_thread_row_to_bucket(
                            threads_by_key, eid=eid, subj=subj, cid=cid,
                            date_str=date_str, sender=sender, is_unread=is_unread, fpath=fpath,
                        )
                    except Exception:
                        continue
        except Exception as e:
            logger.warning(f"Failed listing threads in '{fpath}': {e}")

    return sorted(threads_by_key.values(), key=lambda x: x.get("last_date", ""), reverse=True)


def classify_discovered_threads(
    threads: List[Dict[str, Any]],
    tracked_db: TrackedThreadsDatabase,
    general_email_db: Optional[Any] = None,
    *,
    meeting_id: str = "",
    valid_tdocs: Optional[set] = None,
    agenda_map: Optional[Dict[str, Any]] = None,
    valid_wis: Optional[List[str]] = None,
    ignored_keys: Optional[set] = None,
    tdoc_meta_map: Optional[Dict[str, Dict[str, Any]]] = None,
) -> List[Dict[str, Any]]:
    """
    Annotates discovered Outlook threads with tracking coverage:
    - PINNED: already tracked in TrackedThreadsDatabase for this meeting
    - SUBJECT_TDOC: unpinned, subject contains an official/meeting TDoc
    - BODY_TDOC: unpinned, no TDoc in subject, but matched in GeneralEmailDatabase body/quoted
    - STRICT_UNTRACKED: unpinned, zero TDoc matches in subject or GeneralEmailDatabase
    And marks `is_ignored` when `norm_key` is in `ignored_keys`.
    """
    pinned_threads = tracked_db.get_threads_for_meeting(meeting_id) if tracked_db else []
    pinned_norms = set()
    pinned_cids = set()
    for pt in pinned_threads:
        if pt.get("normalized_subject"):
            pinned_norms.add(str(pt["normalized_subject"]).strip().lower())
        for pat in (pt.get("subject_patterns") or []):
            if pat:
                pinned_norms.add(str(pat).strip().lower())
        for c in (pt.get("conversation_ids") or []):
            if c:
                pinned_cids.add(str(c).strip())

    match_locs_by_eid = (
        general_email_db.get_tdoc_match_locations_by_email_id()
        if general_email_db is not None and hasattr(general_email_db, "get_tdoc_match_locations_by_email_id")
        else {}
    )
    ignored_set = {str(k).strip().lower() for k in (ignored_keys or set()) if str(k).strip()}
    valid_tdocs_up = {str(t).strip().upper() for t in (valid_tdocs or set()) if str(t).strip()}

    classified: List[Dict[str, Any]] = []
    for raw in (threads or []):
        item = dict(raw)
        nkey = str(item.get("norm_key") or normalize_subject(item.get("title", "")).lower()).strip()
        cids = {str(c).strip() for c in (item.get("conversation_ids") or ([item.get("conversation_id")] if item.get("conversation_id") else [])) if str(c).strip()}
        eids = [str(e).strip() for e in (item.get("entry_ids") or []) if str(e).strip()]
        subj_tdocs = [str(t).strip().upper() for t in (item.get("subject_tdocs") or []) if str(t).strip()]
        if valid_tdocs_up:
            subj_tdocs = [t for t in subj_tdocs if t in valid_tdocs_up]

        is_pinned = (nkey in pinned_norms) or bool(cids & pinned_cids)
        is_ignored = nkey in ignored_set

        all_locs = set()
        for eid in eids:
            all_locs.update(match_locs_by_eid.get(eid, ()))

        if is_pinned:
            cov = "PINNED"
            cov_label = "📌 Pinned"
        elif subj_tdocs or ("Subject" in all_locs):
            cov = "SUBJECT_TDOC"
            cov_label = "🟢 Subject TDoc"
        elif all_locs:
            cov = "BODY_TDOC"
            cov_label = "🟡 TDoc in Body Only"
        else:
            cov = "STRICT_UNTRACKED"
            cov_label = "🔴 No TDoc"

        sugg = auto_suggest_links(
            item.get("raw_subject") or item.get("title", ""),
            agenda_map=agenda_map,
            valid_wis=valid_wis,
            tdoc_meta_map=tdoc_meta_map,
            extra_tdocs=subj_tdocs,
        )
        item["coverage"] = cov
        item["coverage_label"] = "🚫 Ignored" if (is_ignored and not is_pinned) else cov_label
        item["is_pinned"] = is_pinned
        item["is_ignored"] = is_ignored
        item["suggested_ais"] = sugg.get("ais") or []
        item["suggested_sub_items"] = sugg.get("sub_items") or []
        item["suggested_wis"] = sugg.get("wis") or []
        item["suggested_tdocs"] = subj_tdocs or (sugg.get("tdocs") or [])
        classified.append(item)
    return classified


def pin_thread_by_subject_or_cid(
    db: TrackedThreadsDatabase,
    title: str,
    normalized_subject: str,
    conversation_id: str,
    meeting_id: str,
    wg: str,
    agenda_map: Optional[Dict[str, Any]] = None,
    valid_wis: Optional[List[str]] = None,
    tdoc_meta_map: Optional[Dict[str, Dict[str, Any]]] = None,
    linked_ais: Optional[List[str]] = None,
    linked_wis: Optional[List[str]] = None,
    linked_tdocs: Optional[List[str]] = None,
    linked_sub_items: Optional[List[str]] = None,
    seed_messages: Optional[List[Dict[str, Any]]] = None,
    refresh_from_outlook: bool = True,
) -> str:
    """Pins a thread selected from the configured Outlook folder browser or inbox and pulls matching messages."""
    folders = _extract_folder_paths(load_wg_email_config(wg)) if wg else []
    suggested = auto_suggest_links(
        title,
        agenda_map=agenda_map,
        valid_wis=valid_wis,
        tdoc_meta_map=tdoc_meta_map,
        extra_tdocs=linked_tdocs,
    )
    thread_id = db.create_or_get_thread(
        title=title,
        normalized_subject=normalized_subject,
        conversation_id=conversation_id,
        outlook_folders=folders,
        meeting_id=meeting_id,
        wg=wg,
        linked_ais=linked_ais if linked_ais is not None else suggested["ais"],
        linked_wis=linked_wis if linked_wis is not None else suggested["wis"],
        linked_tdocs=linked_tdocs if linked_tdocs is not None else suggested["tdocs"],
        linked_sub_items=linked_sub_items if linked_sub_items is not None else suggested.get("sub_items"),
    )
    if seed_messages:
        db.upsert_messages(thread_id, seed_messages)
    if refresh_from_outlook:
        refresh_thread(db, thread_id)
    return thread_id


def open_outlook_message(entry_id: str) -> bool:
    """Opens a specific Outlook MailItem in its native Outlook inspector window by EntryID."""
    if not entry_id or str(entry_id).startswith("MSG-"):
        return False
    from modules.emails.core.outlook_client import OutlookClient

    ns = OutlookClient.get_namespace()
    if not ns:
        return False
    try:
        item = ns.GetItemFromID(str(entry_id).strip())
        if item:
            item.Display()
            return True
    except Exception as e:
        logger.error(f"Failed to display Outlook message {entry_id}: {e}")
    return False


def reload_outlook_message(
    db: TrackedThreadsDatabase,
    thread_id: str,
    entry_id: str,
) -> bool:
    """Re-fetches a single cached Outlook message by EntryID and re-parses its body, tables, links, and flags."""
    eid = str(entry_id or "").strip()
    if not eid or eid.startswith("MSG-"):
        return False
    from modules.emails.core.outlook_client import OutlookClient

    ns = OutlookClient.get_namespace()
    if not ns:
        return False
    try:
        item = ns.GetItemFromID(eid)
        if not item:
            return False
        msg_dict = _mail_item_to_dict(item)
        db.upsert_messages(thread_id, [msg_dict])
        return True
    except Exception as e:
        logger.warning(f"Failed to reload Outlook message {eid}: {e}")
    return False


def reload_all_thread_messages_from_outlook(
    db: TrackedThreadsDatabase,
    thread_id: str,
    progress_cb: Optional[Callable[[int, int, str, List[str]], None]] = None,
    cancel_check: Optional[Callable[[], bool]] = None,
) -> int:
    """Re-fetches every Outlook-backed message in a tracked thread via EntryID (re-parsing tables/links/state)."""
    from modules.emails.core.outlook_client import OutlookClient

    msgs = [
        m for m in db.get_thread_messages(thread_id)
        if str(m.get("id") or "").strip() and not str(m.get("id") or "").startswith("MSG-")
    ]
    if not msgs:
        return 0
    ns = OutlookClient.get_namespace()
    if not ns:
        return 0

    reloaded: List[Dict[str, Any]] = []
    total = len(msgs)
    labels = [f"{m.get('sender_name') or 'Unknown'} ({m.get('date_sent', '')[:16]})" for m in msgs]
    for idx, m in enumerate(msgs):
        if cancel_check and cancel_check():
            break
        if progress_cb:
            progress_cb(idx, total, labels[idx], labels[idx + 1 :])
        try:
            item = ns.GetItemFromID(str(m["id"]).strip())
            if item:
                reloaded.append(_mail_item_to_dict(item))
        except Exception as e:
            logger.debug(f"Could not reload message {m.get('id')}: {e}")
    if reloaded:
        db.upsert_messages(thread_id, reloaded)
    return len(reloaded)


def open_outlook_attachment(entry_id: str, filename: str, save_dir: Optional[Path] = None) -> Optional[Path]:
    """Saves an attachment from an Outlook MailItem to disk and opens it via document_opener (applying sensitivity labels)."""
    if not entry_id or not filename or str(entry_id).startswith("MSG-"):
        return None
    from modules.emails.core.outlook_client import OutlookClient
    from modules.word_tools.core.document_opener import open_document

    ns = OutlookClient.get_namespace()
    if not ns:
        return None
    try:
        item = ns.GetItemFromID(str(entry_id).strip())
        atts = getattr(item, "Attachments", None)
        if not atts:
            return None
        target_dir = Path(save_dir) if save_dir else (Path.home() / ".3gpp_tools_tmp" / "email_attachments")
        target_dir.mkdir(parents=True, exist_ok=True)
        for i in range(1, atts.Count + 1):
            att = atts.Item(i)
            fname = getattr(att, "FileName", "") or ""
            if fname.lower() == filename.lower():
                safe_name = "".join(c for c in fname if c not in '<>:"/\\|?*').strip() or "attachment.bin"
                out_path = target_dir / safe_name
                att.SaveAsFile(str(out_path.resolve()))
                open_document(out_path.resolve())
                return out_path
    except Exception as e:
        logger.error(f"Failed to open attachment '{filename}' from {entry_id}: {e}")
    return None


def _build_candidate_document_urls(url: str) -> List[str]:
    """Builds candidate download URLs translating between 10.10.10.10 local server and 3gpp.org/ftp."""
    candidates = [url]
    # 10.10.10.10/ftp/SA/SA2/Inbox/... -> https://www.3gpp.org/ftp/Meetings_3GPP_SYNC/SA2/Inbox/...
    m_local = re.match(r"^https?://10\.10\.10\.10/ftp/(?:SA|RAN|CT)/([A-Z0-9]+)/(.+)$", url, re.IGNORECASE)
    if m_local:
        wg, rest = m_local.group(1).upper(), m_local.group(2)
        candidates.append(f"https://www.3gpp.org/ftp/Meetings_3GPP_SYNC/{wg}/{rest}")
    # https://www.3gpp.org/ftp/tsg_sa/WG2_Arch/TSGS2_178_.../INBOX/... -> http://10.10.10.10/ftp/SA/SA2/Inbox/...
    m_remote = re.match(r"^https?://(?:www\.)?3gpp\.org/ftp/tsg_(sa|ran|ct)/[^/]+/[^/]+/(.+)$", url, re.IGNORECASE)
    if m_remote:
        tsg = m_remote.group(1).upper()
        rest = m_remote.group(2)
        if tsg == "SA":
            candidates.insert(0, f"http://10.10.10.10/ftp/SA/SA2/{rest}")
    return list(dict.fromkeys(candidates))


def get_drafts_cache_dir(agenda_dir: Path) -> Path:
    """Returns canonical DraftsCache directory under `<meeting_dir>/Agenda/DraftsCache`."""
    return Path(agenda_dir) / "DraftsCache"


def _infer_agenda_dir_from_cache_dir(cache_dir: Optional[Path]) -> Optional[Path]:
    if not cache_dir:
        return None
    p = Path(cache_dir)
    if p.name == "InboxAttachments":
        return p.parent
    if p.parent.name in ("ThreadAttachments", "TDocEmailCache"):
        return p.parent.parent
    if p.name == "DraftsCache":
        return p.parent
    return None


def resolve_draft_cache_target(url: str, drafts_root: Path) -> Path:
    """
    Maps a 10.10.10.10 / Drafts URL to a deterministic local file path inside `drafts_root`,
    preserving subfolders after `/Drafts/` to prevent same-filename collisions across KIs.
    """
    import urllib.parse

    drafts_root = Path(drafts_root)
    parsed = urllib.parse.urlsplit((url or "").strip())
    raw_segs = [urllib.parse.unquote(s) for s in parsed.path.split("/") if s]
    low_segs = [s.strip().lower() for s in raw_segs]
    rel_segs = raw_segs[low_segs.index("drafts") + 1 :] if "drafts" in low_segs else raw_segs[-1:]

    clean_segs = []
    for seg in rel_segs:
        cleaned = "".join(c for c in seg if c not in '<>:"/\\|?*').strip()
        if cleaned and cleaned not in (".", ".."):
            clean_segs.append(cleaned)

    if not clean_segs:
        clean_segs = ["draft_document.docx"]

    candidate = drafts_root.joinpath(*clean_segs)
    if len(str(candidate)) < 220:
        return candidate
    h = hashlib.sha1(url.encode("utf-8", errors="ignore")).hexdigest()[:8]
    return drafts_root / f"{h}_{clean_segs[-1]}"


def _download_document_candidates(
    url: str,
    target_path: Path,
    local_ip_unreachable: Optional[List[bool]] = None,
) -> bool:
    """Downloads a 3GPP / 10.10.10.10 URL to `target_path` with fast 10.10.10.10 connect timeout & failover."""
    from core.network.session import NetworkSession

    session = NetworkSession()
    target_path.parent.mkdir(parents=True, exist_ok=True)
    skip_local = bool(local_ip_unreachable and local_ip_unreachable[0])

    for cand_url in _build_candidate_document_urls(url):
        is_local = "://10.10.10.10/" in cand_url.lower()
        if is_local and skip_local:
            continue
        timeout = (1.5, 12) if is_local else (4, 15)
        try:
            resp = session.get(cand_url, timeout=timeout)
            if resp.status_code == 200 and resp.content:
                target_path.write_bytes(resp.content)
                return True
        except Exception as e:
            if is_local and local_ip_unreachable is not None:
                local_ip_unreachable[0] = True
                skip_local = True
            logger.debug(f"Candidate URL download failed ({cand_url}): {e}")
    return False


def cache_drafts_for_messages(
    messages: List[Dict[str, Any]],
    agenda_dir: Path,
    progress_cb: Optional[Callable[[str], None]] = None,
    cancel_check: Optional[Callable[[], bool]] = None,
) -> List[Path]:
    """
    Proactively downloads uncached 10.10.10.10 /Drafts/ document URLs referenced by emails into
    `<agenda_dir>/DraftsCache/`, immediately skipping files that already exist locally.
    """
    draft_urls = extract_drafts_document_urls(messages, prefer_clean_body=True)
    if not draft_urls:
        return []

    drafts_root = get_drafts_cache_dir(agenda_dir)
    pending: List[tuple] = []
    cached_paths: List[Path] = []
    for u in draft_urls:
        tgt = resolve_draft_cache_target(u, drafts_root)
        if tgt.exists() and tgt.stat().st_size > 0:
            cached_paths.append(tgt)
        else:
            pending.append((u, tgt))

    if not pending:
        return cached_paths

    local_unreachable = [False]
    total = len(pending)
    for idx, (u, tgt) in enumerate(pending, 1):
        if cancel_check and cancel_check():
            break
        if progress_cb:
            progress_cb(f"Caching draft ({idx}/{total}): {tgt.name[:32]}")
        if _download_document_candidates(u, tgt, local_ip_unreachable=local_unreachable):
            cached_paths.append(tgt)
    return cached_paths


def _auto_cache_db_drafts(
    db: Any,
    messages: List[Dict[str, Any]],
    progress_cb: Optional[Callable[[str], None]] = None,
    cancel_check: Optional[Callable[[], bool]] = None,
) -> List[Path]:
    db_path = getattr(db, "db_path", None)
    if not db_path or str(db_path) == ":memory:":
        return []
    agenda_dir = Path(db_path).parent
    return cache_drafts_for_messages(messages, agenda_dir, progress_cb=progress_cb, cancel_check=cancel_check)


def cache_and_open_thread_document_url(
    url: str,
    thread_dir: Path,
    open_after: bool = True,
    drafts_root: Optional[Path] = None,
) -> Optional[Path]:
    """
    Downloads and caches a document linked in an email thread (e.g. 10.10.10.10 or 3gpp.org/ftp)
    into `DraftsCache/` (for Drafts URLs) or `thread_dir` and optionally opens it via sensitivity handler.
    """
    import urllib.parse
    import zipfile

    if not url:
        return None
    thread_dir = Path(thread_dir)
    thread_dir.mkdir(parents=True, exist_ok=True)

    raw_name = urllib.parse.unquote(urllib.parse.urlsplit(url).path.rsplit("/", 1)[-1]) or "document.docx"
    safe_name = "".join(c for c in raw_name if c not in '<>:"/\\|?*').strip() or "document.docx"
    flat_target = thread_dir / safe_name

    eff_drafts_root = drafts_root
    if eff_drafts_root is None:
        inferred_agenda = _infer_agenda_dir_from_cache_dir(thread_dir)
        if inferred_agenda is not None:
            eff_drafts_root = get_drafts_cache_dir(inferred_agenda)

    if is_drafts_document_url(url) and eff_drafts_root is not None:
        draft_target = resolve_draft_cache_target(url, eff_drafts_root)
        if draft_target.exists() and draft_target.stat().st_size > 0:
            target_path = draft_target
        elif flat_target.exists() and flat_target.stat().st_size > 0:
            target_path = flat_target
        else:
            target_path = draft_target
    else:
        target_path = flat_target

    if not (target_path.exists() and target_path.stat().st_size > 0):
        if not _download_document_candidates(url, target_path):
            return None

    open_target = target_path
    if target_path.suffix.lower() == ".zip" and zipfile.is_zipfile(target_path):
        try:
            with zipfile.ZipFile(target_path, "r") as zf:
                for member in zf.namelist():
                    if member.lower().endswith((".docx", ".doc", ".pptx", ".pdf", ".xlsx")) and not Path(member).name.startswith("~$"):
                        extracted = target_path.parent / Path(member).name
                        if not extracted.exists():
                            extracted.write_bytes(zf.read(member))
                        open_target = extracted
                        break
        except Exception:
            pass

    if open_after:
        try:
            from modules.word_tools.core.document_opener import open_document
            open_document(open_target.resolve())
        except Exception as e:
            logger.error(f"Could not open cached document {open_target}: {e}")
    return open_target


def cache_all_thread_documents_and_attachments(
    db: TrackedThreadsDatabase,
    thread_id: str,
    thread_dir: Path,
) -> List[Path]:
    """Caches all linked documents (10.10.10.10 / 3GPP FTP) and Outlook attachments for a thread into thread_dir."""
    messages = db.get_thread_messages(thread_id)
    cached: List[Path] = []
    thread_dir = Path(thread_dir)
    thread_dir.mkdir(parents=True, exist_ok=True)

    # 1. Linked URLs in message bodies
    for doc_url in extract_thread_document_urls(messages):
        p = cache_and_open_thread_document_url(doc_url, thread_dir, open_after=False)
        if p:
            cached.append(p)

    # 2. Outlook attachments
    from modules.emails.core.outlook_client import OutlookClient
    ns = OutlookClient.get_namespace()
    if ns:
        for msg in messages:
            eid = str(msg.get("id") or "").strip()
            atts_names = msg.get("attachments") or []
            if not eid or eid.startswith("MSG-") or not atts_names:
                continue
            try:
                item = ns.GetItemFromID(eid)
                atts = getattr(item, "Attachments", None)
                if not atts:
                    continue
                for i in range(1, atts.Count + 1):
                    att = atts.Item(i)
                    fname = getattr(att, "FileName", "") or ""
                    if not fname or fname.lower().endswith((".jpg", ".jpeg", ".png", ".gif")):
                        continue
                    safe_name = "".join(c for c in fname if c not in '<>:"/\\|?*').strip()
                    if safe_name:
                        out_path = thread_dir / safe_name
                        if not out_path.exists():
                            att.SaveAsFile(str(out_path.resolve()))
                        cached.append(out_path)
            except Exception:
                continue
    return cached


def reply_to_outlook_message(entry_id: str, reply_all: bool = True) -> bool:
    """Opens a native Outlook reply window for a specific MailItem by EntryID."""
    if not entry_id or str(entry_id).startswith("MSG-"):
        return False
    from modules.emails.core.outlook_client import OutlookClient

    ns = OutlookClient.get_namespace()
    if not ns:
        return False
    try:
        item = ns.GetItemFromID(str(entry_id).strip())
        if item:
            reply_item = item.ReplyAll() if reply_all else item.Reply()
            reply_item.Display()
            return True
    except Exception as e:
        logger.error(f"Failed to reply to Outlook message {entry_id}: {e}")
    return False


def mark_thread_read_with_outlook(
    db: TrackedThreadsDatabase,
    thread_id: str,
    sync_outlook: bool = True,
) -> int:
    """
    Marks all messages in a tracked thread as read in SQLite and synchronizes
    the read status back to Microsoft Outlook. Returns count of Outlook items updated.
    """
    msgs = db.get_thread_messages(thread_id)
    unread_ids = [
        str(m.get("id") or "")
        for m in msgs
        if int(m.get("is_read", 0) or 0) == 0 and m.get("id")
    ]
    db.mark_thread_read(thread_id)
    if sync_outlook and unread_ids:
        from modules.emails.core.outlook_client import OutlookClient
        return OutlookClient.set_messages_read_status(unread_ids, is_read=True)
    return 0


def set_thread_message_read_with_outlook(
    db: TrackedThreadsDatabase,
    message_id: str,
    is_read: bool,
    sync_outlook: bool = True,
) -> int:
    """Marks a single message in a tracked thread read/unread and syncs to Outlook."""
    msg_id = str(message_id or "").strip()
    if not msg_id:
        return 0
    db.set_message_read_status(msg_id, bool(is_read))
    if sync_outlook and not msg_id.startswith("MSG-"):
        from modules.emails.core.outlook_client import OutlookClient
        return OutlookClient.set_messages_read_status([msg_id], is_read=bool(is_read))
    return 0


def set_thread_message_flag_with_outlook(
    db: TrackedThreadsDatabase,
    message_id: str,
    is_flagged: bool,
    sync_outlook: bool = True,
) -> int:
    """Sets/clears the flag on a single tracked thread message and syncs to Outlook."""
    msg_id = str(message_id or "").strip()
    if not msg_id:
        return 0
    db.set_message_flag_status(msg_id, bool(is_flagged))
    if sync_outlook and not msg_id.startswith("MSG-"):
        from modules.emails.core.outlook_client import OutlookClient
        return OutlookClient.set_messages_flag_status([msg_id], is_flagged=bool(is_flagged))
    return 0


def prepare_thread_llm_workspace(
    db: TrackedThreadsDatabase,
    thread_id: str,
    cache_dir: Path,
    deduplicate_quotes: bool = True,
    newest_first: bool = False,
    cache_linked_docs: bool = True,
) -> Dict[str, Any]:
    """
    Prepares a self-contained LLM workspace directory containing thread.md plus any
    cached linked documents/attachments converted to Markdown when possible.
    Returns {'workspace_dir': Path, 'thread_md': Path, 'doc_mds': List[Path]}.
    """
    from modules.emails.core.thread_markdown_exporter import export_thread_to_markdown

    cache_dir = Path(cache_dir)
    cache_dir.mkdir(parents=True, exist_ok=True)
    ws_dir = cache_dir / "llm_workspace"
    ws_dir.mkdir(parents=True, exist_ok=True)

    thread = db.get_thread(thread_id) or {"thread_id": thread_id}
    messages = db.get_thread_messages(thread_id)
    md_text = export_thread_to_markdown(
        thread,
        messages,
        deduplicate_quotes=deduplicate_quotes,
        interactive_links=False,
        newest_first=newest_first,
    )
    thread_md = ws_dir / f"{thread_id}.md"
    thread_md.write_text(md_text, encoding="utf-8")

    cached_files: List[Path] = []
    if cache_linked_docs:
        try:
            cached_files = cache_all_thread_documents_and_attachments(db, thread_id, cache_dir)
        except Exception as e:
            logger.debug(f"Could not cache linked docs during LLM workspace prep: {e}")

    # Also include any already-cached .docx/.docm/.md files in cache_dir
    for existing in cache_dir.iterdir():
        if existing.is_file() and existing not in cached_files:
            cached_files.append(existing)

    doc_mds: List[Path] = []
    for fpath in cached_files:
        if fpath.suffix.lower() in (".docx", ".docm"):
            out_md = ws_dir / f"{fpath.stem}.md"
            if not out_md.exists():
                try:
                    from modules.meetings.core.tdoc_markdown_converter import convert_document_to_markdown
                    convert_document_to_markdown(str(fpath), str(out_md))
                except Exception as e:
                    logger.debug(f"Markdown conversion skipped for {fpath.name}: {e}")
            if out_md.exists():
                doc_mds.append(out_md)
        elif fpath.suffix.lower() == ".md" and fpath.parent != ws_dir:
            doc_mds.append(fpath)

    return {"workspace_dir": ws_dir, "thread_md": thread_md, "doc_mds": doc_mds}



