# --- File: src/modules/emails/core/tracked_threads_service.py ---
import logging
import re
from pathlib import Path
from typing import Any, Dict, List, Optional

from modules.emails.core.email_config import load_wg_email_config
from modules.emails.core.thread_markdown_exporter import unwrap_outlook_safelinks
from modules.emails.core.tracked_threads_db import TrackedThreadsDatabase

logger = logging.getLogger(__name__)

_PREFIX_RE = re.compile(r"^(?:\s*(?:re|fwd|fw|aw|wg|sv|antw)\s*:\s*)+", re.IGNORECASE)
_TAG_RE = re.compile(r"^\s*\[[^\]]{2,80}\]\s*")
_TDOC_RE = re.compile(r"\b(?:S[1-6]|R[1-5]|C[1-6]|SP|RP|CP)-\d{6,7}\b", re.IGNORECASE)
_KI_AI_RE = re.compile(r"\b(?:KI\s*#?\s*\d+(?:[._]\d+)?|AI\s*\d+(?:\.\d+)+)\b", re.IGNORECASE)
_WI_RE = re.compile(r"\b(?:FS|TEI\d+|e?[A-Z0-9]{2,})_[A-Z0-9_]{2,}\b")


def auto_suggest_links(subject: str, body: str = "") -> Dict[str, List[str]]:
    """Extracts candidate AIs/KIs, WIs, and TDocs from an email subject and body."""
    text_subj = subject or ""
    text_all = f"{text_subj}\n{(body or '')[:2000]}"
    ais = sorted({re.sub(r"\s+", "", m.upper()) for m in _KI_AI_RE.findall(text_subj)})
    wis = sorted({m.strip() for m in _WI_RE.findall(text_subj)})
    tdocs = sorted({m.upper() for m in _TDOC_RE.findall(text_all)})
    return {"ais": ais, "wis": wis, "tdocs": tdocs}


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


_QUOTE_SPLIT_RE = re.compile(
    r"\r?\n(?:_{5,}|-{5,}|From:\s+.+?\r?\nSent:\s+|On\s+.+?wrote:\s*\r?\n)",
    re.IGNORECASE,
)


def _strip_trailing_quotes(body_text: str) -> str:
    """Strips pure trailing quoted reply chains while preserving inline replies."""
    if not body_text:
        return ""
    parts = _QUOTE_SPLIT_RE.split(body_text, maxsplit=1)
    return parts[0].strip() if parts else body_text.strip()


def _mail_item_to_dict(item: Any) -> Dict[str, Any]:
    """Extracts standardized message fields from an Outlook COM MailItem."""
    subject = getattr(item, "Subject", "") or ""
    sender_name = getattr(item, "SenderName", "") or ""
    try:
        sender_email = getattr(item, "SenderEmailAddress", "") or ""
    except Exception:
        sender_email = ""
    body_text = unwrap_outlook_safelinks(getattr(item, "Body", "") or "")
    conv_id = getattr(item, "ConversationID", "") or ""
    entry_id = getattr(item, "EntryID", "") or f"MSG-{abs(hash((subject, sender_name, body_text[:100])))}"

    try:
        from modules.emails.core.general_email_sync import GeneralEmailSyncThread
        sender_name, sender_email = GeneralEmailSyncThread._resolve_sender(item, sender_name, sender_email, body_text)
    except Exception:
        pass

    dt = getattr(item, "ReceivedTime", None) or getattr(item, "SentOn", None)
    date_sent = dt.strftime("%Y-%m-%d %H:%M:%S") if hasattr(dt, "strftime") else str(dt or "")

    attachments = []
    try:
        atts = getattr(item, "Attachments", None)
        if atts:
            for i in range(1, atts.Count + 1):
                attachments.append(atts.Item(i).FileName)
    except Exception:
        pass

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


def _collect_conversation_items(seed_item: Any, max_items: int = 250) -> List[Dict[str, Any]]:
    """Uses Outlook's GetConversation() table to fetch all emails in a thread across all folders."""
    from modules.emails.core.outlook_client import OutlookClient

    results: List[Dict[str, Any]] = []
    seen_ids = set()
    try:
        conv = seed_item.GetConversation()
        if not conv:
            return results
        tbl = conv.GetTable()
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
                item = ns.GetItemFromID(eid)
                if getattr(item, "Class", 43) != 43:
                    continue
                d = _mail_item_to_dict(item)
                seen_ids.add(d["id"])
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
) -> List[Dict[str, Any]]:
    """Scans configured Outlook folders via fast MAPI GetTable and Outlook GetConversation() for all thread messages."""
    from modules.emails.core.outlook_client import OutlookClient

    matches: List[Dict[str, Any]] = []
    seen_ids = set()
    norm_patterns = {p.lower() for p in subject_patterns if p}
    conv_set = {c for c in conversation_ids if c}

    # 1. If we already have a live Outlook MailItem, pull its full cross-folder conversation first
    if seed_live_item is not None:
        for d in _collect_conversation_items(seed_live_item):
            if d["id"] not in seen_ids:
                seen_ids.add(d["id"])
                matches.append(d)
                if d.get("conversation_id"):
                    conv_set.add(d["conversation_id"])

    # 2. Scan configured folders using bulk MAPI GetTable (sub-second even on 5,000+ item folders)
    ns = OutlookClient.get_namespace()
    conv_pulled_for = set(conv_set)
    for fpath in _extract_folder_paths(folder_paths):
        folder = OutlookClient.get_folder_by_path(fpath)
        if not folder or not ns:
            continue
        try:
            tbl = folder.GetTable()
            tbl.Sort("[ReceivedTime]", True)
            cols = tbl.Columns
            for col_name in ("EntryID", "Subject", "MessageClass"):
                try:
                    cols.Add(col_name)
                except Exception:
                    pass
            scanned = 0
            while not tbl.EndOfTable and scanned < max_scan_per_folder:
                scanned += 1
                try:
                    row = tbl.GetNextRow()
                    msg_cls = str(row("MessageClass") or "")
                    if msg_cls and not msg_cls.startswith("IPM.Note"):
                        continue
                    eid = row("EntryID")
                    subj = str(row("Subject") or "")
                    norm = normalize_subject(subj).lower()
                    if norm and norm in norm_patterns:
                        if eid in seen_ids:
                            continue
                        item = ns.GetItemFromID(eid)
                        cid = getattr(item, "ConversationID", "") or ""
                        if cid and cid not in conv_pulled_for:
                            conv_pulled_for.add(cid)
                            conv_set.add(cid)
                            for cd in _collect_conversation_items(item):
                                if cd["id"] not in seen_ids:
                                    seen_ids.add(cd["id"])
                                    matches.append(cd)
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
    seed_live_item: Any = None,
) -> str:
    norm_subj = normalize_subject(seed["subject"])
    folders = _extract_folder_paths(load_wg_email_config(wg)) if wg else []
    suggested = auto_suggest_links(seed["subject"], seed.get("clean_body") or seed.get("body_text", ""))
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
    return thread_id


def pin_thread_from_msg(
    db: TrackedThreadsDatabase,
    msg_path: Path,
    meeting_id: str,
    wg: str,
    linked_ais: Optional[List[str]] = None,
    linked_wis: Optional[List[str]] = None,
    linked_tdocs: Optional[List[str]] = None,
) -> Optional[str]:
    """Parses a dropped .msg, locates the full thread in configured Outlook folders, and pins it."""
    seed = parse_msg_file(msg_path)
    if not seed:
        return None
    return _pin_from_seed_dict(
        db, seed, meeting_id=meeting_id, wg=wg,
        linked_ais=linked_ais, linked_wis=linked_wis, linked_tdocs=linked_tdocs,
    )


def pin_thread_from_outlook_selection(
    db: TrackedThreadsDatabase,
    meeting_id: str,
    wg: str,
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
            tid = _pin_from_seed_dict(db, seed, meeting_id=meeting_id, wg=wg, seed_live_item=item)
            if tid and tid not in pinned:
                pinned.append(tid)
    except Exception as e:
        logger.error(f"Failed to read Outlook ActiveExplorer.Selection: {e}")
    return pinned


def refresh_thread(db: TrackedThreadsDatabase, thread_id: str) -> int:
    """Refreshes a single tracked thread from its configured Outlook folders. Returns new message count."""
    thread = db.get_thread(thread_id)
    if not thread:
        return 0
    folders = thread.get("outlook_folders") or (load_wg_email_config(thread.get("wg", "")) if thread.get("wg") else [])
    if not folders:
        return 0
    found = _collect_matching_from_folders(
        folders,
        thread.get("conversation_ids") or [],
        thread.get("subject_patterns") or [thread.get("normalized_subject", "")],
    )
    discovered_cids = [m["conversation_id"] for m in found if m.get("conversation_id")]
    if discovered_cids:
        db.add_conversation_ids(thread_id, discovered_cids)
    return db.upsert_messages(thread_id, found)


def refresh_meeting_threads(db: TrackedThreadsDatabase, meeting_id: str) -> Dict[str, int]:
    """Refreshes all tracked threads for a meeting. Returns {thread_id: new_count} for threads with new messages."""
    results: Dict[str, int] = {}
    for thread in db.get_threads_for_meeting(meeting_id):
        tid = thread["thread_id"]
        new_count = refresh_thread(db, tid)
        if new_count > 0:
            results[tid] = new_count
    return results


def list_recent_threads_in_folders(wg: str, max_items_per_folder: int = 300) -> List[Dict[str, Any]]:
    """Lists recent unique conversation threads across configured Outlook folders for a WG."""
    from modules.emails.core.outlook_client import OutlookClient

    folders = _extract_folder_paths(load_wg_email_config(wg)) if wg else []
    threads_by_key: Dict[str, Dict[str, Any]] = {}

    for fpath in folders:
        folder = OutlookClient.get_folder_by_path(fpath)
        if not folder:
            continue
        try:
            items = folder.Items
            items.Sort("[ReceivedTime]", True)
            total = min(len(items), max_items_per_folder)
            for i in range(1, total + 1):
                try:
                    item = items.Item(i)
                    if getattr(item, "Class", 43) != 43:
                        continue
                    subj = getattr(item, "Subject", "") or ""
                    norm = normalize_subject(subj)
                    if not norm:
                        continue
                    cid = getattr(item, "ConversationID", "") or ""
                    key = norm.lower()
                    dt = getattr(item, "ReceivedTime", None) or getattr(item, "SentOn", None)
                    date_str = dt.strftime("%Y-%m-%d %H:%M") if hasattr(dt, "strftime") else str(dt or "")
                    sender = getattr(item, "SenderName", "") or ""
                    if key not in threads_by_key:
                        threads_by_key[key] = {
                            "title": norm,
                            "normalized_subject": norm,
                            "conversation_id": cid,
                            "last_date": date_str,
                            "last_sender": sender,
                            "count": 1,
                            "folder": fpath,
                        }
                    else:
                        threads_by_key[key]["count"] += 1
                except Exception:
                    continue
        except Exception as e:
            logger.warning(f"Failed listing threads in '{fpath}': {e}")

    return sorted(threads_by_key.values(), key=lambda x: x.get("last_date", ""), reverse=True)


def pin_thread_by_subject_or_cid(
    db: TrackedThreadsDatabase,
    title: str,
    normalized_subject: str,
    conversation_id: str,
    meeting_id: str,
    wg: str,
) -> str:
    """Pins a thread selected from the configured Outlook folder browser and pulls all matching messages."""
    folders = _extract_folder_paths(load_wg_email_config(wg)) if wg else []
    suggested = auto_suggest_links(title)
    thread_id = db.create_or_get_thread(
        title=title,
        normalized_subject=normalized_subject,
        conversation_id=conversation_id,
        outlook_folders=folders,
        meeting_id=meeting_id,
        wg=wg,
        linked_ais=suggested["ais"],
        linked_wis=suggested["wis"],
        linked_tdocs=suggested["tdocs"],
    )
    refresh_thread(db, thread_id)
    return thread_id

