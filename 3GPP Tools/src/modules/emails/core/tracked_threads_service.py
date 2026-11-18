# --- File: src/modules/emails/core/tracked_threads_service.py ---
import logging
import re
from pathlib import Path
from typing import Any, Dict, List, Optional

from modules.emails.core.email_config import load_wg_email_config
from modules.emails.core.thread_markdown_exporter import (
    extract_thread_document_urls,
    unwrap_outlook_safelinks,
)
from modules.emails.core.tracked_threads_db import TrackedThreadsDatabase

logger = logging.getLogger(__name__)

_PREFIX_RE = re.compile(r"^(?:\s*(?:re|fwd|fw|aw|wg|sv|antw|回复|答复|转发)\s*[:：]\s*)+", re.IGNORECASE)
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


def open_outlook_attachment(entry_id: str, filename: str, save_dir: Optional[Path] = None) -> Optional[Path]:
    """Saves an attachment from an Outlook MailItem to disk and opens it via os.startfile."""
    import os

    if not entry_id or not filename or str(entry_id).startswith("MSG-"):
        return None
    from modules.emails.core.outlook_client import OutlookClient

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
                if hasattr(os, "startfile"):
                    os.startfile(str(out_path.resolve()))
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


def cache_and_open_thread_document_url(
    url: str,
    thread_dir: Path,
    open_after: bool = True,
) -> Optional[Path]:
    """
    Downloads and caches a document linked in an email thread (e.g. 10.10.10.10 or 3gpp.org/ftp)
    into thread_dir (<meeting_dir>/Agenda/ThreadAttachments/<thread_id>/) and optionally opens it.
    """
    import os
    import urllib.parse
    import zipfile
    from core.network.session import NetworkSession

    if not url:
        return None
    thread_dir = Path(thread_dir)
    thread_dir.mkdir(parents=True, exist_ok=True)

    raw_name = urllib.parse.unquote(urllib.parse.urlsplit(url).path.rsplit("/", 1)[-1]) or "document.docx"
    safe_name = "".join(c for c in raw_name if c not in '<>:"/\\|?*').strip() or "document.docx"
    target_path = thread_dir / safe_name

    if not (target_path.exists() and target_path.stat().st_size > 0):
        session = NetworkSession()
        downloaded = False
        for cand_url in _build_candidate_document_urls(url):
            try:
                resp = session.get(cand_url, timeout=12)
                if resp.status_code == 200 and resp.content:
                    target_path.write_bytes(resp.content)
                    downloaded = True
                    break
            except Exception as e:
                logger.debug(f"Candidate URL download failed ({cand_url}): {e}")
        if not downloaded:
            return None

    open_target = target_path
    if target_path.suffix.lower() == ".zip" and zipfile.is_zipfile(target_path):
        try:
            with zipfile.ZipFile(target_path, "r") as zf:
                for member in zf.namelist():
                    if member.lower().endswith((".docx", ".doc", ".pptx", ".pdf", ".xlsx")) and not Path(member).name.startswith("~$"):
                        extracted = thread_dir / Path(member).name
                        if not extracted.exists():
                            extracted.write_bytes(zf.read(member))
                        open_target = extracted
                        break
        except Exception:
            pass

    if open_after and hasattr(os, "startfile"):
        try:
            os.startfile(str(open_target.resolve()))
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



