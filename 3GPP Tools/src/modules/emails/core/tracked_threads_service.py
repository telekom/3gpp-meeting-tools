# --- File: src/modules/emails/core/tracked_threads_service.py ---
import logging
import re
from pathlib import Path
from typing import Any, Dict, List, Optional

from modules.emails.core.email_config import load_wg_email_config
from modules.emails.core.tracked_threads_db import TrackedThreadsDatabase

logger = logging.getLogger(__name__)

_PREFIX_RE = re.compile(r"^(?:\s*(?:re|fwd|fw|aw|wg|sv|antw)\s*:\s*)+", re.IGNORECASE)
_TAG_RE = re.compile(r"^\s*\[[^\]]{2,30}\]\s*")


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


def _extract_company(sender_name: str, sender_email: str) -> str:
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
    sender_email = getattr(item, "SenderEmailAddress", "") or ""
    body_text = getattr(item, "Body", "") or ""
    conv_id = getattr(item, "ConversationID", "") or ""
    entry_id = getattr(item, "EntryID", "") or f"MSG-{abs(hash((subject, sender_name, body_text[:100])))}"

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


def _collect_matching_from_folders(
    folder_paths: List[str],
    conversation_ids: List[str],
    subject_patterns: List[str],
) -> List[Dict[str, Any]]:
    """Scans configured Outlook folders for messages matching any ConversationID or normalized subject."""
    from modules.emails.core.outlook_client import OutlookClient

    matches: List[Dict[str, Any]] = []
    seen_ids = set()
    norm_patterns = {p.lower() for p in subject_patterns if p}
    conv_set = {c for c in conversation_ids if c}

    for fpath in folder_paths:
        folder = OutlookClient.get_folder_by_path(fpath)
        if not folder:
            continue
        try:
            items = folder.Items
            items.Sort("[ReceivedTime]", True)
            for item in items:
                try:
                    subj = getattr(item, "Subject", "") or ""
                    cid = getattr(item, "ConversationID", "") or ""
                    norm = normalize_subject(subj).lower()
                    if (cid and cid in conv_set) or (norm and norm in norm_patterns):
                        d = _mail_item_to_dict(item)
                        if d["id"] not in seen_ids:
                            seen_ids.add(d["id"])
                            matches.append(d)
                except Exception:
                    continue
        except Exception as e:
            logger.warning(f"Error scanning Outlook folder '{fpath}': {e}")
    return matches


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

    norm_subj = normalize_subject(seed["subject"])
    folders = load_wg_email_config(wg) if wg else []
    thread_id = db.create_or_get_thread(
        title=norm_subj or seed["subject"],
        normalized_subject=norm_subj,
        conversation_id=seed.get("conversation_id", ""),
        outlook_folders=folders,
        meeting_id=meeting_id,
        wg=wg,
        linked_ais=linked_ais,
        linked_wis=linked_wis,
        linked_tdocs=linked_tdocs,
    )

    conv_ids = [seed["conversation_id"]] if seed.get("conversation_id") else []
    patterns = [norm_subj] if norm_subj else []
    found = _collect_matching_from_folders(folders, conv_ids, patterns) if folders else []
    if not any(m["id"] == seed["id"] for m in found):
        found.append(seed)

    db.upsert_messages(thread_id, found)
    return thread_id


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
