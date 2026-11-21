"""Interactive Outlook COM actions: open, reload, attach, reply, read/flag mutations."""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional

from modules.emails.core.tracked_threads_db import TrackedThreadsDatabase
from .parsing import _mail_item_to_dict

logger = logging.getLogger(__name__)


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
