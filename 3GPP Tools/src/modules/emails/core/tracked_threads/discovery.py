"""Per-meeting discovery config, untracked folder scanning, classification, and subject pinning."""

from __future__ import annotations

import json
import logging
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional

from modules.emails.core.tracked_threads_db import TrackedThreadsDatabase
from .parsing import (
    _TDOC_RE,
    _extract_folder_paths,
    _extract_folder_sub_flags,
    auto_suggest_links,
    call_load_wg_email_config as load_wg_email_config,
    normalize_subject,
)
from .sync import call_refresh_thread as refresh_thread

logger = logging.getLogger(__name__)


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

    raw_cfg = load_wg_email_config(wg) if wg else []
    folders = _extract_folder_paths(raw_cfg)
    sub_flags = _extract_folder_sub_flags(raw_cfg)
    ns = OutlookClient.get_namespace()
    if ns:
        try:
            sent_path = getattr(ns.GetDefaultFolder(5), "FolderPath", "") or ""
            if sent_path and sent_path.lower() not in {p.lower() for p in folders}:
                folders.append(sent_path)
                sub_flags.setdefault(sent_path.lower(), False)
        except Exception:
            pass

    threads_by_key: Dict[str, Dict[str, Any]] = {}
    total_folders = len(folders)
    visited_folder_keys: set = set()
    expanded_folders: List[tuple] = []

    for idx, fpath in enumerate(folders):
        if cancel_check and cancel_check():
            break
        if progress_cb:
            progress_cb(idx, total_folders, fpath, folders[idx + 1 :])
        root_folder = OutlookClient.get_folder_by_path(fpath)
        if not root_folder:
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
