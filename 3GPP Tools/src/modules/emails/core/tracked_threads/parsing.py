"""Subject/token parsing, MAPI MailItem normalization, .msg parsing, and folder helpers."""

from __future__ import annotations

import hashlib
import logging
import re
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional

from modules.emails.core.email_config import (
    load_wg_email_config as _ORIG_LOAD_WG_EMAIL_CONFIG,
)
from modules.emails.core.thread_markdown_exporter import (
    inject_html_tables_into_body,
    strip_trailing_quotes,
    unwrap_outlook_safelinks,
)

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


def call_load_wg_email_config(*args: Any, **kwargs: Any) -> Any:
    """Calls load_wg_email_config, honoring monkeypatches on the legacy tracked_threads_service shim."""
    for mod_name in (
        "modules.emails.core.tracked_threads_service",
        "src.modules.emails.core.tracked_threads_service",
    ):
        shim = sys.modules.get(mod_name)
        if shim is not None:
            shim_fn = getattr(shim, "load_wg_email_config", None)
            if shim_fn is not None and shim_fn is not _ORIG_LOAD_WG_EMAIL_CONFIG:
                return shim_fn(*args, **kwargs)
    return _ORIG_LOAD_WG_EMAIL_CONFIG(*args, **kwargs)
