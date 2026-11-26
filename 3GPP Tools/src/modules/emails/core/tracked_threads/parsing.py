"""Subject/token parsing, MAPI MailItem normalization, .msg parsing, and folder helpers."""

from __future__ import annotations

import hashlib
import logging
import re
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence

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
_TDOC_RE = re.compile(
    r"(?<![A-Za-z0-9])(?:S[1-6]|R[1-5]|C[1-6]|SP|RP|CP)-\d{6,8}(?:(?:r|rev)\d{1,2}[a-z]?)?(?![A-Za-z0-9])",
    re.IGNORECASE,
)
_KI_AI_RE = re.compile(
    r"(?<![A-Za-z0-9])(?:KI\s*#?\s*\d+(?:[._]\d+)?|AI\s*#?\s*\d+(?:\.\d+)+)(?![A-Za-z0-9])",
    re.IGNORECASE,
)
_WI_RE = re.compile(
    r"(?<![A-Za-z0-9])(?:FS|TEI\d+|e?[A-Z][A-Za-z0-9]{1,})_[A-Za-z0-9][A-Za-z0-9_]{1,}(?![A-Za-z0-9])"
)


def _extract_wi_tokens(text: str) -> List[str]:
    """Extracts 3GPP Work Item / Study Item tokens (e.g. FS_enIMS_ARC, FS_6G_ARC, eV2X_Ph2)."""
    out: List[str] = []
    for raw in _WI_RE.findall(text or ""):
        tok = raw.strip()
        if not tok:
            continue
        up = tok.upper()
        if up.startswith(("AI_", "KI_")) or re.match(r"^(?:S[1-6]|R[1-5]|C[1-6]|SP|RP|CP)-", up):
            continue
        letters = [c for c in tok if c.isalpha()]
        upper_cnt = sum(1 for c in letters if c.isupper())
        is_wi_shape = up.startswith(("FS_", "TEI")) or (len(letters) >= 3 and (upper_cnt / len(letters)) >= 0.6)
        if is_wi_shape and tok not in out:
            out.append(tok)
    return out


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
        # Direct AI number match (e.g. "AI20.6.1" or "AI#20.10.1" -> "20.6.1" / "20.10.1")
        num_part = re.sub(r"^AI\s*#?\s*", "", t_clean, flags=re.IGNORECASE).strip()
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


def infer_wis_for_ais(
    ais: Sequence[str],
    *,
    agenda_map: Optional[Dict[str, Any]] = None,
    tdoc_meta_map: Optional[Dict[str, Dict[str, Any]]] = None,
    valid_wis: Optional[List[str]] = None,
) -> List[str]:
    """
    Infers Work Item (WI) tokens associated with one or more Agenda Items using:
    1. Explicit WI tokens mentioned inside `agenda_map[ai].description`
    2. Unambiguous WI(s) from `tdoc_meta_map` entries under that AI (when 1..3 distinct WIs exist)
    """
    clean_ais = [str(a).strip() for a in (ais or ()) if str(a).strip()]
    if not clean_ais:
        return []
    wi_by_lower = (
        {str(w).strip().lower(): str(w).strip() for w in valid_wis if str(w).strip()}
        if valid_wis is not None
        else None
    )
    inferred: List[str] = []

    def _add_wi(raw_w: str):
        wc = str(raw_w or "").strip()
        if not wc or wc.lower() in ("(no wi)", "none", "nan", "-"):
            return
        if wi_by_lower is not None:
            if wc.lower() not in wi_by_lower:
                return
            wc = wi_by_lower[wc.lower()]
        if wc not in inferred:
            inferred.append(wc)

    for ai in clean_ais:
        # 1. Check Agenda Item description in agenda_map
        if agenda_map:
            ag_item = agenda_map.get(ai)
            desc = str(getattr(ag_item, "description", "") or "") if ag_item is not None else ""
            if desc:
                for tok in _extract_wi_tokens(desc):
                    _add_wi(tok)
                if wi_by_lower:
                    desc_low = desc.lower()
                    for w_low, w_orig in wi_by_lower.items():
                        if len(w_low) >= 4 and re.search(rf"(?<![A-Za-z0-9_]){re.escape(w_low)}(?![A-Za-z0-9_])", desc_low):
                            _add_wi(w_orig)

        # 2. Check TDocs under this AI in tdoc_meta_map (skip catch-all AIs with > 3 distinct WIs)
        if tdoc_meta_map:
            ai_wis: List[str] = []
            for meta in tdoc_meta_map.values():
                if not isinstance(meta, dict):
                    continue
                if str(meta.get("ai") or "").strip().lower() != ai.lower():
                    continue
                for w in (meta.get("wis") or []):
                    wc = str(w or "").strip()
                    if not wc or wc.lower() in ("(no wi)", "none", "nan", "-"):
                        continue
                    if wi_by_lower is not None:
                        wc = wi_by_lower.get(wc.lower(), wc)
                    if wc not in ai_wis:
                        ai_wis.append(wc)
            if 1 <= len(ai_wis) <= 3:
                for w in ai_wis:
                    _add_wi(w)

    return sorted(set(inferred))


def auto_suggest_links(
    subject: str,
    body: str = "",
    *,
    agenda_map: Optional[Dict[str, Any]] = None,
    valid_ais: Optional[List[str]] = None,
    valid_wis: Optional[List[str]] = None,
    tdoc_meta_map: Optional[Dict[str, Dict[str, Any]]] = None,
    extra_tdocs: Optional[List[str]] = None,
    attachments: Optional[List[str]] = None,
) -> Dict[str, List[str]]:
    """
    Extracts candidate AIs/KIs, Sub-Topics, WIs, and TDocs from an email subject, body, attachments, and linked TDocs.
    When meeting metadata (`agenda_map`/`valid_ais`/`valid_wis`/`tdoc_meta_map`) is provided,
    tokens are resolved against real meeting Agenda Items, Sub-Items, and WIs (including AI -> WI inference).
    """
    text_subj = subject or ""
    att_text = "\n".join(str(a or "") for a in (attachments or []) if str(a or "").strip())
    text_all = f"{text_subj}\n{att_text}\n{(body or '')[:2000]}"
    raw_ais = sorted({re.sub(r"\s+", "", m.upper()) for m in _KI_AI_RE.findall(text_subj)})
    raw_wis = _extract_wi_tokens(text_subj)
    primary_tdocs = (
        {m.upper() for m in _TDOC_RE.findall(f"{text_subj}\n{att_text}")}
        | {str(t).strip().upper() for t in (extra_tdocs or []) if str(t).strip()}
    )
    tdocs = sorted(
        {m.upper() for m in _TDOC_RE.findall(text_all)}
        | primary_tdocs
    )
    body_only_tdocs = sorted(set(tdocs) - primary_tdocs)

    if agenda_map is not None or valid_ais is not None:
        ais = _resolve_ais_against_agenda(raw_ais, raw_wis, agenda_map, valid_ais)
    else:
        ais = [
            re.sub(r"^AI\s*#?\s*", "", a, flags=re.IGNORECASE).strip()
            if re.match(r"^AI\s*#?\s*\d+", a, flags=re.IGNORECASE)
            else a
            for a in raw_ais
        ]

    if valid_wis is not None:
        wi_by_lower = {str(w).strip().lower(): str(w).strip() for w in valid_wis if str(w).strip()}
        wis = sorted({wi_by_lower[w.lower()] for w in raw_wis if w.lower() in wi_by_lower})
    else:
        wis = list(raw_wis)

    sub_items: List[str] = []
    if tdoc_meta_map:
        wi_by_lower = (
            {str(w).strip().lower(): str(w).strip() for w in valid_wis if str(w).strip()}
            if valid_wis is not None
            else None
        )

        def _collect_meta_for(td_list: Sequence[str]):
            found_ais: List[str] = []
            found_subs: List[str] = []
            found_wis: List[str] = []
            for td in td_list:
                base_td = re.sub(r"(?:R|REV)\d{1,2}[A-Z]?$", "", td, flags=re.IGNORECASE)
                meta = tdoc_meta_map.get(td) or tdoc_meta_map.get(base_td)
                if not isinstance(meta, dict):
                    continue
                ai_val = str(meta.get("ai") or "").strip()
                if ai_val and ai_val not in ("N/A", "-", "None") and ai_val not in found_ais:
                    found_ais.append(ai_val)
                sub_val = str(meta.get("sub_item") or "").strip()
                if sub_val and sub_val not in ("N/A", "-", "None") and sub_val not in found_subs:
                    found_subs.append(sub_val)
                for w in (meta.get("wis") or []):
                    wc = str(w or "").strip()
                    if not wc or wc.lower() in ("(no wi)", "none", "nan", "-"):
                        continue
                    if wi_by_lower is not None:
                        wc = wi_by_lower.get(wc.lower(), wc)
                    if wc not in found_wis:
                        found_wis.append(wc)
            return found_ais, found_subs, found_wis

        p_ais, p_subs, p_wis = _collect_meta_for(sorted(primary_tdocs))
        b_ais, b_subs, b_wis = _collect_meta_for(body_only_tdocs)
        subject_ais_set = set(ais)

        if primary_tdocs:
            for ai_val in p_ais:
                if (not subject_ais_set or ai_val in subject_ais_set) and ai_val not in ais:
                    ais.append(ai_val)
            if len(p_subs) <= 2:
                for sub_val in p_subs:
                    if sub_val not in sub_items:
                        sub_items.append(sub_val)
            for w in p_wis:
                if w not in wis:
                    wis.append(w)
        elif not subject_ais_set:
            # Body-only TDocs without subject AI/TDoc: only inherit AI/Sub-Topic if unanimous
            if len(b_ais) == 1 and b_ais[0] not in ais:
                ais.append(b_ais[0])
                if len(b_subs) == 1:
                    sub_items.append(b_subs[0])
            if 1 <= len(b_wis) <= 2:
                for w in b_wis:
                    if w not in wis:
                        wis.append(w)
        else:
            # Subject AI already known: never flood unrelated AIs from body quotes/tables
            matching_subs: List[str] = []
            for td in body_only_tdocs:
                base_td = re.sub(r"(?:R|REV)\d{1,2}[A-Z]?$", "", td, flags=re.IGNORECASE)
                meta = tdoc_meta_map.get(td) or tdoc_meta_map.get(base_td)
                if not isinstance(meta, dict):
                    continue
                if str(meta.get("ai") or "").strip() in subject_ais_set:
                    sub_val = str(meta.get("sub_item") or "").strip()
                    if sub_val and sub_val not in ("N/A", "-", "None") and sub_val not in matching_subs:
                        matching_subs.append(sub_val)
            if len(matching_subs) == 1:
                sub_items.append(matching_subs[0])
            if 1 <= len(b_wis) <= 2:
                for w in b_wis:
                    if w not in wis:
                        wis.append(w)

    for w in infer_wis_for_ais(ais, agenda_map=agenda_map, tdoc_meta_map=tdoc_meta_map, valid_wis=valid_wis):
        if w not in wis:
            wis.append(w)

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


def _extract_folder_sub_flags(raw_folders: Optional[List[Any]]) -> Dict[str, bool]:
    """Maps lowercase folder_path -> include_subfolders (defaults to True)."""
    flags: Dict[str, bool] = {}
    for entry in (raw_folders or []):
        if isinstance(entry, dict):
            fp = str(entry.get("folder_path", "") or "").strip().lower()
            if fp:
                flags[fp] = bool(entry.get("include_subfolders", True))
        else:
            fp = str(entry or "").strip().lower()
            if fp and fp not in flags:
                flags[fp] = True
    return flags


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
