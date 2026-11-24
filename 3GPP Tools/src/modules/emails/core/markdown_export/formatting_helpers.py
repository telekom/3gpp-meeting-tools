# --- File: src/modules/emails/core/markdown_export/formatting_helpers.py ---
"""Sender, timestamp, attachment, subject normalization, and localized quote-stripping helpers."""

import re
from typing import Any, List, Optional

_TDOC_RE = re.compile(r"(?<![A-Za-z0-9])(?:S[1-6]|R[1-5]|C[1-6]|SP|RP|CP)-\d{6,8}(?![A-Za-z0-9])", re.IGNORECASE)
_TDOC_URL_PREFIX_ARTIFACT_RE = re.compile(
    r"^(?:2F|20|2520)+(?=(?:S[1-6]|R[1-5]|C[1-6]|SP|RP|CP)-\d{6,8}\b)",
    re.IGNORECASE,
)
_INLINE_SIG_ATT_RE = re.compile(
    r"^(?:image\d{3,4}\.(?:png|jpe?g|gif|bmp)|oledata\.mso|winmail\.dat)$",
    re.IGNORECASE,
)
_SUBJ_PREFIX_RE = re.compile(r"^(?:(?:re|aw|fwd?|wg|回覆|回复|答复|转发)\s*:\s*)+", re.IGNORECASE)
_SUBJ_TAG_RE = re.compile(r"^(?:\[[^\]]{1,80}\]\s*)+")
_QUOTE_SEP_LINE_RE = re.compile(
    r"^(?:-{3,}\s*(?:Original Message|原始邮件|Ursprüngliche Nachricht|Message d'origine|Mensaje original)\s*-{3,}|_{5,})$",
    re.IGNORECASE,
)
_QUOTE_FROM_HEADER_RE = re.compile(
    r"^(?:From|Von|De|Van|Da|От|发件人|寄件者|差出人|보낸\s*사람)\s*[:：]\s*.+$",
    re.IGNORECASE,
)
_QUOTE_COMPANION_HEADER_RE = re.compile(
    r"(?:^|\s)(?:Sent|Date|Gesendet|Datum|Envoyé|To|An|À|Subject|Betreff|Objet|Cc|发送时间|收件人|主题|抄送|傳送時間|收件者|主旨|送信日時|宛先|件名|보낸\s*날짜|받는\s*사람|제목)\s*[:：]",
    re.IGNORECASE,
)
_QUOTE_WROTE_LINE_RE = re.compile(
    r"^(?:On\s+.+?\bwrote|Am\s+.+?\bschrieb\b.+|Le\s+.+?\ba écrit\b.+|在\s*.+?写道)\s*[:：]\s*$",
    re.IGNORECASE,
)


def clean_tdoc_id(raw: Any) -> str:
    """Normalizes a TDoc token and strips URL-encoding prefix artifacts like 2F or 20."""
    td = str(raw or "").strip().upper()
    if not td:
        return ""
    return _TDOC_URL_PREFIX_ARTIFACT_RE.sub("", td)


def split_sender_company(sender: str, company: str = "") -> tuple:
    """Splits sender into (display_name, company_badge) without duplicating '(Company)'."""
    s = str(sender or "").strip() or "Unknown"
    c = str(company or "").strip()
    if c:
        cleaned = re.sub(rf"\s*\({re.escape(c)}\)\s*$", "", s, flags=re.IGNORECASE).strip()
        if cleaned:
            return cleaned, c
        if c.lower() in s.lower():
            return s, ""
        return s, c
    m = re.match(r"^(.*?)\s*\(([^()]{1,48})\)\s*$", s)
    if m and m.group(1).strip():
        return m.group(1).strip(), m.group(2).strip()
    return s, ""


def format_sender_display(sender: str, company: str = "") -> str:
    """Formats 'Sender (Company)' without duplicating '(Company)' when already inside sender."""
    s = str(sender or "").strip() or "Unknown"
    c = str(company or "").strip()
    if not c or c.lower() in s.lower():
        return s
    return f"{s} ({c})"


def format_compact_timestamp(raw_date: Any) -> str:
    """Formats timestamps cleanly as YYYY-MM-DD HH:MM, dropping seconds/microseconds/offsets."""
    text = str(raw_date or "").strip()
    if not text:
        return "Unknown Date"
    m = re.match(r"^(\d{4}-\d{2}-\d{2})[T ](\d{2}:\d{2})(?::\d{2}(?:\.\d+)?)?(?:Z|[+-]\d{2}:?\d{2})?$", text)
    if m:
        return f"{m.group(1)} {m.group(2)}"
    return text


def _normalize_subject_for_comparison(subject: str) -> str:
    """Strips Re:/AW:/[External]/[Meeting Tag] prefixes so true subject forks can be detected."""
    s = " ".join(str(subject or "").split()).strip()
    while s:
        s2 = _SUBJ_PREFIX_RE.sub("", s).strip()
        s2 = _SUBJ_TAG_RE.sub("", s2).strip()
        if s2 == s:
            break
        s = s2
    return s.lower()


def partition_email_attachments(attachments: List[str]) -> tuple:
    """Splits attachments into (real_attachments, inline_signature_images)."""
    real_atts: List[str] = []
    inline_imgs: List[str] = []
    for raw in (attachments or []):
        name = str(raw or "").strip()
        if not name:
            continue
        if _INLINE_SIG_ATT_RE.match(name):
            inline_imgs.append(name)
        else:
            real_atts.append(name)
    return real_atts, inline_imgs


def find_trailing_quote_line_index(lines: List[str], *, require_preceding_text: bool = False) -> Optional[int]:
    """Finds the starting line index of a trailing localized Outlook/mail reply header block."""
    if not lines:
        return None
    for idx, line in enumerate(lines):
        stripped = line.replace("\xa0", " ").strip()
        if not stripped:
            continue
        is_quote_start = False
        if _QUOTE_SEP_LINE_RE.match(stripped) or _QUOTE_WROTE_LINE_RE.match(stripped):
            is_quote_start = True
        elif _QUOTE_FROM_HEADER_RE.match(stripped):
            window = "\n".join(l.replace("\xa0", " ") for l in lines[idx : min(len(lines), idx + 6)])
            if _QUOTE_COMPANION_HEADER_RE.search(window):
                is_quote_start = True
        if is_quote_start:
            if require_preceding_text and not any(l.strip() for l in lines[:idx]):
                continue
            return idx
    return None


def split_direct_and_quoted_body(body_text: str) -> tuple:
    """Splits an email body into (direct_body, quoted_body) at the first localized trailing reply block."""
    if not body_text:
        return "", ""
    lines = body_text.splitlines()
    idx = find_trailing_quote_line_index(lines, require_preceding_text=False)
    if idx is None:
        return body_text, ""
    return "\n".join(lines[:idx]).strip(), "\n".join(lines[idx:]).strip()


def strip_trailing_quotes(body_text: str) -> str:
    """Strips pure trailing quoted reply chains (including localized headers) while preserving non-empty top text."""
    if not body_text:
        return ""
    lines = body_text.splitlines()
    idx = find_trailing_quote_line_index(lines, require_preceding_text=True)
    if idx is None:
        return body_text.strip()
    return "\n".join(lines[:idx]).strip()


def _format_trailing_quoted_history(body: str) -> str:
    """Converts trailing inline reply headers ('From: ... Sent: ...' / '发件人:') into Markdown blockquotes."""
    if not body:
        return ""
    lines = body.splitlines()
    idx = find_trailing_quote_line_index(lines, require_preceding_text=False)
    if idx is None:
        return body
    head = lines[:idx]
    tail = [
        "> " + ql if ql.strip() and not ql.lstrip().startswith(">") else (">" if not ql.strip() else ql)
        for ql in lines[idx:]
    ]
    return "\n".join(head + ([""] if head and head[-1].strip() else []) + tail)
