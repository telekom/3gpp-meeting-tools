# --- File: src/modules/emails/core/markdown_export/url_and_links.py ---
"""Safelinks/Proofpoint unwrapping, Drafts/Docs URL extraction, and Markdown linkification."""

import re
import urllib.parse
from typing import Any, Dict, List

from .formatting_helpers import strip_trailing_quotes

_SAFELINK_RE = re.compile(
    r"<?https?://[a-z0-9.-]*safelinks\.protection\.outlook\.com/\?[^\s<>\"\)\]]+>?",
    re.IGNORECASE,
)
_URLDEFENSE_RE = re.compile(
    r"<?https?://urldefense(?:\.proofpoint)?\.com/v3/__(.+?)__;[^\s<>\"\)\]]+>?",
    re.IGNORECASE,
)
_FIREEYE_RE = re.compile(
    r"<?https?://protect\d*\.fireeye\.com/v1/url\?[^\s<>\"\)\]]+>?",
    re.IGNORECASE,
)
_CLIENT_PSEUDO_ANGLE_URL_RE = re.compile(
    r"[ \t]*<(?:applewebdata|x-apple-[a-z0-9-]+|cid):[^>\n]+>",
    re.IGNORECASE,
)
_BARE_APPLE_URL_RE = re.compile(
    r"\b(?:applewebdata|x-apple-[a-z0-9-]+)://[^\s<>\"\)\]]+",
    re.IGNORECASE,
)
# Matches http(s)://10.10.10.10/... or 10.10.10.10/ftp/... or https://www.3gpp.org/ftp/... that may contain unencoded spaces, '#', '[' or ']' up to a document extension
_DOC_URL_WITH_SPACES_RE = re.compile(
    r"((?:https?://)?(?:10\.10\.10\.10/ftp|(?:www\.)?3gpp\.org/ftp)/(?:(?!\]\()[^\r\n<>\"\)])+?\.(?:docx|docm|doc|zip|pptx|pdf|xlsx|vsdx))\b",
    re.IGNORECASE,
)
_BARE_LOCAL_IP_URL_RE = re.compile(
    r"(?<![/\w])(10\.10\.10\.10/ftp/[^\s<>\"\)\]]+)",
    re.IGNORECASE,
)
_BARE_URL_RE = re.compile(
    r"(?<![\(\[<\"])(https?://[^\s<>\"\)\]]+)",
    re.IGNORECASE,
)
_DOC_EXT_RE = re.compile(r"\.(?:docx|docm|doc|zip|pptx|pdf|xlsx|vsdx)(?:$|\?)", re.IGNORECASE)
_TDOC_LINK_OR_EXISTING_MD_RE = re.compile(
    r"(\[[^\]]*\]\([^\)]*\))|(?<![A-Za-z0-9])((?:S[1-6]|R[1-5]|C[1-6]|SP|RP|CP)-\d{6,8}(?:r\d{1,2}[a-zA-Z]?)?)(?![A-Za-z0-9])",
    re.IGNORECASE,
)


def _normalize_doc_url_spaces(text: str) -> str:
    """Normalizes unencoded spaces, '#', '[' and ']' in 10.10.10.10/ftp and 3gpp.org/ftp document URLs."""
    if not text:
        return ""

    def _fix_doc_url(m: re.Match) -> str:
        raw = m.group(1).strip()
        if not raw.lower().startswith(("http://", "https://")):
            raw = "http://" + raw
        raw = (
            raw.replace(" ", "%20")
            .replace("#", "%23")
            .replace("[", "%5B")
            .replace("]", "%5D")
        )
        return raw

    out = _DOC_URL_WITH_SPACES_RE.sub(_fix_doc_url, text)
    out = _BARE_LOCAL_IP_URL_RE.sub(lambda m: "http://" + m.group(1), out)
    return out


def _decode_query_param_url(target: str) -> str:
    """Unquotes a URL extracted via parse_qs only if it remained doubly percent-encoded."""
    if target and target.lower().startswith(("http%3a", "https%3a")):
        return urllib.parse.unquote(target)
    return target


def unwrap_outlook_safelinks(text: str) -> str:
    """
    Replaces Microsoft Outlook Safelinks, Proofpoint URLDefense, and FireEye URL wrappers
    with the decoded original URL, and normalizes 10.10.10.10 / 3GPP FTP document links.
    """
    if not text:
        return ""

    cleaned = _CLIENT_PSEUDO_ANGLE_URL_RE.sub("", text)
    cleaned = _BARE_APPLE_URL_RE.sub("", cleaned)

    if "safelinks.protection.outlook.com" in cleaned.lower():
        def _replace_safelink(match: re.Match) -> str:
            full = match.group(0)
            trailing = ""
            while full and full[-1] in ".,;:!?" and not full.endswith(">"):
                trailing = full[-1] + trailing
                full = full[:-1]
            raw = full.strip("<>")
            try:
                parsed = urllib.parse.urlsplit(raw)
                qs = urllib.parse.parse_qs(parsed.query)
                target = (qs.get("url") or [""])[0]
                if target:
                    return _decode_query_param_url(target) + trailing
            except Exception:
                pass
            return raw + trailing

        cleaned = _SAFELINK_RE.sub(_replace_safelink, cleaned)

    if "urldefense" in cleaned.lower():
        def _replace_urldefense(match: re.Match) -> str:
            inner = match.group(1)
            # Proofpoint v3 encodes % as * (e.g. *20, *23) and sometimes collapses https:// to https:/
            inner = re.sub(r"\*([0-9A-Fa-f]{2})", r"%\1", inner)
            inner = re.sub(r"^(https?:)/([^/])", r"\1//\2", inner, flags=re.IGNORECASE)
            return inner

        cleaned = _URLDEFENSE_RE.sub(_replace_urldefense, cleaned)

    if "fireeye.com" in cleaned.lower():
        def _replace_fireeye(match: re.Match) -> str:
            raw = match.group(0).strip("<>")
            try:
                parsed = urllib.parse.urlsplit(raw)
                qs = urllib.parse.parse_qs(parsed.query)
                target = (qs.get("u") or [""])[0]
                if target:
                    return _decode_query_param_url(target)
            except Exception:
                pass
            return raw

        cleaned = _FIREEYE_RE.sub(_replace_fireeye, cleaned)

    cleaned = _normalize_doc_url_spaces(cleaned)
    # Collapse Outlook's "https://example.com https://example.com" or "https://example.com<https://example.com>"
    cleaned = re.sub(r"(https?://[^\s<>\"\)\]]+)\s*<+\1>+", r"\1", cleaned)
    cleaned = re.sub(r"(https?://[^\s<>\"\)\]]+)(?:[ \t]*\1)+", r"\1", cleaned)
    return cleaned


def is_cacheable_document_url(url: str) -> bool:
    """Returns True if url points to a downloadable 3GPP / 10.10.10.10 document file."""
    u = (url or "").strip()
    if not u.lower().startswith(("http://", "https://")):
        return False
    path = urllib.parse.urlsplit(u).path
    return bool(_DOC_EXT_RE.search(path))


def is_drafts_document_url(url: str) -> bool:
    """
    Returns True if url points to a downloadable document in 10.10.10.10's Drafts folder
    (or its 3gpp.org/ftp Drafts mirror), excluding /Revisions/ and /Docs/.
    """
    if not is_cacheable_document_url(url):
        return False
    parsed = urllib.parse.urlsplit(url.strip())
    host = (parsed.netloc or "").lower()
    if not (host == "10.10.10.10" or host.endswith("3gpp.org")):
        return False
    segs = [urllib.parse.unquote(s).strip().lower() for s in parsed.path.split("/") if s.strip()]
    if "drafts" not in segs:
        return False
    if "revisions" in segs or "docs" in segs:
        return False
    return True


def _format_short_url_label(url: str) -> str:
    """Returns a concise human-readable label for long URLs and 3GPP/10.10.10.10 document links."""
    u = str(url or "").strip()
    if is_cacheable_document_url(u):
        raw_name = urllib.parse.urlsplit(u).path.rsplit("/", 1)[-1]
        fname = urllib.parse.unquote(raw_name).strip() or u
        fname = fname.replace("[", "(").replace("]", ")")
        if len(fname) > 95:
            stem, dot, ext = fname.rpartition(".")
            if dot and len(ext) <= 8:
                keep = max(20, 95 - len(ext) - 2)
                fname = f"{stem[:keep]}…{dot}{ext}"
            else:
                fname = fname[:94] + "…"
        return f"📄 {fname}"
    if len(u) <= 72:
        return u
    parsed = urllib.parse.urlsplit(u)
    netloc = parsed.netloc or "link"
    tail = urllib.parse.unquote(parsed.path.rstrip("/").rsplit("/", 1)[-1] or parsed.query or "")
    tail = tail.replace("[", "(").replace("]", ")")
    if tail:
        if len(tail) > 32:
            tail = tail[:31] + "…"
        return f"🔗 {netloc}/…/{tail}"
    return f"🔗 {netloc}/…"


def _extract_matching_document_urls(
    messages: List[Dict[str, Any]],
    predicate: Any,
    *,
    prefer_clean_body: bool = False,
) -> List[str]:
    found: List[str] = []
    seen = set()
    for msg in (messages or []):
        if prefer_clean_body:
            raw = msg.get("clean_body") or strip_trailing_quotes(msg.get("body_text") or "")
        else:
            raw = msg.get("body_text") or msg.get("clean_body") or ""
        body = unwrap_outlook_safelinks(raw)
        for m in _BARE_URL_RE.finditer(body):
            url = m.group(1).rstrip(".,;:!?")
            if predicate(url) and url not in seen:
                seen.add(url)
                found.append(url)
    return found


def extract_thread_document_urls(messages: List[Dict[str, Any]]) -> List[str]:
    """Extracts all unique cacheable document URLs (e.g. 10.10.10.10 or 3gpp.org/ftp .docx/.zip) from thread messages."""
    return _extract_matching_document_urls(messages, is_cacheable_document_url, prefer_clean_body=False)


def extract_drafts_document_urls(messages: List[Dict[str, Any]], prefer_clean_body: bool = True) -> List[str]:
    """Extracts unique 10.10.10.10 / Drafts folder document URLs (excluding /Revisions/ and /Docs/)."""
    return _extract_matching_document_urls(messages, is_drafts_document_url, prefer_clean_body=prefer_clean_body)


def linkify_markdown_urls(text: str, shorten_long_urls: bool = False) -> str:
    """Converts bare http(s):// URLs in text into clickable Markdown links [label](url)."""
    if not text:
        return ""
    unwrapped = unwrap_outlook_safelinks(text)

    def _fmt_link(url: str, trailing: str = "") -> str:
        label = _format_short_url_label(url) if shorten_long_urls else url
        return f"[{label}]({url}){trailing}"

    unwrapped = re.sub(
        r"<(https?://[^\s<>\"\)\]]+)>",
        lambda m: _fmt_link(m.group(1)),
        unwrapped,
    )

    def _linkify_match(m: re.Match) -> str:
        url = m.group(1)
        trailing = ""
        while url and url[-1] in ".,;:!?":
            trailing = url[-1] + trailing
            url = url[:-1]
        if not url:
            return m.group(0)
        return _fmt_link(url, trailing)

    return _BARE_URL_RE.sub(_linkify_match, unwrapped)


def linkify_tdoc_tokens(text: str) -> str:
    """Turns bare 3GPP TDoc IDs into clickable Markdown tdoc: links while leaving existing Markdown links intact."""
    if not text:
        return ""

    def _repl(m: re.Match) -> str:
        if m.group(1):
            return m.group(1)
        tok = m.group(2)
        return f"[{tok}](tdoc:{tok.upper()})"

    return _TDOC_LINK_OR_EXISTING_MD_RE.sub(_repl, text)
