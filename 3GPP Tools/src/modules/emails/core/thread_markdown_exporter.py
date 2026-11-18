# --- File: src/modules/emails/core/thread_markdown_exporter.py ---
import re
import urllib.parse
from typing import Any, Dict, List


_TDOC_RE = re.compile(r"\b(?:S[1-6]|R[1-5]|C[1-6]|SP|RP|CP)-\d{6,7}\b", re.IGNORECASE)
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
# Matches http(s)://10.10.10.10/... or 10.10.10.10/ftp/... or https://www.3gpp.org/ftp/... that may contain unencoded spaces up to a document extension
_DOC_URL_WITH_SPACES_RE = re.compile(
    r"((?:https?://)?(?:10\.10\.10\.10/ftp|(?:www\.)?3gpp\.org/ftp)/[^\r\n<>\"\)\]]+?\.(?:docx|docm|doc|zip|pptx|pdf|xlsx|vsdx))\b",
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


def _normalize_doc_url_spaces(text: str) -> str:
    """Normalizes unencoded spaces and '#' in 10.10.10.10/ftp and 3gpp.org/ftp document URLs."""
    if not text:
        return ""

    def _fix_doc_url(m: re.Match) -> str:
        raw = m.group(1).strip()
        if not raw.lower().startswith(("http://", "https://")):
            raw = "http://" + raw
        raw = raw.replace(" ", "%20").replace("#", "%23")
        return raw

    out = _DOC_URL_WITH_SPACES_RE.sub(_fix_doc_url, text)
    out = _BARE_LOCAL_IP_URL_RE.sub(lambda m: "http://" + m.group(1), out)
    return out


def unwrap_outlook_safelinks(text: str) -> str:
    """
    Replaces Microsoft Outlook Safelinks, Proofpoint URLDefense, and FireEye URL wrappers
    with the decoded original URL, and normalizes 10.10.10.10 / 3GPP FTP document links.
    """
    if not text:
        return ""

    cleaned = text

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
                    return urllib.parse.unquote(target) + trailing
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
                    return urllib.parse.unquote(target)
            except Exception:
                pass
            return raw

        cleaned = _FIREEYE_RE.sub(_replace_fireeye, cleaned)

    cleaned = _normalize_doc_url_spaces(cleaned)
    # Collapse Outlook's "https://example.com https://example.com" or "https://example.com<https://example.com>"
    cleaned = re.sub(r"(https?://[^\s<>\"\)\]]+)\s*<+\1>+", r"\1", cleaned)
    cleaned = re.sub(r"(https?://[^\s<>\"\)\]]+)(?:\1)+", r"\1", cleaned)
    return cleaned


def is_cacheable_document_url(url: str) -> bool:
    """Returns True if url points to a downloadable 3GPP / 10.10.10.10 document file."""
    u = (url or "").strip()
    if not u.lower().startswith(("http://", "https://")):
        return False
    path = urllib.parse.urlsplit(u).path
    return bool(_DOC_EXT_RE.search(path))


def extract_thread_document_urls(messages: List[Dict[str, Any]]) -> List[str]:
    """Extracts all unique cacheable document URLs (e.g. 10.10.10.10 or 3gpp.org/ftp .docx/.zip) from thread messages."""
    found: List[str] = []
    seen = set()
    for msg in (messages or []):
        body = unwrap_outlook_safelinks(msg.get("body_text") or msg.get("clean_body") or "")
        for m in _BARE_URL_RE.finditer(body):
            url = m.group(1).rstrip(".,;:!?")
            if is_cacheable_document_url(url) and url not in seen:
                seen.add(url)
                found.append(url)
    return found


def linkify_markdown_urls(text: str) -> str:
    """Converts bare http(s):// URLs in text into clickable Markdown links [url](url)."""
    if not text:
        return ""
    unwrapped = unwrap_outlook_safelinks(text)
    unwrapped = re.sub(
        r"<(https?://[^\s<>\"\)\]]+)>",
        lambda m: f"[{m.group(1)}]({m.group(1)})",
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
        return f"[{url}]({url}){trailing}"

    return _BARE_URL_RE.sub(_linkify_match, unwrapped)


def export_thread_to_markdown(
    thread: Dict[str, Any],
    messages: List[Dict[str, Any]],
    deduplicate_quotes: bool = True,
    interactive_links: bool = False,
) -> str:
    """
    Exports a tracked email thread to a lossless, human- and LLM-readable Markdown document.
    Preserves inline replies, metadata, participants/companies, and TDoc references,
    unwraps Outlook Safelinks/Proofpoint, and formats URLs as clickable Markdown links.
    """
    participants = []
    seen_p = set()
    tdocs_mentioned = set(thread.get("linked_tdocs") or [])

    for msg in messages:
        name = (msg.get("sender_name") or msg.get("sender_email") or "Unknown").strip()
        comp = (msg.get("company") or "").strip()
        label = f"{name} ({comp})" if comp else name
        if label not in seen_p:
            seen_p.add(label)
            participants.append(label)
        for m in _TDOC_RE.findall(f"{msg.get('subject', '')} {msg.get('body_text', '')}"):
            tdocs_mentioned.add(m.upper())

    dates = [m.get("date_sent", "") for m in messages if m.get("date_sent")]
    date_range = f"{min(dates)} to {max(dates)}" if dates else "N/A"
    doc_urls = extract_thread_document_urls(messages)

    lines = [
        f"# Email Thread: {thread.get('title', 'Untitled Thread')}",
        "",
        "## Thread Metadata",
        f"- **Thread ID:** `{thread.get('thread_id', '')}`",
        f"- **Meeting:** `{thread.get('meeting_id', '') or 'N/A'}` ({thread.get('wg', '') or 'N/A'})",
        f"- **Linked AIs:** {', '.join(thread.get('linked_ais') or []) or 'None'}",
        f"- **Linked WIs:** {', '.join(thread.get('linked_wis') or []) or 'None'}",
        f"- **Linked / Mentioned TDocs:** {', '.join(sorted(tdocs_mentioned)) or 'None'}",
        f"- **Date Range:** {date_range}",
        f"- **Messages:** {len(messages)}",
        f"- **Participants:** {', '.join(participants) or 'None'}",
    ]
    if doc_urls:
        lines.append(f"- **Linked Documents ({len(doc_urls)}):**")
        for u in doc_urls:
            fname = urllib.parse.unquote(urllib.parse.urlsplit(u).path.rsplit("/", 1)[-1]) or u
            lines.append(f"  - [📄 {fname}]({u})")
    lines.extend([
        "",
        "---",
        "",
        "## Chronological Discussion",
        "",
    ])

    base_subject = thread.get("normalized_subject", "").lower()
    for idx, msg in enumerate(messages, 1):
        sender = (msg.get("sender_name") or msg.get("sender_email") or "Unknown").strip()
        comp = (msg.get("company") or "").strip()
        sender_disp = f"{sender} ({comp})" if comp else sender
        date_str = msg.get("date_sent", "Unknown Date")
        subj = (msg.get("subject") or "").strip()
        msg_id = str(msg.get("id") or "").strip()
        can_open_outlook = interactive_links and msg_id and not msg_id.startswith("MSG-")

        if can_open_outlook:
            enc_id = urllib.parse.quote(msg_id, safe="")
            lines.append(f"### {idx}. [{date_str}] {sender_disp} · [📧 Open in Outlook](outlook://msg/{enc_id})")
        else:
            lines.append(f"### {idx}. [{date_str}] {sender_disp}")
        if subj and subj.lower() != base_subject:
            lines.append(f"**Subject:** {subj}  ")
        if msg.get("attachments"):
            if can_open_outlook:
                enc_id = urllib.parse.quote(msg_id, safe="")
                att_links = [
                    f"[📎 {a}](outlook://att/{enc_id}/{urllib.parse.quote(str(a), safe='')})"
                    for a in msg["attachments"]
                ]
                lines.append(f"**Attachments:** {' · '.join(att_links)}  ")
            else:
                lines.append(f"**Attachments:** {', '.join(msg['attachments'])}  ")
        lines.append("")

        body = (msg.get("clean_body") if deduplicate_quotes and msg.get("clean_body") else msg.get("body_text")) or ""
        body = linkify_markdown_urls(body.strip())
        if body:
            lines.append(body)
        else:
            lines.append("*(Empty message body)*")
        lines.append("")
        lines.append("---")
        lines.append("")

    return "\n".join(lines).strip() + "\n"
