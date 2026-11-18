# --- File: src/modules/emails/core/thread_markdown_exporter.py ---
import re
import urllib.parse
from typing import Any, Dict, List


_TDOC_RE = re.compile(r"\b(?:S[1-6]|R[1-5]|C[1-6]|SP|RP|CP)-\d{6,7}\b", re.IGNORECASE)
_SAFELINK_RE = re.compile(
    r"<?https?://[a-z0-9.-]*safelinks\.protection\.outlook\.com/\?[^\s<>\"\)\]]+>?",
    re.IGNORECASE,
)
_BARE_URL_RE = re.compile(
    r"(?<![\(\[<\"])(https?://[^\s<>\"\)\]]+)",
    re.IGNORECASE,
)


def unwrap_outlook_safelinks(text: str) -> str:
    """
    Replaces Microsoft Outlook Safelinks (https://*.safelinks.protection.outlook.com/?url=...)
    with the decoded original URL, and collapses duplicate 'url<url>' patterns from Outlook plain text.
    """
    if not text or "safelinks.protection.outlook.com" not in text.lower():
        return text or ""

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

    cleaned = _SAFELINK_RE.sub(_replace_safelink, text)
    # Collapse Outlook's "https://example.com https://example.com" or "https://example.com<https://example.com>"
    cleaned = re.sub(r"(https?://[^\s<>\"\)\]]+)\s*<+\1>+", r"\1", cleaned)
    cleaned = re.sub(r"(https?://[^\s<>\"\)\]]+)(?:\1)+", r"\1", cleaned)
    return cleaned


def linkify_markdown_urls(text: str) -> str:
    """Converts bare http(s):// URLs in text into clickable Markdown links [url](url)."""
    if not text:
        return ""
    unwrapped = unwrap_outlook_safelinks(text)
    # Also unwrap angle-bracketed <https://...> into clean [url](url)
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
) -> str:
    """
    Exports a tracked email thread to a lossless, human- and LLM-readable Markdown document.
    Preserves inline replies, metadata, participants/companies, and TDoc references,
    unwraps Outlook Safelinks, and formats URLs as clickable Markdown links.
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
        "",
        "---",
        "",
        "## Chronological Discussion",
        "",
    ]

    base_subject = thread.get("normalized_subject", "").lower()
    for idx, msg in enumerate(messages, 1):
        sender = (msg.get("sender_name") or msg.get("sender_email") or "Unknown").strip()
        comp = (msg.get("company") or "").strip()
        sender_disp = f"{sender} ({comp})" if comp else sender
        date_str = msg.get("date_sent", "Unknown Date")
        subj = (msg.get("subject") or "").strip()

        lines.append(f"### {idx}. [{date_str}] {sender_disp}")
        if subj and subj.lower() != base_subject:
            lines.append(f"**Subject:** {subj}  ")
        if msg.get("attachments"):
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
