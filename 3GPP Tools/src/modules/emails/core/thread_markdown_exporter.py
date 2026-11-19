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
_ALLOWED_HTML_TAGS = {
    "a", "b", "i", "em", "strong", "code", "pre", "br", "hr", "p", "div", "span",
    "ul", "ol", "li", "h1", "h2", "h3", "h4", "h5", "h6", "sub", "sup", "mark", "blockquote",
}


def escape_non_html_angles(text: str) -> str:
    """Escapes inline delegate tags like <Ashok> or <Jinguo> so Qt's Markdown renderer does not treat them as unclosed HTML."""
    if not text:
        return ""
    # Simplify Outlook "<email <mailto:email> >" artifacts
    cleaned = re.sub(r"\s*<mailto:[^>\s]+>\s*", " ", text, flags=re.IGNORECASE)

    def _replace_tag(m: re.Match) -> str:
        inner = m.group(1)
        first = inner.strip().split()[0].lstrip("/").lower() if inner.strip() else ""
        if first in _ALLOWED_HTML_TAGS:
            return m.group(0)
        return f"&lt;{inner}&gt;"

    return re.sub(r"<([^<>\n]{1,120})>", _replace_tag, cleaned)


_ANGLE_HTTP_RE = re.compile(r"\s*<https?://[^>]+>", re.IGNORECASE)


def _table_cell_to_plain_and_md(cell: Any) -> tuple:
    """Extracts (plain_text_for_matching, markdown_cell_text) preserving clickable hyperlinks inside HTML tables."""
    plain_text = " ".join(cell.get_text(" ", strip=True).replace("\xa0", " ").split())

    parts: List[str] = []
    # Collect anchor replacements without mutating caller soup nodes
    for node in cell.descendants:
        pass
    # Walk children preserving <a> hyperlinks while flattening nested formatting tags
    def _walk(el: Any):
        name = getattr(el, "name", None)
        if name is None:
            parts.append(str(el))
            return
        if name.lower() == "a" and el.get("href"):
            href = unwrap_outlook_safelinks(str(el.get("href") or "").strip())
            label = " ".join(el.get_text(" ", strip=True).replace("\xa0", " ").split()) or href
            if href.lower().startswith(("http://", "https://")) and label:
                parts.append(f"[{label.replace('|', '\\|')}]({href})")
                return
        if name.lower() == "br":
            parts.append(" ")
            return
        for ch in getattr(el, "children", ()):
            _walk(ch)

    _walk(cell)
    md_text = " ".join("".join(parts).replace("\xa0", " ").split()).replace("|", "\\|")
    if not md_text:
        md_text = plain_text.replace("|", "\\|")
    return plain_text, md_text


def inject_html_tables_into_body(body_text: str, html_body: str) -> str:
    """
    Replaces Outlook's flattened plain-text table cells in `body_text` with structured
    GFM Markdown pipe tables extracted from `html_body`. Preserves surrounding text intact.
    """
    if not body_text or not html_body or "<table" not in html_body.lower():
        return body_text
    try:
        from bs4 import BeautifulSoup
    except ImportError:
        return body_text

    soup = BeautifulSoup(html_body, "html.parser")
    tables = [t for t in soup.find_all("table") if not t.find("table")]
    if not tables:
        return body_text

    lines = body_text.splitlines()
    search_from = 0
    for tbl in tables:
        plain_grid: List[List[str]] = []
        md_grid: List[List[str]] = []
        for tr in tbl.find_all("tr"):
            row_pairs = [_table_cell_to_plain_and_md(c) for c in tr.find_all(["td", "th"])]
            plain_cells = [p for p, _m in row_pairs]
            md_cells = [m for _p, m in row_pairs]
            if any(plain_cells):
                plain_grid.append(plain_cells)
                md_grid.append(md_cells)
        if not plain_grid or max(len(r) for r in plain_grid) < 2:
            continue
        tbl_norm = " ".join(" ".join(cell for row in plain_grid for cell in row if cell).split())
        if len(tbl_norm) < 4:
            continue

        matched_span = None
        for i in range(search_from, len(lines)):
            first = " ".join(_ANGLE_HTTP_RE.sub("", lines[i]).replace("\xa0", " ").split())
            if not first or not tbl_norm.startswith(first):
                continue
            acc = first
            if acc == tbl_norm:
                matched_span = (i, i)
                break
            for j in range(i + 1, len(lines)):
                nxt = " ".join(_ANGLE_HTTP_RE.sub("", lines[j]).replace("\xa0", " ").split())
                if not nxt:
                    continue
                cand = acc + " " + nxt
                if not tbl_norm.startswith(cand):
                    break
                acc = cand
                if acc == tbl_norm:
                    matched_span = (i, j)
                    break
            if matched_span:
                break

        if matched_span:
            col_count = max(len(r) for r in md_grid)
            padded = [r + [""] * (col_count - len(r)) for r in md_grid]
            md_rows = [
                "| " + " | ".join(padded[0]) + " |",
                "| " + " | ".join(["---"] * col_count) + " |",
            ] + ["| " + " | ".join(br) + " |" for br in padded[1:]]
            replacement = [""] + md_rows + [""]
            s_i, s_j = matched_span
            lines[s_i : s_j + 1] = replacement
            search_from = s_i + len(replacement)

    return "\n".join(lines)


def _normalize_outlook_list_spacing(text: str) -> str:
    """Normalizes Word/Outlook list markers with 4+ spaces/tabs so CommonMark does not render them as code blocks."""
    if not text:
        return ""
    cleaned = text.replace("\xa0", " ")
    # Collapse wide gaps after bullet/numbered list markers (e.g. "1.\t   Text" -> "1. Text")
    cleaned = re.sub(r"^([ \t]*(?:[-*+•]|\d+\.))[ \t]{2,}", r"\1 ", cleaned, flags=re.MULTILINE)
    # Strip 4+ leading spaces on standalone non-code lines unless part of a fenced block
    cleaned = re.sub(r"^[ \t]{4,}(?=\S)", "", cleaned, flags=re.MULTILINE)
    return cleaned


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
    newest_first: bool = False,
    include_metadata: bool = True,
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

    disc_heading = "## Discussion (Newest First)" if newest_first else "## Chronological Discussion"
    lines: List[str] = []
    if include_metadata:
        lines.extend([
            f"# Email Thread: {thread.get('title', 'Untitled Thread')}",
            "",
            "## Thread Metadata",
            f"- **Thread ID:** `{thread.get('thread_id', '')}`",
            f"- **Meeting:** `{thread.get('meeting_id', '') or 'N/A'}` ({thread.get('wg', '') or 'N/A'})",
            f"- **Linked AIs:** {', '.join(thread.get('linked_ais') or []) or 'None'}",
            f"- **Linked Sub-Topics:** {', '.join(thread.get('linked_sub_items') or []) or 'None'}",
            f"- **Linked WIs:** {', '.join(thread.get('linked_wis') or []) or 'None'}",
            f"- **Linked / Mentioned TDocs:** {', '.join(sorted(tdocs_mentioned)) or 'None'}",
            f"- **Date Range:** {date_range}",
            f"- **Messages:** {len(messages)}",
            f"- **Participants:** {', '.join(participants) or 'None'}",
        ])
        user_status = str(thread.get("user_status") or "").strip()
        user_tags = [str(x).strip() for x in (thread.get("user_tags") or []) if str(x).strip()]
        user_notes = str(thread.get("user_notes") or "").strip()
        if user_status and user_status != "⚪ Neutral":
            lines.append(f"- **My Status:** {user_status}")
        if user_tags:
            lines.append(f"- **My Tags:** {', '.join(user_tags)}")
        if doc_urls:
            lines.append(f"- **Linked Documents ({len(doc_urls)}):**")
            for u in doc_urls:
                fname = urllib.parse.unquote(urllib.parse.urlsplit(u).path.rsplit("/", 1)[-1]) or u
                lines.append(f"  - [📄 {fname}]({u})")
        if user_notes:
            lines.append("")
            lines.append("> [!NOTE] Delegate Notes")
            for nline in user_notes.splitlines():
                lines.append(f"> {nline}" if nline.strip() else ">")
        lines.extend(["", "---", ""])
    lines.extend([disc_heading, ""])

    base_subject = thread.get("normalized_subject", "").lower()
    indexed_msgs = list(enumerate(messages, 1))
    if newest_first:
        indexed_msgs.reverse()

    unread_seq = 0
    for idx, msg in indexed_msgs:
        sender = (msg.get("sender_name") or msg.get("sender_email") or "Unknown").strip()
        comp = (msg.get("company") or "").strip()
        sender_disp = f"{sender} ({comp})" if comp else sender
        date_str = msg.get("date_sent", "Unknown Date")
        subj = (msg.get("subject") or "").strip()
        msg_id = str(msg.get("id") or "").strip()
        is_unread = int(msg.get("is_read", 1) or 0) == 0
        is_flagged = int(msg.get("flag_status", 0) or 0) == 2
        new_badge = " 🔵 **NEW**" if (interactive_links and is_unread) else ""
        flag_badge = " 🚩 **FLAGGED**" if (interactive_links and is_flagged) else ""
        can_open_outlook = interactive_links and msg_id and not msg_id.startswith("MSG-")

        if interactive_links:
            anchors = f'<a name="msg-{idx}"></a>'
            if is_unread:
                unread_seq += 1
                anchors += f'<a name="unread-{unread_seq}"></a>'
            enc_id = urllib.parse.quote(msg_id, safe="")
            action_parts = []
            if can_open_outlook:
                action_parts.append(f"[📧 Open in Outlook](outlook://msg/{enc_id})")
                action_parts.append(f"[↩️ Reply](outlook://reply/{enc_id})")
                action_parts.append(f"[🔄 Reload](outlook://reload/{enc_id})")
            if msg_id:
                action_parts.append(
                    f"[✓ Mark Read](outlook://read/{enc_id})" if is_unread else f"[🔵 Mark Unread](outlook://unread/{enc_id})"
                )
                action_parts.append(
                    f"[🏳️ Clear Flag](outlook://unflag/{enc_id})" if is_flagged else f"[🚩 Flag](outlook://flag/{enc_id})"
                )
            action_parts.append(f"[📋 Copy](outlook://copy/{enc_id})")
            lines.append(f"### {anchors}{idx}. [{date_str}] {sender_disp}{new_badge}{flag_badge} · {' · '.join(action_parts)}")
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
        body = _normalize_outlook_list_spacing(body.strip())
        body = linkify_markdown_urls(body)
        if interactive_links:
            body = escape_non_html_angles(body)
        if body:
            lines.append(body)
        else:
            lines.append("*(Empty message body)*")
        lines.append("")
        lines.append("---")
        lines.append("")

    return "\n".join(lines).strip() + "\n"
