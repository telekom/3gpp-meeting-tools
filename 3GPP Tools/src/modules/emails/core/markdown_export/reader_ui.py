# --- File: src/modules/emails/core/markdown_export/reader_ui.py ---
"""e-Meeting callout rendering and shared QTextBrowser navigation/styling helpers."""

from typing import Any, List

from core.ui.theme import C

from .formatting_helpers import _format_trailing_quoted_history, strip_trailing_quotes
from .html_and_tables import _normalize_outlook_list_spacing, escape_non_html_angles
from .url_and_links import linkify_markdown_urls, linkify_tdoc_tokens


def find_unread_heading_blocks(browser: Any) -> List[Any]:
    """Returns QTextBlock headings representing unread messages in a QTextBrowser rendered with export_thread_to_markdown."""
    doc = browser.document()
    blocks: List[Any] = []
    block = doc.firstBlock()
    while block.isValid():
        if block.blockFormat().headingLevel() == 3 and "🔵 NEW" in block.text():
            blocks.append(block)
        block = block.next()
    return blocks


def find_heading_block_for_message_index(browser: Any, msg_index: int) -> Any:
    """Finds the QTextBlock heading for message `msg_index` (1-based) in a QTextBrowser."""
    if browser is None or not hasattr(browser, "document"):
        return None
    prefix = f"{int(msg_index)}."
    doc = browser.document()
    block = doc.firstBlock()
    while block.isValid():
        if block.blockFormat().headingLevel() == 3 and block.text().lstrip().startswith(prefix):
            return block
        block = block.next()
    return None


def format_emeeting_body_with_callouts(
    raw_body: str,
    *,
    primary_tdoc: str = "",
    interactive_links: bool = False,
    enable_tdoc_links: bool = False,
    deduplicate_quotes: bool = True,
) -> str:
    """
    Formats an email body, extracting any `<<START>> ... <<END>>` e-Meeting formal comment blocks
    into a distinct Option A amber callout box (or structured Markdown blockquote for LLM export)
    and linkifying TDoc tokens + shorthand revisions (`r01`).
    """
    from modules.emails.core.emeeting_signals import (
        extract_start_end_blocks,
        linkify_shorthand_revisions,
    )

    formal_blocks, remaining_body = extract_start_end_blocks(raw_body or "")

    def _transform_chunk(chunk: str, is_formal: bool = False) -> str:
        txt = strip_trailing_quotes(chunk) if (deduplicate_quotes and not is_formal) else chunk
        txt = _normalize_outlook_list_spacing(txt.strip())
        if interactive_links and not deduplicate_quotes and not is_formal:
            txt = _format_trailing_quoted_history(txt)
        txt = linkify_markdown_urls(txt, shorten_long_urls=interactive_links)
        if interactive_links and enable_tdoc_links:
            txt = linkify_tdoc_tokens(txt)
            if is_formal and primary_tdoc:
                txt = linkify_shorthand_revisions(txt, primary_tdoc=primary_tdoc)
        if interactive_links:
            txt = escape_non_html_angles(txt)
        return txt

    out_parts: List[str] = []
    for blk in formal_blocks:
        formatted_blk = _transform_chunk(blk, is_formal=True)
        if not formatted_blk:
            continue
        if interactive_links:
            blk_lines = "<br>".join(ln for ln in formatted_blk.splitlines())
            out_parts.append(
                f'<table width="100%" cellspacing="0" cellpadding="6" '
                f'style="background-color:{C.WARNING_SOFT}; border:1px solid {C.WARNING_BORDER}; '
                f'margin-top:4px; margin-bottom:6px;">'
                f'<tr>'
                f'<td width="4" style="background-color:{C.WARNING_ACCENT};"></td>'
                f'<td><b style="color:{C.WARNING_DARK};">📋 e-Meeting Formal Comment (&lt;&lt;START&gt;&gt;…&lt;&lt;END&gt;&gt;)</b><br>'
                f'{blk_lines}</td>'
                f'</tr></table>'
            )
        else:
            q_lines = "\n".join(f"> {ln}" if ln.strip() else ">" for ln in formatted_blk.splitlines())
            out_parts.append(f"> **[e-Meeting Formal Comment `<<START>>...<<END>>`]:**\n{q_lines}")

    free_formatted = _transform_chunk(remaining_body, is_formal=False)
    if free_formatted:
        out_parts.append(free_formatted)

    return "\n\n".join(out_parts).strip()


def _wrap_hover_url(url: str, width: int = 84) -> str:
    """Wraps long URLs across multiple lines so hover tooltips fit comfortably on screen."""
    u = str(url or "").strip()
    if len(u) <= width:
        return u
    return "\n".join(u[i : i + width] for i in range(0, len(u), width))


def handle_reader_link_hover(browser: Any, href: str) -> bool:
    """Shows full URL tooltip only on http(s) links and hides it otherwise. Returns True if shown."""
    from PyQt5.QtGui import QCursor
    from PyQt5.QtWidgets import QToolTip

    url = str(href or "").strip()
    if url.lower().startswith(("http://", "https://")):
        QToolTip.showText(QCursor.pos(), _wrap_hover_url(url), browser)
        return True
    QToolTip.hideText()
    return False


def apply_email_reader_document_style(browser: Any) -> None:
    """Applies clean readability styling and targeted http(s) link hover tooltips to a QTextBrowser."""
    if browser is None:
        return
    doc = browser.document() if hasattr(browser, "document") else None
    if doc is not None:
        doc.setDefaultStyleSheet(
            "h3 { margin-top: 6px; margin-bottom: 3px; font-size: 105%; } "
            "p { margin-top: 3px; margin-bottom: 4px; } "
            f"blockquote {{ margin: 4px 0 4px 10px; padding-left: 8px; color: {C.TEXT_MUTED}; border-left: 2px solid {C.BORDER_STRONG}; }}"
        )
    if hasattr(browser, "highlighted") and not browser.property("_email_link_hover_installed"):
        browser.setProperty("_email_link_hover_installed", True)
        try:
            browser.highlighted[str].connect(lambda href, b=browser: handle_reader_link_hover(b, href))
        except Exception:
            pass


def scroll_browser_to_block(browser: Any, block: Any) -> int:
    """Deterministically scrolls a QTextBrowser viewport to the vertical coordinate of `block`."""
    if browser is None or block is None or not block.isValid():
        return 0
    doc = browser.document()
    layout = doc.documentLayout()
    layout.documentSize()
    top_y = int(layout.blockBoundingRect(block).top()) - 6
    sb = browser.verticalScrollBar()
    target = max(0, min(top_y, sb.maximum()))
    sb.setValue(target)
    return sb.value()
