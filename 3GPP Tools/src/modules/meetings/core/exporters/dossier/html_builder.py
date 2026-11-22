"""Interactive HTML dossier synthesizer and Markdown-to-HTML converter."""

from __future__ import annotations

import base64
import datetime
import html
import logging
import re
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Tuple

import markdown
from PyQt5.QtGui import QTextDocument

from core.utils.company_sanitizer import CompanySanitizer
from core.utils.portal_utils import get_portal_tdoc_url

from .css import DOSSIER_CSS

logger = logging.getLogger(__name__)


def convert_markdown_to_html(
    md_text: str,
    local_media_dir: Path,
    media_export_dir: Path,
    tdoc_id: str,
    *,
    meeting_dir: Path,
    embed_base64_images: bool = True,
) -> str:
    """
    Converts document Markdown to semantic HTML, styling inline revision marks,
    3GPP change markers, and inlining images as Base64 Data URIs for a self-contained dossier.
    """
    if not md_text.strip():
        return "<p><em>(No text content available or document could not be converted)</em></p>"

    # Pre-process revision tags before markdown parser
    # Convert [INSERTED: ...] and [DELETED: ...] into styled HTML
    processed = md_text

    # Change markers (* * * * Start of changes * * * *)
    processed = re.sub(
        r"(\*\s*\*\s*\*\s*\*?\s*(?:Start|Next|End)\s+of\s+changes\s*\*?\s*\*\s*\*\s*\*?)",
        r'\n\n<div class="change-marker">\1</div>\n\n',
        processed,
        flags=re.IGNORECASE,
    )

    # Style [INSERTED: ...] and [DELETED: ...]
    processed = re.sub(
        r"\[INSERTED:\s*(.*?)\]",
        r'<ins class="diff-ins">\1</ins>',
        processed,
        flags=re.DOTALL,
    )
    processed = re.sub(
        r"\[DELETED:\s*(.*?)\]",
        r'<del class="diff-del">\1</del>',
        processed,
        flags=re.DOTALL,
    )

    # Render Markdown to HTML
    try:
        body_html = markdown.markdown(
            processed,
            extensions=["tables", "fenced_code", "toc", "attr_list"],
        )
    except Exception as e:
        logger.warning(f"[DossierExporter] markdown library parsing error: {e}. Falling back.")
        doc = QTextDocument()
        doc.setMarkdown(processed)
        body_html = doc.toHtml()

    # In-line images as Base64 Data URIs (or resolve paths)
    if embed_base64_images:

        def _replace_img(match):
            full_tag = match.group(0)
            src = match.group(1)
            img_name = Path(src).name

            # Look for image in local_media_dir or media_export_dir
            candidates = [
                local_media_dir / img_name,
                media_export_dir / img_name,
                meeting_dir / tdoc_id / "media" / img_name,
            ]
            found_img = next((c for c in candidates if c.is_file()), None)
            if found_img:
                try:
                    ext = found_img.suffix.lower()
                    mime = (
                        "image/png"
                        if ext == ".png"
                        else ("image/jpeg" if ext in [".jpg", ".jpeg"] else "image/png")
                    )
                    b64 = base64.b64encode(found_img.read_bytes()).decode("ascii")
                    return f'<img src="data:{mime};base64,{b64}" alt="{html.escape(img_name)}" />'
                except Exception as err:
                    logger.debug(f"[DossierExporter] Base64 encode failed for {found_img}: {err}")
            return full_tag

        body_html = re.sub(r'<img[^>]+src=["\']([^"\']+)["\'][^>]*>', _replace_img, body_html)

    return body_html


def build_html_dossier(
    docs: List[Dict[str, Any]],
    media_export_dir: Path,
    mtg_info: Dict[str, Any],
    *,
    meeting_dir: Path,
    embed_base64_images: bool = True,
    convert_md_fn: Optional[Callable[[str, Path, Path, str], str]] = None,
) -> Tuple[str, Dict[str, str]]:
    """
    Synthesizes the complete HTML dossier with:
    - Floating navigation bar with dropdown and quick actions
    - Executive cover card and statistics
    - Master Table of Contents with jump anchors
    - Fully rendered TDoc sections with highlighted diffs and embedded images
    """
    wg_name = str(mtg_info.get("wg_name", "3GPP"))
    mtg_num = str(mtg_info.get("meeting_number", ""))
    timestamp_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")

    # Sort docs naturally by Agenda Item then TDoc ID
    def _sort_key(d):
        ai = str(d.get("ai", ""))
        chunks = [int(c) if c.isdigit() else c.lower() for c in re.split(r"(\d+)", ai)]
        return (chunks, str(d.get("tdoc_id", "")))

    sorted_docs = sorted(docs, key=_sort_key)

    # Count categories
    type_counts: Dict[str, int] = {}
    for d in sorted_docs:
        t = d.get("type", "Other")
        type_counts[t] = type_counts.get(t, 0) + 1

    # Build dropdown options
    options_html = ['<option value="#master-index">📑 Master Index & Summary</option>']
    for d in sorted_docs:
        tid = d["tdoc_id"]
        short_t = html.escape(d["title"][:45] + "..." if len(d["title"]) > 45 else d["title"])
        options_html.append(f'<option value="#tdoc-{tid}">{tid} · {short_t}</option>')
    options_markup = "\n".join(options_html)

    # Build stats chips
    stats_chips = [
        f'<div class="stat-chip"><div class="val">{len(sorted_docs)}</div><div class="lbl">Total Documents</div></div>'
    ]
    for t, cnt in sorted(type_counts.items(), key=lambda x: -x[1])[:4]:
        stats_chips.append(
            f'<div class="stat-chip"><div class="val">{cnt}</div><div class="lbl">{html.escape(t)}</div></div>'
        )
    stats_markup = "\n".join(stats_chips)

    # Build Master Index Table
    table_rows = []
    for d in sorted_docs:
        tid = d["tdoc_id"]
        t_type = d.get("type", "")
        t_status = d.get("status", "")
        badge_type = (
            "badge-cr"
            if "CR" in t_type.upper() and "P" not in t_type.upper()
            else (
                "badge-pcr"
                if "PCR" in t_type.upper() or "P-CR" in t_type.upper()
                else ("badge-ls" if "LS" in t_type.upper() else "badge-disc")
            )
        )
        badge_status = (
            "badge-agreed"
            if any(x in t_status.lower() for x in ["agreed", "approved"])
            else "badge-noted"
        )

        row_html = f"""
            <tr>
              <td><a href="#tdoc-{tid}"><strong>{tid}</strong></a></td>
              <td><span class="badge {badge_type}">{html.escape(t_type)}</span></td>
              <td>{html.escape(d.get("ai", ""))}</td>
              <td>{html.escape(d.get("title", ""))}</td>
              <td>{CompanySanitizer.format_source_html(d.get("source", ""))}</td>
              <td><span class="badge {badge_status}">{html.escape(t_status or "-")}</span></td>
              <td><a href="#tdoc-{tid}">Jump ➔</a></td>
            </tr>
            """
        table_rows.append(row_html)
    index_table_markup = "\n".join(table_rows)

    # Build Document Sections
    sections_html = []
    tdoc_html_map: Dict[str, str] = {}

    for d in sorted_docs:
        tid = d["tdoc_id"]
        raw_md = d.get("md_text", "")
        if convert_md_fn is not None:
            body_html = convert_md_fn(raw_md, d["media_dir"], media_export_dir, tid)
        else:
            body_html = convert_markdown_to_html(
                raw_md,
                d["media_dir"],
                media_export_dir,
                tid,
                meeting_dir=meeting_dir,
                embed_base64_images=embed_base64_images,
            )

        url_html = (
            f'<a href="{d["tdoc_url"]}" target="_blank">🌐 3GPP Server</a>'
            if d.get("tdoc_url")
            else ""
        )
        portal_url = get_portal_tdoc_url(tid)

        section_markup = f"""
            <section id="tdoc-{tid}" class="tdoc-section" data-tdoc="{tid}">
              <!-- BEGIN_TDOC:{tid} -->
              <div class="tdoc-header">
                <div class="tdoc-header-top">
                  <h1 class="tdoc-id-title"><span style="color:#005A9E;">[{tid}]</span> {html.escape(d.get("title", ""))}</h1>
                  <div class="tdoc-actions no-print">
                    <a href="#master-index">⬆ Back to Index</a>
                    <a href="{portal_url}" target="_blank">🌐 3GU Portal</a>
                    {url_html}
                  </div>
                </div>
                <div class="tdoc-meta-row">
                  <div><strong>Source:</strong> {CompanySanitizer.format_source_html(d.get("source", "Unknown"))}</div>
                  <div><strong>Type:</strong> {html.escape(d.get("type", "Unknown"))}</div>
                  <div><strong>Agenda Item:</strong> {html.escape(d.get("ai", "Unassigned"))}</div>
                  <div><strong>Status:</strong> {html.escape(d.get("status", "Unknown"))}</div>
                </div>
              </div>
              <div class="tdoc-body">
                {body_html}
              </div>
              <div style="margin-top: 24px; text-align: right;" class="no-print">
                <a href="#master-index" style="font-size: 11px; color: #64748B; text-decoration: none;">⬆ Back to Top</a>
              </div>
            </section>
            """
        sections_html.append(section_markup)
        tdoc_html_map[tid] = section_markup

    all_sections_markup = "\n".join(sections_html)

    full_html = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>3GPP Dossier - {html.escape(wg_name)} #{html.escape(mtg_num)}</title>
  <style>
{DOSSIER_CSS}
  </style>
</head>
<body>

  <!-- Floating Navigation Toolbar -->
  <div class="floating-nav no-print">
    <div class="title-pill">
      📑 <span>{html.escape(wg_name)} #{html.escape(mtg_num)} Dossier</span>
    </div>
    <div class="nav-controls">
      <select onchange="if(this.value) window.location.hash = this.value;">
        {options_markup}
      </select>
      <a href="#master-index" class="btn">📋 Index</a>
      <button class="primary" onclick="window.print();">🖨️ Print / PDF</button>
    </div>
  </div>

  <div class="dossier-wrapper">

    <!-- Cover & Statistics Card -->
    <div id="master-index" class="cover-card">
      <h1>3GPP Technical Dossier</h1>
      <div class="meta-subtitle">
        <strong>Working Group:</strong> {html.escape(wg_name)} | 
        <strong>Meeting:</strong> #{html.escape(mtg_num)} | 
        <strong>Compiled:</strong> {timestamp_str}
      </div>
      <div class="stats-grid">
        {stats_markup}
      </div>
    </div>

    <!-- Master Index Table -->
    <div class="index-card">
      <h2>📑 Master Table of Contents</h2>
      <table class="index-table">
        <thead>
          <tr>
            <th>TDoc</th>
            <th>Type</th>
            <th>AI</th>
            <th>Title</th>
            <th>Source</th>
            <th>Status</th>
            <th>Action</th>
          </tr>
        </thead>
        <tbody>
          {index_table_markup}
        </tbody>
      </table>
    </div>

    <!-- TDoc Content Sections -->
    {all_sections_markup}

  </div>

</body>
</html>
"""
    return full_html, tdoc_html_map
