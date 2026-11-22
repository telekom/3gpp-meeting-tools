"""Vector PDF renderer with pypdf outline bookmarks and 1-indexed page spans."""

from __future__ import annotations

import logging
import os
import re
import shutil
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Tuple

import pypdf
from PyQt5.QtGui import QPageSize, QPdfWriter, QTextDocument

from core.utils.paths import to_extended_path

from .css import DOSSIER_CSS

logger = logging.getLogger(__name__)


def render_bookmarked_pdf(
    arg1: str,
    arg2: Any,
    arg3: Any = None,
    arg4: Optional[List[Dict[str, Any]]] = None,
    *,
    is_cancelled: Optional[Callable[[], bool]] = None,
    progress_cb: Optional[Callable[[str], None]] = None,
) -> Dict[str, Tuple[int, int]]:
    """
    Renders HTML to a vector PDF via modular QTextDocument + QPdfWriter rendering,
    merging pages via pypdf and injecting hierarchical outline bookmarks
    (Master Index -> Agenda Item -> TDoc ID).

    Supports both:
    - 4-argument modular signature: (cover_html, section_tuples, output_pdf_path, docs)
    - 3-argument legacy signature: (full_html, output_pdf_path, docs)

    Modular rendering ensures O(N) linear time and low memory footprint, preventing
    Windows printer subsystem deadlocks and Qt layout engine freezes on large HTML dossiers.
    """
    if arg4 is not None:
        # Called with 4-argument signature: (cover_html, section_tuples, output_pdf_path, docs)
        cover_html = arg1
        section_tuples = arg2
        actual_pdf_path = Path(arg3)
        actual_docs = arg4
    else:
        # Called with 3-argument signature: (full_html, output_pdf_path, docs)
        full_html = arg1
        actual_pdf_path = Path(arg2)
        actual_docs = arg3 if isinstance(arg3, list) else []

        # Extract sections from full_html if present
        css_match = re.search(r"<style>(.*?)</style>", full_html, flags=re.DOTALL)
        css_str = css_match.group(1) if css_match else DOSSIER_CSS
        sec_idx = full_html.find("<section")
        if sec_idx != -1:
            cover_html = (
                f"<!DOCTYPE html><html><head><style>{css_str}</style></head>"
                f"<body>{full_html[:sec_idx]}</div></body></html>"
            )
            extracted_sections = re.findall(r"<section id=[\s\S]*?</section>", full_html)
            doc_map = {d["tdoc_id"].upper(): d for d in actual_docs}
            section_tuples = []
            for s in extracted_sections:
                m = re.search(r'data-tdoc=["\']([^"\']+)["\']', s)
                tid = m.group(1).upper() if m else ""
                doc_meta = doc_map.get(tid, {"tdoc_id": tid})
                section_tuples.append((doc_meta, s))
        else:
            cover_html = full_html
            section_tuples = []

    temp_dir = actual_pdf_path.parent / f"_temp_pdf_{actual_pdf_path.stem}"
    os.makedirs(to_extended_path(temp_dir), exist_ok=True)
    pdf_writer = pypdf.PdfWriter()

    page_spans: Dict[str, Tuple[int, int]] = {}
    current_page_idx = 0

    try:
        # 1. Render Cover & Master Table of Contents
        cover_pdf = temp_dir / "cover.pdf"
        doc_cover = QTextDocument()
        doc_cover.setHtml(cover_html)
        qpdf_cover = QPdfWriter(str(cover_pdf))
        qpdf_cover.setPageSize(QPageSize(QPageSize.A4))
        doc_cover.print_(qpdf_cover)

        r_cover = pypdf.PdfReader(str(cover_pdf))
        cover_page_count = len(r_cover.pages)
        for p in r_cover.pages:
            pdf_writer.add_page(p)

        # Master index bookmark at page 0
        pdf_writer.add_outline_item("📑 Master Index & Summary", 0)
        current_page_idx += cover_page_count

        # Cleanup cover temp file early
        try:
            cover_pdf.unlink()
        except Exception:
            pass

        # 2. Render each TDoc section individually and record exact start page & span
        temp_sec_pdf = temp_dir / "temp_section.pdf"
        qpdf_sec = QPdfWriter(str(temp_sec_pdf))
        qpdf_sec.setPageSize(QPageSize(QPageSize.A4))

        start_pages: Dict[str, int] = {}
        total_sections = len(section_tuples)

        for idx, (doc_meta, sec_markup) in enumerate(section_tuples):
            if is_cancelled and is_cancelled():
                break

            tid = str(doc_meta.get("tdoc_id", "")).strip().upper()
            if progress_cb and ((idx + 1) % 10 == 0 or idx == total_sections - 1):
                progress_cb(f"Rendering vector PDF pages ({idx + 1}/{total_sections})...")

            sec_html = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <style>
{DOSSIER_CSS}
  </style>
</head>
<body>
  <div class="dossier-wrapper">
    {sec_markup}
  </div>
</body>
</html>
"""
            doc_sec = QTextDocument()
            doc_sec.setHtml(sec_html)
            doc_sec.print_(qpdf_sec)

            r_sec = pypdf.PdfReader(str(temp_sec_pdf))
            sec_pages = len(r_sec.pages)
            start_0 = current_page_idx
            start_pages[tid] = start_0
            page_spans[tid] = (start_0 + 1, start_0 + sec_pages)

            for p in r_sec.pages:
                pdf_writer.add_page(p)

            current_page_idx += sec_pages

        if temp_sec_pdf.exists():
            try:
                temp_sec_pdf.unlink()
            except Exception:
                pass

        # 3. Inject Hierarchical Outline Bookmarks (Agenda Item > TDoc ID)
        ai_groups: Dict[str, List[Dict[str, Any]]] = {}
        for d in actual_docs:
            ai = str(d.get("ai", "Unassigned")).strip()
            ai_groups.setdefault(ai, []).append(d)

        def _ai_sort(ai_name):
            return [int(c) if c.isdigit() else c.lower() for c in re.split(r"(\d+)", ai_name)]

        for ai, group_docs in sorted(ai_groups.items(), key=lambda x: _ai_sort(x[0])):
            first_tid = group_docs[0]["tdoc_id"].upper()
            first_page_0 = start_pages.get(first_tid, 0)

            ai_title = f"📁 Agenda Item {ai} ({len(group_docs)} Docs)"
            ai_bm = pdf_writer.add_outline_item(ai_title, first_page_0)

            for d in group_docs:
                tid = d["tdoc_id"]
                tid_upper = tid.upper()
                p_0 = start_pages.get(tid_upper, 0)
                src = d.get("source", "")
                src_tag = f"[{src}] " if src else ""
                t_title = d.get("title", "")
                short_title = t_title[:40] + "..." if len(t_title) > 40 else t_title
                doc_bm_title = f"📄 {tid} · {src_tag}{short_title}"
                pdf_writer.add_outline_item(doc_bm_title, p_0, parent=ai_bm)

        # 4. Write final bookmarked PDF (if not cancelled)
        if not (is_cancelled and is_cancelled()):
            with open(to_extended_path(actual_pdf_path), "wb") as f_out:
                pdf_writer.write(f_out)
            logger.info(
                f"[DossierExporter] Saved bookmarked PDF dossier: {actual_pdf_path} ({len(pdf_writer.pages)} pages)"
            )
        else:
            logger.info("[DossierExporter] Vector PDF rendering interrupted by user request.")

    finally:
        try:
            shutil.rmtree(temp_dir, ignore_errors=True)
        except Exception:
            pass

    return page_spans
