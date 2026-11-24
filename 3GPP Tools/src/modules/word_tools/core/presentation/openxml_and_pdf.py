"""Tier 2 headless OpenXML (.pptx) and Tier 3 PDF slide extraction engines."""

from __future__ import annotations

import posixpath
import re
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path
from typing import Callable, List, Optional, Tuple

from core.utils.paths import strip_extended_path, to_extended_path
from modules.word_tools.core.presentation.models import (
    OPENXML_NS,
    ExtractedSlide,
    _clean_tdoc_prefix,
    _safe_log,
)


def _extract_via_openxml(
    pptx_path: Path,
    media_dir: Path,
    tdoc_id: Optional[str] = None,
    cancellation_check: Optional[Callable[[], bool]] = None,
    log_fn: Optional[Callable[[str], None]] = None,
) -> Tuple[List[ExtractedSlide], List[str]]:
    """
    Parses a .pptx ZIP archive directly without Microsoft Office.
    Extracts ordered slides, titles, paragraphs, tables, speaker notes, and embedded media.
    """
    slides_data: List[ExtractedSlide] = []
    warnings: List[str] = []

    if pptx_path.suffix.lower() != ".pptx":
        raise ValueError(f"OpenXML direct parsing only supports .pptx archives, not {pptx_path.suffix}")

    media_dir = Path(media_dir).resolve()
    media_dir.mkdir(parents=True, exist_ok=True)

    pptx_path_ext = to_extended_path(pptx_path)
    with zipfile.ZipFile(pptx_path_ext, "r") as z:
        # 1. Extract all embedded media from ppt/media/ into media_dir
        extracted_media_count = 0
        clean_prefix = _clean_tdoc_prefix(tdoc_id)
        for name in z.namelist():
            if name.startswith("ppt/media/") and not name.endswith("/"):
                base_name = posixpath.basename(name)
                dest = media_dir / base_name
                media_bytes = z.read(name)
                dest.write_bytes(media_bytes)
                if clean_prefix:
                    (media_dir / f"{clean_prefix}{base_name}").write_bytes(media_bytes)
                extracted_media_count += 1

        if extracted_media_count > 0:
            _safe_log(f"Extracted {extracted_media_count} embedded figures from presentation package.", log_fn)

        # 2. Determine ordered list of slide XML files from presentation.xml.rels
        slide_targets: List[str] = []
        try:
            pres_rels_xml = z.read("ppt/_rels/presentation.xml.rels")
            rels_root = ET.fromstring(pres_rels_xml)
            rid_map = {}
            for rel in rels_root.findall("pr:Relationship", OPENXML_NS):
                r_id = rel.get("Id")
                r_target = rel.get("Target")
                if r_target and "slides/slide" in r_target:
                    t_norm = r_target.lstrip("/")
                    if not t_norm.startswith("ppt/"):
                        t_norm = "ppt/" + t_norm
                    rid_map[r_id] = t_norm

            pres_xml = z.read("ppt/presentation.xml")
            pres_root = ET.fromstring(pres_xml)
            for sld_id in pres_root.findall(".//p:sldId", OPENXML_NS):
                r_id = sld_id.get(f"{{{OPENXML_NS['r']}}}id")
                if r_id and r_id in rid_map:
                    slide_targets.append(rid_map[r_id])
        except Exception as pres_err:
            warnings.append(f"Could not parse presentation relationships ({pres_err}); scanning zip for slide*.xml")

        if not slide_targets:
            slide_files = [n for n in z.namelist() if re.match(r"^ppt/slides/slide\d+\.xml$", n)]
            slide_files.sort(key=lambda s: int(re.search(r"\d+", posixpath.basename(s)).group(0)))
            slide_targets = slide_files

        for idx, slide_path in enumerate(slide_targets, start=1):
            if cancellation_check and cancellation_check():
                break

            slide_item = ExtractedSlide(index=idx)
            try:
                slide_xml = z.read(slide_path)
                s_root = ET.fromstring(slide_xml)

                title_text = ""
                for sp in s_root.findall(".//p:sp", OPENXML_NS):
                    ph = sp.find(".//p:ph", OPENXML_NS)
                    if ph is not None and ph.get("type") in ("title", "ctrTitle"):
                        t_parts = [t.text for t in sp.findall(".//a:t", OPENXML_NS) if t.text]
                        if t_parts:
                            title_text = "".join(t_parts).strip()
                            break

                for sp in s_root.findall(".//p:sp", OPENXML_NS):
                    ph = sp.find(".//p:ph", OPENXML_NS)
                    if ph is not None and ph.get("type") in ("title", "ctrTitle"):
                        continue

                    for p in sp.findall(".//a:p", OPENXML_NS):
                        p_pr = p.find("a:pPr", OPENXML_NS)
                        lvl = int(p_pr.get("lvl", "0")) if p_pr is not None else 0
                        indent_spaces = "  " * lvl

                        run_texts = [t.text for t in p.findall(".//a:t", OPENXML_NS) if t.text]
                        p_full = "".join(run_texts).strip()
                        if p_full:
                            if not title_text and idx == 1 and lvl == 0:
                                title_text = p_full
                            else:
                                slide_item.body_paragraphs.append(f"{indent_spaces}- {p_full}")

                for tbl in s_root.findall(".//a:tbl", OPENXML_NS):
                    rows_data: List[List[str]] = []
                    for tr in tbl.findall("a:tr", OPENXML_NS):
                        row_cells: List[str] = []
                        for tc in tr.findall("a:tc", OPENXML_NS):
                            cell_text_parts = [t.text for t in tc.findall(".//a:t", OPENXML_NS) if t.text]
                            row_cells.append("".join(cell_text_parts).strip())
                        rows_data.append(row_cells)
                    if rows_data:
                        slide_item.tables.append(rows_data)

                if not title_text and slide_item.body_paragraphs:
                    first_p = slide_item.body_paragraphs[0].lstrip("- ").strip()
                    if len(first_p) < 80:
                        title_text = first_p
                        slide_item.body_paragraphs.pop(0)

                slide_item.title = title_text

                # Extract Slide Media & Speaker Notes
                slide_rel_path = f"ppt/slides/_rels/{posixpath.basename(slide_path)}.rels"
                if slide_rel_path in z.namelist():
                    try:
                        s_rels_xml = z.read(slide_rel_path)
                        s_rels = ET.fromstring(s_rels_xml)

                        # Check for slide embedded image
                        for rel in s_rels.findall("pr:Relationship", OPENXML_NS):
                            rel_type = rel.get("Type", "")
                            rel_target = rel.get("Target", "")
                            if "image" in rel_type.lower() or "media" in rel_target.lower():
                                img_base = posixpath.basename(rel_target)
                                if clean_prefix and (media_dir / f"{clean_prefix}{img_base}").is_file():
                                    slide_item.image_filename = f"{clean_prefix}{img_base}"
                                    break
                                elif (media_dir / img_base).is_file():
                                    slide_item.image_filename = img_base
                                    break

                        for rel in s_rels.findall("pr:Relationship", OPENXML_NS):
                            if "notesSlide" in rel.get("Type", "") or "notesSlide" in rel.get("Target", ""):
                                raw_target = rel.get("Target", "")
                                t_clean = raw_target.lstrip("/")
                                if not t_clean.startswith("ppt/"):
                                    t_clean = posixpath.normpath(posixpath.join("ppt/slides", raw_target))
                                notes_target = t_clean
                                if notes_target in z.namelist():
                                    n_xml = z.read(notes_target)
                                    n_root = ET.fromstring(n_xml)
                                    notes_paras = []
                                    for sp in n_root.findall(".//p:sp", OPENXML_NS):
                                        ph = sp.find(".//p:ph", OPENXML_NS)
                                        if ph is not None and ph.get("type") == "body":
                                            for p in sp.findall(".//a:p", OPENXML_NS):
                                                t_vals = [t.text for t in p.findall(".//a:t", OPENXML_NS) if t.text]
                                                if t_vals:
                                                    notes_paras.append("".join(t_vals).strip())
                                    if notes_paras:
                                        slide_item.speaker_notes = "\n\n".join(notes_paras)
                    except Exception as notes_err:
                        warnings.append(f"Slide {idx}: notes extraction warning ({notes_err})")

            except Exception as slide_err:
                warnings.append(f"Slide {idx} parse failed ({slide_err})")

            slides_data.append(slide_item)

    return slides_data, warnings


def _extract_via_pdf(
    pdf_path: Path,
    media_dir: Path,
    tdoc_id: Optional[str] = None,
    image_width: int = 1920,
    image_height: int = 1080,
    cancellation_check: Optional[Callable[[], bool]] = None,
    log_fn: Optional[Callable[[str], None]] = None,
) -> Tuple[List[ExtractedSlide], List[str]]:
    """
    Extracts slide visuals (via pypdfium2 1080p rendering) and structured text (via pypdf)
    from PDF presentations and discussion decks.
    """
    slides_data: List[ExtractedSlide] = []
    warnings: List[str] = []

    media_dir = Path(strip_extended_path(str(media_dir)))
    media_dir.mkdir(parents=True, exist_ok=True)
    clean_prefix = _clean_tdoc_prefix(tdoc_id)

    # Clean up any leftover corrupted image files from previous regex bug
    try:
        for bad_f in media_dir.glob("__-________*"):
            try:
                bad_f.unlink()
            except Exception:
                pass
    except Exception:
        pass

    import pypdf
    try:
        import pypdfium2 as pdfium
    except ImportError:
        pdfium = None

    pdf_doc = pypdf.PdfReader(str(pdf_path))
    pdfium_doc = None
    if pdfium is not None:
        try:
            pdfium_doc = pdfium.PdfDocument(str(pdf_path))
        except Exception as e:
            warnings.append(f"pdfium open failed: {e}")
            pdfium_doc = None

    total_pages = len(pdf_doc.pages)
    _safe_log(f"Extracting {total_pages} slides from PDF presentation '{pdf_path.name}'...", log_fn)

    for i in range(1, total_pages + 1):
        if cancellation_check and cancellation_check():
            break

        slide_item = ExtractedSlide(index=i)
        page = pdf_doc.pages[i - 1]

        # 1. Slide Visual Rendering
        img_name = f"{clean_prefix}slide_{i:02d}.png"
        img_dest = media_dir / img_name

        rendered_visual = False
        if pdfium_doc is not None and i - 1 < len(pdfium_doc):
            try:
                pil_img = pdfium_doc[i - 1].render(scale=2).to_pil()
                pil_img.save(str(img_dest), "PNG")
                if img_dest.is_file() and img_dest.stat().st_size > 0:
                    slide_item.image_filename = img_name
                    rendered_visual = True
            except Exception as rend_err:
                warnings.append(f"Slide {i}: pdfium render failed ({rend_err})")

        # Fallback to embedded image extraction if visual rendering wasn't possible
        if not rendered_visual:
            try:
                if len(page.images) > 0:
                    first_img = page.images[0]
                    first_img_ext = Path(first_img.name).suffix or ".png"
                    fallback_name = f"{clean_prefix}slide_{i:02d}{first_img_ext}"
                    fallback_dest = media_dir / fallback_name
                    fallback_dest.write_bytes(first_img.data)
                    slide_item.image_filename = fallback_name
            except Exception as img_err:
                warnings.append(f"Slide {i}: embedded image extraction failed ({img_err})")

        # 2. Extract Text & Structured Content
        try:
            raw_text = page.extract_text() or ""
            lines = [ln.strip() for ln in raw_text.splitlines() if ln.strip()]

            title_text = ""
            body_lines = []
            for ln in lines:
                if ln == str(i):
                    continue
                if not title_text:
                    title_text = ln
                else:
                    body_lines.append(ln)

            slide_item.title = title_text or f"Slide {i}"

            for bl in body_lines:
                clean_b = bl.lstrip("•*- \t")
                if clean_b:
                    slide_item.body_paragraphs.append(f"- {clean_b}")

        except Exception as txt_err:
            warnings.append(f"Slide {i}: text extraction failed ({txt_err})")

        slides_data.append(slide_item)

    return slides_data, warnings
