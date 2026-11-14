"""
Presentation to Markdown Converter (PPTX / PPT -> Markdown).

Extracts slide visuals (screenshots / figures), slide titles, hierarchical bullet text,
tables, and speaker notes from PowerPoint presentations (.pptx and .ppt) to rich Markdown.

Conversion Engines:
1. Tier 1 (Primary Windows): Microsoft PowerPoint COM (win32com.client)
   - Exports pixel-perfect 1080p full-slide screenshots (slide_01.png, ...)
   - Extracts structured text, bullet levels, tables, and speaker notes.
2. Tier 2 (Headless / Zero-Office): Pure-Python OpenXML Parser
   - Extracts slide titles, paragraphs, tables, speaker notes, and embedded media
     directly from the .pptx zip archive without requiring PowerPoint or LibreOffice.
3. Tier 3 (Fallback): Headless LibreOffice
   - Converts legacy presentations to PDF if Microsoft PowerPoint is unavailable.
"""

from __future__ import annotations

import logging
import os
import posixpath
import re
import shutil
import tempfile
import xml.etree.ElementTree as ET
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Dict, List, Optional, Tuple, Union

from core.utils.paths import to_extended_path, strip_extended_path

logger = logging.getLogger(__name__)

CONVERTER_VERSION = "2.1.0"

TDOC_RE = re.compile(r"(?<![A-Za-z0-9])([A-Z][A-Z0-9]{0,3}-\d{6,8}(?:r\d{1,2})?)(?![0-9])", re.I)

OPENXML_NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
}


def _clean_tdoc_prefix(tdoc_id: Optional[str]) -> str:
    """Produces a clean, filesystem-safe prefix from a TDoc ID."""
    if not tdoc_id:
        return ""
    clean = re.sub(r'[^\w\-]', '_', str(tdoc_id).strip())
    clean = re.sub(r'_+', '_', clean).strip("._ ")
    return f"{clean}_" if clean else ""


@dataclass
class ExtractedSlide:
    """Representation of an extracted PowerPoint slide."""
    index: int
    title: str = ""
    body_paragraphs: List[str] = field(default_factory=list)
    tables: List[List[List[str]]] = field(default_factory=list)
    speaker_notes: str = ""
    image_filename: Optional[str] = None


def _safe_log(msg: str, log_fn: Optional[Callable[[str], None]] = None) -> None:
    """Invokes logger safely handling character encoding mismatches on Windows."""
    if not log_fn:
        return
    try:
        log_fn(msg)
    except Exception:
        try:
            clean_msg = msg.encode("ascii", errors="replace").decode("ascii")
            log_fn(clean_msg)
        except Exception:
            pass


def _format_markdown_table(rows: List[List[str]]) -> str:
    """Renders a 2D matrix of text cells as a GitHub Flavored Markdown table."""
    if not rows:
        return ""
    col_count = max(len(r) for r in rows)
    if col_count == 0:
        return ""

    norm_rows = []
    for row in rows:
        cells = [c.replace("\r", " ").replace("\n", " ").replace("|", "\\|").strip() for c in row]
        if len(cells) < col_count:
            cells.extend([""] * (col_count - len(cells)))
        norm_rows.append(cells)

    header = norm_rows[0]
    separator = [":---"] * col_count
    body = norm_rows[1:]

    lines = [
        "| " + " | ".join(header) + " |",
        "| " + " | ".join(separator) + " |",
    ]
    for row in body:
        lines.append("| " + " | ".join(row) + " |")
    return "\n".join(lines)


# =============================================================================
# TIER 1: PowerPoint COM Engine (Windows)
# =============================================================================
def _extract_via_powerpoint_com(
    ppt_path: Path,
    media_dir: Path,
    tdoc_id: Optional[str] = None,
    image_width: int = 1920,
    image_height: int = 1080,
    cancellation_check: Optional[Callable[[], bool]] = None,
    log_fn: Optional[Callable[[str], None]] = None,
) -> Tuple[List[ExtractedSlide], List[str]]:
    """
    Automates Microsoft PowerPoint COM to render high-res slide images
    and extract slide titles, body text, tables, and speaker notes.
    """
    slides_data: List[ExtractedSlide] = []
    warnings: List[str] = []

    if os.name != "nt":
        raise RuntimeError("PowerPoint COM automation is only available on Windows.")

    import pythoncom
    import win32com.client

    # Ensure destination media folder exists
    media_dir = Path(strip_extended_path(str(media_dir)))
    media_dir.mkdir(parents=True, exist_ok=True)

    # Clean up any leftover corrupted image files from previous regex bug
    try:
        for bad_f in media_dir.glob("__-________*"):
            try:
                bad_f.unlink()
            except Exception:
                pass
    except Exception:
        pass

    from core.utils.office_printer_guard import temporary_virtual_default_printer

    pythoncom.CoInitialize()
    ppt_app = None
    pres = None
    launched_own_instance = False

    try:
        with temporary_virtual_default_printer(log=logger):
            _safe_log("Connecting to Microsoft PowerPoint COM engine...", log_fn)
            try:
                ppt_app = win32com.client.DispatchEx("PowerPoint.Application")
                launched_own_instance = True
            except Exception as com_launch_err:
                try:
                    ppt_app = win32com.client.GetActiveObject("PowerPoint.Application")
                    launched_own_instance = False
                except Exception:
                    raise RuntimeError(f"Could not initialize PowerPoint COM: {com_launch_err}") from com_launch_err

            resolved_path = strip_extended_path(str(ppt_path.resolve()))
            _safe_log(f"Opening presentation '{ppt_path.name}' in PowerPoint...", log_fn)

            pres = ppt_app.Presentations.Open(
                resolved_path,
                ReadOnly=True,
                Untitled=False,
                WithWindow=False,
            )

            slide_count = pres.Slides.Count
            _safe_log(f"Exporting {slide_count} slides and extracting structured content...", log_fn)

            for i in range(1, slide_count + 1):
                if cancellation_check and cancellation_check():
                    break

                slide = pres.Slides(i)
                slide_item = ExtractedSlide(index=i)

                # 1. Slide Screenshot
                clean_prefix = _clean_tdoc_prefix(tdoc_id)
                img_name = f"{clean_prefix}slide_{i:02d}.png"
                img_dest = media_dir / img_name
                img_dest_str = strip_extended_path(str(img_dest.resolve()))
                try:
                    slide.Export(img_dest_str, "PNG", image_width, image_height)
                    slide_item.image_filename = img_name
                except Exception as exp_err:
                    try:
                        slide.Export(img_dest_str, "PNG")
                        slide_item.image_filename = img_name
                    except Exception as retry_err:
                        warnings.append(f"Slide {i}: Screenshot export failed ({retry_err})")

                # 2. Slide Title
                title_text = ""
                try:
                    if slide.Shapes.HasTitle:
                        title_text = slide.Shapes.Title.TextFrame.TextRange.Text.strip()
                except Exception:
                    pass

                # 3. Slide Shapes (Text & Tables)
                text_candidates: List[Tuple[float, str]] = []  # (top_pos, text)
                try:
                    for shape in slide.Shapes:
                        if getattr(shape, "HasTable", False):
                            try:
                                tbl = shape.Table
                                rows_data: List[List[str]] = []
                                for r in range(1, tbl.Rows.Count + 1):
                                    row_cells: List[str] = []
                                    for c in range(1, tbl.Columns.Count + 1):
                                        cell_txt = tbl.Cell(r, c).Shape.TextFrame.TextRange.Text.strip()
                                        row_cells.append(cell_txt)
                                    rows_data.append(row_cells)
                                if rows_data:
                                    slide_item.tables.append(rows_data)
                            except Exception as tbl_err:
                                warnings.append(f"Slide {i}: Table extraction error ({tbl_err})")
                            continue

                        if getattr(shape, "HasTextFrame", False) and shape.TextFrame.HasText:
                            top_pos = getattr(shape, "Top", 9999.0)
                            raw_text = shape.TextFrame.TextRange.Text.strip()
                            if not raw_text:
                                continue

                            if title_text and raw_text == title_text:
                                continue

                            text_candidates.append((top_pos, raw_text))

                            try:
                                paragraphs = shape.TextFrame.TextRange.Paragraphs()
                                for p in paragraphs:
                                    p_txt = p.Text.strip()
                                    if not p_txt:
                                        continue
                                    if title_text and p_txt == title_text:
                                        continue
                                    indent_lvl = getattr(p, "IndentLevel", 1)
                                    indent_spaces = "  " * max(0, indent_lvl - 1)
                                    slide_item.body_paragraphs.append(f"{indent_spaces}- {p_txt}")
                            except Exception:
                                for line in raw_text.splitlines():
                                    l_clean = line.strip()
                                    if l_clean:
                                        slide_item.body_paragraphs.append(f"- {l_clean}")
                except Exception as shape_err:
                    warnings.append(f"Slide {i}: Shapes inspection warning ({shape_err})")

                # Fallback heuristic for Title if HasTitle was False
                if not title_text and text_candidates:
                    text_candidates.sort(key=lambda x: x[0])
                    first_top, first_text = text_candidates[0]
                    if first_top < 150.0 and len(first_text.splitlines()) <= 2:
                        title_text = first_text.replace("\n", " ")
                        cleaned_body = []
                        for b in slide_item.body_paragraphs:
                            if b.lstrip("- ").strip() != first_text:
                                cleaned_body.append(b)
                        slide_item.body_paragraphs = cleaned_body

                slide_item.title = title_text

                # 4. Speaker Notes
                try:
                    if slide.HasNotesPage:
                        notes_page = slide.NotesPage
                        for n_shape in notes_page.Shapes:
                            if getattr(n_shape, "HasTextFrame", False) and n_shape.TextFrame.HasText:
                                is_body = False
                                try:
                                    if getattr(n_shape, "Type", 0) == 14:  # msoPlaceholder
                                        if n_shape.PlaceholderFormat.Type == 2:  # ppPlaceholderBody
                                            is_body = True
                                except Exception:
                                    pass

                                txt = n_shape.TextFrame.TextRange.Text.strip()
                                if is_body and txt:
                                    slide_item.speaker_notes = txt
                                    break
                                elif not slide_item.speaker_notes and txt and not txt.isdigit():
                                    slide_item.speaker_notes = txt
                except Exception as notes_err:
                    warnings.append(f"Slide {i}: Speaker notes extraction skipped ({notes_err})")

                slides_data.append(slide_item)

    finally:
        if pres is not None:
            try:
                pres.Close()
            except Exception:
                pass
        if ppt_app is not None and launched_own_instance:
            try:
                ppt_app.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()

    return slides_data, warnings


# =============================================================================
# TIER 2: Headless Pure-Python OpenXML Parser (.pptx only)
# =============================================================================
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



# =============================================================================
# TIER 3: PDF Slide Presentation Engine (pypdfium2 visual + pypdf structured text)
# =============================================================================
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


# =============================================================================
# Markdown Assembly & Document Construction
# =============================================================================
def _build_presentation_header(
    tdoc_id: str,
    meta: Dict[str, str],
    title: str,
    original_filename: str,
    slide_count: int,
    tdoc_url: Optional[str] = None,
) -> str:
    """Builds the 3GPP document header or generic presentation header."""
    is_tdoc = bool(
        TDOC_RE.search(tdoc_id)
        or (meta and any(k in meta for k in ("tdoc", "tdoc_id", "agenda item", "work item", "spec", "source")))
    )

    doc_title = meta.get("Title") or meta.get("title") or title or tdoc_id
    doc_source = meta.get("Source") or meta.get("source") or ""
    doc_type = meta.get("Type") or meta.get("type") or "Discussion Paper (Presentation)"
    doc_ai = meta.get("Agenda Item") or meta.get("agenda item") or ""
    doc_wi = meta.get("Work Item") or meta.get("work item") or ""
    doc_status = meta.get("Status") or meta.get("status") or meta.get("tdoc status") or ""

    lines = []
    if is_tdoc:
        header_title = f"# {tdoc_id}: {doc_title}" if doc_title else f"# {tdoc_id}"
        lines.append(header_title)
        lines.append("")
        fields = [
            ("TDoc", tdoc_id),
            ("Title", doc_title),
            ("Source", doc_source),
            ("Type", doc_type),
            ("Agenda Item", doc_ai),
            ("Work Item", doc_wi),
            ("Status", doc_status),
            ("Slides", f"{slide_count} slides"),
            ("Original file", f"`{original_filename}`"),
        ]
        if tdoc_url:
            fields.append(("Authoritative 3GPP Portal Link", f"[{tdoc_id}]({tdoc_url})"))

        for k, v in fields:
            if v:
                lines.append(f"- **{k}:** {v.replace(chr(10), ' ')}")
        lines.append("")
    else:
        lines.append(f"# {doc_title or tdoc_id}")
        lines.append("")
        lines.append(f"- **Original file:** `{original_filename}`")
        lines.append(f"- **Slides:** {slide_count} slides")
        lines.append("")

    lines.append("---")
    lines.append("")
    return "\n".join(lines)


def convert_presentation_to_markdown(
    ppt_path: Union[str, Path],
    out_dir: Optional[Union[str, Path]] = None,
    tdoc_id: Optional[str] = None,
    meta: Optional[Dict[str, str]] = None,
    media_dir: Optional[Path] = None,
    media_rel: Optional[str] = None,
    cancellation_check: Optional[Callable[[], bool]] = None,
    tdoc_url: Optional[str] = None,
    include_header: bool = True,
    logger_fn: Optional[Callable[[str], None]] = None,
) -> "ExportedDoc":
    """
    Primary Entry Point: Converts a PowerPoint presentation (.pptx or .ppt)
    into a structured Markdown document with slide screenshots, bullet points,
    tables, and speaker notes.

    Returns:
        ExportedDoc: Compatible with 3GPP Tools TDoc exporter and AI triage pipelines.
    """
    from modules.meetings.core.exporters.tdoc2md import ExportedDoc, sanitize_filename, guess_tdoc_id

    path = Path(to_extended_path(ppt_path))
    if not path.is_file():
        raise FileNotFoundError(f"Presentation document not found: {path}")

    tid = tdoc_id or guess_tdoc_id(path)
    target_dir = Path(to_extended_path(out_dir)) if out_dir else path.parent
    target_dir.mkdir(parents=True, exist_ok=True)

    companion_media = Path(to_extended_path(media_dir)) if media_dir is not None else target_dir / f"{path.stem}_media"
    companion_media.mkdir(parents=True, exist_ok=True)
    rel_media_path = media_rel if media_rel is not None else f"{path.stem}_media"

    suffix = path.suffix.lower()
    slides: List[ExtractedSlide] = []
    warnings: List[str] = []

    # -------------------------------------------------------------------------
    # Conversion Dispatch: PDF, COM (PPTX/PPT), OpenXML Fallback
    # -------------------------------------------------------------------------
    if suffix == ".pdf":
        _safe_log(f"Extracting presentation slides and visuals from PDF '{path.name}'...", logger_fn)
        slides, pdf_warnings = _extract_via_pdf(
            path,
            media_dir=companion_media,
            tdoc_id=tid,
            cancellation_check=cancellation_check,
            log_fn=logger_fn,
        )
        warnings.extend(pdf_warnings)
    else:
        com_succeeded = False
        if os.name == "nt":
            try:
                _safe_log(f"Initiating PowerPoint COM visual extraction for '{path.name}'...", logger_fn)
                slides, com_warnings = _extract_via_powerpoint_com(
                    path,
                    media_dir=companion_media,
                    tdoc_id=tid,
                    cancellation_check=cancellation_check,
                    log_fn=logger_fn,
                )
                warnings.extend(com_warnings)
                com_succeeded = True
                _safe_log(f"Extracted {len(slides)} slide screenshots and content via PowerPoint COM.", logger_fn)
            except Exception as com_err:
                _safe_log(f"PowerPoint COM extraction unavailable or failed ({com_err}). Trying fallback engines...", logger_fn)
                warnings.append(f"PowerPoint COM failed: {com_err}")

        if not com_succeeded:
            if suffix == ".pptx":
                _safe_log(f"Extracting structured content via OpenXML direct package parser...", logger_fn)
                try:
                    slides, ox_warnings = _extract_via_openxml(
                        path,
                        media_dir=companion_media,
                        tdoc_id=tid,
                        cancellation_check=cancellation_check,
                        log_fn=logger_fn,
                    )
                    warnings.extend(ox_warnings)
                    _safe_log(f"Extracted {len(slides)} slides via OpenXML parser.", logger_fn)
                except Exception as ox_err:
                    raise RuntimeError(f"Failed to parse presentation '{path.name}': {ox_err}") from ox_err
            elif suffix == ".ppt":
                from modules.word_tools.core.libreoffice_converter import is_libreoffice_available, convert_document_libreoffice
                if is_libreoffice_available():
                    _safe_log("Converting legacy .ppt to PDF using Headless LibreOffice...", logger_fn)
                    pdf_path = convert_document_libreoffice(path, target_format="pdf")
                    slides, pdf_warnings = _extract_via_pdf(
                        pdf_path,
                        media_dir=companion_media,
                        tdoc_id=tid,
                        cancellation_check=cancellation_check,
                        log_fn=logger_fn,
                    )
                    warnings.extend(pdf_warnings)
                else:
                    raise RuntimeError("Legacy binary .ppt requires Microsoft PowerPoint (on Windows) or LibreOffice.")

    # -------------------------------------------------------------------------
    # Assemble Markdown Document
    # -------------------------------------------------------------------------
    first_title = slides[0].title if slides else path.stem
    header_text = ""
    if include_header:
        header_text = _build_presentation_header(
            tdoc_id=tid,
            meta=meta or {},
            title=first_title,
            original_filename=path.name,
            slide_count=len(slides),
            tdoc_url=tdoc_url,
        )

    md_lines: List[str] = [header_text] if header_text else []

    exported_images_count = 0
    for s in slides:
        slide_title = s.title.strip() if s.title else f"Slide {s.index}"
        md_lines.append(f"## Slide {s.index}: {slide_title}")
        md_lines.append("")

        if s.image_filename:
            img_rel = f"{rel_media_path}/{s.image_filename}".replace("\\", "/")
            md_lines.append(f"![Slide {s.index}: {slide_title}]({img_rel})")
            md_lines.append("")
            exported_images_count += 1

        if s.body_paragraphs:
            for p in s.body_paragraphs:
                md_lines.append(p)
            md_lines.append("")

        for tbl_data in s.tables:
            tbl_md = _format_markdown_table(tbl_data)
            if tbl_md:
                md_lines.append(tbl_md)
                md_lines.append("")

        if s.speaker_notes:
            md_lines.append("> [!NOTE] Speaker Notes")
            for n_line in s.speaker_notes.splitlines():
                clean_n = n_line.strip()
                if clean_n:
                    md_lines.append(f"> {clean_n}")
            md_lines.append("")

        md_lines.append("---")
        md_lines.append("")

    full_markdown = "\n".join(md_lines)
    stamp = f"<!--- tdoc_converter: version={CONVERTER_VERSION}; type=presentation; mode=presentation --->\n"
    if not full_markdown.startswith("<!--- tdoc_converter:"):
        full_markdown = stamp + full_markdown

    clean_name = sanitize_filename(tid or path.stem)
    md_output_path = target_dir / f"{clean_name}.md"
    md_output_path.write_text(full_markdown, encoding="utf-8")

    _safe_log(f"Wrote Markdown presentation export to '{md_output_path.name}' ({len(slides)} slides, {exported_images_count} figures).", logger_fn)

    doc_type = (meta or {}).get("Type") or (meta or {}).get("type") or "Discussion Paper"
    doc_ai = (meta or {}).get("Agenda Item") or (meta or {}).get("agenda item") or ""
    doc_status = (meta or {}).get("Status") or (meta or {}).get("status") or ""

    return ExportedDoc(
        tdoc_id=tid,
        name=clean_name,
        md_path=md_output_path,
        text=full_markdown,
        meta=meta or {},
        doc_type=doc_type,
        agenda_item=doc_ai,
        title=first_title or tid,
        status=doc_status,
        images=exported_images_count,
        warnings=warnings,
    )
