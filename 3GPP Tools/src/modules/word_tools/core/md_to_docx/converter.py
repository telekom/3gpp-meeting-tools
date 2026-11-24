"""Core OpenXML synthesis engine for converting Markdown documents to .docx."""

from __future__ import annotations

import logging
import re
import tempfile
from pathlib import Path
from typing import Any, Callable, Dict, Optional, Union

import docx
import markdown
from bs4 import BeautifulSoup, Tag
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.image.image import Image as DocxImage
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.shared import Inches, Pt, RGBColor

from modules.word_tools.core.md_to_docx.inline_and_lists import (
    _embed_image_tag,
    _process_list,
    add_runs_from_element,
)
from modules.word_tools.core.md_to_docx.media_renderer import (
    CALLOUT_CONFIG,
    PLANTUML_START_DIRECTIVES,
    _extract_plantuml_title,
    render_plantuml_to_png,
)
from modules.word_tools.core.sensitivity_label import set_sensitivity_label
from modules.word_tools.core.styles_3gpp import ensure_3gpp_package_styles

logger = logging.getLogger(__name__)


def convert_markdown_to_docx(
    md_path: Union[str, Path],
    output_docx_path: Optional[Union[str, Path]] = None,
    companion_media_dir: Optional[Union[str, Path]] = None,
    auto_detect_media: bool = True,
    render_diagrams: bool = True,
    apply_3gpp_styles: bool = True,
    max_image_width_in: float = 6.0,
    log_callback: Optional[Callable[[str], None]] = None,
) -> Dict[str, Any]:
    """
    Converts a Markdown document to Microsoft Word (.docx).
    Embeds all referenced media, compiles PlantUML diagrams on the fly,
    formats tables and callouts, and applies 3GPP contribution styles.
    """
    md = Path(md_path).resolve()
    if not md.is_file():
        raise FileNotFoundError(f"Markdown file does not exist: {md}")

    def log(msg: str):
        if log_callback:
            log_callback(msg)
        logger.info(msg)

    if not output_docx_path:
        out_docx = md.with_suffix(".docx")
    else:
        out_docx = Path(output_docx_path).resolve()

    comp_dir: Optional[Path] = None
    if companion_media_dir:
        comp_dir = Path(companion_media_dir).resolve()
    elif auto_detect_media:
        auto_cand = md.parent / f"{md.stem}_media"
        if auto_cand.is_dir():
            comp_dir = auto_cand

    log(f"⏳ Reading Markdown from '{md.name}'...")
    raw_text = md.read_text(encoding="utf-8", errors="ignore")
    # Normalize Windows line endings to prevent regex mismatch in fenced_code
    normalized_text = raw_text.replace("\r\n", "\n")

    # Parse Markdown into HTML with standard extensions
    html = markdown.markdown(
        normalized_text,
        extensions=["tables", "fenced_code", "sane_lists"],
    )
    soup = BeautifulSoup(html, "html.parser")

    doc = docx.Document()

    # Configure baseline document typography
    normal_style = doc.styles["Normal"]
    normal_style.font.name = "Arial"
    normal_style.font.size = Pt(10)
    normal_style.font.color.rgb = RGBColor(0x1E, 0x29, 0x3B)

    stats = {
        "headings": 0,
        "paragraphs": 0,
        "tables": 0,
        "images": 0,
        "diagrams": 0,
    }

    temp_cache_dir = tempfile.TemporaryDirectory()
    temp_cache_path = Path(temp_cache_dir.name)

    try:
        log("⚙️ Synthesizing Word document structure...")
        for element in soup.children:
            if not isinstance(element, Tag):
                continue
            tag = element.name.lower()

            # 1. Headings (H1 - H6)
            if tag in ("h1", "h2", "h3", "h4", "h5", "h6"):
                level = int(tag[1])
                h_para = doc.add_heading(level=level)
                h_para.paragraph_format.space_before = Pt(12)
                h_para.paragraph_format.space_after = Pt(4)
                h_para.paragraph_format.keep_with_next = True
                add_runs_from_element(h_para, element)
                stats["headings"] += 1

            # 2. Paragraphs & Figures
            elif tag == "p":
                img_tags = element.find_all("img")
                if img_tags:
                    rem_text = element.get_text().strip()
                    has_inline_caption = bool(rem_text) and len(img_tags) == 1
                    for img_tag in img_tags:
                        _embed_image_tag(
                            doc,
                            img_tag,
                            md,
                            comp_dir,
                            auto_detect_media,
                            temp_cache_path,
                            max_image_width_in,
                            stats,
                            suppress_caption=has_inline_caption,
                        )
                    if rem_text:
                        para = doc.add_paragraph()
                        if has_inline_caption:
                            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            para.paragraph_format.space_before = Pt(0)
                            para.paragraph_format.space_after = Pt(8)
                            add_runs_from_element(
                                para,
                                element,
                                base_style={"italic": True, "size": Pt(9), "color": RGBColor(0x64, 0x74, 0x8B)},
                            )
                        else:
                            para.paragraph_format.space_after = Pt(6)
                            para.paragraph_format.line_spacing = 1.15
                            add_runs_from_element(para, element)
                        stats["paragraphs"] += 1
                else:
                    para = doc.add_paragraph()
                    para.paragraph_format.space_after = Pt(6)
                    para.paragraph_format.line_spacing = 1.15
                    add_runs_from_element(para, element)
                    stats["paragraphs"] += 1

            # 3. Blockquotes & GitHub-style Callouts
            elif tag == "blockquote":
                full_text = element.get_text().strip()
                callout_type = None
                callout_title = None
                body_text = full_text

                first_line = full_text.split("\n")[0].strip() if full_text else ""
                match = re.match(
                    r"^\[!(NOTE|TIP|IMPORTANT|WARNING|CAUTION)\](?:\s+(.*))?$",
                    first_line,
                    re.IGNORECASE,
                )
                if match:
                    callout_type = match.group(1).upper()
                    custom_title = match.group(2)
                    border_color, fill_color, default_lbl = CALLOUT_CONFIG.get(
                        callout_type, ("#0284C7", "F0F9FF", "NOTE")
                    )
                    callout_title = custom_title if custom_title else default_lbl
                    body_lines = full_text.split("\n")[1:]
                    body_text = "\n".join(body_lines).strip()
                else:
                    border_color = "#CBD5E1"
                    fill_color = "F8FAFC"

                p = doc.add_paragraph()
                p.paragraph_format.left_indent = Inches(0.4)
                p.paragraph_format.right_indent = Inches(0.2)
                p.paragraph_format.space_before = Pt(6)
                p.paragraph_format.space_after = Pt(6)

                pPr = p._p.get_or_add_pPr()
                bdr_hex = border_color.lstrip("#")
                bdr_xml = f'<w:pBdr {nsdecls("w")}><w:left w:val="single" w:sz="24" w:space="8" w:color="{bdr_hex}"/></w:pBdr>'
                shd_xml = f'<w:shd {nsdecls("w")} w:fill="{fill_color}"/>'
                pPr.append(parse_xml(bdr_xml))
                pPr.append(parse_xml(shd_xml))

                if callout_title:
                    r_title = p.add_run(f"{callout_title}\n")
                    r_title.bold = True
                    r_title.font.size = Pt(9.5)
                    r_title.font.color.rgb = RGBColor.from_string(bdr_hex)
                    if body_text:
                        r_body = p.add_run(body_text)
                        r_body.font.size = Pt(9.5)
                else:
                    # Standard blockquote: preserve rich inline formatting
                    p_children = element.find_all("p", recursive=False)
                    if p_children:
                        for idx_p, p_child in enumerate(p_children):
                            if idx_p > 0:
                                p.add_run().add_break()
                            add_runs_from_element(p, p_child, base_style={"size": Pt(9.5), "italic": True})
                    else:
                        add_runs_from_element(p, element, base_style={"size": Pt(9.5), "italic": True})

                stats["paragraphs"] += 1

            # 4. Tables
            elif tag == "table":
                rows = element.find_all("tr")
                if not rows:
                    continue

                max_cols = max(len(tr.find_all(["th", "td"])) for tr in rows)
                tbl = doc.add_table(rows=len(rows), cols=max_cols)
                tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
                tbl.style = "Table Grid"

                for r_idx, tr in enumerate(rows):
                    is_header = bool(tr.find("th")) or (r_idx == 0 and bool(element.find("thead")))
                    cells = tr.find_all(["th", "td"])

                    trPr = tbl.rows[r_idx]._tr.get_or_add_trPr()
                    if is_header:
                        trPr.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))

                    for c_idx, cell in enumerate(cells):
                        if c_idx >= max_cols:
                            break
                        doc_cell = tbl.cell(r_idx, c_idx)
                        p_cell = doc_cell.paragraphs[0]
                        p_cell.paragraph_format.space_before = Pt(2)
                        p_cell.paragraph_format.space_after = Pt(2)

                        base_style = {"bold": True, "size": Pt(9.5)} if is_header else {"size": Pt(9.5)}
                        add_runs_from_element(p_cell, cell, base_style)

                        tcPr = doc_cell._tc.get_or_add_tcPr()
                        if is_header:
                            tcPr.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="F1F5F9"/>'))
                            tcPr.append(
                                parse_xml(
                                    f'<w:tcBorders {nsdecls("w")}><w:bottom w:val="single" w:sz="12" w:space="0" w:color="94A3B8"/></w:tcBorders>'
                                )
                            )
                        else:
                            tcPr.append(
                                parse_xml(
                                    f'<w:tcBorders {nsdecls("w")}><w:bottom w:val="single" w:sz="4" w:space="0" w:color="E2E8F0"/></w:tcBorders>'
                                )
                            )

                doc.add_paragraph().paragraph_format.space_after = Pt(6)
                stats["tables"] += 1

            # 5. Lists (Unordered & Ordered, including nested sub-lists)
            elif tag in ("ul", "ol"):
                _process_list(
                    doc,
                    element,
                    md,
                    comp_dir,
                    auto_detect_media,
                    temp_cache_path,
                    max_image_width_in,
                    stats,
                    level=0,
                )

            # 6. Fenced Code Blocks & PlantUML Diagrams (<pre>)
            elif tag == "pre":
                code_tag = element.find("code")
                raw_code = code_tag.get_text() if code_tag else element.get_text()
                classes = code_tag.get("class", []) if code_tag else []
                lang = ""
                for cls in classes:
                    if cls.startswith("language-"):
                        lang = cls.replace("language-", "").lower()
                        break

                stripped_code = raw_code.strip()
                is_puml = lang in ("plantuml", "puml", "uml") or stripped_code.lower().startswith(
                    PLANTUML_START_DIRECTIVES
                )

                # PlantUML diagram rendering
                if render_diagrams and is_puml:
                    log(f"🎨 Rendering PlantUML diagram #{stats['diagrams'] + 1}...")
                    png_name = f"diagram_{stats['diagrams'] + 1}.png"
                    diag_png = temp_cache_path / png_name
                    rendered = render_plantuml_to_png(raw_code, diag_png)
                    if rendered and diag_png.exists():
                        try:
                            img_obj = DocxImage.from_file(str(diag_png))
                            max_w = Inches(max_image_width_in)
                            target_w = max_w if img_obj.width > max_w else img_obj.width

                            p_diag = doc.add_paragraph()
                            p_diag.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            p_diag.paragraph_format.space_before = Pt(8)
                            p_diag.paragraph_format.space_after = Pt(4)
                            r_diag = p_diag.add_run()
                            r_diag.add_picture(str(diag_png), width=target_w)

                            extracted_title = _extract_plantuml_title(raw_code)
                            cap_label = (
                                f"Figure: {extracted_title}"
                                if extracted_title
                                else f"Figure: Sequence Diagram ({stats['diagrams'] + 1})"
                            )

                            p_cap = doc.add_paragraph()
                            p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            p_cap.paragraph_format.space_before = Pt(0)
                            p_cap.paragraph_format.space_after = Pt(8)
                            r_cap = p_cap.add_run(cap_label)
                            r_cap.italic = True
                            r_cap.font.size = Pt(9)
                            r_cap.font.color.rgb = RGBColor(0x64, 0x74, 0x8B)

                            stats["diagrams"] += 1
                            continue
                        except Exception as e:
                            log(f"⚠️ Diagram embedding error: {e}")

                # Standard code block fallback
                p_code = doc.add_paragraph()
                p_code.paragraph_format.left_indent = Inches(0.2)
                p_code.paragraph_format.space_before = Pt(4)
                p_code.paragraph_format.space_after = Pt(6)

                pPr = p_code._p.get_or_add_pPr()
                pPr.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="F8FAFC"/>'))
                pPr.append(
                    parse_xml(
                        f'<w:pBdr {nsdecls("w")}><w:left w:val="single" w:sz="12" w:space="6" w:color="CBD5E1"/></w:pBdr>'
                    )
                )

                r_code = p_code.add_run(raw_code.rstrip())
                r_code.font.name = "Consolas"
                r_code.font.size = Pt(9)
                r_code.font.color.rgb = RGBColor(0x33, 0x41, 0x55)
                stats["paragraphs"] += 1

            # 7. Horizontal Rules (<hr>)
            elif tag == "hr":
                p_hr = doc.add_paragraph()
                p_hr.paragraph_format.space_before = Pt(6)
                p_hr.paragraph_format.space_after = Pt(6)
                pPr = p_hr._p.get_or_add_pPr()
                pPr.append(
                    parse_xml(
                        f'<w:pBdr {nsdecls("w")}><w:bottom w:val="single" w:sz="6" w:space="1" w:color="CBD5E1"/></w:pBdr>'
                    )
                )

        out_docx.parent.mkdir(parents=True, exist_ok=True)
        doc.save(str(out_docx))
        log(f"💾 Saved base Word document to '{out_docx.name}'.")

        # 8. Apply canonical 3GPP styles (TableGrid, heading hierarchy)
        if apply_3gpp_styles:
            try:
                ensure_3gpp_package_styles(out_docx)
                log("🏛️ Injected canonical 3GPP styles & TableGrid definitions.")
            except Exception as e:
                logger.warning(f"Could not apply 3GPP package styles: {e}")

        # 9. Apply corporate sensitivity labels
        try:
            set_sensitivity_label(out_docx)
        except Exception:
            pass

    finally:
        temp_cache_dir.cleanup()

    log(
        f"✅ Markdown conversion complete: {out_docx.name} "
        f"({stats['headings']} headings, {stats['paragraphs']} paragraphs, "
        f"{stats['tables']} tables, {stats['images']} images, {stats['diagrams']} diagrams)"
    )

    return {
        "output_path": str(out_docx),
        "stats": stats,
    }
