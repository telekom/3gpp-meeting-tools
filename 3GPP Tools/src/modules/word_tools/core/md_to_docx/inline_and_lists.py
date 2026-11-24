"""Inline run styling, hyperlink synthesis, image embedding, and nested list rendering."""

from __future__ import annotations

from pathlib import Path
from typing import Any, Dict, Optional, Tuple

from bs4 import NavigableString, Tag
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.image.image import Image as DocxImage
from docx.opc.constants import RELATIONSHIP_TYPE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

from modules.word_tools.core.md_to_docx.media_renderer import (
    render_metafile_to_png,
    render_svg_to_png,
    resolve_image_path,
)


def _apply_run_style(run, style: Dict[str, Any]) -> None:
    """Applies formatting attributes from a style dictionary to a python-docx Run."""
    if style.get("bold"):
        run.bold = True
    if style.get("italic"):
        run.italic = True
    if style.get("underline"):
        run.underline = True
    if style.get("strike"):
        run.font.strike = True
    if style.get("superscript"):
        run.font.superscript = True
    if style.get("subscript"):
        run.font.subscript = True
    if style.get("code"):
        run.font.name = "Consolas"
        run.font.size = Pt(9.5)
        run.font.color.rgb = RGBColor(0xBE, 0x18, 0x5D)
    if style.get("color"):
        run.font.color.rgb = style["color"]
    if style.get("size"):
        run.font.size = style["size"]
    if style.get("font_name"):
        run.font.name = style["font_name"]


def add_runs_from_element(
    para,
    element,
    base_style: Optional[Dict[str, Any]] = None,
    skip_tags: Tuple[str, ...] = (),
) -> None:
    """Recursively processes HTML children and adds formatted text runs (and clickable hyperlinks) to a Word paragraph."""
    base_style = base_style or {}

    for child in element.children:
        if isinstance(child, NavigableString):
            text = str(child)
            if not text:
                continue
            run = para.add_run(text)
            _apply_run_style(run, base_style)
            if base_style.get("hyperlink_rid"):
                hl = OxmlElement("w:hyperlink")
                hl.set(qn("r:id"), base_style["hyperlink_rid"])
                hl.append(run._r)
                para._p.append(hl)

        elif isinstance(child, Tag):
            tag_name = child.name.lower()
            if tag_name in skip_tags or tag_name == "img":
                continue
            if tag_name == "br":
                para.add_run().add_break()
                continue

            new_style = dict(base_style)
            if tag_name in ("b", "strong"):
                new_style["bold"] = True
            elif tag_name in ("i", "em"):
                new_style["italic"] = True
            elif tag_name in ("u", "ins"):
                new_style["underline"] = True
            elif tag_name in ("s", "del", "strike"):
                new_style["strike"] = True
            elif tag_name == "sup":
                new_style["superscript"] = True
            elif tag_name == "sub":
                new_style["subscript"] = True
            elif tag_name == "code":
                new_style["code"] = True
            elif tag_name == "a":
                href = (child.get("href") or "").strip()
                new_style["color"] = RGBColor(0x02, 0x84, 0xC7)
                new_style["underline"] = True
                if href.startswith(("http://", "https://", "mailto:", "ftp://")):
                    try:
                        r_id = para.part.relate_to(href, RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
                        new_style["hyperlink_rid"] = r_id
                    except Exception:
                        pass

            add_runs_from_element(para, child, new_style, skip_tags=skip_tags)


def _embed_image_tag(
    doc,
    img_tag: Tag,
    md_path: Path,
    comp_dir: Optional[Path],
    auto_detect_media: bool,
    temp_cache_path: Path,
    max_image_width_in: float,
    stats: Dict[str, int],
    suppress_caption: bool = False,
) -> None:
    """Resolves, converts (if SVG/EMF/WMF), scales, and embeds a single <img> tag into the document."""
    src = img_tag.get("src", "")
    alt = img_tag.get("alt", "")
    title = img_tag.get("title", "")

    resolved = resolve_image_path(src, md_path, comp_dir, auto_detect_media=auto_detect_media)
    img_to_embed = resolved

    if resolved:
        ext_lower = resolved.suffix.lower()
        if ext_lower in (".svg", ".svgz"):
            svg_png = temp_cache_path / f"svg_{stats['images'] + 1}.png"
            if render_svg_to_png(resolved, svg_png):
                img_to_embed = svg_png
        elif ext_lower in (".emf", ".wmf"):
            meta_png = temp_cache_path / f"meta_{stats['images'] + 1}.png"
            if render_metafile_to_png(resolved, meta_png):
                img_to_embed = meta_png

    if img_to_embed and img_to_embed.is_file():
        try:
            img_obj = DocxImage.from_file(str(img_to_embed))
            max_w = Inches(max_image_width_in)
            target_w = max_w if img_obj.width > max_w else img_obj.width

            p_img = doc.add_paragraph()
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_img.paragraph_format.space_before = Pt(8)
            p_img.paragraph_format.space_after = Pt(4)
            r_img = p_img.add_run()
            r_img.add_picture(str(img_to_embed), width=target_w)
            stats["images"] += 1

            caption_text = title or alt
            if caption_text and not suppress_caption:
                p_cap = doc.add_paragraph()
                p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_cap.paragraph_format.space_before = Pt(0)
                p_cap.paragraph_format.space_after = Pt(8)
                r_cap = p_cap.add_run(caption_text)
                r_cap.italic = True
                r_cap.font.size = Pt(9)
                r_cap.font.color.rgb = RGBColor(0x64, 0x74, 0x8B)
        except Exception as e:
            p_err = doc.add_paragraph(f"[Image: {alt or src} (render error: {e})]")
            p_err.runs[0].font.color.rgb = RGBColor(0xDC, 0x26, 0x26)
    else:
        p_miss = doc.add_paragraph(f"[Image not found: {alt or src}]")
        p_miss.runs[0].font.color.rgb = RGBColor(0xDC, 0x26, 0x26)


def _process_list(
    doc,
    list_element: Tag,
    md_path: Path,
    comp_dir: Optional[Path],
    auto_detect_media: bool,
    temp_cache_path: Path,
    max_image_width_in: float,
    stats: Dict[str, int],
    level: int = 0,
) -> None:
    """Recursively renders <ul> and <ol> lists while preserving nested sub-list hierarchy."""
    is_num = list_element.name.lower() == "ol"
    style_name = "List Number" if is_num else "List Bullet"

    for li in list_element.find_all("li", recursive=False):
        p_li = doc.add_paragraph(style=style_name)
        p_li.paragraph_format.space_after = Pt(3)
        p_li.paragraph_format.line_spacing = 1.15
        if level > 0:
            p_li.paragraph_format.left_indent = Inches(0.25 * (level + 1))

        # Add runs from the list item excluding nested <ul>/<ol> sub-lists
        add_runs_from_element(p_li, li, skip_tags=("ul", "ol"))
        stats["paragraphs"] += 1

        # Embed any images inside the list item
        for img_tag in li.find_all("img", recursive=False):
            _embed_image_tag(
                doc,
                img_tag,
                md_path,
                comp_dir,
                auto_detect_media,
                temp_cache_path,
                max_image_width_in,
                stats,
            )

        # Recursively process nested lists
        for sub_list in li.find_all(["ul", "ol"], recursive=False):
            _process_list(
                doc,
                sub_list,
                md_path,
                comp_dir,
                auto_detect_media,
                temp_cache_path,
                max_image_width_in,
                stats,
                level=level + 1,
            )
