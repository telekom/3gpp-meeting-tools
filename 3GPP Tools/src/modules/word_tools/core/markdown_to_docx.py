"""
Core engine and background worker for converting Markdown documents to Microsoft Word (.docx).

Features:
- Headless OpenXML document synthesis using python-docx and lxml.
- Full support for headings (H1-H6), rich text runs (bold, italic, underline, strike, sup/sub, code, clickable hyperlinks, line breaks).
- Recursive nested list rendering (ordered & unordered lists without flattening sub-items).
- Table conversion with canonical 3GPP TableGrid borders, shaded header rows, and page-repeat headers.
- GitHub-style callout alert boxes (> [!NOTE], [!TIP], [!IMPORTANT], [!WARNING], [!CAUTION]) preserving inline formatting.
- Automatic embedded figure and media resolution (relative paths, companion {stem}_media/ folder).
- Image dimension safety: auto-scaling wide figures to fit standard printable page margins.
- On-the-fly vector & metafile conversion: renders SVG images via QtSvg and EMF/WMF via metafile converter to high-resolution PNGs.
- On-the-fly diagram rendering: compiles PlantUML fenced code blocks (```plantuml) into embedded figures with extracted diagram titles.
- Standard 3GPP styles injection (ensure_3gpp_package_styles) and enterprise sensitivity labeling.
"""

from __future__ import annotations

import logging
import os
import re
import shutil
import subprocess
import tempfile
import urllib.parse
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Tuple, Union

from PyQt5.QtCore import QByteArray, QObject, QSize, QThread, pyqtSignal
from PyQt5.QtGui import QImage, QPainter
from PyQt5.QtSvg import QSvgRenderer

import docx
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.image.image import Image as DocxImage
from docx.opc.constants import RELATIONSHIP_TYPE
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Inches, Pt, RGBColor
import markdown
from bs4 import BeautifulSoup, NavigableString, Tag

from core.network.session import NetworkSession
from core.utils.paths import get_project_root, strip_extended_path, to_extended_path
from core.utils.utils import get_best_java
from modules.puml2visio.config.paths import PLANTUML_JAR_NAME
from modules.puml2visio.utils.paths import get_puml2visio_asset_path
from modules.word_tools.core.document_opener import open_document
from modules.word_tools.core.sensitivity_label import set_sensitivity_label
from modules.word_tools.core.styles_3gpp import ensure_3gpp_package_styles

logger = logging.getLogger(__name__)

# Callout styling configuration: (border_hex, fill_hex, default_title)
CALLOUT_CONFIG = {
    "NOTE": ("#0284C7", "F0F9FF", "ℹ️ NOTE"),
    "TIP": ("#10B981", "ECFDF5", "💡 TIP"),
    "IMPORTANT": ("#8B5CF6", "F5F3FF", "❗ IMPORTANT"),
    "WARNING": ("#F59E0B", "FFFBEB", "⚠️ WARNING"),
    "CAUTION": ("#EF4444", "FEF2F2", "🛑 CAUTION"),
}

PLANTUML_START_DIRECTIVES = (
    "@startuml",
    "@startmindmap",
    "@startwbs",
    "@startgantt",
    "@startjson",
    "@startyaml",
    "@startditaa",
    "@startsalt",
)


def render_svg_to_png(svg_path: Path, output_png_path: Path, max_dimension: int = 1600) -> bool:
    """Renders an SVG file to a high-resolution PNG image using QtSvg."""
    try:
        svg_bytes = svg_path.read_bytes()
        renderer = QSvgRenderer(QByteArray(svg_bytes))
        if not renderer.isValid():
            return False

        default_size = renderer.defaultSize()
        w = default_size.width() if default_size.width() > 0 else 800
        h = default_size.height() if default_size.height() > 0 else 600

        # Scale up proportionally for crisp print resolution if small
        if w < 800 and h < 800:
            scale = min(max_dimension / max(w, 1), 2.0)
            w = int(w * scale)
            h = int(h * scale)

        img = QImage(QSize(w, h), QImage.Format_ARGB32)
        img.fill(0)  # Transparent background

        painter = QPainter(img)
        renderer.render(painter)
        painter.end()

        output_png_path.parent.mkdir(parents=True, exist_ok=True)
        return img.save(str(output_png_path), "PNG")
    except Exception as e:
        logger.debug(f"SVG to PNG rendering failed for '{svg_path}': {e}")
        return False


def render_metafile_to_png(metafile_path: Path, output_png_path: Path) -> bool:
    """Converts a Windows Metafile (.emf / .wmf) to PNG using the application metafile converter."""
    try:
        from modules.meetings.core.tdoc_markdown_converter import convert_metafiles

        output_png_path.parent.mkdir(parents=True, exist_ok=True)
        convert_metafiles([(metafile_path, output_png_path)], dpi=150)
        return output_png_path.is_file() and output_png_path.stat().st_size > 0
    except Exception as e:
        logger.debug(f"Metafile to PNG conversion failed for '{metafile_path}': {e}")
        return False


def _extract_plantuml_title(puml_code: str) -> Optional[str]:
    """Extracts an explicit diagram title from PlantUML source code if defined."""
    for line in puml_code.splitlines():
        s = line.strip()
        m_title = re.match(r"^title\s+(.+)$", s, re.IGNORECASE)
        if m_title:
            return m_title.group(1).strip().strip('"')
        m_start = re.match(r"^@startuml\s+(.+)$", s, re.IGNORECASE)
        if m_start:
            cand = m_start.group(1).strip().strip('"')
            if cand and not cand.startswith("-"):
                return cand
    return None


def render_plantuml_to_png(puml_code: str, output_png_path: Path) -> bool:
    """Compiles PlantUML sequence/architecture text into a PNG using the bundled plantuml.jar."""
    java_exe, _ = get_best_java()
    if not java_exe:
        return False

    jar_path = get_puml2visio_asset_path(PLANTUML_JAR_NAME)
    if not jar_path.exists():
        return False

    code = puml_code.strip()
    if not code.lower().startswith("@start"):
        code = f"@startuml\n{code}\n@enduml"

    with tempfile.TemporaryDirectory() as td:
        puml_file = Path(td) / "diagram.puml"
        puml_file.write_text(code, encoding="utf-8")

        cmd = [java_exe, "-jar", str(jar_path), "-tpng", str(puml_file)]
        kwargs = {"creationflags": 0x08000000} if os.name == "nt" else {}
        try:
            res = subprocess.run(cmd, capture_output=True, text=True, timeout=45, **kwargs)
            png_files = sorted(Path(td).glob("*.png"))
            gen_png = Path(td) / "diagram.png"
            if not gen_png.exists() and png_files:
                gen_png = png_files[0]
            if gen_png.exists() and gen_png.stat().st_size > 0:
                output_png_path.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(gen_png, output_png_path)
                return True
            else:
                logger.debug(f"PlantUML did not produce image (exit {res.returncode}): {res.stderr}")
        except Exception as e:
            logger.debug(f"PlantUML compilation error: {e}")
            return False
    return False


def resolve_image_path(
    src: str,
    md_path: Path,
    companion_media_dir: Optional[Path] = None,
    auto_detect_media: bool = True,
) -> Optional[Path]:
    """Resolves an image reference from Markdown into a local file path."""
    if not src:
        return None

    raw = src.strip()
    if raw.lower().startswith("file:///"):
        raw = raw[8:] if len(raw) > 9 and raw[9] == ":" else raw[7:]

    clean_src = urllib.parse.unquote(raw.split("?")[0].split("#")[0].strip())
    if not clean_src:
        return None

    p = Path(clean_src)

    # 1. Absolute path check
    try:
        if p.is_Absolute() if hasattr(p, "is_Absolute") else p.is_absolute():
            if p.is_file():
                return p
    except Exception:
        pass

    # 2. Relative to markdown file parent
    try:
        cand = (md_path.parent / clean_src).resolve()
        if cand.is_file():
            return cand
    except Exception:
        pass

    # 3. Explicit companion media directory
    if companion_media_dir and companion_media_dir.is_dir():
        cand = (companion_media_dir / p.name).resolve()
        if cand.is_file():
            return cand

    # 4. Auto-detected companion media directory ({stem}_media/) and standard subdirectories
    if auto_detect_media:
        cand = (md_path.parent / f"{md_path.stem}_media" / p.name).resolve()
        if cand.is_file():
            return cand

        for sub in ("media", "images", "assets", "figures"):
            cand = (md_path.parent / sub / p.name).resolve()
            if cand.is_file():
                return cand

    return None


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


class MarkdownToWordThread(QThread):
    """
    Background worker for compiling Markdown to Word with embedded media and diagrams.
    Compatible with the QueueManager task registry.
    """

    ui_log_msg = pyqtSignal(str, int)
    finished_path = pyqtSignal(str)
    conversion_failed = pyqtSignal(str)
    progress = pyqtSignal(int, str)

    def __init__(
        self,
        doc_source: str,
        output_dir: Optional[Union[str, Path]] = None,
        companion_media_dir: Optional[Union[str, Path]] = None,
        auto_detect_media: bool = True,
        render_diagrams: bool = True,
        apply_3gpp_styles: bool = True,
        auto_open: bool = False,
    ):
        super().__init__()
        self.doc_source = str(doc_source)
        self.output_dir = Path(output_dir).resolve() if output_dir else None
        self.companion_media_dir = Path(companion_media_dir).resolve() if companion_media_dir else None
        self.auto_detect_media = auto_detect_media
        self.render_diagrams = render_diagrams
        self.apply_3gpp_styles = apply_3gpp_styles
        self.auto_open = auto_open
        self._downloaded_temp: Optional[Path] = None

    def _log(self, message: str, level: int = logging.INFO) -> None:
        logger.log(level, message)
        self.ui_log_msg.emit(message, level)

    def _resolve_source_path(self, input_str: str) -> Path:
        clean_input = input_str.strip()
        if not clean_input:
            raise ValueError("Markdown document path is empty.")

        if clean_input.startswith("http://") or clean_input.startswith("https://"):
            self._log("⏳ Downloading Markdown document from URL...", logging.INFO)
            session = NetworkSession.get_instance()
            r = session.get(clean_input, allow_redirects=True, timeout=30)
            r.raise_for_status()

            fd, tmp_str = tempfile.mkstemp(prefix="puml2word_", suffix=".md")
            os.close(fd)
            tmp_path = Path(tmp_str)
            tmp_path.write_bytes(r.content)
            self._downloaded_temp = tmp_path
            return tmp_path

        p = Path(clean_input).resolve()
        if not p.is_file():
            raise FileNotFoundError(f"Markdown file does not exist: {p}")
        return p

    def run(self) -> None:
        self._downloaded_temp = None
        try:
            source_path = self._resolve_source_path(self.doc_source)

            if self.output_dir:
                self.output_dir.mkdir(parents=True, exist_ok=True)
                target_docx = self.output_dir / f"{source_path.stem}.docx"
            else:
                target_docx = source_path.with_suffix(".docx")

            res = convert_markdown_to_docx(
                md_path=source_path,
                output_docx_path=target_docx,
                companion_media_dir=self.companion_media_dir,
                auto_detect_media=self.auto_detect_media,
                render_diagrams=self.render_diagrams,
                apply_3gpp_styles=self.apply_3gpp_styles,
                log_callback=lambda msg: self._log(msg, logging.INFO),
            )

            out_path = res["output_path"]
            clean_out = strip_extended_path(out_path)

            if self.auto_open:
                try:
                    open_document(clean_out)
                    self._log(f"👁️ Auto-opened '{Path(clean_out).name}' in Microsoft Word.", logging.INFO)
                except Exception as e:
                    self._log(f"⚠️ Could not auto-open Word: {e}", logging.WARNING)

            self.finished_path.emit(clean_out)

        except Exception as e:
            err_msg = str(e)
            self._log(f"❌ Markdown to Word conversion failed: {err_msg}", logging.ERROR)
            self.conversion_failed.emit(err_msg)

        finally:
            if self._downloaded_temp and self._downloaded_temp.exists():
                try:
                    self._downloaded_temp.unlink(missing_ok=True)
                except Exception:
                    pass
