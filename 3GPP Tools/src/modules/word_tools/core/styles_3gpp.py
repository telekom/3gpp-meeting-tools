# --- File: src/modules/word_tools/core/styles_3gpp.py ---
"""
3GPP Word Styles Registry and OpenXML Packaging Helper.

This module provides programmatic access to all official 3GPP Word styles
extracted from the authoritative 3GPP Change Request template (Template_3GPP_CR.docx).
It enables strict enforcement of 3GPP styles in both formal CRs and programmatically
constructed pseudo-CRs (pCRs), preventing delegate contributions from being rejected
by Rapporteurs or MCC automated validation tools.
"""

import io
import logging
import os
import re
import shutil
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Optional, Set, Tuple

from lxml import etree
import requests

from core.utils.paths import to_extended_path

logger = logging.getLogger(__name__)

OFFICIAL_CR_TEMPLATE_URL = (
    "https://www.3gpp.org/ftp/3guInternal/3GPP_ultimate_templates/Template_3GPP_CR.docx"
)
OFFICIAL_PCR_TEMPLATE_ZIP_URL = (
    "https://www.3gpp.org/ftp/Information/All_Templates/Pseudo-CR-Form-v0-1.zip"
)
OFFICIAL_3GPP_DOTX_URL = (
    "https://www.3gpp.org/ftp/Information/All_Templates/3gpp_70.dotx"
)


@dataclass(frozen=True)
class StyleInfo:
    style_id: str
    name: str
    style_type: str
    description: str


class Style3GPP:
    """Canonical 3GPP style identifiers (w:styleId values)."""
    # Headings
    HEADING_1 = "Heading1"
    HEADING_2 = "Heading2"
    HEADING_3 = "Heading3"
    HEADING_4 = "Heading4"
    HEADING_5 = "Heading5"
    HEADING_6 = "Heading6"
    HEADING_7 = "Heading7"
    HEADING_8 = "Heading8"
    HEADING_9 = "Heading9"

    # Body and Text
    NORMAL = "Normal"
    FP = "FP"
    NW = "NW"
    EW = "EW"
    LD = "LD"

    # 3GPP Bullets (Standard)
    B1 = "B1"
    B2 = "B2"
    B3 = "B3"
    B4 = "B4"
    B5 = "B5"

    # Numbered and Bullet Lists (Word compatibility)
    LIST = "List"
    LIST_2 = "List2"
    LIST_3 = "List3"
    LIST_4 = "List4"
    LIST_5 = "List5"
    LIST_BULLET = "ListBullet"
    LIST_BULLET_2 = "ListBullet2"
    LIST_BULLET_3 = "ListBullet3"
    LIST_BULLET_4 = "ListBullet4"
    LIST_BULLET_5 = "ListBullet5"
    LIST_NUMBER = "ListNumber"
    LIST_NUMBER_2 = "ListNumber2"

    # Notes, Examples, and Notices
    NOTE = "NO"
    EXAMPLE = "EX"
    EDITORS_NOTE = "EditorsNote"

    # Tables and Figures
    TABLE_HEADER = "TH"
    TABLE_TITLE = "TT"
    TABLE_FOOTNOTE = "TF"
    TABLE_ALIGN_LEFT = "TAL"
    TABLE_ALIGN_CENTER = "TAC"
    TABLE_ALIGN_RIGHT = "TAR"
    TABLE_ALIGN_NUM = "TAN"
    TABLE_ALIGN_HEADER = "TAH"
    FIGURE = "NF"
    EQUATION = "EQ"
    PROGRAM_LISTING = "PL"

    # 3GPP Contribution & CR Specific
    TDOC_HEADER = "tdoc-header"
    CR_COVER_PAGE = "CRCoverPage"
    CR_SEPARATOR = "CRSeparator"
    CR_SEPARATOR_CHAR = "CRSeparatorChar"

    # Annex Headings
    ANNEX_TITLE = "ZA"
    ANNEX_H1 = "ZB"
    ANNEX_H2 = "ZD"
    ANNEX_H3 = "ZG"
    ANNEX_H4 = "ZH"
    ANNEX_H5 = "ZT"
    ANNEX_H6 = "ZTD"
    ANNEX_H7 = "ZU"
    ANNEX_H8 = "ZV"


# Programmatic registry of all detected official 3GPP styles
STYLES_3GPP_REGISTRY: Dict[str, StyleInfo] = {
    # Core Headings
    Style3GPP.HEADING_1: StyleInfo("Heading1", "heading 1", "paragraph", "Top-level clause heading (e.g. 1. Discussion, 2. Proposal)"),
    Style3GPP.HEADING_2: StyleInfo("Heading2", "heading 2", "paragraph", "Second-level clause heading (e.g. 5.1, 6.x)"),
    Style3GPP.HEADING_3: StyleInfo("Heading3", "heading 3", "paragraph", "Third-level clause heading (e.g. 5.1.1, 6.X.0)"),
    Style3GPP.HEADING_4: StyleInfo("Heading4", "heading 4", "paragraph", "Fourth-level clause heading (e.g. 5.1.1.1)"),
    Style3GPP.HEADING_5: StyleInfo("Heading5", "heading 5", "paragraph", "Fifth-level clause heading"),
    Style3GPP.HEADING_6: StyleInfo("Heading6", "heading 6", "paragraph", "Sixth-level clause heading"),
    Style3GPP.HEADING_7: StyleInfo("Heading7", "heading 7", "paragraph", "Seventh-level clause heading"),
    Style3GPP.HEADING_8: StyleInfo("Heading8", "heading 8", "paragraph", "Eighth-level clause heading"),
    Style3GPP.HEADING_9: StyleInfo("Heading9", "heading 9", "paragraph", "Ninth-level clause heading"),

    # Body text
    Style3GPP.NORMAL: StyleInfo("Normal", "Normal", "paragraph", "Default body text in 3GPP specifications and contributions"),

    # Bullets (B1-B5 are mandatory for 3GPP specification text)
    Style3GPP.B1: StyleInfo("B1", "B1", "paragraph", "3GPP 1st level bullet list"),
    Style3GPP.B2: StyleInfo("B2", "B2", "paragraph", "3GPP 2nd level bullet list"),
    Style3GPP.B3: StyleInfo("B3", "B3", "paragraph", "3GPP 3rd level bullet list"),
    Style3GPP.B4: StyleInfo("B4", "B4", "paragraph", "3GPP 4th level bullet list"),
    Style3GPP.B5: StyleInfo("B5", "B5", "paragraph", "3GPP 5th level bullet list"),

    # Notes & Editor's notes
    Style3GPP.NOTE: StyleInfo("NO", "NO", "paragraph", "3GPP Note style (e.g. NOTE 1: ...)"),
    Style3GPP.EXAMPLE: StyleInfo("EX", "EX", "paragraph", "3GPP Example style"),
    Style3GPP.EDITORS_NOTE: StyleInfo("EditorsNote", "Editor's Note", "paragraph", "3GPP Editor's Note style for study items and draft specs"),

    # Tables & Figures
    Style3GPP.TABLE_HEADER: StyleInfo("TH", "TH", "paragraph", "Table Header text"),
    Style3GPP.TABLE_TITLE: StyleInfo("TT", "TT", "paragraph", "Table Title caption"),
    Style3GPP.TABLE_FOOTNOTE: StyleInfo("TF", "TF", "paragraph", "Table Footnote"),
    Style3GPP.TABLE_ALIGN_LEFT: StyleInfo("TAL", "TAL", "paragraph", "Table cell text aligned left"),
    Style3GPP.TABLE_ALIGN_CENTER: StyleInfo("TAC", "TAC", "paragraph", "Table cell text aligned center"),
    Style3GPP.TABLE_ALIGN_RIGHT: StyleInfo("TAR", "TAR", "paragraph", "Table cell text aligned right"),
    Style3GPP.TABLE_ALIGN_NUM: StyleInfo("TAN", "TAN", "paragraph", "Table cell numbers aligned right"),
    Style3GPP.TABLE_ALIGN_HEADER: StyleInfo("TAH", "TAH", "paragraph", "Table Header aligned center"),
    Style3GPP.FIGURE: StyleInfo("NF", "NF", "paragraph", "Figure caption / non-formula figure title"),
    Style3GPP.EQUATION: StyleInfo("EQ", "EQ", "paragraph", "Equation and mathematical expressions"),
    Style3GPP.PROGRAM_LISTING: StyleInfo("PL", "PL", "paragraph", "Program and protocol message code listing"),

    # Document Header & Change Separators
    Style3GPP.TDOC_HEADER: StyleInfo("tdoc-header", "tdoc-header", "paragraph", "Top-level meeting and TDoc identifier line"),
    Style3GPP.CR_COVER_PAGE: StyleInfo("CRCoverPage", "CR Cover Page", "paragraph", "CR Cover page title and metadata tables"),
    Style3GPP.CR_SEPARATOR: StyleInfo("CRSeparator", "CR_Separator", "paragraph", "3GPP Change separator (* * * First Change * * * *)"),
    Style3GPP.CR_SEPARATOR_CHAR: StyleInfo("CRSeparatorChar", "CR_Separator Char", "character", "Character run style for change separator"),

    # Annex Headings
    Style3GPP.ANNEX_TITLE: StyleInfo("ZA", "ZA", "paragraph", "Annex Title (informative / normative)"),
    Style3GPP.ANNEX_H1: StyleInfo("ZB", "ZB", "paragraph", "Annex Heading 1"),
    Style3GPP.ANNEX_H2: StyleInfo("ZD", "ZD", "paragraph", "Annex Heading 2"),
    Style3GPP.ANNEX_H3: StyleInfo("ZG", "ZG", "paragraph", "Annex Heading 3"),
    Style3GPP.ANNEX_H4: StyleInfo("ZH", "ZH", "paragraph", "Annex Heading 4"),
    Style3GPP.ANNEX_H5: StyleInfo("ZT", "ZT", "paragraph", "Annex Heading 5"),
    Style3GPP.ANNEX_H6: StyleInfo("ZTD", "ZTD", "paragraph", "Annex Heading 6"),
    Style3GPP.ANNEX_H7: StyleInfo("ZU", "ZU", "paragraph", "Annex Heading 7"),
    Style3GPP.ANNEX_H8: StyleInfo("ZV", "ZV", "paragraph", "Annex Heading 8"),

    # Lists, Numbers, and Bullets
    "List": StyleInfo("List", "List", "paragraph", "Standard list paragraph"),
    "List2": StyleInfo("List2", "List 2", "paragraph", "Level 2 list paragraph"),
    "List3": StyleInfo("List3", "List 3", "paragraph", "Level 3 list paragraph"),
    "List4": StyleInfo("List4", "List 4", "paragraph", "Level 4 list paragraph"),
    "List5": StyleInfo("List5", "List 5", "paragraph", "Level 5 list paragraph"),
    "ListBullet": StyleInfo("ListBullet", "List Bullet", "paragraph", "Word bullet list"),
    "ListBullet2": StyleInfo("ListBullet2", "List Bullet 2", "paragraph", "Word bullet list level 2"),
    "ListBullet3": StyleInfo("ListBullet3", "List Bullet 3", "paragraph", "Word bullet list level 3"),
    "ListBullet4": StyleInfo("ListBullet4", "List Bullet 4", "paragraph", "Word bullet list level 4"),
    "ListBullet5": StyleInfo("ListBullet5", "List Bullet 5", "paragraph", "Word bullet list level 5"),
    "ListNumber": StyleInfo("ListNumber", "List Number", "paragraph", "Numbered list level 1"),
    "ListNumber2": StyleInfo("ListNumber2", "List Number 2", "paragraph", "Numbered list level 2"),

    # Table of Contents & Index
    "TOC1": StyleInfo("TOC1", "toc 1", "paragraph", "Table of contents level 1"),
    "TOC2": StyleInfo("TOC2", "toc 2", "paragraph", "Table of contents level 2"),
    "TOC3": StyleInfo("TOC3", "toc 3", "paragraph", "Table of contents level 3"),
    "TOC4": StyleInfo("TOC4", "toc 4", "paragraph", "Table of contents level 4"),
    "TOC5": StyleInfo("TOC5", "toc 5", "paragraph", "Table of contents level 5"),
    "TOC6": StyleInfo("TOC6", "toc 6", "paragraph", "Table of contents level 6"),
    "TOC7": StyleInfo("TOC7", "toc 7", "paragraph", "Table of contents level 7"),
    "TOC8": StyleInfo("TOC8", "toc 8", "paragraph", "Table of contents level 8"),
    "TOC9": StyleInfo("TOC9", "toc 9", "paragraph", "Table of contents level 9"),
    "Index1": StyleInfo("Index1", "index 1", "paragraph", "Index entry level 1"),
    "Index2": StyleInfo("Index2", "index 2", "paragraph", "Index entry level 2"),

    # Special Formatting
    "FP": StyleInfo("FP", "FP", "paragraph", "Full page paragraph style"),
    "LD": StyleInfo("LD", "LD", "paragraph", "List definition paragraph"),
    "NW": StyleInfo("NW", "NW", "paragraph", "No wrap text style"),
    "EW": StyleInfo("EW", "EW", "paragraph", "Equation wrap style"),
    "H6": StyleInfo("H6", "H6", "paragraph", "Special heading 6 style"),
    "Header": StyleInfo("Header", "header", "paragraph", "Page header text"),
    "Footer": StyleInfo("Footer", "footer", "paragraph", "Page footer text"),
    "FootnoteText": StyleInfo("FootnoteText", "footnote text", "paragraph", "Footnote body text"),
    "FootnoteReference": StyleInfo("FootnoteReference", "footnote reference", "character", "Footnote reference number"),
    "DefaultParagraphFont": StyleInfo("DefaultParagraphFont", "Default Paragraph Font", "character", "Default font formatting"),
    "TableNormal": StyleInfo("TableNormal", "Normal Table", "table", "Standard 3GPP table grid"),
    "NoList": StyleInfo("NoList", "No List", "numbering", "List override clear"),
    "Hyperlink": StyleInfo("Hyperlink", "Hyperlink", "character", "Web/cross-reference hyperlink"),
    "FollowedHyperlink": StyleInfo("FollowedHyperlink", "FollowedHyperlink", "character", "Visited hyperlink"),
    "ZGSM": StyleInfo("ZGSM", "ZGSM", "character", "3GPP GSM character style"),
    "BalloonText": StyleInfo("BalloonText", "Balloon Text", "paragraph", "Comment balloon text"),
    "CommentReference": StyleInfo("CommentReference", "annotation reference", "character", "Comment mark"),
    "CommentText": StyleInfo("CommentText", "annotation text", "paragraph", "Comment content"),
    "CommentSubject": StyleInfo("CommentSubject", "annotation subject", "paragraph", "Comment subject heading"),
    "DocumentMap": StyleInfo("DocumentMap", "Document Map", "paragraph", "Word document outline map"),
}

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

# Canonical normalized lower-case set of all official 3GPP styles (IDs and names)
ALL_3GPP_STYLE_NAMES_LOWER: Set[str] = {
    # Headings
    "heading 1", "heading 2", "heading 3", "heading 4", "heading 5",
    "heading 6", "heading 7", "heading 8", "heading 9",
    "heading1", "heading2", "heading3", "heading4", "heading5",
    "heading6", "heading7", "heading8", "heading9",
    # Body
    "normal", "fp", "nw", "ew", "ld",
    # Bullets & Lists
    "b1", "b2", "b3", "b4", "b5",
    "list", "list 2", "list 3", "list 4", "list 5",
    "list2", "list3", "list4", "list5",
    "list bullet", "list bullet 2", "list bullet 3", "list bullet 4", "list bullet 5",
    "listbullet", "listbullet2", "listbullet3", "listbullet4", "listbullet5",
    "list number", "list number 2", "listnumber", "listnumber2",
    # Notes, Examples, Warnings
    "no", "ex", "editor's note", "editorsnote",
    # Tables and Figures
    "th", "tt", "tf", "tal", "tac", "tar", "tan", "tah",
    "nf", "eq", "pl",
    # CR & Contribution Specific
    "cr_separator", "crseparator", "cr_separator char", "crseparatorchar",
    "cr cover page", "crcoverpage", "tdoc-header",
    # Annex styles
    "za", "zb", "zd", "zg", "zgsm", "zh", "zt", "ztd", "zu", "zv",
    # Character companion styles
    "th char", "thchar", "b1 char", "b1char", "no zchn", "nozchn", "no char", "nochar",
    "editor's note char", "editorsnotechar", "ex char", "exchar",
    "tal char", "talchar", "tac char", "tacchar", "tah char", "tahchar",
    "tf char", "tfchar", "tt char", "ttchar",
}

# Essential 3GPP character companion styles
LINKED_CHAR_STYLES = [
    # (styleId, name, linked_para_id, font_name, is_bold, is_italic)
    ("THChar", "TH Char", "TH", "Arial", True, False),
    ("B1Char", "B1 Char", "B1", "Times New Roman", False, False),
    ("NOZchn", "NO Zchn", "NO", "Times New Roman", False, False),
    ("EditorsNoteChar", "Editor's Note Char", "EditorsNote", "Times New Roman", False, True),
    ("EXChar", "EX Char", "EX", "Times New Roman", False, False),
    ("TALChar", "TAL Char", "TAL", "Times New Roman", False, False),
    ("TACChar", "TAC Char", "TAC", "Times New Roman", False, False),
    ("TAHChar", "TAH Char", "TAH", "Arial", True, False),
    ("TFChar", "TF Char", "TF", "Times New Roman", False, False),
    ("CRSeparatorChar", "CR_Separator Char", "CRSeparator", "Arial", True, False),
]

TABLE_GRID_XML = """
<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:type="table" w:styleId="TableGrid">
  <w:name w:val="Table Grid"/>
  <w:basedOn w:val="TableNormal"/>
  <w:uiPriority w:val="39"/>
  <w:pPr>
    <w:spacing w:after="0" w:line="240" w:lineRule="auto"/>
  </w:pPr>
  <w:tblPr>
    <w:tblBorders>
      <w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>
      <w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>
      <w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>
      <w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>
      <w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>
      <w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>
    </w:tblBorders>
  </w:tblPr>
</w:style>
"""


def promote_3gpp_styles_in_xml(xml_bytes: bytes) -> bytes:
    """
    Parses word/styles.xml, ensures that all official 3GPP styles have <w:qFormat/>
    so they are displayed prominently in the Microsoft Word QuickStyle ribbon gallery,
    removes semiHidden/unhideWhenUsed flags, sets uiPriority='1' (top of Styles pane),
    guarantees that TableGrid table style is defined with standard borders,
    and injects missing authentic 3GPP linked character styles.
    """
    ns = {"w": W_NS}
    root = etree.fromstring(xml_bytes)

    # 1. Update existing styles
    existing_style_ids = set()
    existing_names = set()
    has_table_grid = False

    for style in root.findall(".//w:style", ns):
        s_id = style.get(f"{{{W_NS}}}styleId", "")
        if s_id:
            existing_style_ids.add(s_id.lower())
            if s_id.lower() == "tablegrid":
                has_table_grid = True

        name_el = style.find("w:name", ns)
        name_val = name_el.get(f"{{{W_NS}}}val", "").strip().lower() if name_el is not None else ""
        if name_val:
            existing_names.add(name_val)
            if name_val == "table grid":
                has_table_grid = True

        if name_val in ALL_3GPP_STYLE_NAMES_LOWER or s_id.lower() in ALL_3GPP_STYLE_NAMES_LOWER:
            # Ensure qFormat
            if style.find("w:qFormat", ns) is None:
                etree.SubElement(style, f"{{{W_NS}}}qFormat")

            # Remove semiHidden
            sh = style.find("w:semiHidden", ns)
            if sh is not None:
                style.remove(sh)

            # Remove unhideWhenUsed
            uh = style.find("w:unhideWhenUsed", ns)
            if uh is not None:
                style.remove(uh)

            # Set uiPriority to 1 so 3GPP styles appear at the top of the pane
            pri = style.find("w:uiPriority", ns)
            if pri is None:
                pri = etree.SubElement(style, f"{{{W_NS}}}uiPriority")
            pri.set(f"{{{W_NS}}}val", "1")

    # 2. Ensure TableGrid table style exists so table borders and gridlines render properly
    if not has_table_grid:
        tg_elem = etree.fromstring(TABLE_GRID_XML.strip().encode("utf-8"))
        root.append(tg_elem)
        existing_style_ids.add("tablegrid")
        existing_names.add("table grid")

    # 3. Inject missing authentic 3GPP linked character styles
    for s_id, s_name, linked_para_id, font_name, is_bold, is_italic in LINKED_CHAR_STYLES:
        if s_id.lower() not in existing_style_ids and s_name.lower() not in existing_names:
            char_style = etree.SubElement(
                root,
                f"{{{W_NS}}}style",
                {
                    f"{{{W_NS}}}type": "character",
                    f"{{{W_NS}}}customStyle": "1",
                    f"{{{W_NS}}}styleId": s_id,
                },
            )
            etree.SubElement(char_style, f"{{{W_NS}}}name", {f"{{{W_NS}}}val": s_name})
            etree.SubElement(char_style, f"{{{W_NS}}}link", {f"{{{W_NS}}}val": linked_para_id})
            etree.SubElement(char_style, f"{{{W_NS}}}uiPriority", {f"{{{W_NS}}}val": "1"})
            etree.SubElement(char_style, f"{{{W_NS}}}qFormat")
            rPr = etree.SubElement(char_style, f"{{{W_NS}}}rPr")
            etree.SubElement(rPr, f"{{{W_NS}}}rFonts", {f"{{{W_NS}}}ascii": font_name, f"{{{W_NS}}}hAnsi": font_name})
            if is_bold:
                etree.SubElement(rPr, f"{{{W_NS}}}b")
            if is_italic:
                etree.SubElement(rPr, f"{{{W_NS}}}i")
            etree.SubElement(rPr, f"{{{W_NS}}}lang", {f"{{{W_NS}}}val": "en-GB", f"{{{W_NS}}}eastAsia": "en-US"})

        # Also ensure the corresponding paragraph style has <w:link w:val="..."/>
        para_styles = root.xpath(f".//w:style[@w:styleId='{linked_para_id}']", namespaces=ns)
        if para_styles:
            para_st = para_styles[0]
            if para_st.find("w:link", ns) is None:
                etree.SubElement(para_st, f"{{{W_NS}}}link", {f"{{{W_NS}}}val": s_id})

    return etree.tostring(root, xml_declaration=True, encoding="utf-8", standalone=True)


def merge_and_promote_3gpp_styles(target_styles_bytes: bytes, template_styles_bytes: bytes) -> bytes:
    """
    Non-destructively merges authentic 3GPP styles into a document's existing styles.xml.
    Preserves all existing document styles (including table styles like TableGrid,
    custom table borders, formatting, and fonts) while ensuring that:
    1. TableGrid table style exists with proper 0.5pt single borders.
    2. Any missing official 3GPP styles from the template are injected.
    3. All 3GPP styles are promoted with <w:qFormat/> and uiPriority="1".
    4. Essential linked character companion styles are injected.
    """
    import copy

    ns = {"w": W_NS}
    target_root = etree.fromstring(target_styles_bytes)

    # 1. Inspect existing styles in target
    existing_style_ids = set()
    existing_names = set()
    has_table_grid = False

    for style in target_root.findall(".//w:style", ns):
        s_id = style.get(f"{{{W_NS}}}styleId", "")
        if s_id:
            existing_style_ids.add(s_id.lower())
            if s_id.lower() == "tablegrid":
                has_table_grid = True

        name_el = style.find("w:name", ns)
        name_val = name_el.get(f"{{{W_NS}}}val", "").strip().lower() if name_el is not None else ""
        if name_val:
            existing_names.add(name_val)
            if name_val == "table grid":
                has_table_grid = True

    # 2. Ensure TableGrid style is present so table borders/gridlines are preserved
    if not has_table_grid:
        tg_elem = etree.fromstring(TABLE_GRID_XML.strip().encode("utf-8"))
        target_root.append(tg_elem)
        existing_style_ids.add("tablegrid")
        existing_names.add("table grid")

    # 3. Additively merge missing 3GPP styles from template without overwriting existing styles
    if template_styles_bytes:
        try:
            template_root = etree.fromstring(template_styles_bytes)
            for style in template_root.findall(".//w:style", ns):
                s_id = style.get(f"{{{W_NS}}}styleId", "")
                name_el = style.find("w:name", ns)
                name_val = name_el.get(f"{{{W_NS}}}val", "").strip().lower() if name_el is not None else ""

                is_3gpp = (name_val in ALL_3GPP_STYLE_NAMES_LOWER) or (s_id.lower() in ALL_3GPP_STYLE_NAMES_LOWER)
                if is_3gpp:
                    if s_id.lower() not in existing_style_ids and name_val not in existing_names:
                        target_root.append(copy.deepcopy(style))
                        if s_id:
                            existing_style_ids.add(s_id.lower())
                        if name_val:
                            existing_names.add(name_val)
        except Exception as e:
            logger.debug("Failed merging 3GPP styles from template: %s", e)

    # 4. Promote all 3GPP styles in target
    for style in target_root.findall(".//w:style", ns):
        s_id = style.get(f"{{{W_NS}}}styleId", "")
        name_el = style.find("w:name", ns)
        name_val = name_el.get(f"{{{W_NS}}}val", "").strip().lower() if name_el is not None else ""

        if name_val in ALL_3GPP_STYLE_NAMES_LOWER or s_id.lower() in ALL_3GPP_STYLE_NAMES_LOWER:
            if style.find("w:qFormat", ns) is None:
                etree.SubElement(style, f"{{{W_NS}}}qFormat")
            sh = style.find("w:semiHidden", ns)
            if sh is not None:
                style.remove(sh)
            uh = style.find("w:unhideWhenUsed", ns)
            if uh is not None:
                style.remove(uh)
            pri = style.find("w:uiPriority", ns)
            if pri is None:
                pri = etree.SubElement(style, f"{{{W_NS}}}uiPriority")
            pri.set(f"{{{W_NS}}}val", "1")

    # 5. Inject missing linked character companion styles
    for s_id, s_name, linked_para_id, font_name, is_bold, is_italic in LINKED_CHAR_STYLES:
        if s_id.lower() not in existing_style_ids and s_name.lower() not in existing_names:
            char_style = etree.SubElement(
                target_root,
                f"{{{W_NS}}}style",
                {
                    f"{{{W_NS}}}type": "character",
                    f"{{{W_NS}}}customStyle": "1",
                    f"{{{W_NS}}}styleId": s_id,
                },
            )
            etree.SubElement(char_style, f"{{{W_NS}}}name", {f"{{{W_NS}}}val": s_name})
            etree.SubElement(char_style, f"{{{W_NS}}}link", {f"{{{W_NS}}}val": linked_para_id})
            etree.SubElement(char_style, f"{{{W_NS}}}uiPriority", {f"{{{W_NS}}}val": "1"})
            etree.SubElement(char_style, f"{{{W_NS}}}qFormat")
            rPr = etree.SubElement(char_style, f"{{{W_NS}}}rPr")
            etree.SubElement(rPr, f"{{{W_NS}}}rFonts", {f"{{{W_NS}}}ascii": font_name, f"{{{W_NS}}}hAnsi": font_name})
            if is_bold:
                etree.SubElement(rPr, f"{{{W_NS}}}b")
            if is_italic:
                etree.SubElement(rPr, f"{{{W_NS}}}i")
            etree.SubElement(rPr, f"{{{W_NS}}}lang", {f"{{{W_NS}}}val": "en-GB", f"{{{W_NS}}}eastAsia": "en-US"})

        para_styles = target_root.xpath(f".//w:style[@w:styleId='{linked_para_id}']", namespaces=ns)
        if para_styles:
            para_st = para_styles[0]
            if para_st.find("w:link", ns) is None:
                etree.SubElement(para_st, f"{{{W_NS}}}link", {f"{{{W_NS}}}val": s_id})

    return etree.tostring(target_root, xml_declaration=True, encoding="utf-8", standalone=True)


def promote_3gpp_styles_in_word(doc) -> int:
    """
    Ensures that Microsoft Word's live document model marks all 3GPP styles as
    QuickStyles (visible on the Home tab ribbon gallery) and visible in the Styles pane.
    Also ensures doc.UpdateStylesOnOpen is disabled so Word does not wipe them out on open.
    """
    count = 0
    try:
        try:
            doc.UpdateStylesOnOpen = False
        except Exception:
            pass

        for s in doc.Styles:
            try:
                name_clean = s.NameLocal.strip().lower()
                if name_clean in ALL_3GPP_STYLE_NAMES_LOWER:
                    s.QuickStyle = True
                    s.Visibility = False  # Word COM: Visibility=False means NOT hidden in Styles task pane
                    s.Priority = 1
                    s.UnhideWhenUsed = False
                    s.Hidden = False
                    count += 1
            except Exception:
                pass
    except Exception as exc:
        logger.debug("Style promotion loop in Word COM encountered: %s", exc)
    return count


def get_template_cr_path() -> Path:
    """
    Returns the absolute path to the cached Template_3GPP_CR.docx.
    If the template is missing locally, attempts an automatic download from 3GPP.
    """
    assets_dir = Path(__file__).resolve().parent.parent.parent.parent / "assets"
    assets_dir.mkdir(parents=True, exist_ok=True)
    tpl_path = assets_dir / "Template_3GPP_CR.docx"

    if tpl_path.exists() and tpl_path.stat().st_size > 5000:
        return tpl_path

    # Fallback auto-download: try corporate-aware NetworkSession first, then direct requests
    try:
        logger.info("Template_3GPP_CR.docx not found in assets; downloading from 3GPP via NetworkSession...")
        from core.network.session import NetworkSession
        NetworkSession.download_file(
            url=OFFICIAL_CR_TEMPLATE_URL,
            dest_path=tpl_path,
            timeout=30,
            atomic=True,
        )
        if tpl_path.exists() and tpl_path.stat().st_size > 5000:
            logger.info("Successfully downloaded Template_3GPP_CR.docx (%d bytes).", tpl_path.stat().st_size)
            return tpl_path
    except Exception as exc:
        logger.warning("NetworkSession download failed for Template_3GPP_CR.docx: %s. Trying direct requests...", exc)

    try:
        headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
        resp = requests.get(OFFICIAL_CR_TEMPLATE_URL, headers=headers, timeout=20)
        if resp.status_code == 200 and len(resp.content) > 5000:
            with open(to_extended_path(tpl_path), "wb") as f:
                f.write(resp.content)
            logger.info("Successfully downloaded Template_3GPP_CR.docx via requests (%d bytes).", len(resp.content))
            return tpl_path
    except Exception as exc:
        logger.warning("Could not auto-download Template_3GPP_CR.docx: %s", exc)

    if tpl_path.exists():
        return tpl_path
    raise FileNotFoundError(f"3GPP CR Template not found at {tpl_path} and download failed.")


def get_pseudo_cr_template_path() -> Path:
    """
    Returns the absolute path to the cached official Pseudo-CR-Form-v0-1.docx.
    If missing, downloads and extracts the official 3GPP template from 3GPP FTP.
    """
    assets_dir = Path(__file__).resolve().parent.parent.parent.parent / "assets"
    assets_dir.mkdir(parents=True, exist_ok=True)
    pcr_path = assets_dir / "Pseudo-CR-Form-v0-1.docx"

    if pcr_path.exists() and pcr_path.stat().st_size > 5000:
        return pcr_path

    zip_path = assets_dir / "Pseudo-CR-Form-v0-1.zip"
    try:
        logger.info("Pseudo-CR-Form-v0-1.docx not found; downloading from 3GPP...")
        from core.network.session import NetworkSession
        NetworkSession.download_file(
            url=OFFICIAL_PCR_TEMPLATE_ZIP_URL,
            dest_path=zip_path,
            timeout=30,
            atomic=True,
        )
        if zip_path.exists():
            with zipfile.ZipFile(to_extended_path(zip_path), "r") as z:
                z.extractall(assets_dir)
            if pcr_path.exists():
                logger.info("Successfully extracted official Pseudo-CR-Form-v0-1.docx.")
                return pcr_path
    except Exception as exc:
        logger.warning("Could not auto-download official Pseudo-CR template: %s", exc)
    finally:
        if zip_path.exists():
            try:
                zip_path.unlink()
            except OSError:
                pass

    if pcr_path.exists():
        return pcr_path
    return get_template_cr_path()


def get_3gpp_dotx_path() -> Path:
    """
    Returns the absolute path to the cached 3gpp_70.dotx template.
    """
    assets_dir = Path(__file__).resolve().parent.parent.parent.parent / "assets"
    assets_dir.mkdir(parents=True, exist_ok=True)
    dotx_path = assets_dir / "3gpp_70.dotx"

    if dotx_path.exists() and dotx_path.stat().st_size > 5000:
        return dotx_path

    try:
        logger.info("3gpp_70.dotx not found; downloading from 3GPP...")
        from core.network.session import NetworkSession
        NetworkSession.download_file(
            url=OFFICIAL_3GPP_DOTX_URL,
            dest_path=dotx_path,
            timeout=30,
            atomic=True,
        )
        if dotx_path.exists() and dotx_path.stat().st_size > 5000:
            logger.info("Successfully downloaded 3gpp_70.dotx (%d bytes).", dotx_path.stat().st_size)
            return dotx_path
    except Exception as exc:
        logger.warning("Could not auto-download 3gpp_70.dotx: %s", exc)

    return dotx_path


def is_3gpp_dotx_installed() -> bool:
    """Returns True if 3gpp_70.dotx is installed in the user's Word templates folder."""
    appdata = os.environ.get("APPDATA")
    if not appdata:
        return False
    target = Path(appdata) / "Microsoft" / "Templates" / "3gpp_70.dotx"
    return target.exists() and target.stat().st_size > 5000


def get_installed_3gpp_dotx_path() -> Optional[Path]:
    """Returns the Path to 3gpp_70.dotx in the user's Word templates folder if installed."""
    appdata = os.environ.get("APPDATA")
    if not appdata:
        return None
    target = Path(appdata) / "Microsoft" / "Templates" / "3gpp_70.dotx"
    return target if target.exists() and target.stat().st_size > 5000 else None


def install_3gpp_dotx_template(force: bool = False) -> Tuple[bool, str]:
    """
    Installs 3gpp_70.dotx into the user's Word templates folder (%APPDATA%\\Microsoft\\Templates)
    so that Microsoft Word automatically links and attaches the official 3GPP styles and ribbon.
    Returns (success: bool, message: str).
    """
    try:
        appdata = os.environ.get("APPDATA")
        if not appdata:
            return False, "APPDATA environment variable is not defined."

        word_templates = Path(appdata) / "Microsoft" / "Templates"
        word_templates.mkdir(parents=True, exist_ok=True)
        target = word_templates / "3gpp_70.dotx"

        dotx_source = get_3gpp_dotx_path()
        if not dotx_source.exists() or dotx_source.stat().st_size == 0:
            return False, "Could not acquire official 3gpp_70.dotx template from 3GPP."

        if not force and target.exists() and target.stat().st_size == dotx_source.stat().st_size:
            return True, f"3GPP Word template is already installed at {target}."

        shutil.copy2(to_extended_path(dotx_source), to_extended_path(target))
        # Ensure the installed template itself has all authentic 3GPP styles promoted with QuickStyle flags
        try:
            with zipfile.ZipFile(to_extended_path(target), "r") as zin:
                if "word/styles.xml" in zin.namelist():
                    dotx_styles = zin.read("word/styles.xml")
                    dotx_promoted = promote_3gpp_styles_in_xml(dotx_styles)
                    dotx_buf = io.BytesIO()
                    with zipfile.ZipFile(dotx_buf, "w", compression=zipfile.ZIP_DEFLATED) as zout:
                        for item in zin.infolist():
                            if item.filename == "word/styles.xml":
                                zout.writestr(item, dotx_promoted)
                            else:
                                zout.writestr(item, zin.read(item.filename))
                    with open(to_extended_path(target), "wb") as f:
                        f.write(dotx_buf.getvalue())
        except Exception as promote_err:
            logger.debug("Failed promoting styles in installed 3gpp_70.dotx: %s", promote_err)

        logger.info("Installed and promoted 3gpp_70.dotx into %s (%d bytes)", target, target.stat().st_size)
        return True, f"Successfully installed 3GPP Word template to {target}."
    except Exception as exc:
        msg = f"Failed to install 3GPP Word template: {exc}"
        logger.warning(msg)
        return False, msg


def uninstall_3gpp_dotx_template() -> Tuple[bool, str]:
    """
    Removes 3gpp_70.dotx from the user's Word templates folder (%APPDATA%\\Microsoft\\Templates).
    Returns (success: bool, message: str).
    """
    try:
        appdata = os.environ.get("APPDATA")
        if not appdata:
            return False, "APPDATA environment variable is not defined."

        target = Path(appdata) / "Microsoft" / "Templates" / "3gpp_70.dotx"
        if not target.exists():
            return True, "3GPP Word template is not currently installed."

        try:
            target.unlink()
            logger.info("Removed 3gpp_70.dotx from %s", target)
            return True, "Successfully removed 3GPP Word template from Word templates folder."
        except PermissionError:
            return False, "Could not remove 3gpp_70.dotx: file is locked by a running Microsoft Word process."
    except Exception as exc:
        msg = f"Failed to remove 3GPP Word template: {exc}"
        logger.warning(msg)
        return False, msg


def get_3gpp_package_part(part_name: str) -> bytes:
    """
    Extracts a raw package part (e.g. 'word/styles.xml', 'word/theme/theme1.xml')
    directly from the cached official 3GPP template, with full 3GPP QuickStyle promotion.
    """
    tpl_path = get_template_cr_path()
    with zipfile.ZipFile(to_extended_path(tpl_path), "r") as z:
        data = z.read(part_name)
    if part_name == "word/styles.xml":
        return promote_3gpp_styles_in_xml(data)
    return data


def ensure_3gpp_package_styles(docx_path: Path) -> None:
    """
    Non-destructively ensures that the target .docx file contains all official 3GPP styles
    and the standard TableGrid style with proper borders and lines.
    Preserves all existing document styles, tables, themes, and fonts.
    """
    ext_path = to_extended_path(docx_path)
    if not os.path.exists(ext_path):
        raise FileNotFoundError(f"Target document does not exist: {docx_path}")

    template_styles_xml = get_3gpp_package_part("word/styles.xml")
    template_theme_xml = get_3gpp_package_part("word/theme/theme1.xml")
    template_font_table_xml = get_3gpp_package_part("word/fontTable.xml")

    tmp_path = docx_path.with_name(docx_path.name + ".stylepatch.tmp")
    ext_tmp = to_extended_path(tmp_path)

    try:
        with open(ext_path, "rb") as sf:
            source_bytes = sf.read()

        out_buf = io.BytesIO()
        with zipfile.ZipFile(io.BytesIO(source_bytes), "r") as zin, zipfile.ZipFile(
            out_buf, "w", compression=zipfile.ZIP_DEFLATED
        ) as zout:
            namelist = set(zin.namelist())

            # Non-destructive merge: preserve existing document styles, merge 3GPP styles & ensure TableGrid
            if "word/styles.xml" in namelist:
                existing_styles = zin.read("word/styles.xml")
                merged_styles = merge_and_promote_3gpp_styles(existing_styles, template_styles_xml)
            else:
                merged_styles = promote_3gpp_styles_in_xml(template_styles_xml)

            for item in zin.infolist():
                if item.filename == "word/styles.xml":
                    zout.writestr(item, merged_styles)
                elif item.filename == "word/theme/theme1.xml":
                    # Preserve existing theme (colors, fonts, effects)
                    zout.writestr(item, zin.read(item.filename))
                elif item.filename == "word/fontTable.xml":
                    # Preserve existing font table
                    zout.writestr(item, zin.read(item.filename))
                elif item.filename == "word/settings.xml":
                    # Remove <w:linkStyles/> to prevent Microsoft Word from overwriting styles on document open
                    try:
                        settings_text = zin.read(item.filename).decode("utf-8")
                        settings_text = re.sub(r"<(?:\w+:)?linkStyles\b[^>]*/?>", "", settings_text)
                        zout.writestr(item, settings_text.encode("utf-8"))
                    except Exception:
                        zout.writestr(item, zin.read(item.filename))
                else:
                    zout.writestr(item, zin.read(item.filename))

            # Only provide theme/fontTable if completely absent from target document
            if "word/theme/theme1.xml" not in namelist and template_theme_xml:
                zout.writestr("word/theme/theme1.xml", template_theme_xml)
            if "word/fontTable.xml" not in namelist and template_font_table_xml:
                zout.writestr("word/fontTable.xml", template_font_table_xml)

        with open(ext_tmp, "wb") as tf:
            tf.write(out_buf.getvalue())

        os.replace(ext_tmp, ext_path)
        logger.info("Ensured authentic 3GPP styles and table formatting in %s", docx_path.name)
    finally:
        if os.path.exists(ext_tmp):
            try:
                os.remove(ext_tmp)
            except OSError:
                pass
