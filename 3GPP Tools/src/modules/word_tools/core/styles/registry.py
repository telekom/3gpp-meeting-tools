"""Canonical 3GPP Word style identifiers, registry metadata, and OpenXML TableGrid snippet."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, Set

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
