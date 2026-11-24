"""Programmatic 3GPP Pseudo-CR (pCR) OpenXML document builder."""

from __future__ import annotations

import io
import zipfile
from pathlib import Path
from typing import Optional, Sequence
from xml.sax.saxutils import escape

from core.utils.paths import to_extended_path
from modules.word_tools.core.cr_pcr.models import (
    ChangeBlock,
    ContributionMetadata,
    R_NS,
    W_NS,
)
from modules.word_tools.core.styles_3gpp import (
    Style3GPP,
    ensure_3gpp_package_styles,
    get_pseudo_cr_template_path,
)


def build_pcr_base_document(
    meta: ContributionMetadata,
    change_blocks: Sequence[ChangeBlock],
    target_path: Path,
) -> None:
    """
    Programmatically builds the authentic 3GPP pseudo-CR (pCR) base document
    matching the official 3GPP template (Pseudo-CR-Form-v0-1.docx) and authentic
    contributions (e.g. S3-263735).

    Enforces:
    - 3GPP CRCoverPage style with 12pt Bold Arial header line (WG & non-italic TDoc).
    - Underline border rule on Meeting Date/Location line.
    - 3.5cm hanging indent (<w:ind w:left="1985" w:hanging="1985"/>) on metadata
      rows so tabs and multiline titles align in a clean, professional column.
    - Italic Arial for contribution abstract.
    - Official 3GPP 4-sided boxed blue 14pt Arial change separators.
    - XML character escaping on all delegate inputs.
    """
    tpl_path = get_pseudo_cr_template_path()
    with open(to_extended_path(tpl_path), "rb") as f:
        tpl_bytes = f.read()

    # Safe escaping helper
    def esc(text: Optional[str]) -> str:
        return escape(str(text or ""))

    # Date and location formatting
    date_parts = [p.strip() for p in (meta.meeting_dates, meta.meeting_location) if p and p.strip()]
    date_loc_str = ", ".join(date_parts)

    # Construct the body XML
    body_xml_parts = [
        f'<w:document xmlns:w="{W_NS}" xmlns:r="{R_NS}"><w:body>',
        # 1. Header Line: WG#Meeting \t TDoc (CRCoverPage, Arial 12pt Bold, non-italic TDoc)
        f'<w:p><w:pPr><w:pStyle w:val="{Style3GPP.CR_COVER_PAGE}"/>'
        f'<w:tabs><w:tab w:val="right" w:pos="9639"/></w:tabs>'
        f'<w:spacing w:after="0"/><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:b/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:pPr>'
        f'<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:b/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t>3GPP TSG-WG {esc(meta.meeting_wg)}#{esc(meta.meeting_num)}</w:t></w:r>'
        f'<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:b/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:tab/><w:t>{esc(meta.tdoc_id)}</w:t></w:r></w:p>',

        # 2. Date and Location (CRCoverPage, Arial 12pt Bold + Bottom Border dividing line)
        f'<w:p><w:pPr><w:pStyle w:val="{Style3GPP.CR_COVER_PAGE}"/>'
        f'<w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"/></w:pBdr>'
        f'<w:spacing w:after="120"/><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:b/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:pPr>'
        f'<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:b/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t>{esc(date_loc_str)}</w:t></w:r></w:p>',
    ]

    # 3. Metadata block with 3.5cm hanging indent (<w:ind w:left="1985" w:hanging="1985"/>)
    meta_rows = [
        ("Source:", meta.source_company),
        ("Title:", meta.title),
        ("Document for:", meta.doc_for),
        ("Agenda item:", meta.agenda_item),
        ("Work Item:", meta.work_item),
    ]

    for label, val in meta_rows:
        body_xml_parts.append(
            f'<w:p><w:pPr><w:spacing w:after="120"/>'
            f'<w:ind w:left="1985" w:hanging="1985"/>'
            f'<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:b/><w:bCs/></w:rPr></w:pPr>'
            f'<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:b/><w:bCs/></w:rPr><w:t xml:space="preserve">{label}\t</w:t></w:r>'
            f'<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:b/><w:bCs/></w:rPr><w:t>{esc(val)}</w:t></w:r></w:p>'
        )

    # Empty paragraph spacer following metadata block
    body_xml_parts.append(
        '<w:p><w:pPr><w:spacing w:after="120"/></w:pPr></w:p>'
    )

    # 4. Abstract (Arial Italic matching S3-263735)
    if meta.abstract.strip():
        body_xml_parts.append(
            f'<w:p><w:pPr><w:spacing w:after="180"/><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:i/><w:iCs/></w:rPr></w:pPr>'
            f'<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:i/><w:iCs/></w:rPr><w:t xml:space="preserve">Abstract of the contribution: </w:t></w:r>'
            f'<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:i/><w:iCs/></w:rPr><w:t>{esc(meta.abstract)}</w:t></w:r></w:p>'
        )

    # 5. Clause 1. Discussion
    disc_text = meta.discussion_text.strip() or (
        f"This document proposes updates to {meta.spec_type} {meta.spec_num} v{meta.spec_version} "
        f"for {meta.title}."
    )
    body_xml_parts.extend([
        f'<w:p><w:pPr><w:pStyle w:val="{Style3GPP.HEADING_1}"/><w:spacing w:before="240" w:after="120"/></w:pPr>'
        f'<w:r><w:t>1. Discussion</w:t></w:r></w:p>',
        f'<w:p><w:pPr><w:pStyle w:val="{Style3GPP.NORMAL}"/><w:spacing w:after="180"/></w:pPr>'
        f'<w:r><w:t>{esc(disc_text)}</w:t></w:r></w:p>',
    ])

    # 6. Clause 2. Proposal
    prop_text = meta.proposal_intro.strip() or (
        f"It is proposed to include the following changes in {meta.spec_type} {meta.spec_num} v{meta.spec_version}:"
    )
    body_xml_parts.extend([
        f'<w:p><w:pPr><w:pStyle w:val="{Style3GPP.HEADING_1}"/><w:spacing w:before="240" w:after="120"/></w:pPr>'
        f'<w:r><w:t>2. Proposal</w:t></w:r></w:p>',
        f'<w:p><w:pPr><w:pStyle w:val="{Style3GPP.NORMAL}"/><w:spacing w:after="180"/></w:pPr>'
        f'<w:r><w:t>{esc(prop_text)}</w:t></w:r></w:p>',
    ])

    # 7. Injected Change Separators and Anchors (Boxed Blue Arial 14pt)
    for idx, block in enumerate(change_blocks):
        if idx == 0:
            sep_label = "* * * First Change * * * *"
        else:
            sep_label = "* * * Next Change (all new)* * * *" if block.is_all_new else "* * * Next Change * * * *"

        body_xml_parts.extend([
            f'<w:p><w:pPr>'
            f'<w:pBdr>'
            f'<w:top w:val="single" w:sz="4" w:space="1" w:color="auto"/>'
            f'<w:left w:val="single" w:sz="4" w:space="4" w:color="auto"/>'
            f'<w:bottom w:val="single" w:sz="4" w:space="1" w:color="auto"/>'
            f'<w:right w:val="single" w:sz="4" w:space="4" w:color="auto"/>'
            f'</w:pBdr>'
            f'<w:jc w:val="center"/>'
            f'<w:spacing w:before="240" w:after="120"/>'
            f'<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:color w:val="0000FF"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr>'
            f'</w:pPr>'
            f'<w:r>'
            f'<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:color w:val="0000FF"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr>'
            f'<w:t>{sep_label}</w:t></w:r></w:p>',
            f'<w:p><w:r><w:t>{{{{CHANGE_BLOCK_ANCHOR_{idx}}}}}</w:t></w:r></w:p>',
        ])

    # Final End of Changes separator (Boxed Blue Arial 14pt)
    body_xml_parts.extend([
        '<w:p><w:pPr>'
        '<w:pBdr>'
        '<w:top w:val="single" w:sz="4" w:space="1" w:color="auto"/>'
        '<w:left w:val="single" w:sz="4" w:space="4" w:color="auto"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="1" w:color="auto"/>'
        '<w:right w:val="single" w:sz="4" w:space="4" w:color="auto"/>'
        '</w:pBdr>'
        '<w:jc w:val="center"/>'
        '<w:spacing w:before="240" w:after="120"/>'
        '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:color w:val="0000FF"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr>'
        '</w:pPr>'
        '<w:r>'
        '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:color w:val="0000FF"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr>'
        '<w:t>* * * End of Changes * * * *</w:t></w:r></w:p>',
        '<w:p><w:pPr><w:spacing w:after="0"/></w:pPr></w:p>',
        # Page layout / section properties
        '<w:sectPr><w:pgSz w:w="11907" w:h="16840" w:code="9"/>'
        '<w:pgMar w:top="1418" w:right="1134" w:bottom="1134" w:left="1134" w:header="680" w:footer="567" w:gutter="0"/>'
        '<w:cols w:space="720"/></w:sectPr>',
        '</w:body></w:document>',
    ])

    pcr_doc_xml = "".join(body_xml_parts).encode("utf-8")

    # Repackage template with new document.xml
    with zipfile.ZipFile(io.BytesIO(tpl_bytes), "r") as zin:
        out_buf = io.BytesIO()
        with zipfile.ZipFile(out_buf, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == "word/document.xml":
                    zout.writestr(item, pcr_doc_xml)
                else:
                    zout.writestr(item, zin.read(item.filename))

    with open(to_extended_path(target_path), "wb") as f:
        f.write(out_buf.getvalue())
    ensure_3gpp_package_styles(target_path)
