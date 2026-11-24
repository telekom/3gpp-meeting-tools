"""OpenXML style promotion, non-destructive merging, and Word COM QuickStyle promotion."""

from __future__ import annotations

import copy
import logging

from lxml import etree

from modules.word_tools.core.styles.registry import (
    ALL_3GPP_STYLE_NAMES_LOWER,
    LINKED_CHAR_STYLES,
    TABLE_GRID_XML,
    W_NS,
)

logger = logging.getLogger(__name__)


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
