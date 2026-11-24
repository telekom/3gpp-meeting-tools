"""OpenXML tag/run-property comparison helpers for surgical revision analysis."""

from __future__ import annotations

import xml.etree.ElementTree as ET
from typing import List


def _get_tag_name(elem: ET.Element) -> str:
    """Returns local XML tag name without namespace."""
    tag = elem.tag
    return tag.split("}")[-1] if "}" in tag else tag


def _extract_rpr_signature(
    rpr_elem: ET.Element,
    ignore_change: bool = True,
    ignore_lang: bool = True,
) -> List[str]:
    """Generates a canonical signature list of run properties for comparison."""
    sig = []
    for c in rpr_elem:
        tag = _get_tag_name(c)
        if ignore_change and "Change" in tag:
            continue
        if ignore_lang and tag == "lang":
            continue
        attrib_str = ";".join(f"{k}={v}" for k, v in sorted(c.attrib.items()))
        sig.append(f"{tag}:{attrib_str}:{c.text or ''}")
    return sorted(sig)


def _is_pure_language_change(parent_rpr: ET.Element, rpr_change: ET.Element) -> bool:
    """
    Checks if rpr_change is solely tracking a change in w:lang.
    Compares parent_rpr (current properties, ignoring rprChange and lang)
    against inner rPr (previous properties, ignoring lang).
    """
    inner_rpr = None
    for child in rpr_change:
        if _get_tag_name(child) == "rPr":
            inner_rpr = child
            break

    if inner_rpr is None:
        return True

    outer_sig = _extract_rpr_signature(parent_rpr, ignore_change=True, ignore_lang=True)
    inner_sig = _extract_rpr_signature(inner_rpr, ignore_change=False, ignore_lang=True)
    return outer_sig == inner_sig
