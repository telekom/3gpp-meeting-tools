"""OpenXML document sanitization engine."""

from __future__ import annotations

import io
import logging
import os
import shutil
import tempfile
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path
from typing import Tuple, Union

from .models import (
    A_NS,
    CORE_PROPS_XML,
    STORY_XML_PATTERN,
    STYLES_XML_PATTERN,
    SanitizeConfig,
    W_AUTHOR,
    W_BIDI,
    W_EAST_ASIA,
    W_INITIALS,
    W_NS,
    W_VAL,
)
from .xml_helpers import _get_tag_name, _is_pure_language_change

logger = logging.getLogger(__name__)


def _register_part_namespaces(data: bytes) -> None:
    for _, (pfx, uri) in ET.iterparse(io.BytesIO(data), events=["start-ns"]):
        if pfx:
            ET.register_namespace(pfx, uri)


def _apply_lang_to_elem(lang_elem: ET.Element, target_language: str) -> None:
    lang_elem.attrib[W_VAL] = target_language
    if W_EAST_ASIA in lang_elem.attrib:
        lang_elem.attrib[W_EAST_ASIA] = target_language
    if W_BIDI in lang_elem.attrib:
        lang_elem.attrib[W_BIDI] = target_language


def _sanitize_styles_xml(data: bytes, config: SanitizeConfig, src_name: str) -> Tuple[bytes, int]:
    mod_count = 0
    try:
        _register_part_namespaces(data)
        root = ET.fromstring(data)
        modified = False

        # A. Update docDefaults
        doc_defaults = root.find(f".//{{{W_NS}}}docDefaults")
        if doc_defaults is not None:
            rpr_default = doc_defaults.find(f".//{{{W_NS}}}rPrDefault")
            if rpr_default is None:
                rpr_default = ET.SubElement(doc_defaults, f"{{{W_NS}}}rPrDefault")
            rpr = rpr_default.find(f"{{{W_NS}}}rPr")
            if rpr is None:
                rpr = ET.SubElement(rpr_default, f"{{{W_NS}}}rPr")

            lang_elem = rpr.find(f"{{{W_NS}}}lang")
            if lang_elem is None:
                lang_elem = ET.SubElement(rpr, f"{{{W_NS}}}lang")
            lang_elem.attrib[W_VAL] = config.target_language
            lang_elem.attrib[W_BIDI] = config.target_language
            if W_EAST_ASIA in lang_elem.attrib:
                lang_elem.attrib[W_EAST_ASIA] = config.target_language
            modified = True
            mod_count += 1

        # B. Update all styles with rPr/lang
        for style in root.findall(f".//{{{W_NS}}}style"):
            rpr = style.find(f"{{{W_NS}}}rPr")
            if rpr is not None:
                lang_elem = rpr.find(f"{{{W_NS}}}lang")
                if lang_elem is not None:
                    _apply_lang_to_elem(lang_elem, config.target_language)
                    modified = True
                    mod_count += 1
                if config.clear_no_proof:
                    np = rpr.find(f"{{{W_NS}}}noProof")
                    if np is not None:
                        rpr.remove(np)
                        modified = True

        if modified:
            data = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    except Exception as exc:
        logger.warning("Could not update styles.xml in %s: %s", src_name, exc)
    return data, mod_count


def _sanitize_story_xml(item_name: str, data: bytes, config: SanitizeConfig) -> Tuple[bytes, int]:
    mod_count = 0
    try:
        _register_part_namespaces(data)
        root = ET.fromstring(data)
        modified = False

        # A. Language Unification on Runs, Paragraphs, and DrawingML
        if config.unify_language:
            for r in root.findall(f".//{{{W_NS}}}r"):
                rpr = r.find(f"{{{W_NS}}}rPr")
                if rpr is not None:
                    lang_elem = rpr.find(f"{{{W_NS}}}lang")
                    if lang_elem is not None:
                        _apply_lang_to_elem(lang_elem, config.target_language)
                        modified = True
                        mod_count += 1
                    if config.clear_no_proof:
                        np = rpr.find(f"{{{W_NS}}}noProof")
                        if np is not None:
                            rpr.remove(np)
                            modified = True

            for p in root.findall(f".//{{{W_NS}}}p"):
                ppr = p.find(f"{{{W_NS}}}pPr")
                if ppr is not None:
                    rpr = ppr.find(f"{{{W_NS}}}rPr")
                    if rpr is not None:
                        lang_elem = rpr.find(f"{{{W_NS}}}lang")
                        if lang_elem is not None:
                            _apply_lang_to_elem(lang_elem, config.target_language)
                            modified = True
                            mod_count += 1

            for elem in root.iter():
                if elem.tag.startswith(f"{{{A_NS}}}"):
                    tag = _get_tag_name(elem)
                    if tag in ("rPr", "endParaRPr") and "lang" in elem.attrib:
                        elem.attrib["lang"] = config.target_language
                        if "altLang" in elem.attrib:
                            elem.attrib["altLang"] = config.target_language
                        modified = True
                        mod_count += 1

        # B. Surgical Language Revision Purge (<w:rPrChange> / <w:pPrChange>)
        if (
            config.unify_language
            and config.purge_language_changes
            and not config.purge_all_formatting_changes
        ):
            for rpr in root.findall(f".//{{{W_NS}}}rPr"):
                changes_to_remove = []
                for ch in list(rpr):
                    if _get_tag_name(ch) in ("rPrChange", "pPrChange"):
                        inner_rpr = None
                        has_lang = False
                        for sub in ch:
                            if _get_tag_name(sub) == "rPr":
                                inner_rpr = sub
                                for prop in sub:
                                    if _get_tag_name(prop) == "lang":
                                        has_lang = True
                                        break
                                break

                        if has_lang or _is_pure_language_change(rpr, ch):
                            if _is_pure_language_change(rpr, ch):
                                changes_to_remove.append(ch)
                            elif inner_rpr is not None:
                                inner_lang = inner_rpr.find(f"{{{W_NS}}}lang")
                                if inner_lang is not None:
                                    _apply_lang_to_elem(inner_lang, config.target_language)
                                    modified = True
                                    mod_count += 1

                for ch in changes_to_remove:
                    rpr.remove(ch)
                    modified = True
                    mod_count += 1

        # C. Purge ALL Formatting Revisions
        if config.purge_all_formatting_changes:
            for parent in root.iter():
                to_del = [
                    child
                    for child in list(parent)
                    if _get_tag_name(child)
                    in (
                        "rPrChange",
                        "pPrChange",
                        "tblPrChange",
                        "tblGridChange",
                        "tcPrChange",
                        "sectPrChange",
                    )
                ]
                for ch in to_del:
                    parent.remove(ch)
                    modified = True
                    mod_count += 1

        # D. Strip Comments
        if config.remove_all_comments:
            for parent in root.iter():
                to_del = [
                    child
                    for child in list(parent)
                    if _get_tag_name(child)
                    in ("commentRangeStart", "commentRangeEnd", "commentReference")
                ]
                for ch in to_del:
                    parent.remove(ch)
                    modified = True
                    mod_count += 1

            if item_name.lower().endswith("comments.xml"):
                for comm in list(root):
                    root.remove(comm)
                    modified = True
                    mod_count += 1

        # E. Revision Authors Replacement
        if config.update_authors and config.author_mapping:
            for elem in root.iter():
                author_key = None
                for k in elem.attrib:
                    if k == W_AUTHOR or k.endswith("}author") or k == "author":
                        author_key = k
                        break

                if author_key and elem.attrib[author_key] in config.author_mapping:
                    elem.attrib[author_key] = config.author_mapping[elem.attrib[author_key]]
                    mod_count += 1
                    modified = True

                    if config.new_initials is not None:
                        initials_key = None
                        for k in elem.attrib:
                            if k == W_INITIALS or k.endswith("}initials") or k == "initials":
                                initials_key = k
                                break
                        if initials_key:
                            if config.new_initials:
                                elem.attrib[initials_key] = config.new_initials
                            else:
                                del elem.attrib[initials_key]

        if modified:
            data = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    except Exception as exc:
        logger.warning("Could not process story part %s: %s", item_name, exc)
    return data, mod_count


def _sanitize_core_props_xml(data: bytes, config: SanitizeConfig) -> Tuple[bytes, int]:
    mod_count = 0
    try:
        _register_part_namespaces(data)
        root = ET.fromstring(data)
        modified = False

        target_author = ""
        if config.update_authors and config.author_mapping:
            target_author = next(iter(config.author_mapping.values()), "")

        if target_author:
            for elem in root.iter():
                if _get_tag_name(elem) in ("creator", "lastModifiedBy"):
                    elem.text = target_author
                    modified = True
                    mod_count += 1

        if modified:
            data = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    except Exception as exc:
        logger.warning("Could not update docProps/core.xml: %s", exc)
    return data, mod_count


def sanitize_document(
    docx_path: Union[str, Path],
    output_path: Union[str, Path],
    config: SanitizeConfig,
) -> Tuple[bool, int, str]:
    """
    Applies comprehensive document sanitization according to config.
    Returns (success: bool, modifications_count: int, message: str).
    """
    src = Path(docx_path).resolve()
    dst = Path(output_path).resolve()

    if not src.is_file():
        return False, 0, f"Source document not found: {src}"

    is_inplace = src == dst
    if is_inplace and config.backup:
        bak_path = src.with_suffix(src.suffix + ".bak")
        try:
            shutil.copy2(src, bak_path)
            logger.info("Created backup file: %s", bak_path.name)
        except Exception as exc:
            logger.warning("Could not create backup file %s: %s", bak_path.name, exc)

    mod_count = 0
    temp_dir = dst.parent
    temp_dir.mkdir(parents=True, exist_ok=True)

    fd, temp_file_path = tempfile.mkstemp(suffix=".docx", dir=str(temp_dir))
    os.close(fd)
    temp_path = Path(temp_file_path)

    try:
        with zipfile.ZipFile(src, "r") as z_in, zipfile.ZipFile(
            temp_path, "w", compression=zipfile.ZIP_DEFLATED
        ) as z_out:
            for item in z_in.infolist():
                data = z_in.read(item.filename)
                item_name = item.filename

                if (
                    STYLES_XML_PATTERN.match(item_name)
                    and config.unify_language
                    and config.update_styles_and_defaults
                ):
                    data, delta = _sanitize_styles_xml(data, config, src.name)
                    mod_count += delta
                elif STORY_XML_PATTERN.match(item_name):
                    data, delta = _sanitize_story_xml(item_name, data, config)
                    mod_count += delta
                elif config.update_metadata and item_name.lower() == CORE_PROPS_XML.lower():
                    data, delta = _sanitize_core_props_xml(data, config)
                    mod_count += delta

                z_out.writestr(item, data)

        if dst.exists():
            dst.unlink()
        shutil.move(temp_path, dst)
        return (
            True,
            mod_count,
            f"Successfully applied sanitization ({mod_count} operations applied).",
        )

    except Exception as exc:
        if temp_path.exists():
            try:
                temp_path.unlink()
            except Exception:
                pass
        logger.error("Failed to sanitize document %s: %s", src.name, exc, exc_info=True)
        return False, 0, f"Error sanitizing document: {exc}"
