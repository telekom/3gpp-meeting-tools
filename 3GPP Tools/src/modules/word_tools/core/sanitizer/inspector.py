"""Single-pass OpenXML document hygiene auditor."""

from __future__ import annotations

import logging
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path
from typing import Dict, Union

from .models import (
    A_NS,
    AuthorStats,
    CORE_PROPS_XML,
    DocumentHygieneAudit,
    LanguageStats,
    STORY_XML_PATTERN,
    STYLES_XML_PATTERN,
    W_AUTHOR,
    W_BIDI,
    W_EAST_ASIA,
    W_INITIALS,
    W_NS,
    W_VAL,
)
from .xml_helpers import _get_tag_name

logger = logging.getLogger(__name__)


def inspect_document(docx_path: Union[str, Path]) -> DocumentHygieneAudit:
    """
    Performs a high-speed (<150ms) single-pass inspection of an OpenXML document,
    extracting revision authors, proofing languages, change mark counts, and metadata.
    """
    path = Path(docx_path).resolve()
    if not path.is_file():
        raise FileNotFoundError(f"File not found: {path}")

    stats_map: Dict[str, dict] = {}
    lang_counts: Dict[str, int] = {}
    no_proof_count = 0
    lang_change_count = 0
    pure_lang_change_count = 0
    compound_lang_change_count = 0
    default_lang = ""
    creator = ""
    last_modified = ""

    file_size_kb = path.stat().st_size / 1024.0

    with zipfile.ZipFile(path, "r") as z:
        for item in z.infolist():
            # 1. Styles & Defaults
            if STYLES_XML_PATTERN.match(item.filename):
                try:
                    data = z.read(item.filename)
                    root = ET.fromstring(data)
                    # Inspect docDefaults
                    doc_def_lang = root.find(f".//{{{W_NS}}}docDefaults//{{{W_NS}}}lang")
                    if doc_def_lang is not None:
                        val = doc_def_lang.attrib.get(W_VAL) or doc_def_lang.attrib.get(W_EAST_ASIA)
                        if val:
                            default_lang = val
                except Exception as exc:
                    logger.debug("Failed inspecting styles in %s: %s", path.name, exc)

            # 2. Metadata (docProps/core.xml)
            elif item.filename.lower() == CORE_PROPS_XML.lower():
                try:
                    data = z.read(item.filename)
                    root = ET.fromstring(data)
                    for elem in root.iter():
                        tag = _get_tag_name(elem)
                        if tag == "creator" and elem.text:
                            creator = elem.text.strip()
                        elif tag == "lastModifiedBy" and elem.text:
                            last_modified = elem.text.strip()
                except Exception as exc:
                    logger.debug("Failed inspecting core props in %s: %s", path.name, exc)

            # 3. Document Stories & Comments
            elif STORY_XML_PATTERN.match(item.filename):
                try:
                    data = z.read(item.filename)
                    root = ET.fromstring(data)
                except Exception as exc:
                    logger.warning(
                        "Could not parse XML part %s in %s: %s", item.filename, path.name, exc
                    )
                    continue

                for elem in root.iter():
                    tag = _get_tag_name(elem)

                    # Language detection on runs / paragraph properties
                    if tag == "lang":
                        val = (
                            elem.attrib.get(W_VAL)
                            or elem.attrib.get(W_EAST_ASIA)
                            or elem.attrib.get(W_BIDI)
                        )
                        if val:
                            lang_counts[val] = lang_counts.get(val, 0) + 1
                    elif tag == "noProof":
                        no_proof_count += 1

                    # DrawingML language attributes
                    if tag in ("rPr", "endParaRPr") and elem.tag.startswith(f"{{{A_NS}}}"):
                        a_lang = elem.attrib.get("lang")
                        if a_lang:
                            lang_counts[a_lang] = lang_counts.get(a_lang, 0) + 1

                    # Track Changes revisions and authors
                    author = None
                    initials = ""
                    for k, v in elem.attrib.items():
                        if k == W_AUTHOR or k.endswith("}author") or k == "author":
                            author = v
                        elif k == W_INITIALS or k.endswith("}initials") or k == "initials":
                            initials = v

                    if author:
                        if author not in stats_map:
                            stats_map[author] = {
                                "ins": 0,
                                "del": 0,
                                "fmt": 0,
                                "comm": 0,
                                "initials": "",
                            }
                        if tag == "ins":
                            stats_map[author]["ins"] += 1
                        elif tag == "del":
                            stats_map[author]["del"] += 1
                        elif "Change" in tag or tag.startswith("move") or tag.startswith("customXml"):
                            stats_map[author]["fmt"] += 1
                        elif tag == "comment":
                            stats_map[author]["comm"] += 1

                        if initials and not stats_map[author]["initials"]:
                            stats_map[author]["initials"] = initials

                    # Check for language-specific revision marks (rPrChange)
                    if tag == "rPrChange":
                        has_lang_in_change = False
                        inner_rpr = None
                        for ch in elem:
                            if _get_tag_name(ch) == "rPr":
                                inner_rpr = ch
                                for sub in ch:
                                    if _get_tag_name(sub) == "lang":
                                        has_lang_in_change = True
                                        break
                                break

                        if has_lang_in_change:
                            lang_change_count += 1
                            if inner_rpr is not None and len(inner_rpr) == 1:
                                pure_lang_change_count += 1
                            else:
                                compound_lang_change_count += 1

    authors_list = [
        AuthorStats(
            name=name,
            insertions=d["ins"],
            deletions=d["del"],
            formatting=d["fmt"],
            comments=d["comm"],
            initials=d["initials"],
        )
        for name, d in stats_map.items()
    ]
    authors_list.sort(key=lambda s: s.total, reverse=True)

    lang_stats = LanguageStats(
        language_counts=lang_counts,
        no_proof_count=no_proof_count,
        language_change_marks=lang_change_count,
        pure_language_change_marks=pure_lang_change_count,
        compound_language_change_marks=compound_lang_change_count,
        default_language=default_lang,
    )

    tot_ins = sum(a.insertions for a in authors_list)
    tot_del = sum(a.deletions for a in authors_list)
    tot_fmt = sum(a.formatting for a in authors_list)
    tot_comm = sum(a.comments for a in authors_list)

    return DocumentHygieneAudit(
        file_path=path,
        file_name=path.name,
        file_size_kb=file_size_kb,
        authors=authors_list,
        total_insertions=tot_ins,
        total_deletions=tot_del,
        total_formatting=tot_fmt,
        total_comments=tot_comm,
        language_stats=lang_stats,
        doc_creator=creator,
        doc_last_modified_by=last_modified,
    )
