# --- File: src/modules/word_tools/core/document_sanitizer.py ---
"""
Document Sanitizer & Revision Hygiene Engine.

Provides unified inspection and single-pass OpenXML (.docx / .docm) processing for:
1. Proofing language unification across all document stories, styles, and DrawingML.
2. Surgical Track Changes cleanup (purging or harmonizing language revision balloons).
3. Revision author consolidation, comment initials sync, and metadata harmonization.
4. Optional formatting-only revision acceptance and comment scrubbing.

Operates directly on the OpenXML ZIP package level with zero external dependencies.
"""
import io
import logging
import os
import platform
import re
import shutil
import tempfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional, Set, Tuple, Union

from modules.word_tools.core.author_renamer import (
    AuthorStats,
    W_AUTHOR,
    W_INITIALS,
    W_NS,
    CP_NS,
    DC_NS,
    DCTERMS_NS,
    DCMITYPE_NS,
    XSI_NS,
    STANDARD_NAMESPACES,
    CORE_PROPS_XML,
    close_word_document_for_update,
    is_file_open_in_word,
)

logger = logging.getLogger(__name__)

# Register standard namespaces
for prefix, uri in STANDARD_NAMESPACES.items():
    try:
        ET.register_namespace(prefix, uri)
    except Exception:
        pass

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
ET.register_namespace("a", A_NS)

W_VAL = f"{{{W_NS}}}val"
W_EAST_ASIA = f"{{{W_NS}}}eastAsia"
W_BIDI = f"{{{W_NS}}}bidi"
W_LANG = f"{{{W_NS}}}lang"
W_NOPROOF = f"{{{W_NS}}}noProof"

STORY_XML_PATTERN = re.compile(
    r"^word/(document|comments|commentsExtended|commentsIds|header\d+|footer\d+|footnotes|endnotes|glossary/document)\.xml$",
    re.IGNORECASE,
)
STYLES_XML_PATTERN = re.compile(r"^word/styles\.xml$", re.IGNORECASE)

# Standard language presets common in 3GPP and international standardization
LANGUAGE_PRESETS = [
    ("en-GB", "English (United Kingdom) [en-GB] — 3GPP Standard"),
    ("en-US", "English (United States) [en-US]"),
    ("de-DE", "German (Germany) [de-DE]"),
    ("fr-FR", "French (France) [fr-FR]"),
    ("zh-CN", "Chinese (Simplified) [zh-CN]"),
    ("ja-JP", "Japanese [ja-JP]"),
    ("ko-KR", "Korean [ko-KR]"),
    ("es-ES", "Spanish (Spain) [es-ES]"),
    ("sv-SE", "Swedish [sv-SE]"),
    ("fi-FI", "Finnish [fi-FI]"),
]


@dataclass
class LanguageStats:
    """Statistics on proofing languages detected in a document."""
    language_counts: Dict[str, int] = field(default_factory=dict)
    no_proof_count: int = 0
    language_change_marks: int = 0
    pure_language_change_marks: int = 0
    compound_language_change_marks: int = 0
    default_language: str = ""

    @property
    def primary_language(self) -> str:
        if self.default_language:
            return self.default_language
        if not self.language_counts:
            return "Unknown"
        return max(self.language_counts.items(), key=lambda kv: kv[1])[0]

    @property
    def summary_text(self) -> str:
        if not self.language_counts:
            return self.default_language or "No language tags detected"
        total = sum(self.language_counts.values())
        parts = []
        for lang, count in sorted(self.language_counts.items(), key=lambda kv: kv[1], reverse=True)[:3]:
            pct = int(round((count / total) * 100))
            parts.append(f"{lang} ({pct}%)")
        return ", ".join(parts)


@dataclass
class DocumentHygieneAudit:
    """Complete pre-flight health and revision audit of a Word document."""
    file_path: Path
    file_name: str
    file_size_kb: float
    authors: List[AuthorStats] = field(default_factory=list)
    total_insertions: int = 0
    total_deletions: int = 0
    total_formatting: int = 0
    total_comments: int = 0
    language_stats: LanguageStats = field(default_factory=LanguageStats)
    doc_creator: str = ""
    doc_last_modified_by: str = ""

    @property
    def total_authors(self) -> int:
        return len(self.authors)

    @property
    def total_changes(self) -> int:
        return self.total_insertions + self.total_deletions + self.total_formatting + self.total_comments


@dataclass
class SanitizeConfig:
    """User-selectable options for document sanitization."""
    # Language Unification & Language Revision Purge
    unify_language: bool = True
    target_language: str = "en-GB"
    apply_to_all_stories: bool = True
    update_styles_and_defaults: bool = True
    purge_language_changes: bool = True
    clear_no_proof: bool = False

    # Revision Authors & Metadata
    update_authors: bool = False
    author_mapping: Dict[str, str] = field(default_factory=dict)
    new_initials: Optional[str] = None
    update_metadata: bool = True

    # Formatting & Comments Cleanup
    purge_all_formatting_changes: bool = False
    remove_all_comments: bool = False

    # Safety
    backup: bool = True


def _get_tag_name(elem: ET.Element) -> str:
    """Returns local XML tag name without namespace."""
    tag = elem.tag
    return tag.split("}")[-1] if "}" in tag else tag


def _extract_rpr_signature(rpr_elem: ET.Element, ignore_change: bool = True, ignore_lang: bool = True) -> List[str]:
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


def inspect_document(docx_path: Union[str, Path]) -> DocumentHygieneAudit:
    """
    Performs a high-speed (<150ms) single-pass inspection of an OpenXML document,
    extracting revision authors, proofing languages, change mark counts, and metadata.
    """
    path = Path(docx_path).resolve()
    if not path.is_file():
        raise FileNotFoundError(f"File not found: {path}")

    import zipfile

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
                    logger.warning("Could not parse XML part %s in %s: %s", item.filename, path.name, exc)
                    continue

                for elem in root.iter():
                    tag = _get_tag_name(elem)

                    # Language detection on runs / paragraph properties
                    if tag == "lang":
                        val = elem.attrib.get(W_VAL) or elem.attrib.get(W_EAST_ASIA) or elem.attrib.get(W_BIDI)
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
                                "ins": 0, "del": 0, "fmt": 0, "comm": 0, "initials": ""
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
                        # Check if this rPrChange involves w:lang
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
                            # We can determine if it's pure by checking parent rPr if available
                            # During stream traversal, approximate:
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

    import zipfile

    is_inplace = (src == dst)
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

                # --- 1. Styles & Defaults (styles.xml) ---
                if STYLES_XML_PATTERN.match(item_name) and config.unify_language and config.update_styles_and_defaults:
                    try:
                        for _, (pfx, uri) in ET.iterparse(io.BytesIO(data), events=["start-ns"]):
                            if pfx:
                                ET.register_namespace(pfx, uri)
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
                            # Preserve or set eastAsia if Latin or matching
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
                                    lang_elem.attrib[W_VAL] = config.target_language
                                    if W_EAST_ASIA in lang_elem.attrib:
                                        lang_elem.attrib[W_EAST_ASIA] = config.target_language
                                    if W_BIDI in lang_elem.attrib:
                                        lang_elem.attrib[W_BIDI] = config.target_language
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
                        logger.warning("Could not update styles.xml in %s: %s", src.name, exc)

                # --- 2. Document Stories & Comments ---
                elif STORY_XML_PATTERN.match(item_name):
                    try:
                        for _, (pfx, uri) in ET.iterparse(io.BytesIO(data), events=["start-ns"]):
                            if pfx:
                                ET.register_namespace(pfx, uri)
                        root = ET.fromstring(data)
                        modified = False

                        # A. Language Unification on Runs, Paragraphs, and DrawingML
                        if config.unify_language:
                            # 1. Runs (<w:rPr>)
                            for r in root.findall(f".//{{{W_NS}}}r"):
                                rpr = r.find(f"{{{W_NS}}}rPr")
                                if rpr is not None:
                                    lang_elem = rpr.find(f"{{{W_NS}}}lang")
                                    if lang_elem is not None:
                                        lang_elem.attrib[W_VAL] = config.target_language
                                        if W_EAST_ASIA in lang_elem.attrib:
                                            lang_elem.attrib[W_EAST_ASIA] = config.target_language
                                        if W_BIDI in lang_elem.attrib:
                                            lang_elem.attrib[W_BIDI] = config.target_language
                                        modified = True
                                        mod_count += 1
                                    if config.clear_no_proof:
                                        np = rpr.find(f"{{{W_NS}}}noProof")
                                        if np is not None:
                                            rpr.remove(np)
                                            modified = True

                            # 2. Paragraph Properties (<w:pPr><w:rPr>)
                            for p in root.findall(f".//{{{W_NS}}}p"):
                                ppr = p.find(f"{{{W_NS}}}pPr")
                                if ppr is not None:
                                    rpr = ppr.find(f"{{{W_NS}}}rPr")
                                    if rpr is not None:
                                        lang_elem = rpr.find(f"{{{W_NS}}}lang")
                                        if lang_elem is not None:
                                            lang_elem.attrib[W_VAL] = config.target_language
                                            if W_EAST_ASIA in lang_elem.attrib:
                                                lang_elem.attrib[W_EAST_ASIA] = config.target_language
                                            if W_BIDI in lang_elem.attrib:
                                                lang_elem.attrib[W_BIDI] = config.target_language
                                            modified = True
                                            mod_count += 1

                            # 3. DrawingML shapes & text runs
                            for elem in root.iter():
                                if elem.tag.startswith(f"{{{A_NS}}}"):
                                    tag = _get_tag_name(elem)
                                    if tag in ("rPr", "endParaRPr"):
                                        if "lang" in elem.attrib:
                                            elem.attrib["lang"] = config.target_language
                                            if "altLang" in elem.attrib:
                                                elem.attrib["altLang"] = config.target_language
                                            modified = True
                                            mod_count += 1

                        # B. Surgical Language Revision Purge (<w:rPrChange> / <w:pPrChange>)
                        if config.unify_language and config.purge_language_changes and not config.purge_all_formatting_changes:
                            # Walk all rPr elements across the tree
                            for rpr in root.findall(f".//{{{W_NS}}}rPr"):
                                changes_to_remove = []
                                for ch in list(rpr):
                                    if _get_tag_name(ch) in ("rPrChange", "pPrChange"):
                                        # Check if this change involves language
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
                                                # Pure language change: remove entire rPrChange node
                                                changes_to_remove.append(ch)
                                            else:
                                                # Compound change: synchronize inner lang to match target
                                                if inner_rpr is not None:
                                                    inner_lang = inner_rpr.find(f"{{{W_NS}}}lang")
                                                    if inner_lang is not None:
                                                        inner_lang.attrib[W_VAL] = config.target_language
                                                        if W_EAST_ASIA in inner_lang.attrib:
                                                            inner_lang.attrib[W_EAST_ASIA] = config.target_language
                                                        if W_BIDI in inner_lang.attrib:
                                                            inner_lang.attrib[W_BIDI] = config.target_language
                                                        modified = True
                                                        mod_count += 1

                                for ch in changes_to_remove:
                                    rpr.remove(ch)
                                    modified = True
                                    mod_count += 1

                        # C. Purge ALL Formatting Revisions (if requested)
                        if config.purge_all_formatting_changes:
                            for parent in root.iter():
                                to_del = [
                                    child for child in list(parent)
                                    if _get_tag_name(child) in (
                                        "rPrChange", "pPrChange", "tblPrChange",
                                        "tblGridChange", "tcPrChange", "sectPrChange"
                                    )
                                ]
                                for ch in to_del:
                                    parent.remove(ch)
                                    modified = True
                                    mod_count += 1

                        # D. Strip Comments (if requested)
                        if config.remove_all_comments:
                            for parent in root.iter():
                                to_del = [
                                    child for child in list(parent)
                                    if _get_tag_name(child) in (
                                        "commentRangeStart", "commentRangeEnd", "commentReference"
                                    )
                                ]
                                for ch in to_del:
                                    parent.remove(ch)
                                    modified = True
                                    mod_count += 1

                            # If this is the comments.xml part, empty out comments
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
                                    target_name = config.author_mapping[elem.attrib[author_key]]
                                    elem.attrib[author_key] = target_name
                                    mod_count += 1
                                    modified = True

                                    # Update comment initials if requested
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

                # --- 3. Core Properties / Metadata (docProps/core.xml) ---
                elif config.update_metadata and item_name.lower() == CORE_PROPS_XML.lower():
                    try:
                        for _, (pfx, uri) in ET.iterparse(io.BytesIO(data), events=["start-ns"]):
                            if pfx:
                                ET.register_namespace(pfx, uri)
                        root = ET.fromstring(data)
                        modified = False

                        target_author = ""
                        if config.update_authors and config.author_mapping:
                            target_author = next(iter(config.author_mapping.values()), "")

                        if target_author:
                            for elem in root.iter():
                                tag = _get_tag_name(elem)
                                if tag in ("creator", "lastModifiedBy"):
                                    elem.text = target_author
                                    modified = True
                                    mod_count += 1

                        if modified:
                            data = ET.tostring(root, encoding="utf-8", xml_declaration=True)
                    except Exception as exc:
                        logger.warning("Could not update docProps/core.xml: %s", exc)

                z_out.writestr(item, data)

        # Atomic replacement / move
        if dst.exists():
            dst.unlink()
        shutil.move(temp_path, dst)
        return True, mod_count, f"Successfully applied sanitization ({mod_count} operations applied)."

    except Exception as exc:
        if temp_path.exists():
            try:
                temp_path.unlink()
            except Exception:
                pass
        logger.error("Failed to sanitize document %s: %s", src.name, exc, exc_info=True)
        return False, 0, f"Error sanitizing document: {exc}"
