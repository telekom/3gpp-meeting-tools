# --- File: src/modules/word_tools/core/author_renamer.py ---
"""
Core engine for inspecting and replacing Track Changes and comment authors in OpenXML (.docx) files.
Operates directly on the OpenXML ZIP package level with zero external pip dependencies.
"""
import io
import logging
import os
import platform
import re
import shutil
import tempfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Tuple, Union

logger = logging.getLogger(__name__)

# Standard OpenXML namespaces
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
CP_NS = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
DC_NS = "http://purl.org/dc/elements/1.1/"
DCTERMS_NS = "http://purl.org/dc/terms/"
DCMITYPE_NS = "http://purl.org/dc/dcmitype/"
XSI_NS = "http://www.w3.org/2001/XMLSchema-instance"

STANDARD_NAMESPACES = {
    "w": W_NS,
    "cp": CP_NS,
    "dc": DC_NS,
    "dcterms": DCTERMS_NS,
    "dcmitype": DCMITYPE_NS,
    "xsi": XSI_NS,
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "v": "urn:schemas-microsoft-com:vml",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "w10": "urn:schemas-microsoft-com:office:word",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    "w15": "http://schemas.microsoft.com/office/word/2012/wordml",
    "w16cid": "http://schemas.microsoft.com/office/word/2016/wordml/cid",
    "w16se": "http://schemas.microsoft.com/office/word/2015/wordml/symex",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
}

# Pre-register namespaces so ElementTree doesn't invent ns0:, ns1:, etc.
for prefix, uri in STANDARD_NAMESPACES.items():
    try:
        ET.register_namespace(prefix, uri)
    except Exception:
        pass

W_AUTHOR = f"{{{W_NS}}}author"
W_INITIALS = f"{{{W_NS}}}initials"

REVISION_XML_PATTERN = re.compile(
    r"^word/(document|comments|commentsExtended|commentsIds|header\d+|footer\d+|footnotes|endnotes|glossary/document)\.xml$",
    re.IGNORECASE,
)

CORE_PROPS_XML = "docProps/core.xml"


@dataclass
class AuthorStats:
    """Per-author breakdown of tracked changes and review comments."""
    name: str
    insertions: int = 0
    deletions: int = 0
    formatting: int = 0
    comments: int = 0
    initials: str = ""

    @property
    def total(self) -> int:
        return self.insertions + self.deletions + self.formatting + self.comments


def inspect_authors(docx_path: Union[str, Path]) -> List[AuthorStats]:
    """
    Inspects a .docx file and returns a list of AuthorStats for all authors found,
    sorted in descending order by total changes.
    """
    path = Path(docx_path).resolve()
    if not path.is_file():
        raise FileNotFoundError(f"File not found: {path}")

    stats_map: Dict[str, dict] = {}

    import zipfile
    with zipfile.ZipFile(path, "r") as z:
        for item in z.infolist():
            if REVISION_XML_PATTERN.match(item.filename):
                try:
                    xml_data = z.read(item.filename)
                    # Register any namespace discovered dynamically
                    try:
                        for _, (pfx, uri) in ET.iterparse(io.BytesIO(xml_data), events=["start-ns"]):
                            if pfx:
                                ET.register_namespace(pfx, uri)
                    except Exception:
                        pass
                    root = ET.fromstring(xml_data)
                except Exception as exc:
                    logger.warning("Could not parse XML part %s in %s: %s", item.filename, path.name, exc)
                    continue

                for elem in root.iter():
                    author = None
                    initials = ""
                    for k, v in elem.attrib.items():
                        if k == W_AUTHOR or k.endswith("}author") or k == "author":
                            author = v
                        elif k == W_INITIALS or k.endswith("}initials") or k == "initials":
                            initials = v

                    if not author:
                        continue

                    if author not in stats_map:
                        stats_map[author] = {
                            "ins": 0, "del": 0, "fmt": 0, "comm": 0, "initials": ""
                        }

                    tag = elem.tag.split("}")[-1] if "}" in elem.tag else elem.tag
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

    results = [
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
    results.sort(key=lambda s: s.total, reverse=True)
    return results


def replace_authors(
    docx_path: Union[str, Path],
    output_path: Union[str, Path],
    mapping: Dict[str, str],
    new_initials: Optional[str] = None,
    update_metadata: bool = True,
    backup: bool = True,
) -> Tuple[bool, int, str]:
    """
    Replaces author names across all revision XML parts and document metadata.

    Args:
        docx_path: Path to source .docx file.
        output_path: Destination path (can be same as docx_path for in-place update).
        mapping: Dictionary mapping old author name -> new author name.
        new_initials: If string, updates comment initials; if empty string, clears initials;
                      if None, leaves initials unchanged.
        update_metadata: If True, updates <dc:creator> and <cp:lastModifiedBy> in docProps/core.xml.
        backup: If True and updating in-place, creates a .bak backup file before overwriting.

    Returns:
        Tuple of (success: bool, replacement_count: int, message: str)
    """
    src = Path(docx_path).resolve()
    dst = Path(output_path).resolve()

    if not src.is_file():
        return False, 0, f"Source document not found: {src}"

    if not mapping:
        return True, 0, "No authors mapped for replacement."

    import zipfile

    is_inplace = (src == dst)
    if is_inplace and backup:
        bak_path = src.with_suffix(src.suffix + ".bak")
        try:
            shutil.copy2(src, bak_path)
            logger.info("Created backup file: %s", bak_path.name)
        except Exception as exc:
            logger.warning("Could not create backup file %s: %s", bak_path.name, exc)

    replacement_count = 0
    temp_dir = dst.parent
    temp_dir.mkdir(parents=True, exist_ok=True)

    # Use a temporary file in the destination folder to ensure atomic replacement
    fd, temp_file_path = tempfile.mkstemp(suffix=".docx", dir=str(temp_dir))
    os.close(fd)
    temp_path = Path(temp_file_path)

    try:
        with zipfile.ZipFile(src, "r") as z_in, zipfile.ZipFile(
            temp_path, "w", compression=zipfile.ZIP_DEFLATED
        ) as z_out:
            for item in z_in.infolist():
                data = z_in.read(item.filename)

                # Process revision & comment XML parts
                if REVISION_XML_PATTERN.match(item.filename):
                    try:
                        # Dynamically register any discovered prefixes
                        for _, (pfx, uri) in ET.iterparse(io.BytesIO(data), events=["start-ns"]):
                            if pfx:
                                ET.register_namespace(pfx, uri)
                        root = ET.fromstring(data)
                        modified = False

                        for elem in root.iter():
                            author_key = None
                            for k, v in elem.attrib.items():
                                if k == W_AUTHOR or k.endswith("}author") or k == "author":
                                    author_key = k
                                    break

                            if author_key and elem.attrib[author_key] in mapping:
                                target_name = mapping[elem.attrib[author_key]]
                                elem.attrib[author_key] = target_name
                                replacement_count += 1
                                modified = True

                                # Update comment initials if requested
                                if new_initials is not None:
                                    initials_key = None
                                    for k in elem.attrib:
                                        if k == W_INITIALS or k.endswith("}initials") or k == "initials":
                                            initials_key = k
                                            break
                                    if initials_key:
                                        if new_initials:
                                            elem.attrib[initials_key] = new_initials
                                        else:
                                            del elem.attrib[initials_key]

                        if modified:
                            data = ET.tostring(root, encoding="utf-8", xml_declaration=True)
                    except Exception as exc:
                        logger.warning("Could not process XML part %s: %s", item.filename, exc)

                # Process core document properties (metadata)
                elif update_metadata and item.filename.lower() == CORE_PROPS_XML.lower():
                    try:
                        for _, (pfx, uri) in ET.iterparse(io.BytesIO(data), events=["start-ns"]):
                            if pfx:
                                ET.register_namespace(pfx, uri)
                        root = ET.fromstring(data)
                        primary_target = next(iter(mapping.values()), "")
                        modified = False
                        if primary_target:
                            for elem in root.iter():
                                tag = elem.tag.split("}")[-1] if "}" in elem.tag else elem.tag
                                if tag in ("creator", "lastModifiedBy"):
                                    elem.text = primary_target
                                    modified = True
                        if modified:
                            data = ET.tostring(root, encoding="utf-8", xml_declaration=True)
                    except Exception as exc:
                        logger.warning("Could not update docProps/core.xml: %s", exc)

                z_out.writestr(item, data)

        # Atomic replacement / move
        if dst.exists():
            dst.unlink()
        shutil.move(temp_path, dst)
        return True, replacement_count, f"Successfully replaced {replacement_count} revisions & comments."

    except Exception as exc:
        if temp_path.exists():
            try:
                temp_path.unlink()
            except Exception:
                pass
        logger.error("Failed to replace authors in %s: %s", src.name, exc, exc_info=True)
        return False, 0, f"Error replacing authors: {exc}"


def is_file_open_in_word(file_path: Union[str, Path]) -> bool:
    """
    Checks whether Microsoft Word currently holds an open document matching
    file_path on Windows.
    """
    if platform.system() != "Windows":
        return False

    target = Path(file_path).resolve()
    target_name = target.name.casefold()
    target_full = str(target).casefold()

    import gc
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    try:
        try:
            word = win32com.client.GetActiveObject("Word.Application")
        except Exception:
            return False

        if word is None:
            return False

        for doc in word.Documents:
            try:
                d_name = str(doc.Name or "").strip().casefold()
                d_full = str(doc.FullName or "").strip().casefold()
                if d_full == target_full or d_name == target_name:
                    return True
            except Exception:
                continue
        return False
    finally:
        pythoncom.CoUninitialize()


def close_word_document_for_update(file_path: Union[str, Path], save_changes: bool = True) -> bool:
    """
    If the document is open in Microsoft Word, saves and closes it so the
    file lock is released for in-place OpenXML editing.
    """
    if platform.system() != "Windows":
        return False

    target = Path(file_path).resolve()
    target_name = target.name.casefold()
    target_full = str(target).casefold()

    import gc
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    try:
        try:
            word = win32com.client.GetActiveObject("Word.Application")
        except Exception:
            return False

        if word is None:
            return False

        for doc in word.Documents:
            try:
                d_name = str(doc.Name or "").strip().casefold()
                d_full = str(doc.FullName or "").strip().casefold()
                if d_full == target_full or d_name == target_name:
                    if save_changes:
                        doc.Save()
                    doc.Close(SaveChanges=False)
                    logger.info("Closed open Word document '%s' for in-place author replacement.", target.name)
                    return True
            except Exception:
                continue
        return False
    finally:
        pythoncom.CoUninitialize()
