"""Isolated fragment builder for Word excerpt extraction.

Fast pure-lxml + zipfile implementation that bypasses python-docx's eager
PackageReader/Part loading and custom OXML class lookup overhead.
"""
import json
import os
import posixpath
import re
import shutil
import stat
import sys
import time
import zipfile
from pathlib import Path
from typing import List, Set, Tuple

from lxml import etree


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

TAG_BODY = f"{{{W_NS}}}body"
TAG_P = f"{{{W_NS}}}p"
TAG_TBL = f"{{{W_NS}}}tbl"
BLOCK_TAGS = (TAG_P, TAG_TBL)

REL_ATTRS = (
    f"{{{R_NS}}}id",
    f"{{{R_NS}}}embed",
    f"{{{R_NS}}}link",
)

# Fallback regex for any relationship IDs inside VML or AlternateContent attributes
RE_ACTIVE_RIDS = re.compile(rb'\br:(?:id|embed|link)="([^"]+)"')

# Relationship types that carry heavy binary payloads (images, Visio/OLE, embedded packages)
PRUNABLE_REL_MARKERS = (
    "image",
    "oleobject",
    "package",
    "chart",
    "diagram",
    "audio",
    "video",
)


def merge_ranges(ranges: List[Tuple[int, int]], block_count: int) -> List[Tuple[int, int]]:
    normalized = sorted(
        (max(0, int(start)), min(int(end), block_count))
        for start, end in ranges
        if int(start) < int(end) and int(start) < block_count
    )
    merged: List[List[int]] = []
    for start, end in normalized:
        if not merged or start > merged[-1][1]:
            merged.append([start, end])
        else:
            merged[-1][1] = max(merged[-1][1], end)
    return [(start, end) for start, end in merged]


def make_fragment(source: Path, target: Path, ranges: List[Tuple[int, int]]) -> None:
    t0 = time.perf_counter()

    # 1. Read ONLY word/document.xml and word/_rels/document.xml.rels from source ZIP
    with zipfile.ZipFile(source, "r") as zin:
        doc_xml_bytes = zin.read("word/document.xml")
        try:
            rels_xml_bytes = zin.read("word/_rels/document.xml.rels")
        except KeyError:
            rels_xml_bytes = None

    # 2. Parse word/document.xml with raw C-level lxml (no python-docx oxml overhead)
    parser = etree.XMLParser(remove_blank_text=False, resolve_entities=False, huge_tree=True)
    root = etree.fromstring(doc_xml_bytes, parser=parser)
    body = root.find(TAG_BODY)
    if body is None:
        raise RuntimeError(f"No <w:body> found in {source}")

    # Separate paragraphs/tables from body-level structural nodes (e.g. final <w:sectPr>)
    blocks = []
    non_blocks = []
    for child in body:
        if child.tag in BLOCK_TAGS:
            blocks.append(child)
        else:
            non_blocks.append(child)

    total_blocks = len(blocks)
    merged = merge_ranges(ranges, total_blocks)

    retained = []
    for start, end in merged:
        retained.extend(blocks[start:end])

    kept_count = len(retained)
    print(
        f"Preparing {source.name}: keeping {kept_count}/{total_blocks} "
        f"body blocks in {len(merged)} range(s)",
        flush=True,
    )

    retained.extend(non_blocks)
    body[:] = retained

    # Serialize modified word/document.xml
    new_doc_xml = etree.tostring(
        root,
        encoding="utf-8",
        xml_declaration=True,
        standalone="yes",
    )

    # 3. Collect all active rId references in the surviving document.xml
    used_rids: Set[str] = set()
    for elem in body.iter():
        attrib = elem.attrib
        if attrib:
            for attr_name in REL_ATTRS:
                val = attrib.get(attr_name)
                if val is not None:
                    used_rids.add(val)

    for match in RE_ACTIVE_RIDS.findall(new_doc_xml):
        used_rids.add(match.decode("utf-8", errors="ignore"))

    # 4. Prune unused heavy relationships in word/_rels/document.xml.rels
    pruned_rels = 0
    pruned_zip_paths: Set[str] = set()
    kept_zip_paths: Set[str] = set()
    new_rels_xml = rels_xml_bytes

    if rels_xml_bytes is not None:
        rels_root = etree.fromstring(rels_xml_bytes, parser=parser)
        rels_to_remove = []

        for rel in rels_root:
            r_id = rel.get("Id", "")
            r_type = rel.get("Type", "").lower()
            r_target = rel.get("Target", "")
            target_mode = rel.get("TargetMode", "")

            # Resolve internal package path relative to "word/"
            norm_path = ""
            if r_target and target_mode.lower() != "external":
                if r_target.startswith("/"):
                    norm_path = r_target.lstrip("/")
                else:
                    norm_path = posixpath.normpath(posixpath.join("word", r_target))

            if r_id not in used_rids and any(m in r_type for m in PRUNABLE_REL_MARKERS):
                rels_to_remove.append(rel)
                pruned_rels += 1
                if norm_path:
                    pruned_zip_paths.add(norm_path)
            else:
                if norm_path:
                    kept_zip_paths.add(norm_path)

        for rel in rels_to_remove:
            rels_root.remove(rel)

        # Never delete a binary file if another kept relationship still points to it
        pruned_zip_paths -= kept_zip_paths

        new_rels_xml = etree.tostring(
            rels_root,
            encoding="utf-8",
            xml_declaration=True,
            standalone="yes",
        )

    if pruned_rels:
        print(
            f"Pruned {pruned_rels} unreferenced media/OLE relationship(s) from {source.name}",
            flush=True,
        )

    # 5. Copy source -> target first to preserve OS ACLs / IT permissions, then overwrite contents
    shutil.copy2(source, target)
    try:
        os.chmod(target, stat.S_IWRITE | stat.S_IREAD)
    except OSError:
        pass

    # Write the slimmed DOCX package directly, skipping pruned media/OLE binaries
    with zipfile.ZipFile(source, "r") as zin, zipfile.ZipFile(
        target, "w", compression=zipfile.ZIP_DEFLATED
    ) as zout:
        for item in zin.infolist():
            filename = item.filename
            if filename in pruned_zip_paths:
                continue
            if filename == "word/document.xml":
                zout.writestr(item, new_doc_xml)
            elif filename == "word/_rels/document.xml.rels" and new_rels_xml is not None:
                zout.writestr(item, new_rels_xml)
            else:
                zout.writestr(item, zin.read(filename))

    elapsed = time.perf_counter() - t0
    size_kb = target.stat().st_size / 1024.0 if target.exists() else 0.0
    print(
        f"Fragment saved: {target.name} ({size_kb:.1f} KB) in {elapsed:.2f}s",
        flush=True,
    )


def main() -> int:
    if len(sys.argv) != 4:
        print("Usage: word_excerpt_fragment_worker.py SOURCE TARGET RANGES_JSON", file=sys.stderr)
        return 2
    source = Path(sys.argv[1])
    target = Path(sys.argv[2])
    ranges = json.loads(sys.argv[3])
    make_fragment(source, target, ranges)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())