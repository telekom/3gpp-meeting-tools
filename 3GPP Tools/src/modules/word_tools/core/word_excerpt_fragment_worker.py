"""Isolated fragment builder for Word excerpt extraction.

This helper intentionally runs in a separate Python process so serialization of
very large DOCX packages cannot starve the PyQt GUI event loop.
"""
import json
import os
import re
import shutil
import stat
import sys
from pathlib import Path
from typing import List, Tuple

from docx import Document
from docx.oxml.ns import qn
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from lxml import etree


# Pre-resolve namespace-qualified attribute names for fast C-level lookup
_ATTR_R_ID = qn("r:id")
_ATTR_R_EMBED = qn("r:embed")
_ATTR_R_LINK = qn("r:link")
_REL_ATTRS = (_ATTR_R_ID, _ATTR_R_EMBED, _ATTR_R_LINK)

# Fallback regex only if non-standard r:* attributes exist
RE_ACTIVE_RIDS = re.compile(r'\br:(?:id|embed|link)="([^"]+)"')

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


def prune_unused_media(doc) -> int:
    """
    Garbage Collector: Scans the surviving document body for active Relationship IDs (rId).
    Unlinks any heavy media/OLE/package relationships that are no longer referenced by the
    extracted clauses so python-docx omits their binary streams on save.
    """
    body = doc.element.body

    # Fast C-level attribute scan via lxml (avoids pretty-printing body.xml into a huge string)
    used_rids = set()
    for elem in body.iter():
        attrib = elem.attrib
        if attrib:
            for attr_name in _REL_ATTRS:
                val = attrib.get(attr_name)
                if val is not None:
                    used_rids.add(val)

    # Also check unparsed VML/AlternateContent raw XML fragments if any exist via fast tostring
    raw_xml = etree.tostring(body, encoding="unicode")
    used_rids.update(RE_ACTIVE_RIDS.findall(raw_xml))

    rels = doc.part.rels
    pruned_count = 0
    for rId in list(rels.keys()):
        if rId not in used_rids:
            rel = rels[rId]
            reltype_lower = str(getattr(rel, "reltype", "")).lower()
            if any(marker in reltype_lower for marker in PRUNABLE_REL_MARKERS):
                del rels[rId]
                pruned_count += 1

    return pruned_count


def make_fragment(source: Path, target: Path, ranges: List[Tuple[int, int]]) -> None:
    shutil.copy2(source, target)
    # Ensure target is writable in case source had a read-only flag from IT/SharePoint/ZIP
    try:
        os.chmod(target, stat.S_IWRITE | stat.S_IREAD)
    except OSError:
        pass

    doc = Document(target)
    body = doc.element.body

    # Fast single-pass separation of blocks vs body-level structural elements (e.g. final sectPr, bookmarks)
    blocks = []
    non_blocks = []
    for child in body:
        if isinstance(child, (CT_P, CT_Tbl)):
            blocks.append(child)
        else:
            non_blocks.append(child)

    total_blocks = len(blocks)
    merged = merge_ranges(ranges, total_blocks)

    # Slice blocks directly using merged ranges instead of hashing thousands of element IDs
    retained = []
    for start, end in merged:
        retained.extend(blocks[start:end])

    kept_count = len(retained)
    print(
        f"Preparing {source.name}: keeping {kept_count}/{total_blocks} "
        f"body blocks in {len(merged)} range(s)",
        flush=True,
    )

    # Preserve non-block body elements (such as the trailing document <w:sectPr>) at the end
    retained.extend(non_blocks)

    # O(N) C-level bulk replacement of body children
    body[:] = retained

    # Prune orphaned Visio, OLE, and image relationships before saving
    pruned_rels = prune_unused_media(doc)
    if pruned_rels:
        print(
            f"Pruned {pruned_rels} unreferenced media/OLE relationship(s) from {source.name}",
            flush=True,
        )

    doc.save(target)
    size_kb = target.stat().st_size / 1024.0 if target.exists() else 0.0
    print(f"Fragment saved: {target.name} ({size_kb:.1f} KB)", flush=True)


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