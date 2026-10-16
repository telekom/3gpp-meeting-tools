"""Isolated fragment builder for Word excerpt extraction.

This helper intentionally runs in a separate Python process so serialization of
very large DOCX packages cannot starve the PyQt GUI event loop.
"""
import json
import re
import shutil
import sys
from pathlib import Path
from typing import List, Tuple

from docx import Document
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P


# Matches r:id="rId...", r:embed="rId...", r:link="rId..." in surviving body XML
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
    merged = []
    for start, end in normalized:
        if not merged or start > merged[-1][1]:
            merged.append([start, end])
        else:
            merged[-1][1] = max(merged[-1][1], end)
    return [(start, end) for start, end in merged]


def prune_unused_media(doc) -> int:
    """
    Garbage Collector: Scans the surviving document body XML for active Relationship IDs (rId).
    Unlinks any heavy media/OLE/package relationships that are no longer referenced by the
    extracted clauses so python-docx omits their binary streams on save.
    """
    xml_str = doc.element.body.xml
    used_rids = set(RE_ACTIVE_RIDS.findall(xml_str))

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
    doc = Document(target)
    body = doc.element.body
    children = list(body)
    blocks = [x for x in children if isinstance(x, (CT_P, CT_Tbl))]
    merged = merge_ranges(ranges, len(blocks))

    keep_indexes = set()
    for start, end in merged:
        keep_indexes.update(range(start, end))
    keep_ids = {id(blocks[i]) for i in keep_indexes}

    print(
        f"Preparing {source.name}: keeping {len(keep_ids)}/{len(blocks)} "
        f"body blocks in {len(merged)} range(s)",
        flush=True,
    )

    retained = [
        child for child in children
        if not isinstance(child, (CT_P, CT_Tbl)) or id(child) in keep_ids
    ]
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