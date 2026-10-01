"""Isolated fragment builder for Word excerpt extraction.

This helper intentionally runs in a separate Python process so serialization of
very large DOCX packages cannot starve the PyQt GUI event loop.
"""
import json
import shutil
import sys
from pathlib import Path
from typing import List, Tuple

from docx import Document
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P


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
    doc.save(target)
    print(f"Fragment saved: {target.name}", flush=True)


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
