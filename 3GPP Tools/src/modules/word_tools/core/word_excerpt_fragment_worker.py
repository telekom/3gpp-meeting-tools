"""Isolated fragment builder for Word excerpt extraction.

Ultra-fast Read-Only DOM implementation:
- Parses word/document.xml in ~0.25s with lxml.
- NEVER mutates, unlinks, or moves DOM nodes (completely avoiding libxml2's
  catastrophic O(N^2) _moveNodeToDocument / namespace reconciliation storm).
- Serializes ONLY the surviving blocks (~100 blocks out of 12,747) in ~0.01s.
"""
import io
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
BLOCK_TAGS = {TAG_P, TAG_TBL}

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
    t_start = time.perf_counter()

    # 1. Read source ZIP into memory once
    t0 = time.perf_counter()
    source_bytes = source.read_bytes()
    with zipfile.ZipFile(io.BytesIO(source_bytes), "r") as zin:
        doc_xml_bytes = zin.read("word/document.xml")
        try:
            rels_xml_bytes = zin.read("word/_rels/document.xml.rels")
        except KeyError:
            rels_xml_bytes = None
    t_read = time.perf_counter() - t0

    # 2. Parse word/document.xml with C-level lxml (takes ~0.27s on TS 23.501)
    t0 = time.perf_counter()
    parser = etree.XMLParser(remove_blank_text=False, resolve_entities=False, huge_tree=True)
    root = etree.fromstring(doc_xml_bytes, parser=parser)
    body = root.find(TAG_BODY)
    if body is None:
        raise RuntimeError(f"No <w:body> found in {source}")
    t_parse = time.perf_counter() - t0

    # 3. READ-ONLY DOM SLICING:
    #    NEVER remove, unlink, or reparent nodes in `body`! Mutating a 2M-node
    #    lxml tree triggers O(N^2) namespace reconciliation in libxml2.
    #    Instead, count blocks, then serialize ONLY the kept children.
    t0 = time.perf_counter()
    total_blocks = 0
    for child in body:
        if child.tag in BLOCK_TAGS:
            total_blocks += 1

    merged = merge_ranges(ranges, total_blocks)

    keep_block_indices: Set[int] = set()
    for start, end in merged:
        keep_block_indices.update(range(start, end))

    kept_count = len(keep_block_indices)
    print(
        f"Preparing {source.name}: keeping {kept_count}/{total_blocks} "
        f"body blocks in {len(merged)} range(s)",
        flush=True,
    )

    # Extract original XML prefix up to <w:body...> so all root namespaces are preserved intact
    body_tag_pos = doc_xml_bytes.find(b"<w:body")
    if body_tag_pos != -1:
        body_open_end = doc_xml_bytes.find(b">", body_tag_pos) + 1
        xml_prefix = doc_xml_bytes[:body_open_end]
        xml_suffix = b"</w:body></w:document>"
    else:
        # Fallback if default namespace is used without 'w:' prefix
        body_tag_pos = doc_xml_bytes.find(b"<body")
        body_open_end = doc_xml_bytes.find(b">", body_tag_pos) + 1
        xml_prefix = doc_xml_bytes[:body_open_end]
        xml_suffix = b"</body></document>"

    kept_chunks: List[bytes] = [xml_prefix]
    used_rids: Set[str] = set()

    block_idx = 0
    for child in body:
        is_block = child.tag in BLOCK_TAGS
        keep = (block_idx in keep_block_indices) if is_block else True
        if is_block:
            block_idx += 1

        if keep:
            # Serialize only this kept element
            kept_chunks.append(etree.tostring(child, encoding="utf-8"))
            # Collect active rIds from this kept element
            for elem in child.iter():
                attrib = elem.attrib
                if attrib:
                    for attr_name in REL_ATTRS:
                        val = attrib.get(attr_name)
                        if val is not None:
                            used_rids.add(val)

    kept_chunks.append(xml_suffix)
    new_doc_xml = b"".join(kept_chunks)
    t_slice = time.perf_counter() - t0

    # 4. Prune unused heavy relationships in word/_rels/document.xml.rels
    t0 = time.perf_counter()
    for match in RE_ACTIVE_RIDS.findall(new_doc_xml):
        used_rids.add(match.decode("utf-8", errors="ignore"))

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

        pruned_zip_paths -= kept_zip_paths

        new_rels_xml = etree.tostring(
            rels_root,
            encoding="utf-8",
            xml_declaration=True,
            standalone="yes",
        )
    t_prune = time.perf_counter() - t0

    if pruned_rels:
        print(
            f"Pruned {pruned_rels} unreferenced media/OLE relationship(s) from {source.name}",
            flush=True,
        )

    # 5. Assemble output DOCX in RAM, then write atomically
    t0 = time.perf_counter()
    out_buffer = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(source_bytes), "r") as zin, zipfile.ZipFile(
        out_buffer, "w", compression=zipfile.ZIP_DEFLATED
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

    shutil.copy2(source, target)
    try:
        os.chmod(target, stat.S_IWRITE | stat.S_IREAD)
    except OSError:
        pass
    target.write_bytes(out_buffer.getvalue())
    t_write = time.perf_counter() - t0

    elapsed = time.perf_counter() - t_start
    size_kb = target.stat().st_size / 1024.0 if target.exists() else 0.0
    print(
        f"Fragment saved: {target.name} ({size_kb:.1f} KB) in {elapsed:.2f}s "
        f"[read={t_read:.2f}s, parse={t_parse:.2f}s, slice={t_slice:.2f}s, "
        f"prune={t_prune:.2f}s, write={t_write:.2f}s]",
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
    # Fast process exit skips Python GC teardown of the 2M-node read-only XML tree
    sys.stdout.flush()
    sys.stderr.flush()
    os._exit(0)


if __name__ == "__main__":
    raise SystemExit(main())