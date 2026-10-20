"""Isolated fragment builder for Word excerpt extraction.

This helper runs in a separate Python process so that excerpt preparation and
serialization of very large DOCX packages (such as 3GPP specifications with
tens of thousands of paragraphs, large tables, and hundreds of embedded Visio/OLE
diagrams) cannot freeze or starve the PyQt GUI event loop.

ARCHITECTURAL NOTES & LESSONS LEARNED:
-------------------------------------
1. Why not high-level python-docx (Document)?
   Calling `Document(source)` eagerly walks the entire OpenXML package relationship
   graph via `PackageReader._walk_parts` and loads every single image, chart, and
   embedded OLE object into memory. For large 3GPP specifications (e.g. TS 23.501
   with 300+ embedded OLE diagrams), this eager unpacking and subsequent re-serialization
   takes several minutes and consumes hundreds of megabytes of RAM.

2. Why a READ-ONLY DOM slice instead of DOM mutations (`body[:] = ...` or `body.remove`)?
   In `lxml.etree` (backed by C `libxml2`), unlinking or reparenting 12,000+ elements
   from `<w:body>` triggers `_moveNodeToDocument` and `xmlReconciliateNs` on every
   unlinked subtree. When Python holds proxy references to discarded nodes, libxml2
   creates individual `_fakeRootDoc` wrappers and recursively walks descendant nodes,
   leading to an O(N^2) CPU bottleneck (taking 2 to 9 minutes on large specs).
   By keeping the parsed tree 100% read-only and serializing ONLY the kept nodes via
   `etree.tostring(child)`, we achieve near-instantaneous execution (~0.01s).

3. Multi-File Word Merge Compatibility & Unlocking:
   - Strips NTFS `:Zone.Identifier` (Mark-of-the-Web) alternate data streams that
     `shutil.copy2` copies from downloaded 3GPP ZIPs, preventing Word COM from
     opening fragments in Protected View ("locked for editing").
   - Strips `<w:documentProtection>` and `<w:writeProtection>` from `word/settings.xml`
     so Word COM `InsertFile` never fails with error -2146823683 ("locked for editing").
   - Cleans `[Content_Types].xml` overrides for pruned media/OLE parts so Word's
     OPC validator never flags missing package parts during `InsertFile`.
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
_worker_src = Path(__file__).resolve().parent.parent.parent.parent
if str(_worker_src) not in sys.path:
    sys.path.insert(0, str(_worker_src))
from core.utils.paths import to_extended_path
from typing import List, Set, Tuple

from lxml import etree


# ---------------------------------------------------------------------------
# OpenXML XML Namespaces and Qualified Tag Names
# ---------------------------------------------------------------------------
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

TAG_BODY = f"{{{W_NS}}}body"
TAG_P = f"{{{W_NS}}}p"
TAG_TBL = f"{{{W_NS}}}tbl"
TAG_SECTPR = f"{{{W_NS}}}sectPr"

# 3GPP document body flow consists of paragraphs (<w:p>) and tables (<w:tbl>).
# Any other body child (such as the document-level <w:sectPr>) is considered structural.
BLOCK_TAGS = {TAG_P, TAG_TBL}

# Relationship ID attributes used by runs, drawings, and OLE objects to link
# to external or package-internal media parts
REL_ATTRS = (
    f"{{{R_NS}}}id",
    f"{{{R_NS}}}embed",
    f"{{{R_NS}}}link",
)

# Fallback regex for legacy VML shapes or non-standard AlternateContent blocks
RE_ACTIVE_RIDS = re.compile(rb'\br:(?:id|embed|link)="([^"]+)"')

# Regexes to strip editing locks from word/settings.xml so Word COM InsertFile never fails
RE_DOC_PROTECTION = re.compile(
    rb"<(?:w:)?(?:documentProtection|writeProtection)\b[^>]*/?>",
    re.IGNORECASE,
)

# Relationship types carrying heavy binary payloads (images, Visio diagrams, OLE binaries)
# that should be removed if no longer referenced in the extracted excerpt
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
    """Normalize, clamp, sort, and merge overlapping or adjacent block index ranges.

    Args:
        ranges: List of (start_idx, end_idx) half-open tuples representing blocks to keep.
        block_count: Total number of content blocks (<w:p> and <w:tbl>) in the document body.

    Returns:
        List of disjoint, non-overlapping (start, end) tuples sorted in ascending order.
    """
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


def _clean_content_types(content_types_bytes: bytes, pruned_zip_paths: Set[str], parser) -> bytes:
    """Remove <Override> entries in [Content_Types].xml for pruned binary parts."""
    if not pruned_zip_paths:
        return content_types_bytes
    try:
        ct_root = etree.fromstring(content_types_bytes, parser=parser)
        to_remove = []
        for child in ct_root:
            part_name = child.get("PartName", "")
            if part_name.startswith("/"):
                norm_part = part_name.lstrip("/")
            else:
                norm_part = part_name
            if norm_part in pruned_zip_paths:
                to_remove.append(child)
        if not to_remove:
            return content_types_bytes
        for child in to_remove:
            ct_root.remove(child)
        return etree.tostring(
            ct_root,
            encoding="utf-8",
            xml_declaration=True,
            standalone="yes",
        )
    except Exception:
        return content_types_bytes


def make_fragment(source: Path, target: Path, ranges: List[Tuple[int, int]]) -> None:
    """Extract selected block ranges into a self-contained DOCX fragment.

    Args:
        source: Path to the original full DOCX document.
        target: Destination path for the extracted excerpt DOCX.
        ranges: List of block index ranges (start, end) to extract.
    """
    t_start = time.perf_counter()

    # -----------------------------------------------------------------------
    # Step 1: Read source DOCX into memory
    # -----------------------------------------------------------------------
    # Reading the archive into an in-memory byte buffer once prevents repeated disk
    # seek operations and avoids triggering synchronous endpoint security (EDR/AV)
    # file-system scanning on individual ZIP part reads.
    t0 = time.perf_counter()
    with open(to_extended_path(source), 'rb') as _sf:
        source_bytes = _sf.read()
    with zipfile.ZipFile(io.BytesIO(source_bytes), "r") as zin:
        doc_xml_bytes = zin.read("word/document.xml")
        try:
            rels_xml_bytes = zin.read("word/_rels/document.xml.rels")
        except KeyError:
            rels_xml_bytes = None
    t_read = time.perf_counter() - t0

    # -----------------------------------------------------------------------
    # Step 2: Parse word/document.xml with raw C-level lxml
    # -----------------------------------------------------------------------
    # `huge_tree=True` is required because large specifications (e.g. TS 23.501)
    # easily exceed libxml2's default XML entity/node depth and size safety limits.
    t0 = time.perf_counter()
    parser = etree.XMLParser(remove_blank_text=False, resolve_entities=False, huge_tree=True)
    root = etree.fromstring(doc_xml_bytes, parser=parser)
    body = root.find(TAG_BODY)
    if body is None:
        raise RuntimeError(f"No <w:body> found in {source}")
    t_parse = time.perf_counter() - t0

    # -----------------------------------------------------------------------
    # Step 3: Read-Only DOM Slicing
    # -----------------------------------------------------------------------
    # IMPORTANT: DO NOT modify, remove, or reparent nodes inside `body`.
    # Any tree mutation while Python references exist triggers libxml2's recursive
    # `xmlReconciliateNs` across millions of descendant nodes.
    # Instead, we perform two fast linear passes over `body`:
    #   Pass A: Count the total number of content blocks (<w:p> and <w:tbl>).
    #   Pass B: Serialize ONLY the kept blocks to byte chunks.
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

    # Extract the original root XML header and opening <w:body ...> tag directly
    # from raw bytes so that all original namespace definitions (xmlns:w, xmlns:r,
    # xmlns:v, etc.) on <w:document> and <w:body> are preserved exactly.
    body_tag_pos = doc_xml_bytes.find(b"<w:body")
    if body_tag_pos != -1:
        body_open_end = doc_xml_bytes.find(b">", body_tag_pos) + 1
        xml_prefix = doc_xml_bytes[:body_open_end]
        xml_suffix = b"</w:body></w:document>"
    else:
        body_tag_pos = doc_xml_bytes.find(b"<body")
        body_open_end = doc_xml_bytes.find(b">", body_tag_pos) + 1
        xml_prefix = doc_xml_bytes[:body_open_end]
        xml_suffix = b"</body></document>"

    kept_chunks: List[bytes] = [xml_prefix]
    used_rids: Set[str] = set()

    # Pass B: Iterate through the body children.
    # - Keep selected content blocks (<w:p>, <w:tbl>).
    # - Keep ONLY <w:sectPr> among non-block body children (omitting any orphaned
    #   body-level <w:permStart>, <w:bookmarkStart>, or <w:sdt> lock markers).
    block_idx = 0
    for child in body:
        tag = child.tag
        if tag in BLOCK_TAGS:
            keep = block_idx in keep_block_indices
            block_idx += 1
        elif tag == TAG_SECTPR:
            keep = True
        else:
            keep = False

        if keep:
            # Serialize only the surviving element to UTF-8 XML bytes
            kept_chunks.append(etree.tostring(child, encoding="utf-8"))

            # Scan the kept subtree for active relationship IDs (rId) in C via lxml attribs
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

    # -----------------------------------------------------------------------
    # Step 4: Prune Unused Heavy Media/OLE Relationships
    # -----------------------------------------------------------------------
    # Scans word/_rels/document.xml.rels and unlinks any relationship pointing
    # to images, OLE embeddings, or charts that are not referenced by the surviving blocks.
    # The actual unreferenced binary files are also flagged so they can be omitted
    # during archive assembly in Step 5.
    t0 = time.perf_counter()

    # Fallback scan for rIds embedded inside unparsed VML or drawing markup
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

            # Resolve internal package path relative to the "word/" directory
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

        # Safety check: Never drop a physical media file if another active relationship
        # still links to it (e.g. shared logo or recurring diagram)
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

    # -----------------------------------------------------------------------
    # Step 5: Assemble Output DOCX in RAM & Write Unlocked File to Disk
    # -----------------------------------------------------------------------
    # 1. We construct the entire slimmed DOCX inside an in-memory buffer (`io.BytesIO`).
    #    Any unreferenced heavy media binaries (often 20MB - 100MB+) are skipped entirely.
    # 2. We use `shutil.copy2` to initialize `target` so Windows ACLs, enterprise security
    #    attributes, and permissions are preserved.
    # 3. We ensure write permissions on `target` in case the source file was marked read-only.
    # 4. We overwrite `target` in a single atomic `write_bytes()` call to avoid triggering
    #    repetitive file-filter hooks in cloud-sync drivers (OneDrive) or antivirus engines.
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
            elif filename == "[Content_Types].xml" and pruned_zip_paths:
                ct_bytes = zin.read(filename)
                zout.writestr(item, _clean_content_types(ct_bytes, pruned_zip_paths, parser))
            elif filename == "word/settings.xml":
                # Strip <w:documentProtection> and <w:writeProtection> so Word COM
                # never locks the fragment during multi-file InsertFile merges
                settings_bytes = zin.read(filename)
                settings_bytes = RE_DOC_PROTECTION.sub(b"", settings_bytes)
                zout.writestr(item, settings_bytes)
            else:
                # Stream remaining parts (styles, numbering, themes, settings) untouched
                zout.writestr(item, zin.read(filename))

    ext_target = to_extended_path(target)
    shutil.copy2(to_extended_path(source), ext_target)
    try:
        os.chmod(ext_target, stat.S_IWRITE | stat.S_IREAD)
    except OSError:
        pass

    # Remove Windows Mark-of-the-Web (Zone.Identifier alternate data stream) copied
    # by shutil.copy2 from downloaded 3GPP ZIPs, preventing Word Protected View locks
    try:
        os.remove(f"{ext_target}:Zone.Identifier")
    except OSError:
        pass

    with open(ext_target, 'wb') as _tf:
        _tf.write(out_buffer.getvalue())
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
    """CLI entry point for the isolated fragment worker process."""
    if len(sys.argv) != 4:
        print("Usage: word_excerpt_fragment_worker.py SOURCE TARGET RANGES_JSON", file=sys.stderr)
        return 2

    source = Path(sys.argv[1])
    target = Path(sys.argv[2])
    ranges = json.loads(sys.argv[3])

    make_fragment(source, target, ranges)

    # Flush all output buffers before terminating
    sys.stdout.flush()
    sys.stderr.flush()

    # Fast process exit:
    # Bypasses Python runtime garbage collection traversal over the parsed 2M-node
    # XML tree. The OS reclaims the entire virtual address space instantaneously on exit.
    os._exit(0)


if __name__ == "__main__":
    raise SystemExit(main())