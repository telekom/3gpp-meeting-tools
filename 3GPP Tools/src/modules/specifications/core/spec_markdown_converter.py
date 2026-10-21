"""
3GPP Specification Markdown Converter
Converts single or multi-part 3GPP Specification documents (.docx / .doc)
into unified or clause-split Markdown documents with version-isolated media.
"""
from __future__ import annotations

import logging
import os
import re
import shutil
import zipfile
from pathlib import Path
from typing import Callable, Dict, List, Optional, Tuple, Union

from core.utils.paths import to_extended_path
from modules.meetings.core.tdoc_markdown_converter import (
    convert_document_to_markdown,
    convert_docx_file,
    convert_metafiles,
    sanitize_filename,
)

logger = logging.getLogger(__name__)

# Pattern to identify and order 3GPP multi-part document chunks (e.g. _s00_s04, _s05_s08, or _1_)
RE_PART_INDEX = re.compile(r"_s(\d+)_s(\d+)|_(\d+)_", re.IGNORECASE)

# Pattern to identify top-level clause boundaries in generated Markdown (# 1 Scope, # 4 General, # Annex A ...)
RE_TOPLEVEL_CLAUSE = re.compile(
    r"^#\s+(([1-9]\d*)\.?\s+([^\n\r]+)|(Annex\s+[A-Z](?:\s*\([^)]*\))?[:\s]*([^\n\r]*)))",
    re.MULTILINE | re.IGNORECASE,
)


def extract_spec_archive(zip_path: Path, dest_dir: Path) -> List[Path]:
    """Extracts Word documents (.docx/.doc) from a 3GPP specification ZIP archive."""
    extracted: List[Path] = []
    ext_zip_path = to_extended_path(zip_path)
    if not os.path.exists(ext_zip_path):
        return extracted

    dest_dir.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(ext_zip_path, "r") as z:
        for member in z.namelist():
            if "__MACOSX" in member or member.startswith("._"):
                continue
            if member.lower().endswith((".docx", ".doc")):
                target = dest_dir / Path(member).name
                ext_target = to_extended_path(target)
                if not os.path.exists(ext_target):
                    with open(ext_target, "wb") as f:
                        f.write(z.read(member))
                if target not in extracted:
                    extracted.append(target)
    return extracted


def get_part_sort_key(path: Path) -> int:
    """Extracts numeric sort key from multi-part document name (e.g. _s05_s08 -> 5)."""
    m = RE_PART_INDEX.search(path.stem)
    if m:
        if m.group(1):
            return int(m.group(1))
        if m.group(3):
            return int(m.group(3))
    return 0


def discover_spec_parts(spec_dir: Path, stem: str) -> List[Path]:
    """
    Finds and chronologically orders all Word parts for a specification version.
    If only a ZIP exists on disk, extracts the documents first.
    """
    spec_dir = Path(to_extended_path(spec_dir))
    candidates = list(spec_dir.glob(f"{stem}*.doc*"))

    # If no word docs present, attempt extraction from ZIP
    if not candidates:
        zip_candidates = list(spec_dir.glob(f"{stem}*.zip"))
        if zip_candidates:
            candidates = extract_spec_archive(zip_candidates[0], spec_dir)

    if not candidates:
        return []

    # Filter out temporary word lock files (~$...)
    valid_docs = [p for p in candidates if not p.name.startswith("~$") and p.suffix.lower() in (".docx", ".doc")]
    if len(valid_docs) > 1:
        # Multi-part specification: sort chronologically by starting clause index
        valid_docs.sort(key=get_part_sort_key)

    return valid_docs


def build_spec_metadata_header(meta: Optional[Dict[str, str]] = None) -> str:
    """Builds a structured 3GPP Specification header table in Markdown."""
    m = meta or {}
    spec = m.get("spec", m.get("spec_num", "Specification"))
    title = m.get("title", "")
    ver = m.get("version", "")
    rel = m.get("release", "")
    wg = m.get("wg", "")
    date = m.get("upload_date", "")

    title_display = f"3GPP TS {spec}: {title}" if title else f"3GPP TS {spec}"
    lines = [
        f"# {title_display}",
        "",
        "| Specification | Release | Version | Working Group | Date |",
        "| :--- | :--- | :--- | :--- | :--- |",
        f"| TS {spec} | {rel or '-'} | v{ver or '-'} | {wg or '-'} | {date or '-'} |",
        "",
        "---",
        "",
    ]
    return "\n".join(lines)


def split_markdown_by_clauses(
    full_markdown: str,
    out_dir: Path,
    meta_header: str,
    spec_title: str,
) -> Path:
    """
    Splits continuous Markdown specification text into individual top-level clause files
    under out_dir and generates an index.md table of contents linking all clauses.
    """
    out_dir = Path(to_extended_path(out_dir))
    out_dir.mkdir(parents=True, exist_ok=True)

    matches = list(RE_TOPLEVEL_CLAUSE.finditer(full_markdown))
    index_entries: List[Tuple[str, str, str]] = []  # (clause_num, title, filename)

    if not matches:
        # No recognized top-level clauses found; save as single comprehensive document
        main_file = out_dir / "specification.md"
        main_file.write_text(meta_header + full_markdown, encoding="utf-8")
        index_entries.append(("-", "Complete Specification", "specification.md"))
    else:
        # 1. Front Matter (everything before Clause 1)
        first_clause_pos = matches[0].start()
        front_matter = full_markdown[:first_clause_pos].strip()
        fm_path = out_dir / "00_Front_Matter.md"
        fm_content = (meta_header + front_matter).strip() + "\n"
        fm_path.write_text(fm_content, encoding="utf-8")
        index_entries.append(("-", "Front Matter & Table of Contents", "00_Front_Matter.md"))

        # 2. Individual top-level clauses
        for i, match in enumerate(matches):
            start_pos = match.start()
            end_pos = matches[i + 1].start() if i + 1 < len(matches) else len(full_markdown)
            clause_text = full_markdown[start_pos:end_pos].strip() + "\n"

            # Parse heading info
            matched_line = match.group(0).lstrip("#").strip()
            c_num = match.group(2) or (match.group(4).split(":")[0].strip() if match.group(4) else "")
            c_title = match.group(3) or match.group(5) or matched_line
            c_title = c_title.strip()

            # Sanitize filename: e.g. "04_Architecture_model.md" or "Annex_A.md"
            if c_num.isdigit():
                prefix = f"{int(c_num):02d}"
                clean_title = sanitize_filename(c_title.replace(" ", "_"))[:40].strip("_")
                filename = f"{prefix}_{clean_title}.md" if clean_title else f"{prefix}_Clause.md"
            else:
                clean_name = sanitize_filename(matched_line.replace(" ", "_"))[:50].strip("_")
                filename = f"{clean_name}.md" if clean_name else f"Clause_{i+1:02d}.md"

            clause_path = out_dir / filename
            clause_path.write_text(clause_text, encoding="utf-8")
            index_entries.append((c_num, c_title, filename))

    # 3. Generate index.md table of contents
    idx_lines = [
        f"# {spec_title}",
        "",
        "## Table of Contents",
        "",
        "| Clause | Title | File |",
        "| :---: | :--- | :--- |",
    ]
    for c_num, title, fname in index_entries:
        escaped_title = title.replace("|", "\\|")
        idx_lines.append(f"| {c_num} | {escaped_title} | [{fname}]({fname}) |")

    idx_path = out_dir / "index.md"
    idx_path.write_text("\n".join(idx_lines) + "\n", encoding="utf-8")
    return idx_path


def convert_spec_to_markdown(
    spec_dir: Path,
    stem: str,
    mode: str = "single",
    meta: Optional[Dict[str, str]] = None,
    progress_callback: Optional[Callable[[str], None]] = None,
    cancellation_check: Optional[Callable[[], bool]] = None,
) -> Path:
    """
    Main Specification Conversion Controller:
    - Automatically discovers single or multi-part specification documents.
    - Isolates media per version to avoid collisions across releases.
    - Generates Mode 1 (single unified document) or Mode 2 (clause-split directory).

    Returns:
        Path to generated .md file (Mode 1) or index.md (Mode 2).
    """
    spec_dir = Path(to_extended_path(spec_dir))
    docs = discover_spec_parts(spec_dir, stem)
    if not docs:
        raise FileNotFoundError(f"No Word documents (.docx/.doc) or ZIP archives found for {stem} in {spec_dir}")

    total_parts = len(docs)
    logger.info(f"[Spec Markdown] Converting {stem} ({total_parts} part(s)) in mode='{mode}'")

    meta_header = build_spec_metadata_header(meta)
    spec_num = (meta or {}).get("spec", (meta or {}).get("spec_num", stem))
    spec_title = f"3GPP TS {spec_num}" + (f": {meta.get('title')}" if meta and meta.get("title") else "")

    # Configure output directories and version-isolated media paths
    if mode == "split":
        export_container = spec_dir / f"{stem}_clauses"
        media_dir = export_container / "media"
        media_rel = "media"
    else:
        export_container = spec_dir
        media_dir = spec_dir / f"{stem}_media"
        media_rel = f"{stem}_media"

    os.makedirs(to_extended_path(media_dir), exist_ok=True)

    part_markdowns: List[str] = []

    for part_idx, part_doc in enumerate(docs):
        if cancellation_check and cancellation_check():
            raise InterruptedError("Conversion cancelled by user.")

        if progress_callback:
            progress_callback(f"Extracting {part_doc.name} ({part_idx + 1}/{total_parts})...")

        # Give each part a unique prefix for images (e.g. p0_image1.png, p1_image1.png)
        part_prefix = f"p{part_idx}_{stem}" if total_parts > 1 else stem

        exported = convert_document_to_markdown(
            part_doc,
            out_dir=export_container,
            tdoc_id=part_prefix,
            meta={},
            revisions="accept",
            media_dir=media_dir,
            media_rel=media_rel,
            cancellation_check=cancellation_check,
        )

        md_text = exported.text

        # Strip any generic auto-generated tdoc header from tdoc2md
        header_end = md_text.find("\n---\n")
        if header_end != -1 and md_text.startswith("#"):
            body_text = md_text[header_end + 5:].lstrip()
        else:
            body_text = md_text

        # For parts after Part 0, skip front matter / cover page until first heading
        if part_idx > 0:
            first_heading_match = re.search(r"^#\s+", body_text, re.MULTILINE)
            if first_heading_match:
                body_text = body_text[first_heading_match.start():]

        part_markdowns.append(body_text.strip())

    if cancellation_check and cancellation_check():
        raise InterruptedError("Conversion cancelled by user.")

    # Assemble all parts together
    stitched_body = "\n\n".join(part_markdowns).strip()

    if mode == "split":
        if progress_callback:
            progress_callback("Partitioning clauses and writing index manifest...")
        result_path = split_markdown_by_clauses(
            stitched_body,
            export_container,
            meta_header,
            spec_title,
        )
        logger.info(f"[Spec Markdown] Successfully generated clause dossier at: {result_path}")
        return result_path
    else:
        if progress_callback:
            progress_callback("Writing unified Markdown document...")
        target_file = spec_dir / f"{stem}.md"
        full_document = (meta_header + stitched_body).strip() + "\n"
        with open(to_extended_path(target_file), "w", encoding="utf-8") as f:
            f.write(full_document)
        logger.info(f"[Spec Markdown] Successfully generated unified document: {target_file}")
        return target_file
