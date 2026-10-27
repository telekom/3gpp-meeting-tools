"""Application adapter for the bundled, replaceable TDoc Markdown converter.

Keep application integration here. Only the sibling tdoc2md.py is copied
from the 3gpp-tdoc-markdown-exporter skill; office tools belong to this application.
"""
from __future__ import annotations

import argparse
import datetime
import logging
import os
import re
from pathlib import Path
import shutil
import sys
import tempfile
from typing import Callable, Dict, List, Optional, Tuple, Union

from modules.word_tools.core.word_converter import convert_doc_to_docx
from modules.word_tools.core.libreoffice_converter import (
    convert_document_libreoffice, is_libreoffice_available,
)
from modules.meetings.core.tdoc_url_resolver import resolve_tdoc_download_url
from modules.meetings.core import tdoc2md as _core
from modules.meetings.core.tdoc2md import (
    DEFAULT_CORPUS_PROMPT, MARKUP_LEGEND, DOC_EXTENSIONS,
    DocResult, ExportedDoc, ParaRender, Styles, Numbering,
    sanitize_filename, safe_media_name, guess_tdoc_id, base_tdoc_id,
    to_extended_path, strip_extended_path, to_tool_path,
    build_header, category_for, write_bulk_corpus,
    read_xlsx_records, load_tdoc_metadata, discover_documents, extract_docs_from_zip,
    _convert_with_powershell, _convert_with_pillow,
)

logger = logging.getLogger(__name__)

CONVERTER_VERSION = "2.1.0"


def _convert_with_libreoffice(jobs: List[Tuple[Path, Path]]) -> None:
    """Uses application's existing profile-isolated LibreOffice converter."""
    if not is_libreoffice_available():
        return
    for src, dst in jobs:
        try:
            convert_document_libreoffice(src, target_format="png", output_path=dst, logger=logger)
        except Exception as e:
            logger.debug(f"LibreOffice conversion failed for {src.name}: {e}")


def convert_metafiles(jobs: List[Tuple[Path, Path]], dpi: int = 150) -> None:
    """Convert EMF/WMF files to PNG using PowerShell, Pillow, or headless LibreOffice."""
    for _s, d in jobs:
        d_ext = Path(to_extended_path(d))
        if d_ext.exists():
            try:
                d_ext.unlink()
            except OSError:
                pass

    def pending() -> List[Tuple[Path, Path]]:
        return [(s, d) for s, d in jobs if not (Path(to_extended_path(d)).exists() and Path(to_extended_path(d)).stat().st_size > 0)]

    if os.name == "nt":
        _convert_with_powershell(pending(), dpi)
    if pending():
        _convert_with_pillow(pending(), dpi)
    if pending():
        _convert_with_libreoffice(pending())


class DocxConverter(_core.DocxConverter):
    """Use application office rendering for direct converter callers, too."""

    def __init__(self, *args, **kwargs):
        kwargs.setdefault("metafile_converter", convert_metafiles)
        super().__init__(*args, **kwargs)


def is_plain_document_type(doc_type: Optional[str]) -> bool:
    """Check if a 3GPP document type represents a plain document that does not
    rely on specification change marks (such as LS IN, LS OUT, or liaison statements).
    """
    if not doc_type:
        return False
    s = str(doc_type).strip().lower()
    norm = re.sub(r"[\(\)\[\]_\-/\\]+", " ", s)
    tokens = set(norm.split())

    if "ls" in tokens or "liaison" in norm:
        return True
    if any(k in norm for k in ["ls in", "ls out", "incoming ls", "outgoing ls", "reply ls"]):
        return True
    return False


def resolve_revision_mode_for_type(
    doc_type: Optional[str] = None,
    default_mode: str = "markup",
) -> str:
    """Determine the appropriate revision rendering mode ('markup' or 'accept') based on 3GPP document type.

    - CR and pCR documents represent specification modifications and rely on 'markup' mode
      (with [INSERTED] / [DELETED] markers and markup legend).
    - LS IN and LS OUT documents (Liaison Statements) are self-contained communications
      that do not rely on Word change marks and are exported as plain documents ('accept' mode).
    - If default_mode is explicitly 'accept' or 'reject', it is honored.
    """
    if default_mode != "markup":
        return default_mode
    if is_plain_document_type(doc_type):
        return "accept"
    return default_mode


def convert_docx_file(
    docx_path: Path,
    out_dir: Path,
    tdoc_id: Optional[str] = None,
    name: Optional[str] = None,
    meta: Optional[Dict[str, str]] = None,
    original_name: Optional[str] = None,
    revisions: str = "markup",
    image_dpi: int = 150,
    convert_images: bool = True,
    keep_metafiles: bool = False,
    media_dir: Optional[Path] = None,
    media_rel: str = "media",
    cancellation_check: Optional[Callable[[], bool]] = None,
    tdoc_url: Optional[str] = None,
    include_header: bool = True,
) -> ExportedDoc:
    """Convert DOCX using the shared parser and application image rendering."""
    doc_type = None
    if meta:
        doc_type = meta.get("Type") or meta.get("type") or meta.get("doc_type")
    effective_revisions = resolve_revision_mode_for_type(doc_type, default_mode=revisions)

    res = _core.convert_docx_file(
        docx_path, out_dir, tdoc_id=tdoc_id, name=name, meta=meta,
        original_name=original_name, revisions=effective_revisions, image_dpi=image_dpi,
        convert_images=convert_images, keep_metafiles=keep_metafiles,
        media_dir=media_dir, media_rel=media_rel,
        cancellation_check=cancellation_check, metafile_converter=convert_metafiles,
        include_header=include_header,
    )

    # Inject 3GPP server download URL into the metadata header if resolvable
    effective_id = tdoc_id or name or guess_tdoc_id(docx_path)
    download_url = tdoc_url or resolve_tdoc_download_url(effective_id, meta=meta)
    if download_url and "- **3GPP Server URL:**" not in res.text:
        url_bullet = f"- **3GPP Server URL:** [{download_url}]({download_url})\n"
        if "- **Original file:**" in res.text:
            res.text = res.text.replace("- **Original file:**", f"{url_bullet}- **Original file:**", 1)
        elif "- **Exporter:**" in res.text:
            res.text = res.text.replace("- **Exporter:**", f"{url_bullet}- **Exporter:**", 1)
        elif "\n---\n" in res.text:
            res.text = res.text.replace("\n---\n", f"\n{url_bullet}\n---\n", 1)

    # Stamp machine-readable metadata header for multi-tier stale detection
    stamp = f"<!--- tdoc_converter: version={CONVERTER_VERSION}; mode={effective_revisions} --->\n"
    if not res.text.startswith("<!--- tdoc_converter:"):
        res.text = stamp + res.text
    try:
        with open(to_extended_path(res.md_path), "w", encoding="utf-8") as f:
            f.write(res.text)
    except Exception as e:
        logger.debug(f"Could not prepend version stamp to {res.md_path}: {e}")
    return res


def convert_document_to_markdown(
    doc_path: Union[str, Path],
    out_dir: Optional[Union[str, Path]] = None,
    tdoc_id: Optional[str] = None,
    meta: Optional[Dict[str, str]] = None,
    revisions: str = "markup",
    media_dir: Optional[Path] = None,
    media_rel: str = "media",
    cancellation_check: Optional[Callable[[], bool]] = None,
    tdoc_url: Optional[str] = None,
    convert_images: bool = True,
    include_header: bool = True,
) -> ExportedDoc:
    """
    Primary Entry Point: Accepts .docx, .docm, or legacy binary .doc.
    For .doc, automatically applies sensitivity labeling and invokes Word COM / LibreOffice
    conversion before extracting Markdown.
    """
    path = Path(to_extended_path(doc_path))
    if not path.is_file():
        raise FileNotFoundError(f"Document not found: {path}")

    target_dir = Path(to_extended_path(out_dir)) if out_dir else path.parent
    target_dir.mkdir(parents=True, exist_ok=True)
    tid = tdoc_id or guess_tdoc_id(path)
    suffix = path.suffix.lower()

    doc_type = None
    if meta:
        doc_type = meta.get("Type") or meta.get("type") or meta.get("doc_type")
    effective_revisions = resolve_revision_mode_for_type(doc_type, default_mode=revisions)

    resolved_tdoc_url = tdoc_url or resolve_tdoc_download_url(tid, meta=meta)

    if suffix == ".doc":
        logger.info(f"Converting legacy .doc to .docx for {tid} using application converters...")
        tmp_dir = Path(tempfile.mkdtemp(prefix="tdoc_conv_"))
        try:
            tmp_docx = convert_doc_to_docx(path, output_path=tmp_dir / f"{path.stem}.docx", logger=logger)
            return convert_docx_file(
                tmp_docx, target_dir, tdoc_id=tid, name=tid, meta=meta,
                original_name=path.name, revisions=effective_revisions,
                convert_images=convert_images,
                media_dir=media_dir, media_rel=media_rel,
                cancellation_check=cancellation_check,
                tdoc_url=resolved_tdoc_url,
                include_header=include_header,
            )
        finally:
            shutil.rmtree(to_extended_path(tmp_dir), ignore_errors=True)
    else:
        return convert_docx_file(
            path, target_dir, tdoc_id=tid, name=tid, meta=meta,
            original_name=path.name, revisions=effective_revisions,
            convert_images=convert_images,
            media_dir=media_dir, media_rel=media_rel,
            cancellation_check=cancellation_check,
            tdoc_url=resolved_tdoc_url,
            include_header=include_header,
        )


def is_conversion_stale(
    source_path: Optional[Union[str, Path]],
    md_path: Union[str, Path],
    expected_version: str = CONVERTER_VERSION,
    expected_mode: str = "markup",
) -> bool:
    """
    Determine whether a cached Markdown dossier needs re-conversion.
    Checks:
      1. File existence (is it present and non-empty?)
      2. Source Word document mtime (was the .docx/.doc updated on disk?)
      3. Converter engine version (has CONVERTER_VERSION been bumped?)
      4. Revision mode match (markup vs accept/reject)
    """
    p_md = Path(to_extended_path(md_path))
    if not p_md.is_file() or p_md.stat().st_size == 0:
        return True

    # 1. Source Word file updated?
    if source_path:
        p_src = Path(to_extended_path(source_path))
        if p_src.is_file():
            try:
                if p_src.stat().st_mtime > p_md.stat().st_mtime:
                    return True
            except OSError:
                pass

    # 2. Check header metadata for version and mode match
    try:
        with open(p_md, "r", encoding="utf-8") as f:
            header_sample = f.read(512)

        # Check explicit machine-readable comment
        m_ver = re.search(r"tdoc_converter:\s*version=([0-9.]+)", header_sample)
        if m_ver:
            if m_ver.group(1) != expected_version:
                return True
        else:
            # Fallback to standard Exporter line: "- **Exporter:** ... vX.Y.Z"
            m_exp = re.search(r"\*\*Exporter:\*\*.*?v([0-9.]+)", header_sample)
            if not m_exp or m_exp.group(1) != expected_version:
                return True

        m_mode = re.search(r"mode=([a-z]+)", header_sample)
        if m_mode:
            if m_mode.group(1) != expected_mode:
                return True
        else:
            m_rev = re.search(r"revisions:\s*([a-z]+)", header_sample)
            if m_rev and m_rev.group(1) != expected_mode:
                return True
    except Exception:
        return True

    return False


def write_index(out_dir: Path, docs: List[ExportedDoc]) -> Path:
    """Generate index.md manifest supporting relative paths to subfolder dossiers and 3GPP server download links."""
    out_dir = Path(to_extended_path(out_dir))
    out_dir.mkdir(parents=True, exist_ok=True)
    clean_out = Path(strip_extended_path(os.path.abspath(str(out_dir))))
    lines = [
        "# TDoc Markdown Export",
        "",
        f"Generated by tdoc_markdown_converter v{CONVERTER_VERSION}. {len(docs)} document(s).",
        "",
        "| TDoc | Title | Type | Agenda Item | Status | Images | 3GPP Server |",
        "| --- | --- | --- | --- | --- | --- | --- |",
    ]
    for d in sorted(docs, key=lambda x: x.name):
        title = (d.title or "").replace("|", "\\|")
        try:
            clean_md = Path(strip_extended_path(os.path.abspath(str(d.md_path))))
            rel_link = clean_md.relative_to(clean_out).as_posix()
        except Exception:
            rel_link = d.md_path.name

        tdoc_url = resolve_tdoc_download_url(d.name, meta=d.meta)
        url_link = f"[📥 Download ZIP]({tdoc_url})" if tdoc_url else "—"

        lines.append(
            f"| [{d.name}]({rel_link}) | {title} | {d.doc_type} | {d.agenda_item} | {d.status} | {d.images} | {url_link} |"
        )
    path = out_dir / "index.md"
    with open(to_extended_path(path), "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")
    return path


def _convert_legacy_document(src: Path, tmpdir: Path, config: dict) -> Optional[Path]:
    """Adapt the application's Word converter to the shared batch interface."""
    return convert_doc_to_docx(src, output_path=tmpdir / (src.stem + ".docx"), logger=logger)


def export_corpus(
    input_path: Path,
    output_dir: Optional[Path] = None,
    excel: Optional[Path] = None,
    bulk: bool = True,
    max_chars: int = 200000,
    revisions: str = "markup",
    image_dpi: int = 150,
    convert_images: bool = True,
    keep_metafiles: bool = False,
    prompt: str = DEFAULT_CORPUS_PROMPT,
    progress_callback: Optional[Callable[[str], None]] = None,
    cancellation_check: Optional[Callable[[], bool]] = None,
) -> Tuple[List[ExportedDoc], List[str]]:
    """Export with application defaults, callbacks, and office backends."""
    if output_dir is None:
        source = Path(input_path)
        base = source if source.is_dir() else source.parent
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H%M%S")
        output_dir = base / "Export" / f"Markdown_{timestamp}"
    return _core.run(
        input_path, output_dir=output_dir, excel=excel, bulk=bulk,
        max_chars=max_chars, revisions=revisions, image_dpi=image_dpi,
        convert_images=convert_images, keep_metafiles=keep_metafiles, prompt=prompt,
        progress_callback=progress_callback, cancellation_check=cancellation_check,
        document_converter=_convert_legacy_document, metafile_converter=convert_metafiles,
        word_config={},
    )


def main(argv: Optional[List[str]] = None) -> int:
    ap = argparse.ArgumentParser(
        description="Convert 3GPP TDocs (.docx/.doc/.zip) to Markdown with track changes, tables, and images."
    )
    ap.add_argument("input", help="A .docx/.doc/.zip file, a TDoc folder, or a meeting folder")
    ap.add_argument("-o", "--output-dir", help="Output directory (default: <input>/Export/Markdown_YYYY-MM-DD_HHMMSS)")
    ap.add_argument("--excel", help="3GPP TDoc list (.xlsx) for metadata")
    ap.add_argument("--bulk", action="store_true", help="Generate per-Agenda Item corpus mega-files")
    ap.add_argument("--max-chars", type=int, default=200000, help="Max characters per bulk corpus file")
    ap.add_argument("--revisions", choices=["markup", "accept", "reject"], default="markup",
                    help="markup: show [INSERTED]/[DELETED]; accept: final text; reject: original text")
    ap.add_argument("--image-dpi", type=int, default=150, help="Resolution for EMF/WMF to PNG conversion")
    ap.add_argument("--no-image-convert", action="store_true", help="Keep EMF/WMF/TIFF images as extracted")
    ap.add_argument("--keep-metafiles", action="store_true", help="Keep original EMF/WMF next to converted PNGs")
    ap.add_argument("-v", "--verbose", action="store_true")
    args = ap.parse_args(argv)

    logging.basicConfig(level=logging.DEBUG if args.verbose else logging.INFO, format="%(levelname)s: %(message)s")

    try:
        exported, failures = export_corpus(
            Path(args.input),
            output_dir=Path(args.output_dir) if args.output_dir else None,
            excel=Path(args.excel) if args.excel else None,
            bulk=args.bulk,
            max_chars=args.max_chars,
            revisions=args.revisions,
            image_dpi=args.image_dpi,
            convert_images=not args.no_image_convert,
            keep_metafiles=args.keep_metafiles,
        )
    except FileNotFoundError as e:
        logger.error(f"Input not found: {e}")
        return 2

    out = exported[0].md_path.parent if exported else None
    disp = strip_extended_path(out) if out else ""
    print(f"Converted {len(exported)} document(s) -> {disp} ({len(failures)} failed)")
    for f in failures:
        print(f"  FAILED: {f}")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
