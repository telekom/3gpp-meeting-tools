"""Application adapter for the bundled, replaceable TDoc Markdown converter.

Keep application integration here. Only the sibling tdoc2md.py is copied
from the 3gpp-tdoc-markdown-exporter skill; office tools belong to this application.
"""
from __future__ import annotations

import argparse
import logging
import os
from pathlib import Path
import shutil
import sys
import tempfile
from typing import Callable, Dict, List, Optional, Tuple, Union

from modules.word_tools.core.word_converter import convert_doc_to_docx
from modules.word_tools.core.libreoffice_converter import (
    convert_document_libreoffice, is_libreoffice_available,
)
from modules.meetings.core import tdoc2md as _core
from modules.meetings.core.tdoc2md import (
    DEFAULT_CORPUS_PROMPT, MARKUP_LEGEND, DOC_EXTENSIONS,
    DocResult, ExportedDoc, ParaRender, Styles, Numbering,
    sanitize_filename, safe_media_name, guess_tdoc_id, base_tdoc_id,
    to_extended_path, strip_extended_path, to_tool_path,
    build_header, category_for, write_index, write_bulk_corpus,
    read_xlsx_records, load_tdoc_metadata, discover_documents, extract_docs_from_zip,
    _convert_with_powershell, _convert_with_pillow,
)

logger = logging.getLogger(__name__)


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
) -> ExportedDoc:
    """Convert DOCX using the shared parser and application image rendering."""
    return _core.convert_docx_file(
        docx_path, out_dir, tdoc_id=tdoc_id, name=name, meta=meta,
        original_name=original_name, revisions=revisions, image_dpi=image_dpi,
        convert_images=convert_images, keep_metafiles=keep_metafiles,
        media_dir=media_dir, media_rel=media_rel,
        cancellation_check=cancellation_check, metafile_converter=convert_metafiles,
    )


def convert_document_to_markdown(
    doc_path: Union[str, Path],
    out_dir: Optional[Union[str, Path]] = None,
    tdoc_id: Optional[str] = None,
    meta: Optional[Dict[str, str]] = None,
    revisions: str = "markup",
    media_dir: Optional[Path] = None,
    media_rel: str = "media",
    cancellation_check: Optional[Callable[[], bool]] = None,
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

    if suffix == ".doc":
        logger.info(f"Converting legacy .doc to .docx for {tid} using application converters...")
        tmp_dir = Path(tempfile.mkdtemp(prefix="tdoc_conv_"))
        try:
            tmp_docx = convert_doc_to_docx(path, output_path=tmp_dir / f"{path.stem}.docx", logger=logger)
            return convert_docx_file(
                tmp_docx, target_dir, tdoc_id=tid, name=tid, meta=meta,
                original_name=path.name, revisions=revisions,
                media_dir=media_dir, media_rel=media_rel,
                cancellation_check=cancellation_check,
            )
        finally:
            shutil.rmtree(to_extended_path(tmp_dir), ignore_errors=True)
    else:
        return convert_docx_file(
            path, target_dir, tdoc_id=tid, name=tid, meta=meta,
            original_name=path.name, revisions=revisions,
            media_dir=media_dir, media_rel=media_rel,
            cancellation_check=cancellation_check,
        )

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
        output_dir = base / "Export" / "LLM_Corpus"
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
    ap.add_argument("-o", "--output-dir", help="Output directory (default: <input>/Export/LLM_Corpus)")
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
