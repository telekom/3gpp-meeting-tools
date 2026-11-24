"""Markdown header assembly and multi-tier presentation-to-Markdown orchestrator."""

from __future__ import annotations

import os
import sys
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, TYPE_CHECKING, Union

from core.utils.paths import to_extended_path
from modules.word_tools.core.presentation.com_extractor import _extract_via_powerpoint_com
from modules.word_tools.core.presentation.models import (
    CONVERTER_VERSION,
    TDOC_RE,
    ExtractedSlide,
    _format_markdown_table,
    _safe_log,
)
from modules.word_tools.core.presentation.openxml_and_pdf import (
    _extract_via_openxml,
    _extract_via_pdf,
)

if TYPE_CHECKING:
    from modules.meetings.core.exporters.tdoc2md import ExportedDoc

_SHIM_MODULE_NAME = "modules.word_tools.core.presentation_to_markdown"


def _resolve_facade(attr_name: str, default: Any) -> Any:
    facade = sys.modules.get(_SHIM_MODULE_NAME)
    if facade is not None and hasattr(facade, attr_name):
        return getattr(facade, attr_name)
    return default


def _build_presentation_header(
    tdoc_id: str,
    meta: Dict[str, str],
    title: str,
    original_filename: str,
    slide_count: int,
    tdoc_url: Optional[str] = None,
) -> str:
    """Builds the 3GPP document header or generic presentation header."""
    is_tdoc = bool(
        TDOC_RE.search(tdoc_id)
        or (meta and any(k in meta for k in ("tdoc", "tdoc_id", "agenda item", "work item", "spec", "source")))
    )

    doc_title = meta.get("Title") or meta.get("title") or title or tdoc_id
    doc_source = meta.get("Source") or meta.get("source") or ""
    doc_type = meta.get("Type") or meta.get("type") or "Discussion Paper (Presentation)"
    doc_ai = meta.get("Agenda Item") or meta.get("agenda item") or ""
    doc_wi = meta.get("Work Item") or meta.get("work item") or ""
    doc_status = meta.get("Status") or meta.get("status") or meta.get("tdoc status") or ""

    lines = []
    if is_tdoc:
        header_title = f"# {tdoc_id}: {doc_title}" if doc_title else f"# {tdoc_id}"
        lines.append(header_title)
        lines.append("")
        fields = [
            ("TDoc", tdoc_id),
            ("Title", doc_title),
            ("Source", doc_source),
            ("Type", doc_type),
            ("Agenda Item", doc_ai),
            ("Work Item", doc_wi),
            ("Status", doc_status),
            ("Slides", f"{slide_count} slides"),
            ("Original file", f"`{original_filename}`"),
        ]
        if tdoc_url:
            fields.append(("Authoritative 3GPP Portal Link", f"[{tdoc_id}]({tdoc_url})"))

        for k, v in fields:
            if v:
                lines.append(f"- **{k}:** {v.replace(chr(10), ' ')}")
        lines.append("")
    else:
        lines.append(f"# {doc_title or tdoc_id}")
        lines.append("")
        lines.append(f"- **Original file:** `{original_filename}`")
        lines.append(f"- **Slides:** {slide_count} slides")
        lines.append("")

    lines.append("---")
    lines.append("")
    return "\n".join(lines)


def convert_presentation_to_markdown(
    ppt_path: Union[str, Path],
    out_dir: Optional[Union[str, Path]] = None,
    tdoc_id: Optional[str] = None,
    meta: Optional[Dict[str, str]] = None,
    media_dir: Optional[Path] = None,
    media_rel: Optional[str] = None,
    cancellation_check: Optional[Callable[[], bool]] = None,
    tdoc_url: Optional[str] = None,
    include_header: bool = True,
    logger_fn: Optional[Callable[[str], None]] = None,
) -> "ExportedDoc":
    """
    Primary Entry Point: Converts a PowerPoint presentation (.pptx or .ppt)
    into a structured Markdown document with slide screenshots, bullet points,
    tables, and speaker notes.

    Returns:
        ExportedDoc: Compatible with 3GPP Tools TDoc exporter and AI triage pipelines.
    """
    from modules.meetings.core.exporters.tdoc2md import ExportedDoc, guess_tdoc_id, sanitize_filename

    path = Path(to_extended_path(ppt_path))
    if not path.is_file():
        raise FileNotFoundError(f"Presentation document not found: {path}")

    tid = tdoc_id or guess_tdoc_id(path)
    target_dir = Path(to_extended_path(out_dir)) if out_dir else path.parent
    target_dir.mkdir(parents=True, exist_ok=True)

    companion_media = Path(to_extended_path(media_dir)) if media_dir is not None else target_dir / f"{path.stem}_media"
    companion_media.mkdir(parents=True, exist_ok=True)
    rel_media_path = media_rel if media_rel is not None else f"{path.stem}_media"

    suffix = path.suffix.lower()
    slides: List[ExtractedSlide] = []
    warnings: List[str] = []

    extract_pdf_fn = _resolve_facade("_extract_via_pdf", _extract_via_pdf)
    extract_com_fn = _resolve_facade("_extract_via_powerpoint_com", _extract_via_powerpoint_com)
    extract_ox_fn = _resolve_facade("_extract_via_openxml", _extract_via_openxml)

    # -------------------------------------------------------------------------
    # Conversion Dispatch: PDF, COM (PPTX/PPT), OpenXML Fallback
    # -------------------------------------------------------------------------
    if suffix == ".pdf":
        _safe_log(f"Extracting presentation slides and visuals from PDF '{path.name}'...", logger_fn)
        slides, pdf_warnings = extract_pdf_fn(
            path,
            media_dir=companion_media,
            tdoc_id=tid,
            cancellation_check=cancellation_check,
            log_fn=logger_fn,
        )
        warnings.extend(pdf_warnings)
    else:
        com_succeeded = False
        if os.name == "nt":
            try:
                _safe_log(f"Initiating PowerPoint COM visual extraction for '{path.name}'...", logger_fn)
                slides, com_warnings = extract_com_fn(
                    path,
                    media_dir=companion_media,
                    tdoc_id=tid,
                    cancellation_check=cancellation_check,
                    log_fn=logger_fn,
                )
                warnings.extend(com_warnings)
                com_succeeded = True
                _safe_log(f"Extracted {len(slides)} slide screenshots and content via PowerPoint COM.", logger_fn)
            except Exception as com_err:
                _safe_log(f"PowerPoint COM extraction unavailable or failed ({com_err}). Trying fallback engines...", logger_fn)
                warnings.append(f"PowerPoint COM failed: {com_err}")

        if not com_succeeded:
            if suffix == ".pptx":
                _safe_log("Extracting structured content via OpenXML direct package parser...", logger_fn)
                try:
                    slides, ox_warnings = extract_ox_fn(
                        path,
                        media_dir=companion_media,
                        tdoc_id=tid,
                        cancellation_check=cancellation_check,
                        log_fn=logger_fn,
                    )
                    warnings.extend(ox_warnings)
                    _safe_log(f"Extracted {len(slides)} slides via OpenXML parser.", logger_fn)
                except Exception as ox_err:
                    raise RuntimeError(f"Failed to parse presentation '{path.name}': {ox_err}") from ox_err
            elif suffix == ".ppt":
                from modules.word_tools.core.libreoffice_converter import (
                    convert_document_libreoffice,
                    is_libreoffice_available,
                )
                if is_libreoffice_available():
                    _safe_log("Converting legacy .ppt to PDF using Headless LibreOffice...", logger_fn)
                    pdf_path = convert_document_libreoffice(path, target_format="pdf")
                    slides, pdf_warnings = extract_pdf_fn(
                        pdf_path,
                        media_dir=companion_media,
                        tdoc_id=tid,
                        cancellation_check=cancellation_check,
                        log_fn=logger_fn,
                    )
                    warnings.extend(pdf_warnings)
                else:
                    raise RuntimeError("Legacy binary .ppt requires Microsoft PowerPoint (on Windows) or LibreOffice.")

    # -------------------------------------------------------------------------
    # Assemble Markdown Document
    # -------------------------------------------------------------------------
    first_title = slides[0].title if slides else path.stem
    header_text = ""
    if include_header:
        header_text = _build_presentation_header(
            tdoc_id=tid,
            meta=meta or {},
            title=first_title,
            original_filename=path.name,
            slide_count=len(slides),
            tdoc_url=tdoc_url,
        )

    md_lines: List[str] = [header_text] if header_text else []

    exported_images_count = 0
    for s in slides:
        slide_title = s.title.strip() if s.title else f"Slide {s.index}"
        md_lines.append(f"## Slide {s.index}: {slide_title}")
        md_lines.append("")

        if s.image_filename:
            img_rel = f"{rel_media_path}/{s.image_filename}".replace("\\", "/")
            md_lines.append(f"![Slide {s.index}: {slide_title}]({img_rel})")
            md_lines.append("")
            exported_images_count += 1

        if s.body_paragraphs:
            for p in s.body_paragraphs:
                md_lines.append(p)
            md_lines.append("")

        for tbl_data in s.tables:
            tbl_md = _format_markdown_table(tbl_data)
            if tbl_md:
                md_lines.append(tbl_md)
                md_lines.append("")

        if s.speaker_notes:
            md_lines.append("> [!NOTE] Speaker Notes")
            for n_line in s.speaker_notes.splitlines():
                clean_n = n_line.strip()
                if clean_n:
                    md_lines.append(f"> {clean_n}")
            md_lines.append("")

        md_lines.append("---")
        md_lines.append("")

    full_markdown = "\n".join(md_lines)
    stamp = f"<!--- tdoc_converter: version={CONVERTER_VERSION}; type=presentation; mode=presentation --->\n"
    if not full_markdown.startswith("<!--- tdoc_converter:"):
        full_markdown = stamp + full_markdown

    clean_name = sanitize_filename(tid or path.stem)
    md_output_path = target_dir / f"{clean_name}.md"
    md_output_path.write_text(full_markdown, encoding="utf-8")

    _safe_log(
        f"Wrote Markdown presentation export to '{md_output_path.name}' ({len(slides)} slides, {exported_images_count} figures).",
        logger_fn,
    )

    doc_type = (meta or {}).get("Type") or (meta or {}).get("type") or "Discussion Paper"
    doc_ai = (meta or {}).get("Agenda Item") or (meta or {}).get("agenda item") or ""
    doc_status = (meta or {}).get("Status") or (meta or {}).get("status") or ""

    return ExportedDoc(
        tdoc_id=tid,
        name=clean_name,
        md_path=md_output_path,
        text=full_markdown,
        meta=meta or {},
        doc_type=doc_type,
        agenda_item=doc_ai,
        title=first_title or tid,
        status=doc_status,
        images=exported_images_count,
        warnings=warnings,
    )
