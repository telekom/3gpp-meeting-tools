"""Tier 1 presentation extraction engine using Windows PowerPoint COM automation."""

from __future__ import annotations

import logging
import os
from pathlib import Path
from typing import Callable, List, Optional, Tuple

from core.utils.paths import strip_extended_path
from modules.word_tools.core.presentation.models import (
    ExtractedSlide,
    _clean_tdoc_prefix,
    _safe_log,
)

logger = logging.getLogger(__name__)


def _extract_via_powerpoint_com(
    ppt_path: Path,
    media_dir: Path,
    tdoc_id: Optional[str] = None,
    image_width: int = 1920,
    image_height: int = 1080,
    cancellation_check: Optional[Callable[[], bool]] = None,
    log_fn: Optional[Callable[[str], None]] = None,
) -> Tuple[List[ExtractedSlide], List[str]]:
    """
    Automates Microsoft PowerPoint COM to render high-res slide images
    and extract slide titles, body text, tables, and speaker notes.
    """
    slides_data: List[ExtractedSlide] = []
    warnings: List[str] = []

    if os.name != "nt":
        raise RuntimeError("PowerPoint COM automation is only available on Windows.")

    import pythoncom
    import win32com.client

    # Ensure destination media folder exists
    media_dir = Path(strip_extended_path(str(media_dir)))
    media_dir.mkdir(parents=True, exist_ok=True)

    # Clean up any leftover corrupted image files from previous regex bug
    try:
        for bad_f in media_dir.glob("__-________*"):
            try:
                bad_f.unlink()
            except Exception:
                pass
    except Exception:
        pass

    from core.utils.office_printer_guard import temporary_virtual_default_printer

    pythoncom.CoInitialize()
    ppt_app = None
    pres = None
    launched_own_instance = False

    try:
        with temporary_virtual_default_printer(log=logger):
            _safe_log("Connecting to Microsoft PowerPoint COM engine...", log_fn)
            try:
                ppt_app = win32com.client.DispatchEx("PowerPoint.Application")
                launched_own_instance = True
            except Exception as com_launch_err:
                try:
                    ppt_app = win32com.client.GetActiveObject("PowerPoint.Application")
                    launched_own_instance = False
                except Exception:
                    raise RuntimeError(f"Could not initialize PowerPoint COM: {com_launch_err}") from com_launch_err

            try:
                ppt_app.DisplayAlerts = 1  # ppAlertsNone
            except Exception:
                pass
            try:
                ppt_app.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
            except Exception:
                pass

            resolved_path = strip_extended_path(str(ppt_path.resolve()))
            _safe_log(f"Opening presentation '{ppt_path.name}' in PowerPoint...", log_fn)

            pres = ppt_app.Presentations.Open(
                resolved_path,
                ReadOnly=True,
                Untitled=False,
                WithWindow=False,
            )

            slide_count = pres.Slides.Count
            _safe_log(f"Exporting {slide_count} slides and extracting structured content...", log_fn)

            for i in range(1, slide_count + 1):
                if cancellation_check and cancellation_check():
                    break

                slide = pres.Slides(i)
                slide_item = ExtractedSlide(index=i)

                # 1. Slide Screenshot
                clean_prefix = _clean_tdoc_prefix(tdoc_id)
                img_name = f"{clean_prefix}slide_{i:02d}.png"
                img_dest = media_dir / img_name
                img_dest_str = strip_extended_path(str(img_dest.resolve()))
                try:
                    slide.Export(img_dest_str, "PNG", image_width, image_height)
                    slide_item.image_filename = img_name
                except Exception:
                    try:
                        slide.Export(img_dest_str, "PNG")
                        slide_item.image_filename = img_name
                    except Exception as retry_err:
                        warnings.append(f"Slide {i}: Screenshot export failed ({retry_err})")

                # 2. Slide Title
                title_text = ""
                try:
                    if slide.Shapes.HasTitle:
                        title_text = slide.Shapes.Title.TextFrame.TextRange.Text.strip()
                except Exception:
                    pass

                # 3. Slide Shapes (Text & Tables)
                text_candidates: List[Tuple[float, str]] = []  # (top_pos, text)
                try:
                    for shape in slide.Shapes:
                        if getattr(shape, "HasTable", False):
                            try:
                                tbl = shape.Table
                                rows_data: List[List[str]] = []
                                for r in range(1, tbl.Rows.Count + 1):
                                    row_cells: List[str] = []
                                    for c in range(1, tbl.Columns.Count + 1):
                                        cell_txt = tbl.Cell(r, c).Shape.TextFrame.TextRange.Text.strip()
                                        row_cells.append(cell_txt)
                                    rows_data.append(row_cells)
                                if rows_data:
                                    slide_item.tables.append(rows_data)
                            except Exception as tbl_err:
                                warnings.append(f"Slide {i}: Table extraction error ({tbl_err})")
                            continue

                        if getattr(shape, "HasTextFrame", False) and shape.TextFrame.HasText:
                            top_pos = getattr(shape, "Top", 9999.0)
                            raw_text = shape.TextFrame.TextRange.Text.strip()
                            if not raw_text:
                                continue

                            if title_text and raw_text == title_text:
                                continue

                            text_candidates.append((top_pos, raw_text))

                            try:
                                paragraphs = shape.TextFrame.TextRange.Paragraphs()
                                for p in paragraphs:
                                    p_txt = p.Text.strip()
                                    if not p_txt:
                                        continue
                                    if title_text and p_txt == title_text:
                                        continue
                                    indent_lvl = getattr(p, "IndentLevel", 1)
                                    indent_spaces = "  " * max(0, indent_lvl - 1)
                                    slide_item.body_paragraphs.append(f"{indent_spaces}- {p_txt}")
                            except Exception:
                                for line in raw_text.splitlines():
                                    l_clean = line.strip()
                                    if l_clean:
                                        slide_item.body_paragraphs.append(f"- {l_clean}")
                except Exception as shape_err:
                    warnings.append(f"Slide {i}: Shapes inspection warning ({shape_err})")

                # Fallback heuristic for Title if HasTitle was False
                if not title_text and text_candidates:
                    text_candidates.sort(key=lambda x: x[0])
                    first_top, first_text = text_candidates[0]
                    if first_top < 150.0 and len(first_text.splitlines()) <= 2:
                        title_text = first_text.replace("\n", " ")
                        cleaned_body = []
                        for b in slide_item.body_paragraphs:
                            if b.lstrip("- ").strip() != first_text:
                                cleaned_body.append(b)
                        slide_item.body_paragraphs = cleaned_body

                slide_item.title = title_text

                # 4. Speaker Notes
                try:
                    if slide.HasNotesPage:
                        notes_page = slide.NotesPage
                        for n_shape in notes_page.Shapes:
                            if getattr(n_shape, "HasTextFrame", False) and n_shape.TextFrame.HasText:
                                is_body = False
                                try:
                                    if getattr(n_shape, "Type", 0) == 14:  # msoPlaceholder
                                        if n_shape.PlaceholderFormat.Type == 2:  # ppPlaceholderBody
                                            is_body = True
                                except Exception:
                                    pass

                                txt = n_shape.TextFrame.TextRange.Text.strip()
                                if is_body and txt:
                                    slide_item.speaker_notes = txt
                                    break
                                elif not slide_item.speaker_notes and txt and not txt.isdigit():
                                    slide_item.speaker_notes = txt
                except Exception as notes_err:
                    warnings.append(f"Slide {i}: Speaker notes extraction skipped ({notes_err})")

                slides_data.append(slide_item)

    finally:
        if pres is not None:
            try:
                pres.Close()
            except Exception:
                pass
        if ppt_app is not None and launched_own_instance:
            try:
                ppt_app.Quit()
            except Exception:
                pass
        pres = None
        ppt_app = None
        try:
            import gc
            gc.collect()
        except Exception:
            pass
        try:
            pythoncom.PumpWaitingMessages()
        except Exception:
            pass
        pythoncom.CoUninitialize()

    return slides_data, warnings
