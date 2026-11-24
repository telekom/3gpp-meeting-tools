"""Win32 COM Visio SVG import, shape flattening, page fit, and source embedding pipeline."""

import gc
import logging
from pathlib import Path
from typing import Callable, Optional

import win32com.client

from modules.puml2visio.config.paths import PLANTUML_WATERMARK
from modules.puml2visio.utils.utils import strip_watermark


def flatten_svg_groups(page, log_fn: Callable[[str, int], None]) -> int:
    """Recursively ungroup imported SVG group shapes (`visTypeGroup == 2`)."""
    ungroup_count = 0
    peeling = True

    while peeling:
        peeling = False
        for i in range(page.Shapes.Count, 0, -1):
            s = page.Shapes(i)
            try:
                if s.Type == 2:
                    s.Ungroup()
                    ungroup_count += 1
                    peeling = True
            except Exception as e:
                log_fn(f"   ⚠️ Could not ungroup shape {i}: {e}", logging.WARNING)
    return ungroup_count


def remove_background_shapes(page, orig_w: float, orig_h: float) -> int:
    """Remove full-canvas borderless background rectangles emitted by PlantUML."""
    deleted_backgrounds = 0
    for i in range(page.Shapes.Count, 0, -1):
        s = page.Shapes(i)
        try:
            w = s.CellsU("Width").ResultIU
            h = s.CellsU("Height").ResultIU
            if w >= orig_w * 0.75 and h >= orig_h * 0.75:
                if len(s.Characters.Text.strip()) == 0:
                    if s.CellsU("LinePattern").ResultIU == 0:
                        s.Delete()
                        deleted_backgrounds += 1
        except Exception:
            # Some imported SVG shapes do not expose all ShapeSheet cells.
            pass
    return deleted_backgrounds


def normalize_text_shapes(shapes) -> tuple[int, int]:
    """Zero text margins and convert borderless labels to `TEXTWIDTH`/`TEXTHEIGHT` formulas."""
    text_shape_count = 0
    text_shape_errors = 0

    def _walk(shape_collection):
        nonlocal text_shape_count
        nonlocal text_shape_errors

        for i in range(1, shape_collection.Count + 1):
            s = shape_collection(i)
            try:
                text = s.Characters.Text
                if len(text.strip()) > 0:
                    text_shape_count += 1
                    for margin_cell in ("TopMargin", "BottomMargin", "LeftMargin", "RightMargin"):
                        s.CellsU(margin_cell).FormulaU = "0 pt"

                    line_pattern = s.CellsU("LinePattern").ResultIU
                    fill_pattern = s.CellsU("FillPattern").ResultIU
                    if line_pattern == 0 and fill_pattern == 0:
                        pin_x = s.CellsU("PinX").ResultIU
                        pin_y = s.CellsU("PinY").ResultIU
                        loc_pin_x = s.CellsU("LocPinX").ResultIU
                        loc_pin_y = s.CellsU("LocPinY").ResultIU
                        h = s.CellsU("Height").ResultIU

                        left = pin_x - loc_pin_x
                        top = pin_y + (h - loc_pin_y)

                        s.CellsU("LocPinX").FormulaU = "0 in"
                        s.CellsU("LocPinY").FormulaU = "Height"
                        s.CellsU("PinX").FormulaU = f"{left} in"
                        s.CellsU("PinY").FormulaU = f"{top} in"
                        s.CellsU("Width").FormulaU = "TEXTWIDTH(TheText)"
                        s.CellsU("Height").FormulaU = "TEXTHEIGHT(TheText, Width)"
            except Exception:
                text_shape_errors += 1

            try:
                if s.Type == 2:
                    _walk(s.Shapes)
            except Exception:
                pass

    _walk(shapes)
    return text_shape_count, text_shape_errors


def resize_page_to_contents(page, log_fn: Callable[[str, int], None]):
    """Set minimal 0.05 in margins and resize the diagram page to fit its shapes."""
    page_sheet = page.PageSheet
    for cell in ("PageLeftMargin", "PageRightMargin", "PageTopMargin", "PageBottomMargin"):
        page_sheet.CellsU(cell).FormulaU = "0.05 in"

    try:
        page.ResizeToFitContents()
        log_fn("   🔎 Page resized successfully.", logging.INFO)
    except Exception as e:
        log_fn(f"   ⚠️ ResizeToFitContents failed: {e}", logging.WARNING)
    return page_sheet


def append_plantuml_source_page(doc, final_source_code: str):
    """Create the embedded `PlantUML Source` page containing watermarked PUML source."""
    src_page = doc.Pages.Add()
    src_page.PageSheet.CellsU("PageWidth").FormulaU = "8.27 in"
    src_page.PageSheet.CellsU("PageHeight").FormulaU = "11.69 in"
    src_page.Name = "PlantUML Source"

    text_box = src_page.DrawRectangle(0.5, 0.5, 7.77, 11.19)
    text_box.CellsU("LinePattern").FormulaU = "0"
    text_box.CellsU("FillPattern").FormulaU = "0"
    text_box.CellsU("Para.HorzAlign").FormulaU = "0"
    text_box.CellsU("VerticalAlign").FormulaU = "0"
    text_box.Characters.Text = final_source_code
    return src_page, text_box


def convert_svg_to_vsdx(
    puml_path: Path,
    svg_path: Path,
    log_fn: Optional[Callable[[str, int], None]] = None,
) -> Path:
    """Run Stages 1–10 Visio COM automation to convert `svg_path` + `puml_path` into `.vsdx`."""
    emit = log_fn or (lambda message, level=logging.INFO: logging.getLogger(__name__).log(level, message))
    vsdx_path = svg_path.with_suffix(".vsdx")
    if vsdx_path.exists():
        try:
            vsdx_path.unlink()
        except PermissionError:
            raise PermissionError("Close file in Visio first.")

    with open(puml_path, "r", encoding="utf-8") as f:
        raw_code = f.read()

    final_source_code = PLANTUML_WATERMARK + "\n\n" + strip_watermark(raw_code)

    visio = None
    doc = None
    page = None
    page_sheet = None
    src_page = None
    text_box = None

    def stage(message: str):
        emit(f"   🔎 {message}", logging.INFO)

    try:
        stage("Starting Microsoft Visio COM server...")
        visio = win32com.client.DispatchEx("Visio.Application")
        visio.Visible = False
        visio.AlertResponse = 7
        stage("Visio started successfully.")

        stage("Creating blank Visio document...")
        doc = visio.Documents.Add("")
        stage("Blank Visio document created.")

        page = doc.Pages(1)
        page.Name = "Sequence Diagram"
        stage(f"Importing SVG: {svg_path.name}")
        try:
            page.Import(str(svg_path.resolve()))
        except Exception as e:
            raise RuntimeError(
                "SVG_IMPORT_FAILED\n"
                f"Visio rejected the generated SVG file:\n"
                f"{svg_path.resolve()}\n\n"
                f"Original COM error: {e}"
            ) from e
        stage(f"SVG imported successfully ({page.Shapes.Count} top-level shape(s)).")

        if page.Shapes.Count > 0:
            stage("Reading imported SVG dimensions...")
            orig_w = page.Shapes(1).CellsU("Width").ResultIU
            orig_h = page.Shapes(1).CellsU("Height").ResultIU
            stage(f"Imported SVG dimensions: {orig_w:.3f} × {orig_h:.3f} in.")

            stage("Ungrouping imported SVG shapes...")
            ungroup_count = flatten_svg_groups(page, emit)
            stage(f"Ungrouping completed ({ungroup_count} group operation(s)).")

            stage("Checking for PlantUML background rectangle...")
            deleted_backgrounds = remove_background_shapes(page, orig_w, orig_h)
            stage(f"Background cleanup completed ({deleted_backgrounds} shape(s) removed).")

            stage("Normalizing imported text shapes...")
            text_shape_count, text_shape_errors = normalize_text_shapes(page.Shapes)
            stage(
                f"Text normalization completed "
                f"({text_shape_count} text shape(s), {text_shape_errors} skipped/error shape(s))."
            )

            stage("Resizing Visio page to contents...")
            page_sheet = resize_page_to_contents(page, emit)

        stage("Creating embedded PlantUML source page...")
        src_page, text_box = append_plantuml_source_page(doc, final_source_code)
        stage("PlantUML source page created successfully.")

        try:
            if visio.ActiveWindow:
                visio.ActiveWindow.Page = page
        except Exception:
            pass

        stage(f"Saving Visio document: {vsdx_path.name}")
        try:
            doc.SaveAs(str(vsdx_path.resolve()))
        except Exception as e:
            raise RuntimeError(
                "VSDX_SAVE_FAILED\n"
                f"Visio successfully imported and processed the SVG, but failed while saving:\n"
                f"{vsdx_path.resolve()}\n\n"
                f"Original COM error: {e}"
            ) from e
        stage("Visio document saved successfully.")

        # Explicitly release child COM proxies BEFORE closing the document / terminating Visio.
        for ref_name in ("text_box", "src_page", "page_sheet", "page"):
            try:
                if ref_name == "text_box":
                    text_box = None
                elif ref_name == "src_page":
                    src_page = None
                elif ref_name == "page_sheet":
                    page_sheet = None
                else:
                    page = None
            except Exception:
                pass
        _ = (text_box, src_page, page_sheet, page)

        doc.Close()
        doc = None

        visio.Quit()
        visio = None

        gc.collect()
        emit(f"✅ Saved: {vsdx_path.name}", logging.INFO)
        return vsdx_path

    except Exception as e:
        if doc is not None:
            try:
                doc.Close()
            except Exception:
                pass
        if visio is not None:
            try:
                visio.Quit()
            except Exception:
                pass
            visio = None
        gc.collect()
        raise RuntimeError(f"Visio COM Error: {e}") from e
