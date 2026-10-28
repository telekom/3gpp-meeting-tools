import logging
import platform
import re
from pathlib import Path
from typing import Any, Dict, List, Optional

import pythoncom
import win32com.client
import gc
from PyQt5.QtCore import QThread, pyqtSignal

from modules.puml2visio.utils.utils import strip_watermark, generate_cleaned_svg
from modules.puml2visio.config.paths import PLANTUML_WATERMARK

logger = logging.getLogger(__name__)


def get_open_powerpoint_presentations() -> List[Dict[str, Any]]:
    """Return a list of dicts describing currently open presentations in Microsoft PowerPoint.

    Returns:
        List of dicts with keys:
            - 'name': str (e.g. '5G_Architecture.pptx' or 'Presentation1')
            - 'full_name': str (full disk path, or presentation name if unsaved)
            - 'slides_count': int (current number of slides)
            - 'is_active': bool (True if this presentation is the foreground active window)
            - 'saved': bool (True if there are no unsaved modifications)
        Returns an empty list if PowerPoint is not running or has no presentations open.
    """
    if platform.system() != "Windows":
        return []

    pythoncom.CoInitialize()
    ppt = None
    try:
        try:
            ppt = win32com.client.GetActiveObject("PowerPoint.Application")
        except Exception:
            logger.debug("No active PowerPoint.Application instance detected via COM.")
            return []

        try:
            if ppt.Presentations.Count == 0:
                logger.info("PowerPoint is running, but 0 presentations are open.")
                return []
        except Exception:
            return []

        active_pres = None
        try:
            active_pres = ppt.ActivePresentation
        except Exception:
            active_pres = None

        active_name = ""
        if active_pres is not None:
            try:
                active_name = str(active_pres.Name or "").strip()
            except Exception:
                pass

        results = []
        for pres in ppt.Presentations:
            try:
                name = str(pres.Name or "").strip()
                full_name = str(pres.FullName or "").strip()
                slides_count = int(pres.Slides.Count)
                saved = bool(getattr(pres, "Saved", True))
                is_active = (name == active_name) if active_name else False

                results.append({
                    "name": name,
                    "full_name": full_name or name,
                    "slides_count": slides_count,
                    "is_active": is_active,
                    "saved": saved,
                })
            except Exception as e:
                logger.debug(f"Error inspecting PowerPoint presentation: {e}")
                continue

        return results
    finally:
        ppt = None
        gc.collect()
        pythoncom.CoUninitialize()


class PptxConverterThread(QThread):
    """Background thread to generate a PowerPoint slide using Visio as an EMF translator, with a native SVG fallback."""
    ui_log_msg = pyqtSignal(str)
    finished_path = pyqtSignal(str)

    def __init__(
        self,
        puml_path: Path,
        jar_path: Path,
        append_to_open: bool = False,
        target_presentation: Optional[str] = None,
        insert_at: str = "end",
    ):
        super().__init__()
        self.puml_path = puml_path
        self.jar_path = jar_path
        self.append_to_open = append_to_open
        self.target_presentation = target_presentation
        self.insert_at = insert_at

    def run(self):
        pythoncom.CoInitialize()
        ppt = None
        pres = None
        slide = None
        shape = None
        emf_path = None
        is_appending = False
        try:
            action_desc = "Appending slide to open presentation" if self.append_to_open else "Generating PowerPoint slide"
            self._emit_log(f"⚙️ {action_desc} for: {self.puml_path.name}", logging.INFO)

            # Step 1: Generate standard SVG from PlantUML
            svg_path = generate_cleaned_svg(self.puml_path, self.jar_path, self._emit_log)

            with open(self.puml_path, "r", encoding="utf-8") as f:
                raw_code = f.read()
            final_source_code = PLANTUML_WATERMARK + "\n\n" + strip_watermark(raw_code)

            # Step 2: Attempt Visio EMF Translation with Native PPT Fallback
            target_import_path = svg_path
            try:
                self._emit_log("⏳ Translating SVG to Microsoft EMF via Visio engine...", logging.INFO)
                emf_path = self._create_emf_via_visio(svg_path)
                if emf_path and emf_path.exists():
                    target_import_path = emf_path
            except Exception as e:
                self._emit_log(f"⚠️ Visio EMF Export Failed: {e}", logging.WARNING)
                self._emit_log("🔄 Falling back to native PowerPoint SVG conversion...", logging.INFO)

            if self.append_to_open:
                try:
                    self._emit_log("⏳ Attaching to open Microsoft PowerPoint instance...", logging.INFO)
                    ppt = win32com.client.GetActiveObject("PowerPoint.Application")
                    ppt.Visible = 1
                    if ppt.WindowState == 2:  # ppWindowStateMinimized
                        ppt.WindowState = 1   # ppWindowStateNormal
                    ppt.DisplayAlerts = 1

                    # Match presentation by target_presentation if specified
                    if self.target_presentation:
                        for p in ppt.Presentations:
                            try:
                                if p.Name == self.target_presentation or p.FullName == self.target_presentation:
                                    pres = p
                                    break
                            except Exception:
                                continue

                    if pres is None and ppt.Presentations.Count > 0:
                        try:
                            pres = ppt.ActivePresentation
                        except Exception:
                            pres = ppt.Presentations(1)

                    if pres is not None:
                        is_appending = True
                        self._emit_log(
                            f"🎯 Target presentation: '{pres.Name}' ({pres.Slides.Count} existing slide(s))",
                            logging.INFO,
                        )
                except Exception as e:
                    self._emit_log(
                        f"⚠️ Could not attach to open PowerPoint ({e}). Falling back to new presentation.",
                        logging.WARNING,
                    )
                    ppt = None
                    pres = None

            if not is_appending:
                if ppt is None:
                    ppt = win32com.client.DispatchEx("PowerPoint.Application")
                    ppt.Visible = 1
                    if ppt.WindowState == 2:
                        ppt.WindowState = 1
                    ppt.DisplayAlerts = 1

                pres = ppt.Presentations.Add()
                slide = pres.Slides.Add(1, 12)  # 12 = ppLayoutBlank
            else:
                # Append as a new blank slide at the very end of the presentation
                new_slide_idx = pres.Slides.Count + 1
                slide = pres.Slides.Add(new_slide_idx, 12)  # 12 = ppLayoutBlank

            # Import either the Visio EMF or the raw SVG if Visio failed
            shape = slide.Shapes.AddPicture(str(target_import_path.resolve()), 0, -1, 0, 0, -1, -1)

            file_type = "EMF" if target_import_path.suffix == ".emf" else "SVG"
            self._emit_log(f"⏳ Unpacking {file_type} into native PowerPoint shapes...", logging.INFO)

            try:
                # Calling Ungroup on an SVG in PPT automatically triggers "Convert to Shape"
                sr = shape.Ungroup()

                if sr.Count > 1:
                    shape = sr.Group()
                elif sr.Count == 1:
                    shape = sr(1)

                self._emit_log("✅ Successfully generated native shapes.", logging.INFO)
            except Exception as e:
                self._emit_log(f"⚠️ Could not unpack {file_type}: {e}", logging.WARNING)
                try:
                    _ = shape.Width
                except:
                    if slide.Shapes.Count > 0:
                        shape = slide.Shapes(slide.Shapes.Count)

            slide_w = pres.PageSetup.SlideWidth
            slide_h = pres.PageSetup.SlideHeight

            shape.LockAspectRatio = -1  # msoTrue
            margin = 20

            if shape.Width > slide_w - margin or shape.Height > slide_h - margin:
                width_ratio = (slide_w - margin) / shape.Width
                height_ratio = (slide_h - margin) / shape.Height
                scale_ratio = min(width_ratio, height_ratio)
                shape.Width = shape.Width * scale_ratio

            shape.Left = (slide_w - shape.Width) / 2
            shape.Top = (slide_h - shape.Height) / 2

            try:
                notes = slide.NotesPage
                for i in range(1, notes.Shapes.Count + 1):
                    if notes.Shapes(i).HasTextFrame:
                        notes.Shapes(i).TextFrame.TextRange.Text = final_source_code
                        break
            except Exception as e:
                self._emit_log(f"⚠️ Warning: Could not write to Speaker Notes: {e}", logging.WARNING)

            if svg_path.exists():
                try:
                    svg_path.unlink()
                except:
                    pass
            if emf_path and emf_path.exists():
                try:
                    emf_path.unlink()
                except:
                    pass

            if is_appending:
                # Focus the new slide and bring window to front
                try:
                    slide.Select()
                    if pres.Windows.Count > 0:
                        pres.Windows(1).Activate()
                        pres.Windows(1).View.GotoSlide(slide.SlideIndex)
                except Exception as e:
                    self._emit_log(f"ℹ️ Note on focusing slide: {e}", logging.DEBUG)

                self._emit_log(
                    f"✅ Slide added to presentation '{pres.Name}' (Slide {slide.SlideIndex})!",
                    logging.INFO,
                )
                self.finished_path.emit(str(pres.FullName or pres.Name))
            else:
                self._emit_log("✅ Slide generated! PowerPoint left open for copying.", logging.INFO)
                self.finished_path.emit("OPENED_IN_PPT")

        except Exception as e:
            if pres and not is_appending:
                try:
                    pres.Saved = True
                    pres.Close()
                except Exception:
                    pass
            if ppt and not is_appending:
                try:
                    ppt.Quit()
                except Exception:
                    pass
            err_str = str(e)
            if "0x80010001" in err_str or "rejected" in err_str.lower():
                self._emit_log(
                    "⚠️ PowerPoint is currently busy or editing text. Please exit text-editing mode in PowerPoint and try again.",
                    logging.ERROR,
                )
            else:
                self._emit_log(f"❌ PowerPoint COM Error: {err_str}", logging.ERROR)
            self.finished_path.emit("")
        finally:
            # Destroy any local COM wrappers while this thread's apartment still exists.
            shape = slide = pres = ppt = None
            gc.collect()
            pythoncom.CoUninitialize()

    def _create_emf_via_visio(self, svg_path: Path) -> Path:
        """Use an isolated Visio COM server to translate SVG to EMF safely."""
        visio = None
        doc = None
        page = None
        page_sheet = None
        s = None
        try:
            visio = win32com.client.DispatchEx("Visio.Application")
            visio.Visible = False
            visio.AlertResponse = 7
            doc = visio.Documents.Add("")
            page = doc.Pages(1)
            page.Import(str(svg_path.resolve()))

            if page.Shapes.Count > 0:
                orig_w = page.Shapes(1).CellsU("Width").ResultIU
                orig_h = page.Shapes(1).CellsU("Height").ResultIU

                # --- CANVAS FIX 1: Aggressive Flattening ---
                peeling = True
                while peeling:
                    peeling = False
                    for i in range(page.Shapes.Count, 0, -1):
                        s = page.Shapes(i)
                        try:
                            if s.Type == 2:
                                s.Ungroup()
                                peeling = True
                        except:
                            pass

                # --- CANVAS FIX 2: Background Rect Deletion ---
                for i in range(page.Shapes.Count, 0, -1):
                    s = page.Shapes(i)
                    try:
                        w = s.CellsU("Width").ResultIU
                        h = s.CellsU("Height").ResultIU
                        if w >= orig_w * 0.75 and h >= orig_h * 0.75:
                            if len(s.Characters.Text.strip()) == 0:
                                if s.CellsU("LinePattern").ResultIU == 0:
                                    s.Delete()
                    except:
                        pass

                def clean_and_shrink_text(shapes):
                    for i in range(1, shapes.Count + 1):
                        s = shapes(i)
                        try:
                            if len(s.Characters.Text.strip()) > 0:
                                s.CellsU("TopMargin").FormulaU = "0 pt"
                                s.CellsU("BottomMargin").FormulaU = "0 pt"
                                s.CellsU("LeftMargin").FormulaU = "0 pt"
                                s.CellsU("RightMargin").FormulaU = "0 pt"

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
                        except:
                            pass
                        try:
                            if s.Type == 2:
                                clean_and_shrink_text(s.Shapes)
                        except:
                            pass

                clean_and_shrink_text(page.Shapes)

                page_sheet = page.PageSheet
                page_sheet.CellsU("PageLeftMargin").FormulaU = "0.05 in"
                page_sheet.CellsU("PageRightMargin").FormulaU = "0.05 in"
                page_sheet.CellsU("PageTopMargin").FormulaU = "0.05 in"
                page_sheet.CellsU("PageBottomMargin").FormulaU = "0.05 in"
                try:
                    page.ResizeToFitContents()
                except:
                    pass

            emf_path = svg_path.with_suffix(".emf")
            if emf_path.exists():
                try:
                    emf_path.unlink()
                except:
                    pass

            page.Export(str(emf_path.resolve()))

            # Release child COM proxies before closing their parent document/application.
            s = None
            page_sheet = None
            page = None

            doc.Close()
            doc = None
            visio.Quit()
            visio = None
            gc.collect()
            return emf_path

        except Exception as e:
            s = None
            page_sheet = None
            page = None
            if doc is not None:
                try:
                    doc.Close()
                except Exception:
                    pass
                doc = None
            if visio is not None:
                try:
                    visio.Quit()
                except Exception:
                    pass
                visio = None
            gc.collect()
            raise RuntimeError(f"Visio EMF Export Failed: {e}") from e

    def _emit_log(self, message: str, level: int):
        logger.log(level, message)
