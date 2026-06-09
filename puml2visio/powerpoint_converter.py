import logging
from pathlib import Path

import pythoncom
import win32com.client
from PyQt5.QtCore import QThread, pyqtSignal

from utils import WATERMARK, strip_watermark, generate_cleaned_svg


class PptxConverterThread(QThread):
    """Background thread to generate a PowerPoint slide using Visio as an EMF translator."""
    ui_log_msg = pyqtSignal(str)
    finished_path = pyqtSignal(str)

    def __init__(self, puml_path: Path, jar_path: Path):
        super().__init__()
        self.puml_path = puml_path
        self.jar_path = jar_path

    def run(self):
        pythoncom.CoInitialize()
        ppt = None
        try:
            self._emit_log(f"\n⚙️ Generating PowerPoint slide for: {self.puml_path.name}", logging.INFO)

            svg_path = generate_cleaned_svg(self.puml_path, self.jar_path, self._emit_log)

            with open(self.puml_path, "r", encoding="utf-8") as f:
                raw_code = f.read()
            final_source_code = WATERMARK + "\n\n" + strip_watermark(raw_code)

            self._emit_log("⏳ Translating SVG to Microsoft EMF via Visio engine...", logging.INFO)
            emf_path = self._create_emf_via_visio(svg_path)

            ppt = win32com.client.DispatchEx("PowerPoint.Application")
            ppt.Visible = 1
            if ppt.WindowState == 2:
                ppt.WindowState = 1

            ppt.DisplayAlerts = 1

            pres = ppt.Presentations.Add()
            slide = pres.Slides.Add(1, 12)  # 12 = ppLayoutBlank

            shape = slide.Shapes.AddPicture(str(emf_path.resolve()), 0, -1, 0, 0, -1, -1)

            self._emit_log("⏳ Unpacking EMF into native PowerPoint shapes...", logging.INFO)
            try:
                sr = shape.Ungroup()

                if sr.Count > 1:
                    shape = sr.Group()
                elif sr.Count == 1:
                    shape = sr(1)

                self._emit_log("✅ Successfully generated native shapes.", logging.INFO)
            except Exception as e:
                self._emit_log(f"⚠️ Could not unpack EMF: {e}", logging.WARNING)
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

            ppt = None

            if svg_path.exists():
                try:
                    svg_path.unlink()
                except:
                    pass
            if emf_path.exists():
                try:
                    emf_path.unlink()
                except:
                    pass

            self._emit_log(f"✅ Slide generated! PowerPoint left open for copying.", logging.INFO)
            self.finished_path.emit("OPENED_IN_PPT")

        except Exception as e:
            if ppt:
                try:
                    ppt.Quit()
                except:
                    pass
            self._emit_log(f"❌ PowerPoint COM Error: {str(e)}", logging.ERROR)
            self.finished_path.emit("")
        finally:
            pythoncom.CoUninitialize()

    def _create_emf_via_visio(self, svg_path: Path) -> Path:
        """Silently uses Visio to parse the SVG, fix text padding, and export as a native Microsoft EMF."""
        visio = win32com.client.DispatchEx("Visio.Application")
        visio.Visible = False
        visio.AlertResponse = 7
        doc = None
        try:
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

            if doc: doc.Close()
            visio.Quit()
            return emf_path

        except Exception as e:
            if doc:
                try:
                    doc.Close()
                except:
                    pass
            if visio:
                try:
                    visio.Quit()
                except:
                    pass
            raise RuntimeError(f"Visio EMF Export Failed: {e}")

    def _emit_log(self, message: str, level: int):
        logging.log(level, message)
        self.ui_log_msg.emit(message)