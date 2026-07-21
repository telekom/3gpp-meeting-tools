# --- File: src/modules/puml2visio/core/visio2pptx_converter.py ---
import logging
import tempfile
from pathlib import Path

import pythoncom
import win32com.client
from PyQt5.QtCore import QThread, pyqtSignal


class VisioToPptxConverterThread(QThread):
    """
    Background thread to convert a multi-page Visio document into a PowerPoint presentation.
    Uses a hybrid approach: Deep COM canvas optimization (shrink-wrapping text) combined with
    Enhanced Metafile (EMF) bridging to ensure perfect line attachment and font fidelity.
    """
    ui_log_msg = pyqtSignal(str)
    finished_path = pyqtSignal(str)

    def __init__(self, vsdx_path: Path):
        super().__init__()
        self.vsdx_path = vsdx_path
        self.temp_dir = Path(tempfile.gettempdir()) / "3gpp_emf_exports"
        self.temp_dir.mkdir(parents=True, exist_ok=True)

    def run(self):
        pythoncom.CoInitialize()
        visio_app = None
        ppt_app = None
        emf_paths = []

        try:
            self._emit_log(f"\n⚙️ Starting Visio to PowerPoint conversion for: {self.vsdx_path.name}")

            # ---------------------------------------------------------
            # PHASE 1: Optimize Canvas and Export to EMF
            # ---------------------------------------------------------
            self._emit_log("⏳ Opening Visio to optimize and extract pages...")
            visio_app = win32com.client.DispatchEx("Visio.Application")
            visio_app.Visible = False
            visio_app.AlertResponse = 7  # Auto-answer OK to alerts

            # Open document ReadOnly (visOpenRO = 2) so we don't overwrite their source file
            doc = visio_app.Documents.OpenEx(str(self.vsdx_path.resolve()), 2)

            for i in range(1, doc.Pages.Count + 1):
                page = doc.Pages(i)

                # ---> THE HYBRID FIX: Shrink wrap the text bounds BEFORE EMF export!
                self._apply_canvas_fixes(page)

                emf_file = self.temp_dir / f"page_{i}.emf"
                # Export page as native EMF to preserve strict line coordinates
                page.Export(str(emf_file.resolve()))
                emf_paths.append(emf_file)
                self._emit_log(f"   -> Optimized and exported Page {i} to temporary EMF.")

            # Tell Visio the document is "saved" so it closes cleanly without prompting
            doc.Saved = True
            doc.Close()
            visio_app.Quit()
            visio_app = None

            # ---------------------------------------------------------
            # PHASE 2: Import EMFs into PowerPoint and Ungroup
            # ---------------------------------------------------------
            self._emit_log("⏳ Spinning up PowerPoint engine to build slides...")
            ppt_app = win32com.client.DispatchEx("PowerPoint.Application")
            ppt_app.Visible = 1
            if ppt_app.WindowState == 2:
                ppt_app.WindowState = 1
            ppt_app.DisplayAlerts = 1

            pres = ppt_app.Presentations.Add()

            # ---> THE THEME FIX: Force Arial Theme Font to bypass Aptos fallback
            try:
                pres.SlideMaster.Theme.ThemeFontScheme.MinorFont.Name = "Arial"
                pres.SlideMaster.Theme.ThemeFontScheme.MajorFont.Name = "Arial"
            except Exception as e:
                self._emit_log(f"⚠️ Could not force Arial theme: {e}")

            slide_w = pres.PageSetup.SlideWidth
            slide_h = pres.PageSetup.SlideHeight

            for i, emf_path in enumerate(emf_paths):
                # Add a blank slide (12 = ppLayoutBlank)
                slide = pres.Slides.Add(i + 1, 12)

                # Insert the EMF graphic natively
                shape = slide.Shapes.AddPicture(str(emf_path.resolve()), 0, -1, 0, 0, -1, -1)

                self._emit_log(f"   -> Unpacking EMF into native shapes on Slide {i + 1}...")
                try:
                    sr = shape.Ungroup()
                    if sr.Count > 1:
                        shape = sr.Group()
                    elif sr.Count == 1:
                        shape = sr(1)
                except Exception as e:
                    self._emit_log(f"⚠️ Could not unpack EMF on Slide {i + 1}: {e}")
                    if slide.Shapes.Count > 0:
                        shape = slide.Shapes(slide.Shapes.Count)

                # Scale and center the generated shape group
                shape.LockAspectRatio = -1
                margin = 20
                if shape.Width > slide_w - margin or shape.Height > slide_h - margin:
                    width_ratio = (slide_w - margin) / shape.Width
                    height_ratio = (slide_h - margin) / shape.Height
                    scale_ratio = min(width_ratio, height_ratio)
                    shape.Width = shape.Width * scale_ratio

                shape.Left = (slide_w - shape.Width) / 2
                shape.Top = (slide_h - shape.Height) / 2

            # ---------------------------------------------------------
            # PHASE 3: Save and Cleanup
            # ---------------------------------------------------------
            pptx_path = self.vsdx_path.with_suffix(".pptx")
            if pptx_path.exists():
                try:
                    pptx_path.unlink()
                except PermissionError:
                    raise PermissionError(f"Please close {pptx_path.name} in PowerPoint before converting.")

            pres.SaveAs(str(pptx_path.resolve()))
            pres.Close()
            ppt_app.Quit()
            ppt_app = None

            self._emit_log(f"✅ Success! Saved as: {pptx_path.name}")
            self.finished_path.emit(str(pptx_path.resolve()))

        except Exception as e:
            self._emit_log(f"❌ Conversion Error: {str(e)}")
            self.finished_path.emit("")
        finally:
            for emf in emf_paths:
                if emf.exists():
                    try:
                        emf.unlink()
                    except:
                        pass
            if visio_app:
                try:
                    visio_app.Quit()
                except:
                    pass
            if ppt_app:
                try:
                    ppt_app.Quit()
                except:
                    pass
            pythoncom.CoUninitialize()

    def _apply_canvas_fixes(self, page):
        """
        Recursively zeroes out text margins and dynamically shrinks bounding boxes
        to prevent PowerPoint from prematurely wrapping/splitting text upon ungrouping.
        """

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

                        # Only aggressively shrink-wrap if it's pure text (no borders/fills)
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

                            # Dynamically shrink boundaries to exact visual text size
                            s.CellsU("Width").FormulaU = "TEXTWIDTH(TheText)"
                            s.CellsU("Height").FormulaU = "TEXTHEIGHT(TheText, Width)"
                except:
                    pass

                # Recurse into groups
                try:
                    if s.Type == 2:  # visTypeGroup
                        clean_and_shrink_text(s.Shapes)
                except:
                    pass

        clean_and_shrink_text(page.Shapes)

    def _emit_log(self, message: str):
        logging.info(message)
        self.ui_log_msg.emit(message)