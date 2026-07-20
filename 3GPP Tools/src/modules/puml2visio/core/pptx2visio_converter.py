import logging
import tempfile
from pathlib import Path

import pythoncom
import win32com.client
from PyQt5.QtCore import QThread, pyqtSignal


class PptxToVisioConverterThread(QThread):
    """
    Background thread to convert a PowerPoint presentation into a multi-page Visio document
    using Enhanced Metafile (EMF) as a lossless intermediate vector format.
    """
    ui_log_msg = pyqtSignal(str)
    finished_path = pyqtSignal(str)

    def __init__(self, pptx_path: Path):
        super().__init__()
        self.pptx_path = pptx_path
        self.temp_dir = Path(tempfile.gettempdir()) / "3gpp_emf_exports"
        self.temp_dir.mkdir(parents=True, exist_ok=True)

    def run(self):
        pythoncom.CoInitialize()
        ppt_app = None
        visio_app = None
        emf_paths = []

        try:
            self._emit_log(f"\n⚙️ Starting PowerPoint to Visio conversion for: {self.pptx_path.name}")

            # ---------------------------------------------------------
            # PHASE 1: Export Slides from PowerPoint to EMF
            # ---------------------------------------------------------
            self._emit_log("⏳ Opening PowerPoint to extract slides...")
            ppt_app = win32com.client.DispatchEx("PowerPoint.Application")

            # Open presentation silently
            pres = ppt_app.Presentations.Open(str(self.pptx_path.resolve()), ReadOnly=True, WithWindow=False)

            for i, slide in enumerate(pres.Slides):
                emf_file = self.temp_dir / f"slide_{i + 1}.emf"
                # Export slide as EMF (FilterName="EMF")
                slide.Export(str(emf_file), "EMF")
                emf_paths.append(emf_file)
                self._emit_log(f"   -> Exported Slide {i + 1} to temporary EMF.")

            pres.Close()
            ppt_app.Quit()
            ppt_app = None  # Clear reference

            # ---------------------------------------------------------
            # PHASE 2: Import EMFs into Visio and Clean Up
            # ---------------------------------------------------------
            self._emit_log("⏳ Spinning up Visio engine to build pages...")
            visio_app = win32com.client.DispatchEx("Visio.Application")
            visio_app.Visible = False
            visio_app.AlertResponse = 7  # Auto-answer OK to alerts

            doc = visio_app.Documents.Add("")

            for i, emf_path in enumerate(emf_paths):
                if i == 0:
                    page = doc.Pages(1)
                else:
                    page = doc.Pages.Add()

                page.Name = f"Slide {i + 1}"
                page.Import(str(emf_path.resolve()))

                if page.Shapes.Count > 0:
                    self._emit_log(f"   -> Processing shapes on Page {i + 1}...")
                    self._apply_canvas_fixes(page)

            # ---------------------------------------------------------
            # PHASE 3: Save and Cleanup
            # ---------------------------------------------------------
            vsdx_path = self.pptx_path.with_suffix(".vsdx")
            if vsdx_path.exists():
                try:
                    vsdx_path.unlink()
                except PermissionError:
                    raise PermissionError(f"Please close {vsdx_path.name} in Visio before converting.")

            doc.SaveAs(str(vsdx_path.resolve()))
            doc.Close()
            visio_app.Quit()
            visio_app = None

            self._emit_log(f"✅ Success! Saved as: {vsdx_path.name}")
            self.finished_path.emit(str(vsdx_path.resolve()))

        except Exception as e:
            self._emit_log(f"❌ Conversion Error: {str(e)}")
            self.finished_path.emit("")
        finally:
            # Clean up temporary EMFs
            for emf in emf_paths:
                if emf.exists():
                    try:
                        emf.unlink()
                    except:
                        pass

            # Failsafe COM cleanup
            if ppt_app:
                try:
                    ppt_app.Quit()
                except:
                    pass
            if visio_app:
                try:
                    visio_app.Quit()
                except:
                    pass
            pythoncom.CoUninitialize()

    def _emit_log(self, message: str):
        """Helper to emit logs to the UI."""
        logging.info(message)
        self.ui_log_msg.emit(message)

    def _apply_canvas_fixes(self, page):
        """Reuses the proven ungrouping and text shrinkage logic from PlantUML conversions."""
        # --- CANVAS FIX 1: Aggressive Flattening ---
        peeling = True
        while peeling:
            peeling = False
            for i in range(page.Shapes.Count, 0, -1):
                s = page.Shapes(i)
                try:
                    if s.Type == 2:  # visTypeGroup
                        s.Ungroup()
                        peeling = True
                except:
                    pass

        # --- CANVAS FIX 2: Text Margin & Pin Adjustment ---
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

                        # If it's pure text (no borders/fill), recalculate boundaries
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

        # Shrink page to fit contents
        page_sheet = page.PageSheet
        page_sheet.CellsU("PageLeftMargin").FormulaU = "0.05 in"
        page_sheet.CellsU("PageRightMargin").FormulaU = "0.05 in"
        page_sheet.CellsU("PageTopMargin").FormulaU = "0.05 in"
        page_sheet.CellsU("PageBottomMargin").FormulaU = "0.05 in"
        try:
            page.ResizeToFitContents()
        except:
            pass