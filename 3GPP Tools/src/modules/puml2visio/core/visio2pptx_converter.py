# --- File: src/modules/puml2visio/core/visio2pptx_converter.py ---
import logging
import tempfile
from pathlib import Path

import pythoncom
import win32com.client
from PyQt5.QtCore import QThread, pyqtSignal


class VisioToPptxConverterThread(QThread):
    """
    Background thread to convert a multi-page Visio document into a PowerPoint presentation
    using Enhanced Metafile (EMF) as a lossless intermediate vector format to generate native shapes.
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
            # PHASE 1: Export Pages from Visio to EMF
            # ---------------------------------------------------------
            self._emit_log("⏳ Opening Visio to extract pages...")
            visio_app = win32com.client.DispatchEx("Visio.Application")
            visio_app.Visible = False
            visio_app.AlertResponse = 7  # Auto-answer OK to alerts

            # Open document ReadOnly (2)
            doc = visio_app.Documents.OpenEx(str(self.vsdx_path.resolve()), 2)

            for i in range(1, doc.Pages.Count + 1):
                page = doc.Pages(i)
                emf_file = self.temp_dir / f"page_{i}.emf"
                # Export page directly as EMF
                page.Export(str(emf_file.resolve()))
                emf_paths.append(emf_file)
                self._emit_log(f"   -> Exported Page {i} to temporary EMF.")

            doc.Close()
            visio_app.Quit()
            visio_app = None  # Clear reference

            # ---------------------------------------------------------
            # PHASE 2: Import EMFs into PowerPoint and Ungroup
            # ---------------------------------------------------------
            self._emit_log("⏳ Spinning up PowerPoint engine to build slides...")
            ppt_app = win32com.client.DispatchEx("PowerPoint.Application")
            # PowerPoint usually requires visibility to safely construct presentations via COM
            ppt_app.Visible = 1
            if ppt_app.WindowState == 2:
                ppt_app.WindowState = 1
            ppt_app.DisplayAlerts = 1

            pres = ppt_app.Presentations.Add()
            slide_w = pres.PageSetup.SlideWidth
            slide_h = pres.PageSetup.SlideHeight

            for i, emf_path in enumerate(emf_paths):
                # Add a blank slide (12 = ppLayoutBlank)
                slide = pres.Slides.Add(i + 1, 12)

                # Insert the EMF graphic
                shape = slide.Shapes.AddPicture(str(emf_path.resolve()), 0, -1, 0, 0, -1, -1)

                self._emit_log(f"   -> Unpacking EMF into native shapes on Slide {i + 1}...")
                try:
                    # The Ungrouping Magic!
                    sr = shape.Ungroup()
                    if sr.Count > 1:
                        shape = sr.Group()
                    elif sr.Count == 1:
                        shape = sr(1)
                except Exception as e:
                    self._emit_log(f"⚠️ Could not unpack EMF on Slide {i + 1}: {e}")
                    # Fallback: Just grab the shape if ungrouping fails
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
            # Clean up temporary EMFs
            for emf in emf_paths:
                if emf.exists():
                    try:
                        emf.unlink()
                    except:
                        pass

            # Failsafe COM cleanup
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

    def _emit_log(self, message: str):
        logging.info(message)
        self.ui_log_msg.emit(message)