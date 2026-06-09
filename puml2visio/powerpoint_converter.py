import subprocess
import re
import logging
from pathlib import Path

import pythoncom
import win32com.client
from PyQt5.QtCore import QThread, pyqtSignal


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
            svg_path = self._generate_svg()

            with open(self.puml_path, "r", encoding="utf-8") as f:
                source_code = f.read()

            # --- THE MAGIC PIPELINE ---
            # Pipe the SVG through Visio to generate a flawless Microsoft EMF vector
            self._emit_log("⏳ Translating SVG to Microsoft EMF via Visio engine...", logging.INFO)
            emf_path = self._create_emf_via_visio(svg_path)

            # Now open PowerPoint
            ppt = win32com.client.DispatchEx("PowerPoint.Application")
            ppt.Visible = 1
            if ppt.WindowState == 2:
                ppt.WindowState = 1

                # ppAlertsNone = 1: Automatically answers "Yes" to the Ungroup safety prompt
            ppt.DisplayAlerts = 1

            pres = ppt.Presentations.Add()
            slide = pres.Slides.Add(1, 12)  # 12 = ppLayoutBlank

            # Insert the EMF Vector
            shape = slide.Shapes.AddPicture(str(emf_path.resolve()), 0, -1, 0, 0, -1, -1)

            self._emit_log("⏳ Unpacking EMF into native PowerPoint shapes...", logging.INFO)
            try:
                # EMF perfectly ungroups without relying on the UI Ribbon!
                sr = shape.Ungroup()

                # --- CRITICAL FIX ---
                # PowerPoint crashes if you try to group a single item.
                if sr.Count > 1:
                    shape = sr.Group()
                elif sr.Count == 1:
                    shape = sr(1)

                self._emit_log("✅ Successfully generated native shapes.", logging.INFO)
            except Exception as e:
                self._emit_log(f"⚠️ Could not unpack EMF: {e}", logging.WARNING)
                # Safety Net: If Ungroup succeeded but Group failed, the original shape is dead.
                # To prevent a crash during resize, we grab the last added shape on the slide.
                try:
                    _ = shape.Width
                except:
                    if slide.Shapes.Count > 0:
                        shape = slide.Shapes(slide.Shapes.Count)

            # Center and scale the figure
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

            # Embed source code securely in the Slide's Speaker Notes
            watermark = "' Generated with puml2visio, https://github.com/telekom/3gpp-meeting-tools/tree/master/puml2visio\n\n"
            try:
                notes = slide.NotesPage
                for i in range(1, notes.Shapes.Count + 1):
                    if notes.Shapes(i).HasTextFrame:
                        notes.Shapes(i).TextFrame.TextRange.Text = watermark + source_code
                        break
            except Exception as e:
                self._emit_log(f"⚠️ Warning: Could not write to Speaker Notes: {e}", logging.WARNING)

            # DO NOT SAVE OR CLOSE - Detach COM object
            ppt = None

            # Cleanup temp files
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

                peeling = True
                while peeling:
                    peeling = False
                    for i in range(page.Shapes.Count, 0, -1):
                        s = page.Shapes(i)
                        try:
                            w = s.CellsU("Width").ResultIU
                            h = s.CellsU("Height").ResultIU
                            if abs(w - orig_w) < 0.1 and abs(h - orig_h) < 0.1:
                                if s.Type == 2:
                                    s.Ungroup()
                                    peeling = True
                        except:
                            pass

                for i in range(page.Shapes.Count, 0, -1):
                    s = page.Shapes(i)
                    try:
                        w = s.CellsU("Width").ResultIU
                        h = s.CellsU("Height").ResultIU
                        if abs(w - orig_w) < 0.1 and abs(h - orig_h) < 0.1:
                            if len(s.Characters.Text.strip()) == 0:
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

    def _generate_svg(self) -> Path:
        command = ["java", "-jar", str(self.jar_path), "-tsvg", str(self.puml_path)]
        subprocess.run(command, check=True, capture_output=True, text=True, cwd=self.puml_path.parent)

        svg_path = self.puml_path.with_suffix(".svg")
        if not svg_path.exists():
            raise FileNotFoundError("PlantUML finished, but SVG was not created.")

        try:
            with open(svg_path, 'r', encoding='utf-8') as f:
                svg_content = f.read()

            svg_content = re.sub(r'\s*textLength="[^"]*"', '', svg_content)
            svg_content = re.sub(r'\s*lengthAdjust="[^"]*"', '', svg_content)

            pattern = re.compile(r'(<text\b[^>]*?\by="([0-9.]+)"[^>]*>)(.*?)(</text>)', re.IGNORECASE | re.DOTALL)
            matches = list(pattern.finditer(svg_content))

            if matches:
                result = []
                last_end = 0
                current_y = None
                current_start_tag = ""
                current_text = ""

                for m in matches:
                    start = m.start()
                    end = m.end()
                    full_open_tag = m.group(1)
                    y_val = m.group(2)
                    inner_text = m.group(3)
                    between = svg_content[last_end:start]

                    if current_y == y_val and not between.strip():
                        current_text += inner_text
                    else:
                        if current_y is not None:
                            result.append(current_start_tag)
                            result.append(current_text)
                            result.append("</text>")
                        result.append(between)
                        current_y = y_val
                        current_start_tag = full_open_tag
                        current_text = inner_text
                    last_end = end

                if current_y is not None:
                    result.append(current_start_tag)
                    result.append(current_text)
                    result.append("</text>")

                result.append(svg_content[last_end:])
                svg_content = "".join(result)

            svg_content = svg_content.replace('&#160;', ' ').replace('\xa0', ' ')

            with open(svg_path, 'w', encoding='utf-8') as f:
                f.write(svg_content)
        except Exception as e:
            self._emit_log(f"⚠️ Warning: Could not clean SVG text attributes: {e}", logging.WARNING)

        return svg_path