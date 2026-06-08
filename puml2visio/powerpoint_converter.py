import subprocess
import re
import logging
from pathlib import Path

import pythoncom
import win32com.client
from PyQt5.QtCore import QThread, pyqtSignal


class PptxConverterThread(QThread):
    """Background thread to generate a PowerPoint slide with native Office shapes."""
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
            self._emit_log(f"\n⚙️ Processing: {self.puml_path.name}", logging.INFO)
            svg_path = self._generate_svg()

            pptx_path = self.puml_path.with_suffix(".pptx")
            if pptx_path.exists():
                try:
                    pptx_path.unlink()
                except PermissionError:
                    raise PermissionError("Close file in PowerPoint first.")

            with open(self.puml_path, "r", encoding="utf-8") as f:
                source_code = f.read()

            ppt = win32com.client.DispatchEx("PowerPoint.Application")

            # PowerPoint requires a visible window to securely perform Ungroup operations
            ppt.Visible = 1
            try:
                ppt.WindowState = 2  # ppWindowMinimized (keeps it out of your way)
            except:
                pass

            # ppAlertsNone = 1: Suppress the "Convert to drawing object?" dialog and auto-accept
            ppt.DisplayAlerts = 1

            pres = ppt.Presentations.Add()
            slide = pres.Slides.Add(1, 12)  # 12 = ppLayoutBlank

            # Insert the SVG
            shape = slide.Shapes.AddPicture(str(svg_path.resolve()), 0, -1, 0, 0, -1, -1)

            self._emit_log("⏳ Converting SVG to native PowerPoint shapes...", logging.INFO)
            try:
                shape_range = shape.Ungroup()
                shape = shape_range.Group()
                self._emit_log("✅ Successfully converted to native shapes.", logging.INFO)
            except Exception as e:
                self._emit_log(f"⚠️ Could not convert to native shapes: {e}. Leaving as embedded SVG.", logging.WARNING)

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

            # Embed source code in Speaker Notes
            watermark = "' Generated with puml2visio, https://github.com/telekom/3gpp-meeting-tools/tree/master/puml2visio\n\n"
            try:
                notes = slide.NotesPage
                for i in range(1, notes.Shapes.Count + 1):
                    if notes.Shapes(i).HasTextFrame:
                        notes.Shapes(i).TextFrame.TextRange.Text = watermark + source_code
                        break
            except Exception as e:
                self._emit_log(f"⚠️ Warning: Could not write notes: {e}", logging.WARNING)

            pres.SaveAs(str(pptx_path.resolve()))
            pres.Close()
            ppt.Quit()
            ppt = None

            if svg_path.exists():
                svg_path.unlink()

            self._emit_log(f"✅ Saved: {pptx_path.name}", logging.INFO)
            self.finished_path.emit(str(pptx_path.resolve()))

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