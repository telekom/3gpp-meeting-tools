import subprocess
import re
import logging
import time
from pathlib import Path

import pythoncom
import win32com.client
import win32gui
import win32con
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
            self._emit_log(f"\n⚙️ Generating PowerPoint slide for: {self.puml_path.name}", logging.INFO)
            svg_path = self._generate_svg()

            with open(self.puml_path, "r", encoding="utf-8") as f:
                source_code = f.read()

            ppt = win32com.client.DispatchEx("PowerPoint.Application")

            # PowerPoint must be visible and normal-sized for Ribbon commands
            ppt.Visible = 1
            try:
                if ppt.WindowState == 2:  # If minimized, restore it
                    ppt.WindowState = 1
            except:
                pass

            # --- CRITICAL FIX: Force PowerPoint to the absolute foreground ---
            # Ribbon commands fail if the application doesn't have OS focus
            try:
                hwnd = ppt.HWND
                if hwnd:
                    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                    win32gui.SetForegroundWindow(hwnd)
            except Exception as e:
                pass

            ppt.DisplayAlerts = 1  # Suppress UI popups

            pres = ppt.Presentations.Add()
            slide = pres.Slides.Add(1, 12)  # 12 = ppLayoutBlank

            # Insert the SVG
            shape = slide.Shapes.AddPicture(str(svg_path.resolve()), 0, -1, 0, 0, -1, -1)

            self._emit_log("⏳ Converting SVG to native PowerPoint shapes...", logging.INFO)

            converted = False

            # Method 1: Robust Ribbon Command Execution
            for attempt in range(15):
                try:
                    # Ensure window focus and shape selection
                    ppt.ActiveWindow.Activate()
                    shape.Select()

                    # ONLY execute if PowerPoint confirms the button is currently clickable!
                    # This completely prevents the DISP_E_EXCEPTION crash.
                    if ppt.CommandBars.GetEnabledMso("PictureConvertToShape"):
                        ppt.CommandBars.ExecuteMso("PictureConvertToShape")

                        # Wait for the shape to convert into a native Group (Type 6)
                        for _ in range(15):
                            time.sleep(0.1)
                            if slide.Shapes.Count > 0 and slide.Shapes(1).Type in [6, 5]:
                                shape = slide.Shapes(1)
                                converted = True
                                break
                    if converted:
                        break
                except Exception as e:
                    pass  # Ignore temporary COM lockups and retry
                time.sleep(0.2)

            # Method 2: Fallback to old Ungroup method if Ribbon wasn't available
            if not converted:
                for attempt in range(3):
                    try:
                        ppt.DisplayAlerts = 1
                        shape_range = shape.Ungroup()
                        shape = shape_range.Group()
                        converted = True
                        break
                    except:
                        pass
                    time.sleep(0.5)

            if converted:
                self._emit_log("✅ Successfully converted to native shapes.", logging.INFO)
            else:
                self._emit_log("⚠️ Ribbon conversion timed out. Leaving as embedded SVG.", logging.WARNING)

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

            if svg_path.exists():
                svg_path.unlink()

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