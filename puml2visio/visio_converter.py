import subprocess
import re
import logging
from pathlib import Path

import pythoncom
import win32com.client
from PyQt5.QtCore import QThread, pyqtSignal


class VisioReaderThread(QThread):
    text_extracted = pyqtSignal(str)
    error_occurred = pyqtSignal(str)

    def __init__(self, vsdx_path):
        super().__init__()
        self.vsdx_path = vsdx_path

    def run(self):
        pythoncom.CoInitialize()
        visio = None
        try:
            visio = win32com.client.DispatchEx("Visio.Application")
            visio.Visible = False
            visio.AlertResponse = 7

            doc = visio.Documents.OpenEx(str(Path(self.vsdx_path).resolve()), 2)
            source_code = ""
            for i in range(1, doc.Pages.Count + 1):
                page = doc.Pages(i)
                if page.Name == "PlantUML Source":
                    if page.Shapes.Count > 0:
                        raw_text = page.Shapes(1).Characters.Text

                        # --- NEW: Strip the watermark during extraction ---
                        # This prevents duplication if a user re-exports a generated Visio file
                        lines = raw_text.splitlines()
                        while lines and ("Generated with puml2visio" in lines[0] or lines[0].strip() == ""):
                            lines.pop(0)

                        source_code = "\n".join(lines)
                    break

            doc.Close()
            visio.Quit()

            if source_code:
                self.text_extracted.emit(source_code)
            else:
                self.error_occurred.emit("Could not find 'PlantUML Source' page in this Visio file.")
        except Exception as e:
            if visio: visio.Quit()
            self.error_occurred.emit(f"Error reading Visio file: {str(e)}")
        finally:
            pythoncom.CoUninitialize()


class ConverterThread(QThread):
    ui_log_msg = pyqtSignal(str)
    finished_path = pyqtSignal(str)

    def __init__(self, puml_path: Path, jar_path: Path):
        super().__init__()
        self.puml_path = puml_path
        self.jar_path = jar_path

    def run(self):
        pythoncom.CoInitialize()
        try:
            self._emit_log(f"\n⚙️ Processing: {self.puml_path.name}", logging.INFO)
            svg_path = self._generate_svg()
            self._convert_to_vsdx(svg_path)

            vsdx_path = self.puml_path.with_suffix(".vsdx")
            self.finished_path.emit(str(vsdx_path.resolve()))

            if svg_path.exists():
                svg_path.unlink()

        except Exception as e:
            self._emit_log(f"❌ Error: {str(e)}\n{'-' * 45}", logging.ERROR)
            self.finished_path.emit("")
        finally:
            pythoncom.CoUninitialize()

    def _emit_log(self, message: str, level: int):
        logging.log(level, message)
        self.ui_log_msg.emit(message)

    def _generate_svg(self) -> Path:
        command = ["java", "-jar", str(self.jar_path), "-tsvg", str(self.puml_path)]
        try:
            subprocess.run(command, check=True, capture_output=True, text=True, cwd=self.puml_path.parent)
        except subprocess.CalledProcessError as e:
            raise RuntimeError(f"PlantUML Syntax Error:\n{e.stderr}")

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

    def _convert_to_vsdx(self, svg_path: Path):
        vsdx_path = svg_path.with_suffix(".vsdx")
        if vsdx_path.exists():
            try:
                vsdx_path.unlink()
            except PermissionError:
                raise PermissionError("Close file in Visio first.")

        with open(self.puml_path, "r", encoding="utf-8") as f:
            source_code = f.read()

        visio = None
        try:
            visio = win32com.client.DispatchEx("Visio.Application")
            visio.Visible = False
            visio.AlertResponse = 7

            doc = visio.Documents.Add("")
            page = doc.Pages(1)
            page.Name = "Sequence Diagram"
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

            src_page = doc.Pages.Add()
            src_page.PageSheet.CellsU("PageWidth").FormulaU = "8.27 in"
            src_page.PageSheet.CellsU("PageHeight").FormulaU = "11.69 in"
            src_page.Name = "PlantUML Source"

            text_box = src_page.DrawRectangle(0.5, 0.5, 7.77, 11.19)
            text_box.CellsU("LinePattern").FormulaU = "0"
            text_box.CellsU("FillPattern").FormulaU = "0"
            text_box.CellsU("Para.HorzAlign").FormulaU = "0"
            text_box.CellsU("VerticalAlign").FormulaU = "0"

            watermark = "' Generated with puml2visio, https://github.com/telekom/3gpp-meeting-tools/tree/master/puml2visio\n\n"
            text_box.Characters.Text = watermark + source_code

            try:
                if visio.ActiveWindow:
                    visio.ActiveWindow.Page = page
            except:
                pass

            doc.SaveAs(str(vsdx_path.resolve()))
            doc.Close()
            visio.Quit()

            self.ui_log_msg.emit(f"✅ Saved: {vsdx_path.name}")

        except Exception as e:
            if visio: visio.Quit()
            raise RuntimeError(f"Visio COM Error: {e}")


class SvgConverterThread(QThread):
    ui_log_msg = pyqtSignal(str)
    finished_path = pyqtSignal(str)

    def __init__(self, puml_path: Path, jar_path: Path):
        super().__init__()
        self.puml_path = puml_path
        self.jar_path = jar_path

    def run(self):
        try:
            self.ui_log_msg.emit(f"\n⚙️ Generating SVG for: {self.puml_path.name}")
            command = ["java", "-jar", str(self.jar_path), "-tsvg", str(self.puml_path)]
            subprocess.run(command, check=True, capture_output=True, text=True, cwd=self.puml_path.parent)

            svg_path = self.puml_path.with_suffix(".svg")
            if svg_path.exists():
                self.ui_log_msg.emit(f"✅ Success! SVG saved: {svg_path.name}\n{'-' * 45}")
                self.finished_path.emit(str(svg_path.resolve()))
            else:
                self.ui_log_msg.emit("❌ Error: PlantUML finished but SVG was not created.")
                self.finished_path.emit("")
        except subprocess.CalledProcessError as e:
            self.ui_log_msg.emit(f"❌ PlantUML Syntax Error:\n{e.stderr}\n{'-' * 45}")
            self.finished_path.emit("")
        except Exception as e:
            self.ui_log_msg.emit(f"❌ Error: {str(e)}\n{'-' * 45}")
            self.finished_path.emit("")