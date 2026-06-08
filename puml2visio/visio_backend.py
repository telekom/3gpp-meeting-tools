import subprocess
import urllib.request
import winreg
import re
import logging
import zlib
import base64
import zipfile
import io
from pathlib import Path

import pythoncom
import win32com.client
from PyQt5.QtCore import QThread, pyqtSignal

# ==========================================
# --- CONFIGURATION & UTILS ---
# ==========================================
JAR_NAME = "plantuml.jar"
URL_LATEST = "https://github.com/plantuml/plantuml/releases/latest/download/plantuml.jar"
URL_JAVA_8 = "https://github.com/plantuml/plantuml/releases/download/v1.2023.13/plantuml-1.2023.13.jar"


def encode_plantuml(text: str) -> str:
    """Encodes raw PlantUML text into the Deflate + Base64 format expected by PlantText."""
    compressor = zlib.compressobj(level=9, wbits=-15)
    compressed = compressor.compress(text.encode('utf-8')) + compressor.flush()
    b64 = base64.b64encode(compressed).decode('ascii')

    std_b64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
    puml_b64 = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz-_"
    trans = str.maketrans(std_b64, puml_b64)
    return b64.translate(trans).replace('=', '')


# ==========================================
# --- BACKGROUND THREADS (THE LOGIC) ---
# ==========================================

class InitializationThread(QThread):
    ui_log_msg = pyqtSignal(str)
    init_complete = pyqtSignal(bool)

    def __init__(self, jar_path: Path):
        super().__init__()
        self.jar_path = jar_path

    def run(self):
        try:
            self._emit_log("🔍 Initializing system checks...", logging.INFO)
            try:
                winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Visio.Application")
            except FileNotFoundError:
                self._emit_log("❌ ERROR: Microsoft Visio is not installed or registered.", logging.ERROR)
                self.init_complete.emit(False)
                return

            java_major = None
            try:
                result = subprocess.run(["java", "-version"], capture_output=True, text=True, check=True)
                match = re.search(r'(?:java|openjdk) version "([^"]+)"', result.stderr, re.IGNORECASE)
                if match:
                    parts = match.group(1).split('.')
                    java_major = int(parts[1]) if parts[0] == '1' else int(parts[0])
            except:
                pass

            if not java_major:
                self._emit_log("❌ ERROR: Java is not installed or not in system PATH.", logging.ERROR)
                self.init_complete.emit(False)
                return

            if not self.jar_path.exists():
                self._emit_log(f"⚠️ {JAR_NAME} missing. Attempting download...", logging.WARNING)
                url = URL_LATEST if java_major >= 11 else URL_JAVA_8
                urllib.request.urlretrieve(url, self.jar_path)
                self._emit_log("✅ PlantUML downloaded successfully.", logging.INFO)

            self.init_complete.emit(True)

        except Exception as e:
            self._emit_log(f"❌ System Check Failed: {e}", logging.CRITICAL)
            self.init_complete.emit(False)

    def _emit_log(self, message: str, level: int):
        logging.log(level, message)
        self.ui_log_msg.emit(message)


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
                        source_code = page.Shapes(1).Characters.Text
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


class WordExtractorThread(QThread):
    ui_log_msg = pyqtSignal(str)

    def __init__(self, docx_path: str):
        super().__init__()
        self.docx_path = Path(docx_path)

    def run(self):
        self.ui_log_msg.emit(f"\n📄 Analyzing Word Document: {self.docx_path.name}...")
        output_dir = self.docx_path.parent
        extracted_files = []

        try:
            with zipfile.ZipFile(self.docx_path, 'r') as z:
                direct_vsdx = [f for f in z.namelist() if f.endswith('.vsdx')]
                for f in direct_vsdx:
                    data = z.read(f)
                    out_name = output_dir / f"{self.docx_path.stem}_{Path(f).name}"
                    with open(out_name, 'wb') as out:
                        out.write(data)
                    extracted_files.append(out_name)
                    self.ui_log_msg.emit(f"✅ Extracted native Visio object: {out_name.name}")

                bins = [f for f in z.namelist() if f.startswith('word/embeddings/') and f.endswith('.bin')]
                for i, emb in enumerate(bins):
                    data = z.read(emb)
                    start_idx = data.find(b'PK\x03\x04')
                    if start_idx != -1:
                        vsdx_data = data[start_idx:]
                        try:
                            with zipfile.ZipFile(io.BytesIO(vsdx_data)) as test_z:
                                if 'visio/document.xml' in test_z.namelist() or '[Content_Types].xml' in test_z.namelist():
                                    out_name = output_dir / f"{self.docx_path.stem}_embedded_{i + 1}.vsdx"
                                    counter = 1
                                    while out_name.exists():
                                        out_name = output_dir / f"{self.docx_path.stem}_embedded_{i + 1}_{counter}.vsdx"
                                        counter += 1
                                    with open(out_name, 'wb') as out:
                                        out.write(vsdx_data)
                                    extracted_files.append(out_name)
                                    self.ui_log_msg.emit(f"✅ Extracted OLE Visio object: {out_name.name}")
                        except zipfile.BadZipFile:
                            pass

        except Exception as e:
            self.ui_log_msg.emit(f"❌ Error reading Word file: {e}")

        if not extracted_files:
            self.ui_log_msg.emit("⚠️ No embedded Visio files found in this document.")
        else:
            self.ui_log_msg.emit(
                f"🎉 Successfully extracted {len(extracted_files)} Visio file(s) to the document's folder!")


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

            # Embed source
            src_page = doc.Pages.Add()
            src_page.PageSheet.CellsU("PageWidth").FormulaU = "8.27 in"
            src_page.PageSheet.CellsU("PageHeight").FormulaU = "11.69 in"
            src_page.Name = "PlantUML Source"

            # Draw rectangle with perfect 0.5-inch margins on all sides (Centered on A4)
            text_box = src_page.DrawRectangle(0.5, 0.5, 7.77, 11.19)
            text_box.CellsU("LinePattern").FormulaU = "0"
            text_box.CellsU("FillPattern").FormulaU = "0"

            # RESTORED: Force Top-Left alignment so the code reads correctly
            text_box.CellsU("Para.HorzAlign").FormulaU = "0"
            text_box.CellsU("VerticalAlign").FormulaU = "0"

            text_box.Characters.Text = source_code

            visio.ActiveWindow.Page = page
            doc.SaveAs(str(vsdx_path.resolve()))
            doc.Close()
            visio.Quit()

            self.ui_log_msg.emit(f"✅ Saved: {vsdx_path.name}")

        except Exception as e:
            if visio: visio.Quit()
            raise RuntimeError(f"Visio COM Error: {e}")

class SvgConverterThread(QThread):
    """Background thread to strictly generate an SVG without launching Visio."""
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