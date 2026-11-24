"""Background QThread workers for Visio VSDX reading, SVG conversion, and VSDX synthesis."""

import gc
import logging
from pathlib import Path

import pythoncom
import win32com.client
from PyQt5.QtCore import QThread, pyqtSignal

from modules.puml2visio.utils.utils import generate_cleaned_svg, strip_watermark

from .svg_sanitize import sanitize_svg_for_visio
from .visio_pipeline import convert_svg_to_vsdx

logger = logging.getLogger(__name__)


class VisioReaderThread(QThread):
    """Extracts embedded PlantUML source text from an existing `.vsdx` file."""

    text_extracted = pyqtSignal(str)
    error_occurred = pyqtSignal(str)

    def __init__(self, vsdx_path):
        super().__init__()
        self.vsdx_path = vsdx_path

    def run(self):
        pythoncom.CoInitialize()
        visio = None
        doc = None
        page = None
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
                        source_code = strip_watermark(raw_text)
                    break

            page = None
            doc.Close()
            doc = None
            visio.Quit()
            visio = None

            if source_code:
                self.text_extracted.emit(source_code)
            else:
                self.error_occurred.emit("Could not find 'PlantUML Source' page in this Visio file.")
        except Exception as e:
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
            self.error_occurred.emit(f"Error reading Visio file: {str(e)}")
        finally:
            _ = page
            gc.collect()
            pythoncom.CoUninitialize()


class ConverterThread(QThread):
    """Converts a `.puml` diagram to `.vsdx` via PlantUML SVG generation and Visio COM."""

    ui_log_msg = pyqtSignal(str)
    finished_path = pyqtSignal(str)

    def __init__(self, puml_path: Path, jar_path: Path):
        super().__init__()
        self.puml_path = puml_path
        self.jar_path = jar_path

    def run(self):
        pythoncom.CoInitialize()
        try:
            self._emit_log(f"⚙️ Processing: {self.puml_path.name}", logging.INFO)

            svg_path = generate_cleaned_svg(self.puml_path, self.jar_path, self._emit_log)
            self._sanitize_svg_for_visio(svg_path)
            self._convert_to_vsdx(svg_path)

            vsdx_path = self.puml_path.with_suffix(".vsdx")
            self.finished_path.emit(str(vsdx_path.resolve()))

            if svg_path.exists():
                svg_path.unlink()

        except Exception as e:
            self._emit_log(f"❌ Error: {str(e)}\n{'-' * 45}", logging.ERROR)
            self.finished_path.emit("")
        finally:
            gc.collect()
            pythoncom.CoUninitialize()

    def _sanitize_svg_for_visio(self, svg_path: Path) -> None:
        sanitize_svg_for_visio(svg_path, self._emit_log)

    def _emit_log(self, message: str, level: int = logging.INFO):
        logger.log(level, message)

    def _convert_to_vsdx(self, svg_path: Path):
        convert_svg_to_vsdx(self.puml_path, svg_path, self._emit_log)


class SvgConverterThread(QThread):
    """Converts a `.puml` diagram to a cleaned `.svg` without invoking Visio."""

    ui_log_msg = pyqtSignal(str)
    finished_path = pyqtSignal(str)

    def __init__(self, puml_path: Path, jar_path: Path):
        super().__init__()
        self.puml_path = puml_path
        self.jar_path = jar_path

    def _emit_log(self, message: str, level: int = logging.INFO):
        logger.log(level, message)

    def run(self):
        try:
            self._emit_log(f"⚙️ Generating SVG for: {self.puml_path.name}")

            svg_path = generate_cleaned_svg(self.puml_path, self.jar_path, self._emit_log)

            if svg_path.exists():
                self._emit_log(f"✅ Success! SVG saved: {svg_path.name}\n{'-' * 45}")
                self.finished_path.emit(str(svg_path.resolve()))
            else:
                self._emit_log("❌ Error: PlantUML finished but SVG was not created.", logging.ERROR)
                self.finished_path.emit("")
        except Exception as e:
            self._emit_log(f"❌ Error: {str(e)}\n{'-' * 45}", logging.ERROR)
            self.finished_path.emit("")
