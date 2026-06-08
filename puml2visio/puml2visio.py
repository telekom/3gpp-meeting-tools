import sys
import subprocess
import urllib.request
import winreg
import re
import logging
import datetime
import zlib
import base64
import webbrowser
from pathlib import Path

# Third-party imports (from requirements.txt)
import pythoncom
import win32com.client
from PyQt5.QtWidgets import (QApplication, QMainWindow, QLabel, QVBoxLayout,
                             QWidget, QTextEdit, QDialog, QLineEdit, QPushButton,
                             QFormLayout, QHBoxLayout, QTabWidget, QCheckBox)
from PyQt5.QtCore import Qt, QThread, pyqtSignal

# ==========================================
# --- CONFIGURATION ---
# ==========================================
JAR_NAME = "plantuml.jar"
URL_LATEST = "https://github.com/plantuml/plantuml/releases/latest/download/plantuml.jar"
URL_JAVA_8 = "https://github.com/plantuml/plantuml/releases/download/v1.2023.13/plantuml-1.2023.13.jar"

# ==========================================
# --- LOGGING SETUP ---
# ==========================================
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler("puml2vsdx.log", mode="a", encoding="utf-8"),
        logging.StreamHandler(sys.stdout)
    ]
)


class ProxyDialog(QDialog):
    """Startup dialog to securely capture proxy settings without hardcoding them."""

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Network Configuration")
        self.setModal(True)
        self.resize(500, 220)

        layout = QVBoxLayout()

        title = QLabel("📡 Proxy Configuration")
        title.setStyleSheet("font-size: 16px; font-weight: bold; margin-bottom: 5px;")
        layout.addWidget(title)

        desc = QLabel(
            "To prevent credential leaks in repositories, please enter your proxy details here "
            "instead of hardcoding them in the script.\n\nLeave blank to connect directly."
        )
        desc.setWordWrap(True)
        desc.setStyleSheet("color: #555; margin-bottom: 15px;")
        layout.addWidget(desc)

        form = QFormLayout()

        self.http_input = QLineEdit()
        self.http_input.setPlaceholderText("e.g., http://user:pass@proxy.company.com:8080")
        self.https_input = QLineEdit()
        self.https_input.setPlaceholderText("e.g., http://user:pass@proxy.company.com:8080")

        self.sync_checkbox = QCheckBox("Use the same proxy for HTTPS")
        self.sync_checkbox.setChecked(False)
        self.sync_checkbox.stateChanged.connect(self.on_sync_changed)
        self.http_input.textChanged.connect(self.on_http_changed)

        form.addRow("HTTP Proxy:", self.http_input)
        form.addRow("", self.sync_checkbox)
        form.addRow("HTTPS Proxy:", self.https_input)
        layout.addLayout(form)

        btn_layout = QHBoxLayout()
        self.skip_btn = QPushButton("Continue Without Proxy")
        self.skip_btn.setCursor(Qt.PointingHandCursor)
        self.skip_btn.clicked.connect(self.skip)

        self.save_btn = QPushButton("Set Proxy && Continue")
        self.save_btn.setStyleSheet("background-color: #0078d7; color: white; font-weight: bold;")
        self.save_btn.setCursor(Qt.PointingHandCursor)
        self.save_btn.clicked.connect(self.accept)

        btn_layout.addWidget(self.skip_btn)
        btn_layout.addStretch()
        btn_layout.addWidget(self.save_btn)

        layout.addSpacing(10)
        layout.addLayout(btn_layout)
        self.setLayout(layout)

    def on_sync_changed(self, state):
        if state == Qt.Checked:
            self.https_input.setEnabled(False)
            self.https_input.setText(self.http_input.text())
        else:
            self.https_input.setEnabled(True)

    def on_http_changed(self, text):
        if self.sync_checkbox.isChecked():
            self.https_input.setText(text)

    def skip(self):
        self.http_input.clear()
        self.https_input.clear()
        self.accept()

    def get_proxies(self):
        return self.http_input.text().strip(), self.https_input.text().strip()


class InitializationThread(QThread):
    """Background thread to configure environment and download dependencies."""
    ui_log_msg = pyqtSignal(str)
    init_complete = pyqtSignal(bool)

    def __init__(self, jar_path: Path):
        super().__init__()
        self.jar_path = jar_path

    def run(self):
        try:
            self._emit_log("🔍 Initializing system checks...", level=logging.INFO)

            if not self._check_visio_installed():
                self._emit_log("❌ ERROR: Microsoft Visio is not installed or registered.", level=logging.ERROR)
                self.init_complete.emit(False)
                return
            self._emit_log("✅ Microsoft Visio detected.", level=logging.INFO)

            java_major = self._get_java_version()
            if not java_major:
                self._emit_log("❌ ERROR: Java is not installed or not in system PATH.", level=logging.ERROR)
                self.init_complete.emit(False)
                return
            self._emit_log(f"✅ Java {java_major} detected.", level=logging.INFO)

            if not self.jar_path.exists():
                self._emit_log(f"⚠️ {JAR_NAME} missing. Attempting download...", level=logging.WARNING)
                self._download_plantuml(java_major)
                self._emit_log("✅ PlantUML downloaded successfully.", level=logging.INFO)
            else:
                self._emit_log("✅ PlantUML jar found locally.", level=logging.INFO)

            self.init_complete.emit(True)

        except Exception as e:
            self._emit_log(f"❌ System Check Failed: {e}", level=logging.CRITICAL)
            self.init_complete.emit(False)

    def _emit_log(self, message: str, level: int):
        logging.log(level, message)
        self.ui_log_msg.emit(message)

    def _check_visio_installed(self) -> bool:
        try:
            winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Visio.Application")
            return True
        except FileNotFoundError:
            return False

    def _get_java_version(self):
        try:
            result = subprocess.run(["java", "-version"], capture_output=True, text=True, check=True)
            output = result.stderr
            match = re.search(r'(?:java|openjdk) version "([^"]+)"', output, re.IGNORECASE)
            if match:
                ver_str = match.group(1)
                parts = ver_str.split('.')
                return int(parts[1]) if parts[0] == '1' else int(parts[0])
            return None
        except (subprocess.CalledProcessError, FileNotFoundError):
            return None

    def _download_plantuml(self, java_major: int):
        url = URL_LATEST if java_major >= 11 else URL_JAVA_8
        self._emit_log(f"   ↳ Fetching from: {url}", level=logging.INFO)
        # Uses the global urllib proxy opener installed in __main__
        urllib.request.urlretrieve(url, self.jar_path)


class ConverterThread(QThread):
    """Background thread to process the text-to-Visio conversion safely."""
    ui_log_msg = pyqtSignal(str)
    finished_path = pyqtSignal(str)

    def __init__(self, puml_path: Path, jar_path: Path):
        super().__init__()
        self.puml_path = puml_path
        self.jar_path = jar_path

    def run(self):
        pythoncom.CoInitialize()
        try:
            self._emit_log(f"\n⚙️ Processing: {self.puml_path.name}", level=logging.INFO)

            self._emit_log("⏳ Calling PlantUML to generate SVG...", level=logging.INFO)
            svg_path = self._generate_svg()

            self._emit_log("⏳ Launching Visio to generate .vsdx and embed source...", level=logging.INFO)
            self._convert_to_vsdx(svg_path)

            self._emit_log(f"✅ Success! Diagram saved.\n{'-' * 45}", level=logging.INFO)

            vsdx_path = self.puml_path.with_suffix(".vsdx")
            self.finished_path.emit(str(vsdx_path.resolve()))

            if svg_path.exists():
                svg_path.unlink()

        except Exception as e:
            self._emit_log(f"❌ Error during {self.puml_path.name}: {str(e)}\n{'-' * 45}", level=logging.ERROR)
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

            # --- UNIVERSAL CENTERING LOGIC ---
            if page.Shapes.Count > 0:
                shape = page.Shapes(1)
                page_sheet = page.PageSheet

                # Get dimensions using safer ResultIU (Internal Units - Inches)
                page_w = page_sheet.CellsU("PageWidth").ResultIU
                page_h = page_sheet.CellsU("PageHeight").ResultIU

                # PinX/PinY refer to the center of the shape
                shape.CellsU("PinX").FormulaU = f"{page_w / 2}"
                shape.CellsU("PinY").FormulaU = f"{page_h / 2}"

                # Try to fit contents safely
                try:
                    page.ResizeToFitContents()
                except:
                    pass

            # Embed source
            src_page = doc.Pages.Add()
            src_page.PageSheet.CellsU("PageWidth").FormulaU = "8.27 in"
            src_page.PageSheet.CellsU("PageHeight").FormulaU = "11.69 in"
            src_page.Name = "PlantUML Source"

            text_box = src_page.DrawRectangle(0.5, 0.5, 8.0, 10.5)
            text_box.CellsU("LinePattern").FormulaU = "0"
            text_box.CellsU("FillPattern").FormulaU = "0"
            text_box.CellsU("Para.HorzAlign").FormulaU = "0"
            text_box.CellsU("VerticalAlign").FormulaU = "0"

            text_box.Characters.Text = source_code

            # Restore active view to Page 1 for flawless OLE previews in Word
            visio.ActiveWindow.Page = page

            doc.SaveAs(str(vsdx_path.resolve()))
            doc.Close()
            visio.Quit()

            if svg_path.exists(): svg_path.unlink()
            self.ui_log_msg.emit(f"✅ Saved: {vsdx_path.name}")

        except Exception as e:
            if visio: visio.Quit()
            raise RuntimeError(f"Visio COM Error: {e}")


class DragDropUI(QMainWindow):
    """Main PyQt5 GUI Window with Tabs for File Drops and Code Pasting."""

    def __init__(self):
        super().__init__()
        self.setWindowTitle("PlantUML to Visio Converter (3GPP)")
        self.resize(750, 650)
        self.setAcceptDrops(True)

        self.jar_path = Path(__file__).parent.resolve() / JAR_NAME

        self.file_queue = []
        self.is_processing = False
        self.last_visio_path = ""

        self._setup_ui()

        self.init_thread = InitializationThread(self.jar_path)
        self.init_thread.ui_log_msg.connect(self.log_message)
        self.init_thread.init_complete.connect(self.on_init_complete)
        self.init_thread.start()

    def _setup_ui(self):
        main_layout = QVBoxLayout()
        self.tabs = QTabWidget()

        # ==========================================
        # Tab 1: Code Paste Area (MOVED TO