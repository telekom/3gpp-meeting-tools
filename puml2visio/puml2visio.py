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

        # --- SVG PRE-PROCESSING: THE ULTIMATE TEXT MERGER ---
        try:
            with open(svg_path, 'r', encoding='utf-8') as f:
                svg_content = f.read()

            # 1. Strip PlantUML's length attributes that confuse Visio
            svg_content = re.sub(r'\s*textLength="[^"]*"', '', svg_content)
            svg_content = re.sub(r'\s*lengthAdjust="[^"]*"', '', svg_content)

            # 2. Merge adjacent <text> tags that share the exact same 'y' coordinate
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

                    # Check what XML exists between the last text tag and this one
                    between = svg_content[last_end:start]

                    # If they are on the same Y line AND there are no structural XML tags between them
                    if current_y == y_val and not between.strip():
                        current_text += inner_text
                    else:
                        # Flush the previous group to the result
                        if current_y is not None:
                            result.append(current_start_tag)
                            result.append(current_text)
                            result.append("</text>")

                        result.append(between)

                        # Start a new group
                        current_y = y_val
                        current_start_tag = full_open_tag
                        current_text = inner_text

                    last_end = end

                # Flush the final group
                if current_y is not None:
                    result.append(current_start_tag)
                    result.append(current_text)
                    result.append("</text>")

                result.append(svg_content[last_end:])
                svg_content = "".join(result)

            # 3. Replace &#160; and \xa0 with standard spaces so Visio renders them with proper widths
            svg_content = svg_content.replace('&#160;', ' ').replace('\xa0', ' ')

            with open(svg_path, 'w', encoding='utf-8') as f:
                f.write(svg_content)

        except Exception as e:
            self._emit_log(f"⚠️ Warning: Could not clean SVG text attributes: {e}", level=logging.WARNING)

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

            # --- INTELLIGENT UNGROUP & CLEANUP LOGIC ---
            if page.Shapes.Count > 0:
                orig_w = page.Shapes(1).CellsU("Width").ResultIU
                orig_h = page.Shapes(1).CellsU("Height").ResultIU

                # 1. Peel ONLY structural wrappers
                peeling = True
                while peeling:
                    peeling = False
                    for i in range(page.Shapes.Count, 0, -1):
                        s = page.Shapes(i)
                        try:
                            w = s.CellsU("Width").ResultIU
                            h = s.CellsU("Height").ResultIU

                            if abs(w - orig_w) < 0.1 and abs(h - orig_h) < 0.1:
                                if s.Type == 2:  # visTypeGroup
                                    s.Ungroup()
                                    peeling = True
                        except:
                            pass

                # 2. Destroy the invisible PlantUML background rectangle
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

                # 3. Recursively strip margins AND perfectly shrink text boxes
                def clean_and_shrink_text(shapes):
                    for i in range(1, shapes.Count + 1):
                        s = shapes(i)
                        try:
                            # Read text safely
                            txt = s.Characters.Text
                            if len(txt.strip()) > 0:

                                # Strip Visio's default padding
                                s.CellsU("TopMargin").FormulaU = "0 pt"
                                s.CellsU("BottomMargin").FormulaU = "0 pt"
                                s.CellsU("LeftMargin").FormulaU = "0 pt"
                                s.CellsU("RightMargin").FormulaU = "0 pt"

                                # Check if this is an invisible SVG text container
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

                # 4. Snap the Page Canvas tightly around the final layout
                page_sheet = page.PageSheet
                page_sheet.CellsU("PageLeftMargin").FormulaU = "0.05 in"
                page_sheet.CellsU("PageRightMargin").FormulaU = "0.05 in"
                page_sheet.CellsU("PageTopMargin").FormulaU = "0.05 in"
                page_sheet.CellsU("PageBottomMargin").FormulaU = "0.05 in"
                try:
                    page.ResizeToFitContents()
                except:
                    pass

            # Embed source code
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

            self.ui_log_msg.emit(f"✅ Saved: {vsdx_path.name}")

        except Exception as e:
            if visio: visio.Quit()
            raise RuntimeError(f"Visio COM Error: {e}")


class VisioReaderThread(QThread):
    """Background thread to extract embedded PlantUML code from a Visio file."""
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

            # OpenEx with flag '2' = Read-Only. Prevents file lock crashes if Visio is already open.
            doc = visio.Documents.OpenEx(str(Path(self.vsdx_path).resolve()), 2)

            source_code = ""
            # Hunt for the embedded source page
            for i in range(1, doc.Pages.Count + 1):
                page = doc.Pages(i)
                if page.Name == "PlantUML Source":
                    if page.Shapes.Count > 0:
                        # Grab the raw text from the bounding box
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


class CodeDropTextEdit(QTextEdit):
    """Custom TextEdit that intercepts Visio file drops while allowing normal text."""
    file_dropped = pyqtSignal(str)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            if any(url.toLocalFile().lower().endswith('.vsdx') for url in urls):
                event.acceptProposedAction()
                return
        # Fall back to standard text drag-and-drop
        super().dragEnterEvent(event)

    def dropEvent(self, event):
        if event.mimeData().hasUrls():
            for url in event.mimeData().urls():
                file_path = url.toLocalFile()
                if file_path.lower().endswith('.vsdx'):
                    self.file_dropped.emit(file_path)
                    event.acceptProposedAction()
                    return  # Only read the first Visio file dropped
        # Fall back to standard text drop
        super().dropEvent(event)

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
        # Tab 1: Code Paste Area (MOVED TO FIRST)
        # ==========================================
        self.tab_text = QWidget()
        tab_text_layout = QVBoxLayout()
        self.text_input = CodeDropTextEdit()
        self.text_input.setPlaceholderText(
            "Paste PlantUML code OR drop a generated .vsdx file here to extract its source...")
        self.text_input.file_dropped.connect(self.extract_code_from_visio)
        self.text_input.setStyleSheet(
            "font-family: Consolas, Courier New, monospace; font-size: 13px; border: 1px solid #ccc;")

        btn_layout = QHBoxLayout()

        self.clear_btn = QPushButton("🗑️ Clear")
        self.clear_btn.setStyleSheet(
            "background-color: #777; color: white; font-weight: bold; padding: 10px; font-size: 13px; border-radius: 5px;")
        self.clear_btn.setCursor(Qt.PointingHandCursor)
        self.clear_btn.clicked.connect(self.text_input.clear)

        self.planttext_btn = QPushButton("🌐 Show in planttext.com")
        self.planttext_btn.setStyleSheet(
            "background-color: #0078d7; color: white; font-weight: bold; padding: 10px; font-size: 13px; border-radius: 5px;")
        self.planttext_btn.setCursor(Qt.PointingHandCursor)
        self.planttext_btn.clicked.connect(self.show_in_planttext)

        self.copy_btn = QPushButton("📋 Copy Visio Path")
        self.copy_btn.setStyleSheet(
            "background-color: #ccc; color: #666; font-weight: bold; padding: 10px; font-size: 13px; border-radius: 5px;")
        self.copy_btn.setEnabled(False)
        self.copy_btn.setCursor(Qt.PointingHandCursor)
        self.copy_btn.clicked.connect(self.copy_visio_path)

        self.convert_btn = QPushButton("Convert to Visio")
        self.convert_btn.setStyleSheet(
            "background-color: #4af626; color: #000; font-weight: bold; padding: 10px; font-size: 13px; border-radius: 5px;")
        self.convert_btn.setCursor(Qt.PointingHandCursor)
        self.convert_btn.clicked.connect(self.convert_pasted_text)

        btn_layout.addWidget(self.clear_btn)
        btn_layout.addWidget(self.planttext_btn)
        btn_layout.addWidget(self.copy_btn)
        btn_layout.addWidget(self.convert_btn)

        tab_text_layout.addWidget(self.text_input)
        tab_text_layout.addLayout(btn_layout)
        self.tab_text.setLayout(tab_text_layout)

        # ==========================================
        # Tab 2: Drag & Drop Area
        # ==========================================
        self.tab_file = QWidget()
        tab_file_layout = QVBoxLayout()
        self.drop_label = QLabel("⏳ Initializing system checks... Please wait.")
        self.drop_label.setAlignment(Qt.AlignCenter)
        self.drop_label.setStyleSheet("""
            QLabel {
                border: 3px dashed #888;
                border-radius: 10px;
                background-color: #e0e0e0;
                font-size: 15px;
                font-weight: bold;
                color: #555;
                padding: 20px;
            }
        """)
        tab_file_layout.addWidget(self.drop_label)
        self.tab_file.setLayout(tab_file_layout)

        # Add Tabs (Paste Code is first, so it opens by default)
        self.tabs.addTab(self.tab_text, "📝 Paste Code")
        self.tabs.addTab(self.tab_file, "📂 Drag && Drop Files")
        self.tabs.setEnabled(False)

        main_layout.addWidget(self.tabs, stretch=1)

        # --- CONSOLE AREA ---
        self.console = QTextEdit()
        self.console.setReadOnly(True)
        self.console.setStyleSheet("""
            QTextEdit {
                background-color: #1e1e1e; 
                color: #4af626; 
                font-family: Consolas, Courier New, monospace;
                font-size: 13px;
                border-radius: 5px;
            }
        """)
        main_layout.addWidget(self.console, stretch=1)

        container = QWidget()
        container.setLayout(main_layout)
        self.setCentralWidget(container)

    def on_init_complete(self, success: bool):
        if success:
            self.tabs.setEnabled(True)
            self._set_drop_zone_ready()
            self.log_message("\n🚀 System Ready. Paste code or drop files to begin.\n" + "-" * 45)
        else:
            self.drop_label.setText("❌ Initialization Failed.\nPlease fix the errors below and restart.")
            self.drop_label.setStyleSheet("border: 3px solid #f64a4a; background-color: #ffeaea; color: #d32f2f;")

    def _set_drop_zone_ready(self):
        self.drop_label.setText("📥 Drag & Drop your .puml or .txt file(s) here\n\n(Batch processing supported)")
        self.drop_label.setStyleSheet(
            "border: 3px dashed #4af626; background-color: #f4f4f4; font-size: 15px; font-weight: bold; color: #333;")

        self.convert_btn.setEnabled(True)
        self.convert_btn.setText("Convert to Visio")
        self.convert_btn.setStyleSheet(
            "background-color: #4af626; color: #000; font-weight: bold; padding: 10px; font-size: 13px; border-radius: 5px;")

        self.clear_btn.setEnabled(True)
        self.planttext_btn.setEnabled(True)

    def _set_drop_zone_busy(self):
        self.drop_label.setText("⚙️ Processing Queue...\n\nPlease wait until finished.")
        self.drop_label.setStyleSheet(
            "border: 3px dashed #f6a826; background-color: #fff4e5; font-size: 15px; font-weight: bold; color: #d37e00;")

        self.convert_btn.setEnabled(False)
        self.convert_btn.setText("Processing...")
        self.convert_btn.setStyleSheet(
            "background-color: #ccc; color: #666; font-weight: bold; padding: 10px; font-size: 13px; border-radius: 5px;")

        self.clear_btn.setEnabled(False)
        self.planttext_btn.setEnabled(False)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event):
        urls = event.mimeData().urls()
        if not urls:
            return

        added_count = 0
        for url in urls:
            file_path = Path(url.toLocalFile())
            if file_path.suffix.lower() in [".puml", ".txt"]:
                self.file_queue.append(file_path)
                added_count += 1
            else:
                self.log_message(f"⚠️ Skipped invalid file type: {file_path.name}")

        if added_count > 0:
            self.log_message(f"📥 Added {added_count} file(s) to the queue. Total pending: {len(self.file_queue)}")
            if not self.is_processing:
                self.process_next_in_queue()

    def show_in_planttext(self):
        raw_text = self.text_input.toPlainText().strip()
        if not raw_text:
            self.log_message("⚠️ Error: The text box is empty. Paste some code first.")
            return

        try:
            # PlantUML Deflate + Base64 encoding algorithm
            compressor = zlib.compressobj(level=9, wbits=-15)
            compressed = compressor.compress(raw_text.encode('utf-8')) + compressor.flush()
            b64 = base64.b64encode(compressed).decode('ascii')

            std_b64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
            puml_b64 = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz-_"
            trans = str.maketrans(std_b64, puml_b64)
            encoded = b64.translate(trans).replace('=', '')

            url = f"https://www.planttext.com/?text={encoded}"
            webbrowser.open(url)
            self.log_message("🌐 Opened code in planttext.com")
        except Exception as e:
            self.log_message(f"❌ Failed to open in PlantText: {e}")

    def copy_visio_path(self):
        if hasattr(self, 'last_visio_path') and self.last_visio_path:
            QApplication.clipboard().setText(self.last_visio_path)
            self.log_message(f"📋 Copied to clipboard: {self.last_visio_path}")

    def extract_code_from_visio(self, file_path):
        self.text_input.clear()
        self.text_input.setPlaceholderText(f"⏳ Extracting source from {Path(file_path).name}...\nPlease wait...")
        self.text_input.setEnabled(False)
        self.log_message(f"📂 Reading embedded source from: {Path(file_path).name}")

        self.reader_thread = VisioReaderThread(file_path)
        self.reader_thread.text_extracted.connect(self.on_visio_code_read)
        self.reader_thread.error_occurred.connect(self.on_visio_code_error)
        self.reader_thread.start()

    def on_visio_code_read(self, source_code):
        self.text_input.setEnabled(True)
        self.text_input.setPlainText(source_code)
        self.log_message("✅ Successfully extracted PlantUML source from Visio file.")

    def on_visio_code_error(self, error_msg):
        self.text_input.setEnabled(True)
        self.text_input.setPlaceholderText(
            "Paste PlantUML code OR drop a generated .vsdx file here to extract its source...")
        self.log_message(f"❌ {error_msg}")

    def convert_pasted_text(self):
        raw_text = self.text_input.toPlainText().strip()
        if not raw_text:
            self.log_message("❌ Error: The text box is empty.")
            return

        if "@startuml" not in raw_text:
            self.log_message("⚠️ Warning: Could not find '@startuml'. PlantUML may fail, but proceeding anyway.")

        base_dir = Path(__file__).parent.resolve()
        timestamp = datetime.datetime.now().strftime("%Y.%m.%d %H-%M-%S")
        base_name = f"{timestamp} diagram"
        puml_path = base_dir / f"{base_name}.puml"

        counter = 1
        while puml_path.exists() or puml_path.with_suffix(".vsdx").exists():
            puml_path = base_dir / f"{base_name}_{counter}.puml"
            counter += 1

        try:
            with open(puml_path, "w", encoding="utf-8") as f:
                f.write(raw_text)
            self.log_message(f"💾 Captured text saved to: {puml_path.name}")
        except Exception as e:
            self.log_message(f"❌ Failed to save pasted text: {e}")
            return

        self.file_queue.append(puml_path)
        if not self.is_processing:
            self.process_next_in_queue()

    def process_next_in_queue(self):
        if not self.file_queue:
            self.is_processing = False
            self.log_message("🏁 All batch processing complete! Waiting for new files.\n" + "-" * 45)
            self._set_drop_zone_ready()
            return

        self.is_processing = True
        self._set_drop_zone_busy()

        next_file = self.file_queue.pop(0)

        self.conv_thread = ConverterThread(next_file, self.jar_path)
        self.conv_thread.ui_log_msg.connect(self.log_message)

        # Connect the successful path signal to update our UI copy button
        self.conv_thread.finished_path.connect(self.on_conversion_success)

        # Trigger the next item in the queue when the thread completely finishes
        self.conv_thread.finished.connect(self.process_next_in_queue)

        self.conv_thread.start()

    def on_conversion_success(self, vsdx_path: str):
        if vsdx_path:
            self.last_visio_path = vsdx_path
            self.copy_btn.setEnabled(True)
            self.copy_btn.setStyleSheet(
                "background-color: #0078d7; color: white; font-weight: bold; padding: 10px; font-size: 13px; border-radius: 5px;")

    def log_message(self, message: str):
        self.console.append(message)
        QApplication.processEvents()
        scrollbar = self.console.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())


if __name__ == '__main__':
    logging.info("Starting PlantUML to Visio Converter application...")
    app = QApplication(sys.argv)
    app.setStyle("Fusion")

    jar_path = Path(__file__).parent.resolve() / JAR_NAME

    # Requirement: Only trigger proxy dialog if an HTTP connection is needed (jar missing)
    if not jar_path.exists():
        proxy_dialog = ProxyDialog()
        if proxy_dialog.exec_() == QDialog.Accepted:
            http_val, https_val = proxy_dialog.get_proxies()

            proxies = {}
            if http_val: proxies['http'] = http_val
            if https_val: proxies['https'] = https_val

            # Requirement: Install the proxy opener globally to catch all urllib requests
            if proxies:
                proxy_handler = urllib.request.ProxyHandler(proxies)
                opener = urllib.request.build_opener(proxy_handler)
                urllib.request.install_opener(opener)
                logging.info("Global proxy installed for all HTTP/HTTPS requests.")

    window = DragDropUI()
    window.show()
    sys.exit(app.exec_())