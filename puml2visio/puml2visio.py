import sys
import subprocess
import urllib.request
import winreg
import re
import logging
import datetime
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

    def __init__(self, jar_path: Path, http_proxy: str, https_proxy: str):
        super().__init__()
        self.jar_path = jar_path
        self.http_proxy = http_proxy
        self.https_proxy = https_proxy

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
        
        proxies = {}
        if self.http_proxy: proxies['http'] = self.http_proxy
        if self.https_proxy: proxies['https'] = self.https_proxy

        if proxies:
            self._emit_log(f"   ↳ Using proxy configuration...", level=logging.INFO)
            proxy_handler = urllib.request.ProxyHandler(proxies)
            opener = urllib.request.build_opener(proxy_handler)
            urllib.request.install_opener(opener)

        urllib.request.urlretrieve(url, self.jar_path)


class ConverterThread(QThread):
    """Background thread to process the text-to-Visio conversion safely."""
    ui_log_msg = pyqtSignal(str)
    
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
            
            self._emit_log(f"✅ Success! Diagram saved.\n{'-'*45}", level=logging.INFO)
            
            if svg_path.exists():
                svg_path.unlink()
                
        except Exception as e:
            self._emit_log(f"❌ Error during {self.puml_path.name}: {str(e)}\n{'-'*45}", level=logging.ERROR)
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
                raise PermissionError(f"Cannot overwrite '{vsdx_path.name}'. Close it in Visio first.")
        
        # 1. Parse the pure mathematical dimensions from the PlantUML SVG file
        svg_w_px, svg_h_px = None, None
        try:
            with open(svg_path, 'r', encoding='utf-8') as f:
                svg_content = f.read()
                # Find the viewBox (e.g., viewBox="0 0 540 820")
                match = re.search(r'viewBox=["\']0 0 ([\d\.]+) ([\d\.]+)["\']', svg_content)
                if match:
                    svg_w_px = float(match.group(1))
                    svg_h_px = float(match.group(2))
        except Exception as e:
            self._emit_log(f"⚠️ Warning: Could not parse SVG dimensions: {e}", level=logging.WARNING)

        # 2. Extract original source code for embedding
        with open(self.puml_path, "r", encoding="utf-8") as f:
            source_code = f.read()
            
        visio = None
        try:
            visio = win32com.client.DispatchEx("Visio.Application")
            visio.Visible = False
            visio.AlertResponse = 7 
            
            doc = visio.Documents.Add("")
            
            # --- Page 1: Diagram Image ---
            page1 = doc.Pages(1)
            page1.Name = "Sequence Diagram"
            
            # Remove all invisible margins
            page_sheet = page1.PageSheet
            page_sheet.CellsU("PageLeftMargin").FormulaU = "0 in"
            page_sheet.CellsU("PageRightMargin").FormulaU = "0 in"
            page_sheet.CellsU("PageTopMargin").FormulaU = "0 in"
            page_sheet.CellsU("PageBottomMargin").FormulaU = "0 in"
            
            # Import the SVG
            page1.Import(str(svg_path.resolve()))
            
            # NEW LOGIC: Perfect Crop
            if page1.Shapes.Count > 0:
                imported_shape = page1.Shapes(1)
                
                if svg_w_px and svg_h_px:
                    # Force the Visio page to match PlantUML's exact dimensions
                    # This prevents Visio's text bounding box engine from blowing up the canvas
                    page_sheet.CellsU("PageWidth").FormulaU = f"{svg_w_px} px"
                    page_sheet.CellsU("PageHeight").FormulaU = f"{svg_h_px} px"
                    
                    # Snap the diagram shape perfectly to the center of our new page
                    imported_shape.CellsU("PinX").FormulaU = "ThePage!PageWidth * 0.5"
                    imported_shape.CellsU("PinY").FormulaU = "ThePage!PageHeight * 0.5"
                else:
                    # Fallback if the regex fails to find the viewBox
                    page1.ResizeToFitContents()
            
            # --- Page 2: Source Code Embedding ---
            page2 = doc.Pages.Add()
            page2.Name = "PlantUML Source"
            
            text_box = page2.DrawRectangle(0.5, 0.5, 8.0, 10.5)
            text_box.CellsU("LinePattern").FormulaU = "0"
            text_box.CellsU("FillPattern").FormulaU = "0"
            text_box.CellsU("Para.HorzAlign").FormulaU = "0"
            text_box.CellsU("VerticalAlign").FormulaU = "0"
            
            text_box.Characters.Text = source_code
            
            # Restore active view to Page 1 for flawless OLE previews in Word
            visio.ActiveWindow.Page = page1
            doc.SaveAs(str(vsdx_path.resolve()))
            doc.Close()
            
        except Exception as e:
            raise RuntimeError(f"Visio COM Error: {e}")
        finally:
            if visio:
                visio.Quit()


class DragDropUI(QMainWindow):
    """Main PyQt5 GUI Window with Tabs for File Drops and Code Pasting."""
    def __init__(self, http_proxy: str, https_proxy: str):
        super().__init__()
        self.setWindowTitle("PlantUML to Visio Converter (3GPP)")
        self.resize(650, 600)
        self.setAcceptDrops(True)
        
        self.jar_path = Path(__file__).parent.resolve() / JAR_NAME
        
        self.file_queue = []
        self.is_processing = False
        
        self._setup_ui()
        
        self.init_thread = InitializationThread(self.jar_path, http_proxy, https_proxy)
        self.init_thread.ui_log_msg.connect(self.log_message)
        self.init_thread.init_complete.connect(self.on_init_complete)
        self.init_thread.start()

    def _setup_ui(self):
        main_layout = QVBoxLayout()
        
        # --- TAB WIDGET ---
        self.tabs = QTabWidget()
        
        # Tab 1: Drag & Drop Area
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
        
        # Tab 2: Code Paste Area
        self.tab_text = QWidget()
        tab_text_layout = QVBoxLayout()
        self.text_input = QTextEdit()
        self.text_input.setPlaceholderText("Paste your PlantUML code here (ensure it starts with @startuml)...\n\nA file named 'YYYY.MM.DD hh-mm-ss diagram.vsdx' will be created in this tool's folder.")
        self.text_input.setStyleSheet("font-family: Consolas, Courier New, monospace; font-size: 13px; border: 1px solid #ccc;")
        
        btn_layout = QHBoxLayout()
        
        self.clear_btn = QPushButton("🗑️ Clear Text")
        self.clear_btn.setStyleSheet("background-color: #777; color: white; font-weight: bold; padding: 10px; font-size: 14px; border-radius: 5px;")
        self.clear_btn.setCursor(Qt.PointingHandCursor)
        self.clear_btn.clicked.connect(self.text_input.clear)
        
        self.convert_btn = QPushButton("Convert to Visio")
        self.convert_btn.setStyleSheet("background-color: #4af626; color: #000; font-weight: bold; padding: 10px; font-size: 14px; border-radius: 5px;")
        self.convert_btn.setCursor(Qt.PointingHandCursor)
        self.convert_btn.clicked.connect(self.convert_pasted_text)
        
        btn_layout.addWidget(self.clear_btn)
        btn_layout.addWidget(self.convert_btn)
        
        tab_text_layout.addWidget(self.text_input)
        tab_text_layout.addLayout(btn_layout)
        self.tab_text.setLayout(tab_text_layout)
        
        self.tabs.addTab(self.tab_file, "📂 Drag & Drop Files")
        self.tabs.addTab(self.tab_text, "📝 Paste Code")
        self.tabs.setEnabled(False)
        
        # Make "Paste Code" the default active tab
        self.tabs.setCurrentIndex(1)
        
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
            self.log_message("\n🚀 System Ready. Paste code or drop files to begin.\n" + "-"*45)
        else:
            self.drop_label.setText("❌ Initialization Failed.\nPlease fix the errors below and restart.")
            self.drop_label.setStyleSheet("border: 3px solid #f64a4a; background-color: #ffeaea; color: #d32f2f;")

    def _set_drop_zone_ready(self):
        self.drop_label.setText("📥 Drag & Drop your .puml or .txt file(s) here\n\n(Batch processing supported)")
        self.drop_label.setStyleSheet("border: 3px dashed #4af626; background-color: #f4f4f4; font-size: 15px; font-weight: bold; color: #333;")
        
        self.convert_btn.setEnabled(True)
        self.convert_btn.setText("Convert to Visio")
        self.convert_btn.setStyleSheet("background-color: #4af626; color: #000; font-weight: bold; padding: 10px; font-size: 14px; border-radius: 5px;")
        
        self.clear_btn.setEnabled(True)
        self.clear_btn.setStyleSheet("background-color: #777; color: white; font-weight: bold; padding: 10px; font-size: 14px; border-radius: 5px;")

    def _set_drop_zone_busy(self):
        self.drop_label.setText("⚙️ Processing Queue...\n\nPlease wait until finished.")
        self.drop_label.setStyleSheet("border: 3px dashed #f6a826; background-color: #fff4e5; font-size: 15px; font-weight: bold; color: #d37e00;")
        
        self.convert_btn.setEnabled(False)
        self.convert_btn.setText("Processing...")
        self.convert_btn.setStyleSheet("background-color: #ccc; color: #666; font-weight: bold; padding: 10px; font-size: 14px; border-radius: 5px;")
        
        self.clear_btn.setEnabled(False)
        self.clear_btn.setStyleSheet("background-color: #ccc; color: #666; font-weight: bold; padding: 10px; font-size: 14px; border-radius: 5px;")

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
            self.log_message("🏁 All batch processing complete! Waiting for new files.\n" + "-"*45)
            self._set_drop_zone_ready()
            return

        self.is_processing = True
        self._set_drop_zone_busy()
        
        next_file = self.file_queue.pop(0)
        
        self.conv_thread = ConverterThread(next_file, self.jar_path)
        self.conv_thread.ui_log_msg.connect(self.log_message)
        self.conv_thread.finished.connect(self.process_next_in_queue)
        self.conv_thread.start()

    def log_message(self, message: str):
        self.console.append(message)
        QApplication.processEvents()
        scrollbar = self.console.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())


if __name__ == '__main__':
    logging.info("Starting PlantUML to Visio Converter application...")
    app = QApplication(sys.argv)
    app.setStyle("Fusion") 
    
    proxy_dialog = ProxyDialog()
    proxy_dialog.exec_()
    
    http_val, https_val = proxy_dialog.get_proxies()
    
    window = DragDropUI(http_proxy=http_val, https_proxy=https_val)
    window.show()
    sys.exit(app.exec_())