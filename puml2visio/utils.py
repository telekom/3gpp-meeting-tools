import subprocess
import urllib.request
import winreg
import re
import logging
import zlib
import base64
from pathlib import Path
from PyQt5.QtCore import QThread, pyqtSignal

JAR_NAME = "plantuml.jar"
URL_LATEST = "https://github.com/plantuml/plantuml/releases/latest/download/plantuml.jar"
URL_JAVA_8 = "https://github.com/plantuml/plantuml/releases/download/v1.2023.13/plantuml-1.2023.13.jar"

WATERMARK = "' Generated with puml2visio, https://github.com/telekom/3gpp-meeting-tools/tree/master/puml2visio"


def strip_watermark(raw_text: str) -> str:
    """Removes existing puml2visio watermarks and leading empty lines to prevent duplication."""
    lines = raw_text.splitlines()
    while lines and ("Generated with puml2visio" in lines[0] or lines[0].strip() == ""):
        lines.pop(0)
    return "\n".join(lines)


def generate_cleaned_svg(puml_path: Path, jar_path: Path, log_callback=None) -> Path:
    """Centralized function to generate an SVG via PlantUML and clean text attributes."""
    command = ["java", "-jar", str(jar_path), "-tsvg", str(puml_path)]
    try:
        subprocess.run(command, check=True, capture_output=True, text=True, cwd=puml_path.parent)
    except subprocess.CalledProcessError as e:
        raise RuntimeError(f"PlantUML Syntax Error:\n{e.stderr}")

    svg_path = puml_path.with_suffix(".svg")
    if not svg_path.exists():
        raise FileNotFoundError("PlantUML finished, but SVG was not created.")

    try:
        with open(svg_path, 'r', encoding='utf-8') as f:
            svg_content = f.read()

        # Strip problematic tags for Office COM importing
        svg_content = re.sub(r'\s*textLength="[^"]*"', '', svg_content)
        svg_content = re.sub(r'\s*lengthAdjust="[^"]*"', '', svg_content)

        # Merge shattered SVG text blocks
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
        if log_callback:
            log_callback(f"⚠️ Warning: Could not clean SVG text attributes: {e}", logging.WARNING)

    return svg_path


def encode_plantuml(text: str) -> str:
    """Encodes raw PlantUML text into the Deflate + Base64 format expected by PlantText."""
    compressor = zlib.compressobj(level=9, wbits=-15)
    compressed = compressor.compress(text.encode('utf-8')) + compressor.flush()
    b64 = base64.b64encode(compressed).decode('ascii')

    std_b64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
    puml_b64 = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz-_"
    trans = str.maketrans(std_b64, puml_b64)
    return b64.translate(trans).replace('=', '')


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