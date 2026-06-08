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