import os
import subprocess
import winreg
import re
import logging
import zlib
import base64
from pathlib import Path
from PyQt5.QtCore import QThread, pyqtSignal

from core.utils.utils import get_best_java
from modules.puml2visio.config.paths import PLANTUML_URL_LATEST, PLANTUML_URL_JAVA_8
from core.network.session import NetworkSession  # <--- NEW: Import the shared session


# --- CORE UTILITIES ---
def strip_watermark(raw_text: str) -> str:
    lines = raw_text.splitlines()
    while lines and ("Generated with 3GPP Tools" in lines[0] or lines[0].strip() == ""):
        lines.pop(0)
    return "\n".join(lines)


def generate_cleaned_svg(puml_path: Path, jar_path: Path, log_callback=None) -> Path:
    java_exe, _ = get_best_java()
    command = [java_exe, "-jar", str(jar_path), "-tsvg", str(puml_path)]
    kwargs = {'creationflags': 0x08000000} if os.name == 'nt' else {}

    try:
        subprocess.run(command, check=True, capture_output=True, text=True, cwd=puml_path.parent, **kwargs)
    except subprocess.CalledProcessError as e:
        raise RuntimeError(f"PlantUML Syntax Error:\n{e.stderr}")

    svg_path = puml_path.with_suffix(".svg")
    if not svg_path.exists():
        raise FileNotFoundError("PlantUML finished, but SVG was not created.")

    try:
        with open(svg_path, 'r', encoding='utf-8') as f:
            svg_content = f.read()

        # --- INTELLIGENT 2D TEXT MERGE ---
        pattern = re.compile(r'(<text\b[^>]*?>)(.*?)(</text>)', re.IGNORECASE | re.DOTALL)
        matches = list(pattern.finditer(svg_content))

        if matches:
            result = []
            last_end = 0

            current_y = None
            current_end_x = None
            current_start_tag = ""
            current_text = ""

            for m in matches:
                start = m.start()
                end = m.end()
                full_open_tag = m.group(1)
                inner_text = m.group(2)
                between = svg_content[last_end:start]

                # Extract coordinates and physical lengths
                y_match = re.search(r'\by="([0-9.]+)"', full_open_tag)
                x_match = re.search(r'\bx="([0-9.]+)"', full_open_tag)
                tl_match = re.search(r'\btextLength="([0-9.]+)"', full_open_tag)

                y_val = float(y_match.group(1)) if y_match else None
                x_val = float(x_match.group(1)) if x_match else None
                tl_val = float(tl_match.group(1)) if tl_match else (
                        len(inner_text.strip()) * 7.0)  # Fallback to pixel estimation

                # Check both Y alignment and X distance gap
                should_merge = False
                if current_y is not None and y_val == current_y and x_val is not None and current_end_x is not None:
                    if not between.strip():  # Only merge if no other physical SVG elements exist between them
                        gap = x_val - current_end_x
                        # Threshold: if tags are within 25 pixels, they are a sentence. If > 25, they are distinct labels.
                        if -10 <= gap <= 25:
                            should_merge = True

                if should_merge:
                    # Intelligently inject spaces for shattered sentences to maintain Visio legibility
                    if gap > 4 and not current_text.endswith(' ') and not inner_text.startswith(' '):
                        current_text += " " + inner_text
                    else:
                        current_text += inner_text
                    current_end_x = x_val + tl_val
                else:
                    # Flush the previous object and start a new distinct text box
                    if current_y is not None:
                        result.append(current_start_tag)
                        result.append(current_text)
                        result.append("</text>")
                    result.append(between)

                    current_y = y_val
                    current_end_x = x_val + tl_val if x_val is not None else None
                    current_start_tag = full_open_tag
                    current_text = inner_text

                last_end = end

            if current_y is not None:
                result.append(current_start_tag)
                result.append(current_text)
                result.append("</text>")

            result.append(svg_content[last_end:])
            svg_content = "".join(result)

        # --- STRIP OFFICE COM ATTRIBUTES ---
        # Now that we've analyzed the text lengths, we can safely delete them so Visio doesn't crash!
        svg_content = re.sub(r'\s*textLength="[^"]*"', '', svg_content)
        svg_content = re.sub(r'\s*lengthAdjust="[^"]*"', '', svg_content)
        svg_content = svg_content.replace('&#160;', ' ').replace('\xa0', ' ')

        with open(svg_path, 'w', encoding='utf-8') as f:
            f.write(svg_content)

    except Exception as e:
        if log_callback:
            log_callback(f"⚠️ Warning: Could not clean SVG text attributes: {e}", logging.WARNING)

    return svg_path


def encode_plantuml(text: str) -> str:
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
    network_error = pyqtSignal()

    def __init__(self, jar_path: Path, check_updates: bool = False):
        super().__init__()
        self.jar_path = jar_path
        self.check_updates = check_updates

    def _get_local_version(self, java_exe):
        try:
            kwargs = {'creationflags': 0x08000000} if os.name == 'nt' else {}
            res = subprocess.run([java_exe, "-jar", str(self.jar_path), "-version"], capture_output=True, text=True,
                                 timeout=5, **kwargs)
            match = re.search(r'version\s+(\d+\.\d+\.\d+)', res.stdout + res.stderr, re.IGNORECASE)
            if match:
                return match.group(1)
        except:
            pass
        return None

    def _get_remote_version(self):
        """Uses the shared NetworkSession to check GitHub for the latest release."""
        try:
            session = NetworkSession.get_instance()
            # A HEAD request with redirects allows us to grab the final URL (which contains the version tag)
            # without actually downloading the HTML body of the page.
            response = session.head("https://github.com/plantuml/plantuml/releases/latest", allow_redirects=True,
                                    timeout=10)

            match = re.search(r'/tag/v?(\d+\.\d+\.\d+)', response.url)
            if match:
                return match.group(1)
        except Exception:
            pass
        return None

    def run(self):
        try:
            self._emit_log("🔍 Initializing system checks...", logging.INFO)
            try:
                winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Visio.Application")
            except FileNotFoundError:
                self._emit_log("❌ ERROR: Microsoft Visio is not installed or registered.", logging.ERROR)
                self.init_complete.emit(False)
                return

            java_exe, java_major = get_best_java(self._emit_log if not self.check_updates else None)

            if java_major == 0:
                self._emit_log("❌ ERROR: Java is not installed or could not be found.", logging.ERROR)
                self.init_complete.emit(False)
                return
            elif not self.check_updates:
                self._emit_log(f"✅ Active Engine: Java {java_major} ({java_exe})", logging.INFO)

            required_type = "modern" if java_major >= 11 else "legacy"
            version_file = self.jar_path.with_suffix('.version')

            current_type = None
            if version_file.exists():
                try:
                    current_type = version_file.read_text(encoding="utf-8").strip()
                except:
                    pass

            download_reason = None

            if not self.jar_path.exists():
                download_reason = "File is missing"
            elif current_type != required_type:
                download_reason = f"Java mismatch (Required: {required_type}, Found: {current_type})"
            elif self.check_updates:
                if required_type == "legacy":
                    self._emit_log("ℹ️ Legacy Java 8 version is pinned. No automated updates available.", logging.INFO)
                else:
                    self._emit_log("🌐 Checking GitHub for the latest PlantUML release...", logging.INFO)
                    local_v = self._get_local_version(java_exe)
                    try:
                        remote_v = self._get_remote_version()
                        if local_v and remote_v:
                            def v_tuple(v_str):
                                return tuple(int(x) for x in re.findall(r'\d+', v_str))

                            if v_tuple(remote_v) > v_tuple(local_v):
                                download_reason = f"Update available ({local_v} → {remote_v})"
                            else:
                                self._emit_log(f"✅ PlantUML is up-to-date (Version {local_v}).", logging.INFO)
                        else:
                            self._emit_log("⚠️ Could not parse version data from GitHub.", logging.WARNING)
                    except Exception as e:
                        self._emit_log(f"⚠️ Network check blocked: {e}", logging.WARNING)
                        self.network_error.emit()
                        self.init_complete.emit(False)
                        return

            if download_reason:
                url_to_download = PLANTUML_URL_LATEST if required_type == "modern" else PLANTUML_URL_JAVA_8
                self._emit_log(f"⚠️ Downloading PlantUML. Reason: {download_reason}...", logging.WARNING)
                try:
                    # ---> REUSING THE SHARED DOWNLOAD UTILITY <---
                    NetworkSession.download_file(url_to_download, self.jar_path)

                    version_file.write_text(required_type, encoding="utf-8")
                    self._emit_log("✅ PlantUML downloaded and installed successfully.", logging.INFO)

                except Exception as e:
                    self._emit_log(f"❌ Download failed. Please check your network or proxy settings: {e}",
                                   logging.ERROR)
                    self.network_error.emit()
                    self.init_complete.emit(False)
                    return

            self.init_complete.emit(True)

        except Exception as e:
            self._emit_log(f"❌ System Check Failed: {e}", logging.CRITICAL)
            self.init_complete.emit(False)

    def _emit_log(self, message: str, level: int):
        logging.log(level, message)
        self.ui_log_msg.emit(message)