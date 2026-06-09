import os
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

# --- SMART JAVA DISCOVERY ENGINE ---
_BEST_JAVA_CACHE = None


def get_best_java(log_callback=None):
    """Scans the environment for all Java executables and returns the path to the newest one."""
    global _BEST_JAVA_CACHE
    if _BEST_JAVA_CACHE is not None:
        return _BEST_JAVA_CACHE

    candidates = set()

    def add_candidate(path_str):
        if not path_str: return
        clean_p = os.path.expandvars(path_str.strip(' "'))
        if clean_p:
            exe = os.path.join(clean_p, 'java.exe')
            if os.path.exists(exe):
                candidates.add(os.path.normpath(exe))

    # 1. JAVA_HOME
    java_home = os.environ.get('JAVA_HOME')
    if java_home:
        add_candidate(os.path.join(java_home, 'bin'))

    # 2. Live Environment PATH
    for p in os.environ.get('PATH', '').split(os.pathsep):
        add_candidate(p)

    # 3. Registry User PATH
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Environment") as key:
            val, _ = winreg.QueryValueEx(key, "Path")
            for p in val.split(os.pathsep):
                add_candidate(p)
    except Exception:
        pass

    # 4. Registry System PATH
    try:
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,
                            r"System\CurrentControlSet\Control\Session Manager\Environment") as key:
            val, _ = winreg.QueryValueEx(key, "Path")
            for p in val.split(os.pathsep):
                add_candidate(p)
    except Exception:
        pass

    best_exe = "java"
    best_ver = 0

    if log_callback and candidates:
        log_callback(f"🔎 Scanning {len(candidates)} Java locations from System/User paths...", logging.INFO)

    def check_version(cmd):
        try:
            kwargs = {'creationflags': 0x08000000} if os.name == 'nt' else {}
            result = subprocess.run([cmd, "-version"], capture_output=True, text=True, timeout=5, **kwargs)

            output = result.stderr + "\n" + result.stdout
            match = re.search(r'"(\d[^"]*)"', output)
            if not match:
                match = re.search(r'version\s+([^\s]+)', output, re.IGNORECASE)

            if match:
                ver_str = match.group(1)
                nums = re.findall(r'\d+', ver_str)
                if nums:
                    v = int(nums[1]) if (nums[0] == '1' and len(nums) > 1) else int(nums[0])
                    if log_callback:
                        log_callback(f"  ✓ Found Java {v} at: {cmd}", logging.INFO)
                    return v

            if log_callback:
                clean_out = output.replace('\n', ' ').strip()[:50]
                log_callback(f"  ⚠️ Unrecognized version format for {cmd} (Output: {clean_out}...)", logging.WARNING)
        except Exception as e:
            if log_callback:
                log_callback(f"  ❌ Failed to test {cmd}: {e}", logging.ERROR)
        return 0

    for exe in candidates:
        v = check_version(exe)
        if v > best_ver:
            best_ver = v
            best_exe = exe

    bare_v = check_version("java")
    if bare_v > best_ver:
        best_ver = bare_v
        best_exe = "java"

    _BEST_JAVA_CACHE = (best_exe, best_ver)
    return _BEST_JAVA_CACHE


# --- CORE UTILITIES ---
def strip_watermark(raw_text: str) -> str:
    lines = raw_text.splitlines()
    while lines and ("Generated with puml2visio" in lines[0] or lines[0].strip() == ""):
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
    network_error = pyqtSignal()  # --- NEW: Signal to alert the GUI of firewalls ---

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
        req = urllib.request.Request("https://github.com/plantuml/plantuml/releases/latest", method="HEAD")
        req.add_header("User-Agent", "Mozilla/5.0")
        with urllib.request.urlopen(req, timeout=5) as response:
            match = re.search(r'/tag/v?(\d+\.\d+\.\d+)', response.url)
            if match:
                return match.group(1)
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
                        self.network_error.emit()  # Alert the GUI!
                        self.init_complete.emit(False)
                        return

            if download_reason:
                url_to_download = URL_LATEST if required_type == "modern" else URL_JAVA_8
                self._emit_log(f"⚠️ Downloading PlantUML. Reason: {download_reason}...", logging.WARNING)
                try:
                    urllib.request.urlretrieve(url_to_download, self.jar_path)
                    version_file.write_text(required_type, encoding="utf-8")
                    self._emit_log("✅ PlantUML downloaded and installed successfully.", logging.INFO)
                except Exception as e:
                    self._emit_log(f"❌ Download failed. Please check your network or proxy settings: {e}",
                                   logging.ERROR)
                    self.network_error.emit()  # Alert the GUI!
                    self.init_complete.emit(False)
                    return

            self.init_complete.emit(True)

        except Exception as e:
            self._emit_log(f"❌ System Check Failed: {e}", logging.CRITICAL)
            self.init_complete.emit(False)

    def _emit_log(self, message: str, level: int):
        logging.log(level, message)
        self.ui_log_msg.emit(message)