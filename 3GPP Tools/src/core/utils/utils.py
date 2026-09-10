import json
import logging
import os
import re
import shutil
import subprocess
import winreg
from pathlib import Path
from typing import Any, Dict, Optional, Tuple

_BEST_JAVA_CACHE: Optional[Tuple[str, int]] = None


def get_java_config_path() -> Path:
    """Returns the path to the persistent Java configuration file."""
    try:
        from core.utils.paths import get_project_root
        return get_project_root() / "config" / "java_config.json"
    except Exception:
        return Path("config") / "java_config.json"


def load_java_config() -> Dict[str, Any]:
    """Loads Java preferences from config/java_config.json."""
    config_file = get_java_config_path()
    default_cfg = {
        "use_custom_path": False,
        "custom_java_path": ""
    }
    if config_file.exists():
        try:
            with open(config_file, "r", encoding="utf-8") as f:
                data = json.load(f)
                default_cfg.update(data)
        except Exception as e:
            logging.warning(f"⚠️ Could not read java_config.json: {e}")
    return default_cfg


def save_java_config(cfg: Dict[str, Any]) -> bool:
    """Saves Java preferences to config/java_config.json."""
    config_file = get_java_config_path()
    try:
        config_file.parent.mkdir(parents=True, exist_ok=True)
        with open(config_file, "w", encoding="utf-8") as f:
            json.dump(cfg, f, indent=4)
        return True
    except Exception as e:
        logging.error(f"❌ Failed to save java_config.json: {e}")
        return False


def reset_java_cache():
    """Clears the cached Java binary to allow immediate re-detection."""
    global _BEST_JAVA_CACHE
    _BEST_JAVA_CACHE = None


def get_detailed_java_info(java_exe: Optional[str] = None) -> Dict[str, Any]:
    """
    Probes the Java executable and returns a diagnostic dictionary.

    Returns:
        dict with keys:
            - 'installed': bool
            - 'path': str (resolved path)
            - 'major': int (e.g. 8, 11, 17, 21)
            - 'version_str': str (e.g. "21.0.2")
            - 'vendor': str (e.g. "OpenJDK Runtime Environment (Temurin)")
            - 'arch': str ("64-Bit", "32-Bit", or "Unknown")
            - 'raw_output': str
    """
    if not java_exe:
        best_exe, _ = get_best_java()
        java_exe = best_exe if best_exe else "java"

    # Resolve executable path
    resolved_path = shutil.which(java_exe) or java_exe
    try:
        resolved_path = str(Path(resolved_path).resolve())
    except Exception:
        resolved_path = str(java_exe)

    kwargs = {'creationflags': 0x08000000} if os.name == 'nt' else {}
    try:
        res = subprocess.run(
            [java_exe, "-version"],
            capture_output=True,
            text=True,
            timeout=5,
            **kwargs
        )
        raw_output = ((res.stderr or "") + "\n" + (res.stdout or "")).strip()
    except Exception as e:
        return {
            "installed": False,
            "path": resolved_path,
            "major": 0,
            "version_str": f"Execution error: {e}",
            "vendor": "Unknown",
            "arch": "Unknown",
            "raw_output": "",
        }

    # 1. Parse Version String
    v_match = re.search(r'(?:java|openjdk)\s+version\s+"([^"]+)"', raw_output, re.IGNORECASE)
    if not v_match:
        v_match = re.search(r'version\s+"?([0-9._]+(?:-[a-zA-Z0-9.+]+)?)"?', raw_output, re.IGNORECASE)

    version_str = v_match.group(1) if v_match else "Unknown"

    # 2. Determine Major Version Integer
    major = 0
    if version_str != "Unknown":
        try:
            if version_str.startswith("1."):
                major = int(version_str.split(".")[1])
            else:
                major = int(re.split(r'[._-]', version_str)[0])
        except (IndexError, ValueError):
            major = 0

    if major == 0:
        return {
            "installed": False,
            "path": resolved_path,
            "major": 0,
            "version_str": version_str,
            "vendor": "Unknown",
            "arch": "Unknown",
            "raw_output": raw_output,
        }

    # 3. Parse Architecture
    if re.search(r'64-Bit|x86_64|amd64', raw_output, re.IGNORECASE):
        arch = "64-Bit"
    elif re.search(r'32-Bit|x86|i386', raw_output, re.IGNORECASE):
        arch = "32-Bit"
    else:
        arch = "Unknown"

    # 4. Parse Vendor / Runtime
    vendor = "Unknown"
    lines = [line.strip() for line in raw_output.splitlines() if line.strip()]
    for line in lines:
        if "Runtime Environment" in line:
            vendor = line
            break
        elif "VM" in line:
            vendor = line
            break

    if vendor == "Unknown" and lines:
        vendor = lines[0]

    return {
        "installed": True,
        "path": resolved_path,
        "major": major,
        "version_str": version_str,
        "vendor": vendor,
        "arch": arch,
        "raw_output": raw_output,
    }


def get_best_java(log_callback=None) -> Tuple[str, int]:
    """
    Scans for the best available Java executable.
    Priority:
      1. User-configured custom Java path in config/java_config.json
      2. Portable folder in project root (e.g., ./jre/bin/java.exe or ./tools/java/bin/java.exe)
      3. JAVA_HOME, User/System PATHs, and Registry
    """
    global _BEST_JAVA_CACHE
    if _BEST_JAVA_CACHE is not None:
        return _BEST_JAVA_CACHE

    # Helper to evaluate an executable
    def test_candidate(cmd):
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
                    return v
        except Exception:
            pass
        return 0

    # 1. Check custom path from java_config.json
    cfg = load_java_config()
    if cfg.get("use_custom_path") and cfg.get("custom_java_path"):
        custom_p = cfg["custom_java_path"].strip(' "')
        custom_exe = custom_p
        if os.path.isdir(custom_p):
            candidate = os.path.join(custom_p, "bin", "java.exe" if os.name == 'nt' else "java")
            if os.path.exists(candidate):
                custom_exe = candidate
        if os.path.exists(custom_exe):
            v = test_candidate(custom_exe)
            if v > 0:
                if log_callback:
                    log_callback(f"✓ Using configured custom Java {v} at: {custom_exe}", logging.INFO)
                _BEST_JAVA_CACHE = (os.path.normpath(custom_exe), v)
                return _BEST_JAVA_CACHE

    candidates = set()

    def add_candidate(path_str):
        if not path_str:
            return
        clean_p = os.path.expandvars(path_str.strip(' "'))
        if clean_p:
            exe = os.path.join(clean_p, 'java.exe' if os.name == 'nt' else 'java')
            if os.path.exists(exe):
                candidates.add(os.path.normpath(exe))

    # 2. Check Portable Project Folders (Zero-Admin local installations)
    try:
        from core.utils.paths import get_project_root
        root = get_project_root()
        portable_dirs = [
            root / "jre" / "bin",
            root / "tools" / "java" / "bin",
            root / "modules" / "puml2visio" / "assets" / "jre" / "bin",
        ]
        # Also check one folder level down (e.g. root/jre/jdk-21.0.2+13-jre/bin)
        for base in [root / "jre", root / "tools"]:
            if base.exists() and base.is_dir():
                for sub in base.iterdir():
                    if sub.is_dir() and (sub / "bin").exists():
                        portable_dirs.append(sub / "bin")

        for p_dir in portable_dirs:
            if p_dir.exists():
                add_candidate(str(p_dir))
    except Exception:
        pass

    # 3. Standard Environment & PATH
    java_home = os.environ.get('JAVA_HOME')
    if java_home:
        add_candidate(os.path.join(java_home, 'bin'))

    for p in os.environ.get('PATH', '').split(os.pathsep):
        add_candidate(p)

    # 4. Registry Paths on Windows
    if os.name == 'nt':
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Environment") as key:
                val, _ = winreg.QueryValueEx(key, "Path")
                for p in val.split(os.pathsep):
                    add_candidate(p)
        except Exception:
            pass

        try:
            with winreg.OpenKey(
                    winreg.HKEY_LOCAL_MACHINE,
                    r"System\CurrentControlSet\Control\Session Manager\Environment"
            ) as key:
                val, _ = winreg.QueryValueEx(key, "Path")
                for p in val.split(os.pathsep):
                    add_candidate(p)
        except Exception:
            pass

    best_exe = "java"
    best_ver = 0

    if log_callback and candidates:
        log_callback(f"🔎 Scanning {len(candidates)} Java candidate locations...", logging.INFO)

    for exe in sorted(candidates):
        v = test_candidate(exe)
        if log_callback and v > 0:
            log_callback(f"  ✓ Found Java {v} at: {exe}", logging.INFO)
        if v > best_ver:
            best_ver = v
            best_exe = exe

    bare_v = test_candidate("java")
    if bare_v > best_ver:
        best_ver = bare_v
        best_exe = "java"

    _BEST_JAVA_CACHE = (best_exe, best_ver)
    return _BEST_JAVA_CACHE


def get_proxies() -> Dict[str, Optional[str]]:
    """Safely fetches the proxy configuration from the application's environment."""
    return {
        "http": os.environ.get("HTTP_PROXY") or os.environ.get("http_proxy"),
        "https": os.environ.get("HTTPS_PROXY") or os.environ.get("https_proxy"),
    }