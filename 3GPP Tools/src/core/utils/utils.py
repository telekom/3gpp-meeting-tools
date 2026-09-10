import logging
import os
import re
import shutil
import subprocess
import winreg
from pathlib import Path
from typing import Any, Dict, Optional, Tuple

_BEST_JAVA_CACHE: Optional[Tuple[str, int]] = None


def get_detailed_java_info(java_exe: Optional[str] = None) -> Dict[str, Any]:
    """
    Probes the Java executable and returns a diagnostic dictionary.

    Returns:
        dict with keys:
            - 'installed': bool
            - 'path': str (fully resolved path to java.exe)
            - 'major': int (e.g. 8, 11, 17, 21)
            - 'version_str': str (e.g. "21.0.2")
            - 'vendor': str (e.g. "OpenJDK Runtime Environment (build 21.0.2+13)")
            - 'arch': str ("64-Bit", "32-Bit", or "Unknown")
            - 'raw_output': str (verbatim output from java -version)
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

    # 4. Parse Vendor / Runtime description
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
    """Scans the environment for all Java executables and returns (path, major_version)."""
    global _BEST_JAVA_CACHE
    if _BEST_JAVA_CACHE is not None:
        return _BEST_JAVA_CACHE

    candidates = set()

    def add_candidate(path_str):
        if not path_str:
            return
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
    if os.name == 'nt':
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Environment") as key:
                val, _ = winreg.QueryValueEx(key, "Path")
                for p in val.split(os.pathsep):
                    add_candidate(p)
        except Exception:
            pass

        # 4. Registry System PATH
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

    for exe in sorted(candidates):
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


def get_proxies() -> Dict[str, Optional[str]]:
    """Safely fetches the proxy configuration from the application's environment."""
    return {
        "http": os.environ.get("HTTP_PROXY") or os.environ.get("http_proxy"),
        "https": os.environ.get("HTTPS_PROXY") or os.environ.get("https_proxy"),
    }