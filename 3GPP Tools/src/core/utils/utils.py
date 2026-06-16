import logging
import os
import re
import subprocess
import winreg

import core.utils

_BEST_JAVA_CACHE = None

# --- SMART JAVA DISCOVERY ENGINE ---
def get_best_java(log_callback=None):
    """Scans the environment for all Java executables and returns the path to the newest one."""
    if core.utils.utils._BEST_JAVA_CACHE is not None:
        return core.utils.utils._BEST_JAVA_CACHE

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

    core.utils.utils._BEST_JAVA_CACHE = (best_exe, best_ver)
    return core.utils.utils._BEST_JAVA_CACHE
