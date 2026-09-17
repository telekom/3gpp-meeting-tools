# --- File: src/modules/word_tools/core/libreoffice_converter.py ---
import configparser
import json
import logging
import os
import platform
import shutil
import stat
import subprocess
import tempfile
import time
import socket
import uuid
from dataclasses import dataclass
from pathlib import Path
from typing import Optional, Union

from modules.word_tools.core.word_config import WordConfig

LIBREOFFICE_DOWNLOAD_URL = "https://portableapps.com/apps/office/libreoffice_portable"

PDF_FILTER_OPTIONS = {
    "ExportBookmarks": {"type": "boolean", "value": "true"},
    "ExportBookmarksToPDFDestination": {"type": "boolean", "value": "true"},
    "ConvertOOoTargetToPDFTarget": {"type": "boolean", "value": "true"},
    "OpenBookmarkLevels": {"type": "integer", "value": "-1"},
}

LO_EXPORT_FILTERS = {
    "docx": "docx:MS Word 2007 XML",
    "doc": "doc:MS Word 97",
    "pdf": f"pdf:writer_pdf_Export:{json.dumps(PDF_FILTER_OPTIONS, separators=(',', ':'))}",
    "html": "html:HTML (StarWriter)",
    "htm": "html:HTML (StarWriter)",
    "rtf": "rtf:Rich Text Format",
    "txt": "txt:Text (encoded)",
}

LO_NO_PRINTER_CONFIG = """<?xml version="1.0" encoding="UTF-8"?>
<oor:items xmlns:oor="http://openoffice.org/2001/registry"
           xmlns:xs="http://www.w3.org/2001/XMLSchema"
           xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <item oor:path="/org.openoffice.Office.Common/Save/Document">
    <prop oor:name="LoadPrinterSetup" oor:type="xs:boolean"><value>false</value></prop>
  </item>
  <item oor:path="/org.openoffice.Office.Writer/Layout/Other">
    <prop oor:name="LoadPrinterSetup" oor:type="xs:boolean"><value>false</value></prop>
  </item>
  <item oor:path="/org.openoffice.Office.Common/Print/Option">
    <prop oor:name="PrinterIndependentLayout" oor:type="xs:string"><value>enabled</value></prop>
  </item>

  <!-- Conservative rendering for headless conversion.  These values live only
       in the temporary conversion profile and never modify the user's portable
       LibreOffice profile. -->
  <item oor:path="/org.openoffice.Office.Common/VCL">
    <prop oor:name="UseSkia" oor:type="xs:boolean"><value>false</value></prop>
    <prop oor:name="ForceSkiaRaster" oor:type="xs:boolean"><value>false</value></prop>
  </item>
</oor:items>
"""


@dataclass(frozen=True)
class LibreOfficeRuntime:
    executable: Path
    portable: bool
    portable_root: Optional[Path] = None


def _normalize_target_format(raw_format: str) -> str:
    fmt = raw_format.lower().replace(".", "").strip()
    for suffix in ("_libreoffice", "_word", "_com", "_headless"):
        if fmt.endswith(suffix):
            fmt = fmt[: -len(suffix)]
    return fmt or "docx"


def _find_portable_root(path: Path) -> Optional[Path]:
    """Find a PortableApps package root without launching anything."""
    try:
        current = path if path.is_dir() else path.parent
    except Exception:
        return None

    for _ in range(10):
        appinfo = current / "App" / "AppInfo"
        launcher = current / "LibreOfficePortable.exe"
        if appinfo.is_dir() and launcher.is_file():
            return current.resolve()

        parent = current.parent
        if parent == current:
            break
        current = parent

    return None


def _read_portable_program_executable(root: Path) -> Optional[Path]:
    """
    Resolve the packaged program from PortableApps launcher metadata.

    This deliberately reads the package metadata instead of assuming a
    particular LibreOffice version or App\\libreoffice layout.
    """
    launcher_dir = root / "App" / "AppInfo" / "Launcher"
    if not launcher_dir.is_dir():
        return None

    preferred = launcher_dir / "LibreOfficePortable.ini"
    ini_files = [preferred] if preferred.is_file() else []
    ini_files.extend(
        p for p in sorted(launcher_dir.glob("*.ini"))
        if p not in ini_files
    )

    for ini_path in ini_files:
        parser = configparser.ConfigParser(interpolation=None)
        try:
            parser.read(ini_path, encoding="utf-8")
        except UnicodeDecodeError:
            try:
                parser.read(ini_path, encoding="utf-8-sig")
            except Exception:
                continue
        except Exception:
            continue

        if not parser.has_section("Launch"):
            continue

        # PortableApps may define architecture/parameter-specific executables.
        # Prefer a soffice-like executable that actually exists.
        keys = (
            "ProgramExecutableWhenParameters64",
            "ProgramExecutable64",
            "ProgramExecutableWhenParameters",
            "ProgramExecutable",
        )
        for key in keys:
            value = parser.get("Launch", key, fallback="").strip().strip('"')
            if not value:
                continue

            # PAL ProgramExecutable paths are relative to AppDir (= root/App).
            candidate = (root / "App" / Path(value.replace("\\", os.sep))).resolve()
            if candidate.is_file() and candidate.name.lower().startswith("soffice"):
                return candidate

    # Defensive structural fallbacks for older/newer LO Portable layouts.
    for candidate in (
        root / "App" / "libreoffice" / "program" / "soffice.exe",
        root / "App" / "LibreOffice64" / "program" / "soffice.exe",
        root / "App" / "LibreOffice" / "program" / "soffice.exe",
    ):
        if candidate.is_file():
            return candidate.resolve()

    return None


def resolve_libreoffice_runtime(
    candidate_path: Optional[Union[str, Path]] = None,
) -> Optional[LibreOfficeRuntime]:
    """
    Resolve installed or Portable LibreOffice without starting a process.

    For PortableApps, the launcher is used only as a package locator. Conversion
    executes the packaged soffice binary directly with a clean environment and
    an isolated LibreOffice profile. This avoids PortableApps launcher lifecycle
    popups/cleanup state while remaining version-independent.
    """
    if candidate_path:
        try:
            path = Path(candidate_path).expanduser().resolve()
        except Exception:
            return None

        root = _find_portable_root(path)
        if root:
            program = _read_portable_program_executable(root)
            if program:
                return LibreOfficeRuntime(program, True, root)

        if path.is_file() and path.name.lower().startswith(("soffice", "libreoffice")):
            return LibreOfficeRuntime(path, False, None)

        if path.is_dir():
            for candidate in (
                path / "program" / "soffice.exe",
                path / "soffice.exe",
            ):
                if candidate.is_file():
                    return LibreOfficeRuntime(candidate.resolve(), False, None)

        return None

    custom_path = WordConfig.get_libreoffice_path()
    if custom_path:
        runtime = resolve_libreoffice_runtime(custom_path)
        if runtime:
            return runtime

    for binary_name in ("soffice", "soffice.exe", "libreoffice"):
        found = shutil.which(binary_name)
        if found:
            runtime = resolve_libreoffice_runtime(found)
            if runtime:
                return runtime

    if platform.system() == "Windows":
        for candidate in (
            Path(os.environ.get("ProgramFiles", r"C:\Program Files"))
            / "LibreOffice" / "program" / "soffice.exe",
            Path(os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)"))
            / "LibreOffice" / "program" / "soffice.exe",
            Path(os.environ.get("LOCALAPPDATA", ""))
            / "Programs" / "LibreOffice" / "program" / "soffice.exe",
        ):
            if candidate.is_file():
                return LibreOfficeRuntime(candidate.resolve(), False, None)

    elif platform.system() == "Darwin":
        candidate = Path("/Applications/LibreOffice.app/Contents/MacOS/soffice")
        if candidate.is_file():
            return LibreOfficeRuntime(candidate.resolve(), False, None)

    elif platform.system() == "Linux":
        for candidate in (
            Path("/usr/bin/soffice"),
            Path("/usr/bin/libreoffice"),
            Path("/usr/local/bin/soffice"),
            Path("/usr/local/bin/libreoffice"),
            Path("/opt/libreoffice/program/soffice"),
        ):
            if candidate.is_file():
                return LibreOfficeRuntime(candidate.resolve(), False, None)

    return None


def resolve_soffice_binary(candidate_path: Union[str, Path]) -> Optional[Path]:
    """
    Backwards-compatible resolver used by the UI.

    Preserve the user's PortableApps launcher/root in configuration so future
    package updates can move the internal soffice binary without making the
    saved path stale. Installed LibreOffice resolves to its soffice executable.
    """
    if not candidate_path:
        return None

    try:
        path = Path(candidate_path).expanduser().resolve()
    except Exception:
        return None

    root = _find_portable_root(path)
    if root:
        launcher = root / "LibreOfficePortable.exe"
        return launcher.resolve() if launcher.is_file() else root

    runtime = resolve_libreoffice_runtime(path)
    return runtime.executable if runtime else None


def find_libreoffice_executable() -> Optional[Path]:
    """
    Return the configured/display path passively.

    PortableApps returns its launcher path for UI/config compatibility, although
    conversion itself deliberately does not execute that launcher.
    """
    custom_path = WordConfig.get_libreoffice_path()
    if custom_path:
        resolved = resolve_soffice_binary(custom_path)
        if resolved:
            return resolved

    runtime = resolve_libreoffice_runtime()
    if runtime:
        if runtime.portable and runtime.portable_root:
            launcher = runtime.portable_root / "LibreOfficePortable.exe"
            if launcher.is_file():
                return launcher.resolve()
        return runtime.executable
    return None


def is_libreoffice_available() -> bool:
    return resolve_libreoffice_runtime() is not None


def get_libreoffice_missing_msg() -> str:
    return (
        "❌ LibreOffice is not installed or could not be found.\n\n"
        "Headless document conversion requires LibreOffice (Installed or Portable).\n"
        f"Download LibreOffice Portable from: {LIBREOFFICE_DOWNLOAD_URL}\n"
        "For PortableApps, select LibreOfficePortable.exe; the converter will "
        "resolve the packaged soffice executable automatically."
    )


def _sanitize_file_attributes(file_path: Path) -> None:
    try:
        if file_path.exists():
            os.chmod(file_path, stat.S_IWRITE | stat.S_IREAD)
    except Exception:
        pass


def _prepare_isolated_profile(profile_dir: Path) -> None:
    user_config_dir = profile_dir / "user"
    user_config_dir.mkdir(parents=True, exist_ok=True)
    (user_config_dir / "registrymodifications.xcu").write_text(
        LO_NO_PRINTER_CONFIG, encoding="utf-8"
    )


def _windows_creation_flags() -> int:
    if platform.system() != "Windows":
        return 0
    return getattr(subprocess, "CREATE_NO_WINDOW", 0x08000000)


def _clean_runtime_environment(runtime: LibreOfficeRuntime) -> dict:
    """
    Build a clean environment for LibreOffice.

    The host Python application may provide PYTHONHOME/PYTHONPATH values which
    are invalid for LibreOffice's bundled Python/runtime and can trigger
    'Could not find platform independent libraries'. PortableApps normally
    prepends LibreOffice's program directory to PATH, so reproduce that generic
    requirement directly for the resolved packaged binary.
    """
    env = os.environ.copy()

    for key in (
        "PYTHONHOME",
        "PYTHONPATH",
        "PYTHONEXECUTABLE",
        "__PYVENV_LAUNCHER__",
        "VIRTUAL_ENV",
    ):
        env.pop(key, None)

    program_dir = runtime.executable.parent

    # LibreOffice Portable may ship its own Python runtime.  soffice can reach
    # PDF export far enough to print "convert ... using filter ..." while a
    # bundled Python child/runtime still fails with:
    #
    #     Could not find platform independent libraries
    #
    # Do not set PYTHONHOME: that can force LibreOffice onto the wrong Python
    # layout.  Instead discover runtime/library directories relative to the
    # resolved LibreOffice installation and make their DLL/program locations
    # available through PATH.
    path_entries = [str(program_dir)]

    portable_app_dir = (
        runtime.portable_root / "App"
        if runtime.portable and runtime.portable_root
        else None
    )

    search_roots = [program_dir]
    if portable_app_dir and portable_app_dir.is_dir():
        search_roots.append(portable_app_dir)

    seen = {str(program_dir).lower()}
    for root in search_roots:
        try:
            # Keep this bounded to directories LibreOffice commonly uses for
            # runtime DLLs; avoid recursively walking the entire portable app.
            candidates = list(root.glob("python*"))
            candidates += list(root.glob("*/python*"))
            candidates += list(root.glob("*/program"))
        except OSError:
            continue

        for candidate in candidates:
            try:
                if not candidate.is_dir():
                    continue
                key = str(candidate.resolve()).lower()
            except OSError:
                continue
            if key in seen:
                continue
            seen.add(key)
            path_entries.append(str(candidate.resolve()))

            # Python distributions commonly place extension DLLs/scripts below
            # these directories. Adding existing directories to PATH is safe
            # and does not impose a Python version.
            for child_name in ("DLLs", "Scripts"):
                child = candidate / child_name
                if child.is_dir():
                    child_key = str(child.resolve()).lower()
                    if child_key not in seen:
                        seen.add(child_key)
                        path_entries.append(str(child.resolve()))

    # Do NOT inherit the host application's full PATH for Portable LibreOffice.
    # 3GPP Tools may be running under a Python environment containing several
    # unrelated Python installations/architectures.  LibreOffice Portable ships
    # its own Python runtime and should not resolve DLLs/executables through the
    # host Python PATH.
    if platform.system() == "Windows":
        system_root = Path(env.get("SystemRoot", r"C:\Windows"))
        for system_dir in (
            system_root / "System32",
            system_root,
            system_root / "System32" / "Wbem",
        ):
            if system_dir.is_dir():
                key = str(system_dir.resolve()).lower()
                if key not in seen:
                    seen.add(key)
                    path_entries.append(str(system_dir.resolve()))
    elif not runtime.portable:
        # Installed LibreOffice on non-Windows platforms can still require the
        # normal system PATH for shared system tools/libraries.
        existing_path = env.get("PATH", "")
        if existing_path:
            path_entries.append(existing_path)

    env["PATH"] = os.pathsep.join(path_entries)

    # LibreOffice uses URE/bootstrap files relative to its program directory.
    # Explicitly point UNO at the resolved installation without inventing a
    # version-specific path.
    fundamental = program_dir / "fundamental.ini"
    if not fundamental.is_file():
        fundamental = program_dir / "fundamentalrc"
    if fundamental.is_file():
        env["URE_BOOTSTRAP"] = fundamental.resolve().as_uri()

    return env


def _apply_conservative_rendering_environment(env: dict) -> dict:
    """
    Return a copy of env configured for a conservative headless retry.

    SAL_USE_VCLPLUGIN=svp selects LibreOffice's headless/virtual VCL backend.
    SAL_DISABLESKIA disables Skia for this child process only.  This does not
    modify the user's LibreOffice/PortableApps profile.
    """
    retry_env = env.copy()
    retry_env["SAL_USE_VCLPLUGIN"] = "svp"
    retry_env["SAL_DISABLESKIA"] = "1"
    retry_env["SAL_DISABLE_OPENGL"] = "1"
    return retry_env


def _process_details(proc: subprocess.CompletedProcess) -> str:
    parts = []
    stdout = (proc.stdout or "").strip()
    stderr = (proc.stderr or "").strip()
    if stdout:
        parts.append(f"stdout: {stdout}")
    if stderr:
        parts.append(f"stderr: {stderr}")
    return "\n".join(parts) or "LibreOffice produced no console output."


def _wait_for_output(path: Path, timeout: float = 10.0) -> bool:
    deadline = time.monotonic() + timeout
    previous_size = -1
    stable = 0

    while time.monotonic() < deadline:
        try:
            if path.exists():
                size = path.stat().st_size
                if size > 0:
                    if size == previous_size:
                        stable += 1
                        if stable >= 2:
                            return True
                    else:
                        stable = 0
                    previous_size = size
        except OSError:
            pass
        time.sleep(0.25)

    try:
        return path.exists() and path.stat().st_size > 0
    except OSError:
        return False



def _diagnostic_runtime_paths(runtime: LibreOfficeRuntime) -> list:
    """Return useful runtime paths without recursively scanning the installation."""
    paths = []
    roots = [runtime.executable.parent]
    if runtime.portable and runtime.portable_root:
        roots.append(runtime.portable_root / "App")

    seen = set()
    for root in roots:
        if not root.exists():
            continue
        candidates = [
            root,
            root / "fundamental.ini",
            root / "fundamentalrc",
            root / "python.exe",
            root / "python-core",
        ]
        try:
            candidates.extend(root.glob("python*"))
            candidates.extend(root.glob("*/python*"))
        except OSError:
            pass

        for p in candidates:
            try:
                rp = p.resolve()
                key = str(rp).lower()
                if key not in seen and rp.exists():
                    seen.add(key)
                    paths.append(rp)
            except OSError:
                pass
    return paths


def _write_failure_diagnostics(
    diagnostic_dir: Path,
    runtime: LibreOfficeRuntime,
    cmd: list,
    runtime_env: dict,
    source: Path,
    expected_file: Path,
    temp_profile_dir: Path,
    temp_out_dir: Path,
    proc: subprocess.CompletedProcess,
    elapsed: float,
) -> Path:
    """Persist a compact diagnostic report and failed working directories."""
    diagnostic_dir.mkdir(parents=True, exist_ok=True)
    report = diagnostic_dir / "diagnostics.txt"

    interesting_env = (
        "PATH", "PYTHONHOME", "PYTHONPATH", "PYTHONEXECUTABLE",
        "VIRTUAL_ENV", "URE_BOOTSTRAP", "UNO_PATH", "UNO_TYPES",
        "UNO_SERVICES", "TEMP", "TMP",
    )

    def list_dir(path: Path) -> list:
        if not path.exists():
            return ["<missing>"]
        rows = []
        try:
            for child in sorted(path.iterdir(), key=lambda p: p.name.lower()):
                try:
                    kind = "DIR" if child.is_dir() else "FILE"
                    size = child.stat().st_size if child.is_file() else ""
                    rows.append(f"{kind:4} {size!s:>12} {child}")
                except OSError as exc:
                    rows.append(f"ERR               {child}: {exc}")
        except OSError as exc:
            rows.append(f"<unable to list: {exc}>")
        return rows or ["<empty>"]

    lines = [
        "3GPP Tools - LibreOffice conversion diagnostics",
        "=" * 56,
        f"Portable: {runtime.portable}",
        f"Portable root: {runtime.portable_root or '<none>'}",
        f"Executable: {runtime.executable}",
        f"Program directory: {runtime.executable.parent}",
        f"Source: {source}",
        f"Expected output: {expected_file}",
        f"Profile: {temp_profile_dir}",
        f"Output directory: {temp_out_dir}",
        f"Elapsed seconds: {elapsed:.3f}",
        f"Return code: {proc.returncode}",
        "",
        "Command:",
        subprocess.list2cmdline([str(x) for x in cmd]),
        "",
        "Runtime paths discovered:",
    ]
    lines.extend(f"  {p}" for p in _diagnostic_runtime_paths(runtime))
    lines.extend(["", "Relevant environment:"])
    for key in interesting_env:
        lines.append(f"{key}={runtime_env.get(key, '<not set>')}")

    lines.extend([
        "",
        "stdout:",
        (proc.stdout or "<empty>").strip(),
        "",
        "stderr:",
        (proc.stderr or "<empty>").strip(),
        "",
        "Output directory contents:",
    ])
    lines.extend(list_dir(temp_out_dir))
    lines.extend(["", "Profile directory contents:"])
    lines.extend(list_dir(temp_profile_dir))

    report.write_text("\n".join(lines), encoding="utf-8", errors="replace")
    return report


UNO_PDF_HELPER = r"""
import sys
import time
import uno
from com.sun.star.beans import PropertyValue

host, port, source_url, target_url = sys.argv[1:5]

def prop(name, value):
    p = PropertyValue()
    p.Name = name
    p.Value = value
    return p

local_ctx = uno.getComponentContext()
resolver = local_ctx.ServiceManager.createInstanceWithContext(
    "com.sun.star.bridge.UnoUrlResolver", local_ctx
)

ctx = None
last_error = None
for _ in range(80):
    try:
        ctx = resolver.resolve(
            "uno:socket,host=%s,port=%s;urp;StarOffice.ComponentContext"
            % (host, port)
        )
        break
    except Exception as exc:
        last_error = exc
        time.sleep(0.25)

if ctx is None:
    raise RuntimeError("Could not connect to LibreOffice UNO server: %s" % last_error)

smgr = ctx.ServiceManager
desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)

load_props = (
    prop("Hidden", True),
    prop("ReadOnly", True),
    prop("UpdateDocMode", 3),  # NO_UPDATE
)

doc = desktop.loadComponentFromURL(source_url, "_blank", 0, load_props)
if doc is None:
    raise RuntimeError("LibreOffice UNO failed to load the source document")

try:
    pdf_data = (
        prop("ExportBookmarks", True),
        prop("ExportBookmarksToPDFDestination", True),
        prop("ConvertOOoTargetToPDFTarget", True),
        prop("OpenBookmarkLevels", -1),
    )
    store_props = (
        prop("FilterName", "writer_pdf_Export"),
        prop("FilterData", uno.Any("[]com.sun.star.beans.PropertyValue", pdf_data)),
        prop("Overwrite", True),
    )
    doc.storeToURL(target_url, store_props)
finally:
    try:
        doc.close(True)
    except Exception:
        try:
            doc.dispose()
        except Exception:
            pass
"""


def _find_free_local_port() -> int:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.bind(("127.0.0.1", 0))
        return sock.getsockname()[1]


def _find_bundled_python(runtime: LibreOfficeRuntime) -> tuple:
    """
    Find LibreOffice's bundled python.exe and its versioned Python home.
    Returns (python_exe, python_home) or (None, None).
    """
    program = runtime.executable.parent
    python_exe = program / "python.exe"
    if not python_exe.is_file():
        return None, None

    homes = sorted(
        (p for p in program.glob("python-core-*") if p.is_dir()),
        reverse=True,
    )
    return python_exe.resolve(), (homes[0].resolve() if homes else None)


def _file_url(path: Path) -> str:
    return path.resolve().as_uri()


def _run_uno_pdf_recovery(
    runtime: LibreOfficeRuntime,
    source: Path,
    target_file: Path,
    work_dir: Path,
    log: logging.Logger,
) -> tuple:
    """
    Start a dedicated headless LibreOffice UNO server and export PDF through
    storeToURL(). The helper runs with LibreOffice's bundled Python so it has
    the matching pyuno module; 3GPP Tools itself does not need a pyuno dependency.

    Returns (success, details).
    """
    python_exe, python_home = _find_bundled_python(runtime)
    if not python_exe:
        return False, "LibreOffice bundled python.exe was not found."

    profile = work_dir / "profile_uno"
    _prepare_isolated_profile(profile)

    port = _find_free_local_port()
    accept = f"socket,host=127.0.0.1,port={port};urp;StarOffice.ServiceManager"

    server_env = _apply_conservative_rendering_environment(
        _clean_runtime_environment(runtime)
    )

    server_cmd = [
        str(runtime.executable),
        f"-env:UserInstallation={profile.as_uri()}",
        "--headless",
        "--invisible",
        "--nodefault",
        "--nofirststartwizard",
        "--nologo",
        "--norestore",
        f"--accept={accept}",
    ]

    helper_path = work_dir / f"uno_pdf_export_{uuid.uuid4().hex}.py"
    helper_path.write_text(UNO_PDF_HELPER, encoding="utf-8")

    helper_env = server_env.copy()
    if python_home:
        helper_env["PYTHONHOME"] = str(python_home)
    helper_env["PYTHONPATH"] = str(runtime.executable.parent)

    log.info(
        "Starting LibreOffice UNO PDF recovery using server=%s, python=%s",
        runtime.executable,
        python_exe,
    )

    server = None
    try:
        server = subprocess.Popen(
            server_cmd,
            stdin=subprocess.DEVNULL,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            creationflags=_windows_creation_flags(),
            env=server_env,
            cwd=str(runtime.executable.parent),
        )

        helper_cmd = [
            str(python_exe),
            str(helper_path),
            "127.0.0.1",
            str(port),
            _file_url(source),
            _file_url(target_file),
        ]

        helper_started = time.monotonic()
        helper = subprocess.run(
            helper_cmd,
            stdin=subprocess.DEVNULL,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            timeout=150,
            creationflags=_windows_creation_flags(),
            env=helper_env,
            cwd=str(runtime.executable.parent),
        )
        elapsed = time.monotonic() - helper_started

        if _wait_for_output(target_file, timeout=5.0):
            return True, (
                f"UNO PDF export succeeded in {elapsed:.1f}s "
                f"(helper exit code {helper.returncode})."
            )

        server_code = server.poll()
        details = [
            f"UNO helper exit code: {helper.returncode}",
            f"UNO helper stdout: {(helper.stdout or '<empty>').strip()}",
            f"UNO helper stderr: {(helper.stderr or '<empty>').strip()}",
            f"LibreOffice server exit code: {server_code}",
        ]
        return False, "\n".join(details)

    except subprocess.TimeoutExpired as exc:
        return False, f"UNO PDF helper timed out after {exc.timeout} seconds."
    except Exception as exc:
        return False, f"UNO PDF recovery failed to start/run: {exc}"
    finally:
        if server is not None and server.poll() is None:
            try:
                server.terminate()
                server.wait(timeout=5)
            except Exception:
                try:
                    server.kill()
                except Exception:
                    pass

def convert_document_libreoffice(
    source_path: Union[str, Path],
    target_format: str = "docx",
    output_path: Optional[Union[str, Path]] = None,
    logger: Optional[logging.Logger] = None,
) -> Path:
    """
    Convert with installed or Portable LibreOffice.

    PortableApps launcher metadata is used to locate the packaged program, but
    LibreOfficePortable.exe itself is never started. This avoids launcher
    cleanup/single-instance dialogs and interactive console behavior.
    """
    log = logger or logging.getLogger(__name__)
    source = Path(source_path).resolve()
    target_ext = _normalize_target_format(target_format)

    if not source.exists():
        raise FileNotFoundError(f"Source document not found: {source}")

    if source.suffix.lower() == f".{target_ext}":
        return source

    target = (
        Path(output_path).resolve()
        if output_path
        else source.with_suffix(f".{target_ext}")
    )
    if target.exists() and target.stat().st_size > 0:
        return target

    runtime = resolve_libreoffice_runtime()
    if not runtime:
        msg = get_libreoffice_missing_msg()
        log.error(msg)
        raise RuntimeError(msg)

    filter_spec = LO_EXPORT_FILTERS.get(target_ext, target_ext)
    _sanitize_file_attributes(source)
    target.parent.mkdir(parents=True, exist_ok=True)

    if target.exists():
        try:
            _sanitize_file_attributes(target)
            target.unlink()
        except Exception:
            pass

    temp_work_dir = Path(tempfile.mkdtemp(prefix="3gpp_lo_conv_"))
    temp_profile_dir = temp_work_dir / "profile"
    temp_out_dir = temp_work_dir / "out"
    temp_out_dir.mkdir(parents=True, exist_ok=True)
    preserve_work_dir = False

    try:
        # Always isolate the conversion profile. For Portable this intentionally
        # avoids touching Data/settings and therefore avoids PortableApps'
        # normal profile migration/cleanup lifecycle.
        _prepare_isolated_profile(temp_profile_dir)

        args = [
            f"-env:UserInstallation={temp_profile_dir.as_uri()}",
            "--headless",
            "--invisible",
            "--nodefault",
            "--nofirststartwizard",
            "--nolockcheck",
            "--nologo",
            "--norestore",
            "--convert-to",
            filter_spec,
            str(source),
            "--outdir",
            str(temp_out_dir),
        ]

        expected_file = temp_out_dir / f"{source.stem}.{target_ext}"

        log.info(
            "Starting LibreOffice %s conversion: %s -> .%s using %s",
            "Portable" if runtime.portable else "installed",
            source.name,
            target_ext,
            runtime.executable,
        )

        runtime_env = _clean_runtime_environment(runtime)
        log.info(
            "LibreOffice runtime environment: program=%s, URE_BOOTSTRAP=%s",
            runtime.executable.parent,
            runtime_env.get("URE_BOOTSTRAP", "<not set>"),
        )

        cmd = [str(runtime.executable), *args]

        attempts = [
            ("normal", runtime_env),
            (
                "conservative headless (Skia disabled, svp VCL)",
                _apply_conservative_rendering_environment(runtime_env),
            ),
        ]

        failure_records = []
        conversion_succeeded = False

        for attempt_index, (attempt_name, attempt_env) in enumerate(attempts, start=1):
            # Each retry gets a completely fresh LibreOffice profile.  A crash
            # can leave lock/state files behind, so never reuse the failed one.
            if attempt_index > 1:
                retry_profile = temp_work_dir / f"profile_retry_{attempt_index}"
                _prepare_isolated_profile(retry_profile)
                args[0] = f"-env:UserInstallation={retry_profile.as_uri()}"
                cmd = [str(runtime.executable), *args]
                active_profile = retry_profile

                # Remove partial output/lock files left by the crashed attempt.
                for child in temp_out_dir.iterdir():
                    try:
                        if child.is_file():
                            child.unlink()
                    except OSError:
                        pass
            else:
                active_profile = temp_profile_dir

            log.info(
                "LibreOffice PDF export attempt %d/%d: %s",
                attempt_index,
                len(attempts),
                attempt_name,
            )
            log.info(
                "LibreOffice runtime environment: program=%s, "
                "URE_BOOTSTRAP=%s, SAL_USE_VCLPLUGIN=%s, SAL_DISABLESKIA=%s, PATH=%s",
                runtime.executable.parent,
                attempt_env.get("URE_BOOTSTRAP", "<not set>"),
                attempt_env.get("SAL_USE_VCLPLUGIN", "<not set>"),
                attempt_env.get("SAL_DISABLESKIA", "<not set>"),
                attempt_env.get("PATH", "<not set>"),
            )

            started = time.monotonic()
            try:
                proc = subprocess.run(
                    cmd,
                    stdin=subprocess.DEVNULL,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    text=True,
                    timeout=120,
                    creationflags=_windows_creation_flags(),
                    env=attempt_env,
                    cwd=str(runtime.executable.parent),
                )
                elapsed = time.monotonic() - started
            except subprocess.TimeoutExpired as exc:
                failure_records.append(
                    f"Attempt {attempt_index} ({attempt_name}) timed out after "
                    f"120 seconds."
                )
                if attempt_index < len(attempts):
                    log.warning(
                        "LibreOffice attempt %d timed out. Retrying with "
                        "conservative headless rendering...",
                        attempt_index,
                    )
                    continue
                raise RuntimeError(
                    f"LibreOffice conversion timed out after 120 seconds while "
                    f"converting {source.name} to .{target_ext}."
                ) from exc

            if _wait_for_output(expected_file):
                conversion_succeeded = True
                if attempt_index > 1:
                    log.info(
                        "LibreOffice conversion recovered successfully using "
                        "conservative headless rendering."
                    )
                break

            failure_records.append(
                f"Attempt {attempt_index} ({attempt_name}), exit code "
                f"{proc.returncode}:\n{_process_details(proc)}"
            )

            if attempt_index < len(attempts):
                log.warning(
                    "LibreOffice attempt %d did not produce a valid PDF. "
                    "Retrying once with Skia disabled and the svp headless "
                    "VCL backend...",
                    attempt_index,
                )
                continue

            # CLI conversion has now failed twice. For PDF only, make one
            # architecturally different recovery attempt through LibreOffice's
            # UNO API. This avoids --convert-to and explicitly calls
            # storeToURL(writer_pdf_Export).
            if target_ext == "pdf":
                log.warning(
                    "LibreOffice CLI PDF export failed twice. "
                    "Attempting UNO storeToURL recovery..."
                )
                uno_ok, uno_details = _run_uno_pdf_recovery(
                    runtime=runtime,
                    source=source,
                    target_file=expected_file,
                    work_dir=temp_work_dir,
                    log=log,
                )
                log.info("LibreOffice UNO recovery result: %s", uno_details)
                if uno_ok:
                    conversion_succeeded = True
                    break
                failure_records.append(
                    "UNO storeToURL recovery:\n" + uno_details
                )

            preserve_work_dir = True
            diagnostic_dir = temp_work_dir / "diagnostics"
            report = _write_failure_diagnostics(
                diagnostic_dir=diagnostic_dir,
                runtime=runtime,
                cmd=cmd,
                runtime_env=attempt_env,
                source=source,
                expected_file=expected_file,
                temp_profile_dir=active_profile,
                temp_out_dir=temp_out_dir,
                proc=proc,
                elapsed=elapsed,
            )
            raise RuntimeError(
                f"LibreOffice failed to convert {source.name} to .{target_ext} "
                f"after {len(attempts)} attempts.\n"
                + "\n\n".join(failure_records)
                + f"\nDiagnostic artifacts preserved at: {temp_work_dir}"
                + f"\nDiagnostic report: {report}"
            )

        if not conversion_succeeded:
            raise RuntimeError(
                f"LibreOffice did not produce a valid .{target_ext} output."
            )

        shutil.copy2(expected_file, target)
        _sanitize_file_attributes(target)

        log.info(
            "Successfully converted via LibreOffice%s: %s -> %s",
            " Portable" if runtime.portable else "",
            source.name,
            target.name,
        )
        return target

    finally:
        if preserve_work_dir:
            log.warning(
                "LibreOffice diagnostic working directory preserved: %s",
                temp_work_dir,
            )
        else:
            shutil.rmtree(temp_work_dir, ignore_errors=True)


def convert_doc_to_docx_libreoffice(
    doc_path: Union[str, Path],
    output_path: Optional[Union[str, Path]] = None,
    logger: Optional[logging.Logger] = None,
) -> Path:
    return convert_document_libreoffice(
        source_path=doc_path,
        target_format="docx",
        output_path=output_path,
        logger=logger,
    )
