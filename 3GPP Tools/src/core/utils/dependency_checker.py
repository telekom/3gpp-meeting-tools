# --- File: src/core/utils/dependency_checker.py ---
"""
Startup Dependency Checker for 3GPP Tools.

Inspects the runtime Python environment before heavy or third-party imports
are evaluated in entry points (main_tools, cli_tools, mcp_server).

If any required dependencies are missing or outdated:
  1. Outputs a clean, actionable ASCII banner to stderr/stdout with the exact
     Python interpreter path and pip install command.
  2. Copies the recommended pip command directly to the Windows clipboard.
  3. On Windows (when GUI is permitted), displays an interactive MessageBox
     offering 1-click automatic installation via pip.
  4. Exits cleanly without dumping cryptic Python tracebacks.
"""

from __future__ import annotations

import importlib.util
import logging
import os
import re
import subprocess
import sys
from pathlib import Path
from typing import Dict, List, NamedTuple, Optional, TextIO, Tuple, Union

logger = logging.getLogger(__name__)

# Map distribution names (as found in requirements.txt or PyPI) to the
# top-level Python module name imported in code.
PACKAGE_IMPORT_MAP: Dict[str, str] = {
    "pyqt5": "PyQt5",
    "pywin32": "win32com",
    "python-docx": "docx",
    "requests": "requests",
    "beautifulsoup4": "bs4",
    "openpyxl": "openpyxl",
    "pandas": "pandas",
    "plotly": "plotly",
    "networkx": "networkx",
    "lxml": "lxml",
    "pillow": "PIL",
    "pypdf": "pypdf",
    "pypdfium2": "pypdfium2",
    "markdown": "markdown",
}

# Default canonical requirement list used when requirements.txt cannot be read.
DEFAULT_REQUIREMENTS: List[Tuple[str, str, str]] = [
    ("PyQt5", ">=5.15.0", "5.15.0"),
    ("pywin32", ">=306", "306"),
    ("python-docx", ">=1.1.0", "1.1.0"),
    ("requests", ">=2.25.1", "2.25.1"),
    ("beautifulsoup4", ">=4.12.0", "4.12.0"),
    ("openpyxl", ">=3.1.0", "3.1.0"),
    ("pandas", ">=1.3.0", "1.3.0"),
    ("plotly", ">=5.10.0", "5.10.0"),
    ("networkx", ">=2.8.0", "2.8.0"),
    ("lxml", ">=4.9.0", "4.9.0"),
    ("Pillow", ">=9.0.0", "9.0.0"),
    ("pypdf", ">=5.0.0", "5.0.0"),
    ("pypdfium2", ">=4.0.0", "4.0.0"),
    ("markdown", ">=3.5.0", "3.5.0"),
]


class Requirement(NamedTuple):
    dist_name: str
    module_name: str
    specifier: str
    min_version: Optional[str]


def parse_version_tuple(v_str: str) -> Tuple[int, ...]:
    """Extracts integer version components (e.g. '5.15.11' -> (5, 15, 11))."""
    parts = re.findall(r"\d+", str(v_str or ""))
    return tuple(int(p) for p in parts) if parts else (0,)


def parse_requirement_line(line: str) -> Optional[Requirement]:
    """Parses a single line from requirements.txt into a Requirement object."""
    cleaned = line.split("#", 1)[0].strip()
    if not cleaned:
        return None

    # Match distribution name and optional specifier (e.g., PyQt5>=5.15.0)
    match = re.match(r"^([A-Za-z0-9_\-\.]+)(?:([><=\!~]+)(.+))?$", cleaned)
    if not match:
        return None

    dist_name = match.group(1).strip()
    op = match.group(2) or ""
    ver = (match.group(3) or "").strip()
    specifier = f"{op}{ver}" if op and ver else ""

    min_version: Optional[str] = None
    if ">=" in specifier or "==" in specifier:
        min_version = ver

    dist_key = dist_name.lower()
    module_name = PACKAGE_IMPORT_MAP.get(dist_key, dist_name.replace("-", "_"))

    return Requirement(
        dist_name=dist_name,
        module_name=module_name,
        specifier=specifier,
        min_version=min_version,
    )


def load_requirements(requirements_path: Optional[Path] = None) -> List[Requirement]:
    """Loads requirements from a file or falls back to DEFAULT_REQUIREMENTS."""
    reqs: List[Requirement] = []
    if requirements_path and requirements_path.is_file():
        try:
            text = requirements_path.read_text(encoding="utf-8")
            for line in text.splitlines():
                parsed = parse_requirement_line(line)
                if parsed:
                    reqs.append(parsed)
        except Exception as exc:
            logger.debug("Could not read requirements file %s: %s", requirements_path, exc)

    if not reqs:
        for dist, spec, min_v in DEFAULT_REQUIREMENTS:
            dist_key = dist.lower()
            mod = PACKAGE_IMPORT_MAP.get(dist_key, dist.replace("-", "_"))
            reqs.append(
                Requirement(
                    dist_name=dist,
                    module_name=mod,
                    specifier=spec,
                    min_version=min_v,
                )
            )

    return reqs


def get_installed_version(dist_name: str) -> Optional[str]:
    """Queries importlib.metadata for the installed distribution version."""
    try:
        import importlib.metadata

        return importlib.metadata.version(dist_name)
    except Exception:
        return None


def is_module_available(module_name: str) -> bool:
    """Checks whether a module can be imported without executing its body."""
    try:
        spec = importlib.util.find_spec(module_name)
        return spec is not None
    except (ImportError, AttributeError, ValueError):
        return False


def copy_text_to_clipboard(text: str) -> bool:
    """Copies text to the system clipboard on Windows using clip.exe."""
    if os.name != "nt":
        return False
    try:
        subprocess.run(
            ["clip"],
            input=text,
            text=True,
            check=True,
            capture_output=True,
            creationflags=0x08000000 if hasattr(subprocess, "CREATE_NO_WINDOW") else 0,
        )
        return True
    except Exception:
        return False


def find_requirements_file() -> Optional[Path]:
    """Searches for requirements.txt in the project root or parent folders."""
    current = Path(__file__).resolve().parent
    for parent in [current] + list(current.parents):
        candidate = parent / "requirements.txt"
        if candidate.is_file():
            return candidate
    return None


def check_missing_dependencies(
    requirements_path: Optional[Path] = None,
) -> Tuple[List[Requirement], List[Tuple[Requirement, str]]]:
    """
    Checks the current Python environment against requirements.
    Returns:
        missing_packages: List of Requirement where module is absent.
        outdated_packages: List of (Requirement, installed_version) where version < min_version.
    """
    req_file = requirements_path or find_requirements_file()
    reqs = load_requirements(req_file)

    missing: List[Requirement] = []
    outdated: List[Tuple[Requirement, str]] = []

    for req in reqs:
        if not is_module_available(req.module_name):
            missing.append(req)
            continue

        # If module is available, optionally verify minimum version if metadata exists
        if req.min_version:
            inst_ver = get_installed_version(req.dist_name)
            if inst_ver:
                if parse_version_tuple(inst_ver) < parse_version_tuple(req.min_version):
                    outdated.append((req, inst_ver))

    return missing, outdated


def format_dependency_report(
    missing: List[Requirement],
    outdated: List[Tuple[Requirement, str]],
    requirements_path: Optional[Path] = None,
) -> str:
    """Formats a detailed, human-readable terminal report with instructions."""
    python_exe = sys.executable
    py_version = (
        f"{sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}"
    )

    req_file = requirements_path or find_requirements_file()
    if req_file and req_file.exists():
        req_cmd = f'"{python_exe}" -m pip install -r "{req_file}"'
    else:
        dist_names = [r.dist_name for r in missing] + [r.dist_name for r, _ in outdated]
        req_cmd = f'"{python_exe}" -m pip install {" ".join(dist_names)}'

    lines = [
        "=" * 72,
        "  [3GPP Tools] Missing Required Python Dependencies",
        "=" * 72,
        "The application cannot start because one or more required Python",
        "packages are missing or outdated in this Python environment.",
        "",
        f"  Active Interpreter : {python_exe}",
        f"  Python Version     : {py_version}",
        "",
    ]

    if missing:
        lines.append("Missing Packages:")
        for r in missing:
            spec_str = f" ({r.specifier})" if r.specifier else ""
            lines.append(f"  * {r.dist_name}{spec_str} [module: {r.module_name}]")
        lines.append("")

    if outdated:
        lines.append("Outdated Packages:")
        for r, inst_v in outdated:
            lines.append(
                f"  * {r.dist_name} (installed: {inst_v}, required: {r.specifier})"
            )
        lines.append("")

    lines.extend([
        "Recommended Fix:",
        "  Run the following command in your terminal or PowerShell:",
        "",
        f"    {req_cmd}",
        "",
        "Note: Installing packages using the exact python path above ensures they",
        "are placed into the active Python environment.",
        "=" * 72,
    ])

    return "\n".join(lines)


def run_pip_install(requirements_path: Optional[Path] = None) -> bool:
    """Executes pip install for requirements.txt in the active Python environment."""
    req_file = requirements_path or find_requirements_file()
    cmd = [sys.executable, "-m", "pip", "install"]
    if req_file and req_file.exists():
        cmd.extend(["-r", str(req_file)])
    else:
        for dist, _, _ in DEFAULT_REQUIREMENTS:
            cmd.append(dist)

    print("\n[3GPP Tools] Running automatic dependency installation:")
    print(" ".join(cmd))
    print("-" * 72)
    try:
        proc = subprocess.run(cmd, check=False)
        return proc.returncode == 0
    except Exception as exc:
        print(f"[3GPP Tools] Failed to launch pip: {exc}", file=sys.stderr)
        return False


def _show_windows_message_box(
    missing: List[Requirement],
    outdated: List[Tuple[Requirement, str]],
    requirements_path: Optional[Path] = None,
) -> Optional[bool]:
    """
    On Windows, displays a native modal dialog explaining the missing packages.
    Returns:
        True: User clicked Yes to install automatically.
        False: User clicked No to cancel.
        None: Dialog could not be displayed.
    """
    if os.name != "nt":
        return None

    try:
        import ctypes

        user32 = ctypes.windll.user32
    except Exception:
        return None

    pkg_lines: List[str] = []
    for r in missing:
        spec = f" {r.specifier}" if r.specifier else ""
        pkg_lines.append(f"  • {r.dist_name}{spec}")
    for r, inst_v in outdated:
        pkg_lines.append(f"  • {r.dist_name} ({inst_v} -> {r.specifier})")

    summary_pkgs = "\n".join(pkg_lines[:12])
    if len(pkg_lines) > 12:
        summary_pkgs += f"\n  • ... and {len(pkg_lines) - 12} more"

    python_exe = sys.executable
    req_file = requirements_path or find_requirements_file()
    pip_cmd = (
        f'"{python_exe}" -m pip install -r "{req_file}"'
        if req_file
        else f'"{python_exe}" -m pip install -r requirements.txt'
    )

    # Automatically copy command to clipboard for delegate convenience
    copy_text_to_clipboard(pip_cmd)

    msg = (
        "3GPP Tools cannot start because required Python packages are missing:\n\n"
        f"{summary_pkgs}\n\n"
        f"Active Python:\n{python_exe}\n\n"
        "Would you like 3GPP Tools to install the missing packages now?\n\n"
        "• Click 'Yes' to run pip installation automatically.\n"
        "• Click 'No' to exit (command copied to clipboard)."
    )

    # MB_YESNO (0x04) | MB_ICONWARNING (0x30) | MB_TOPMOST (0x40000) | MB_SETFOREGROUND (0x10000)
    flags = 0x00000004 | 0x00000030 | 0x00040000 | 0x00010000
    res = user32.MessageBoxW(0, msg, "3GPP Tools - Missing Dependencies", flags)

    IDYES = 6
    if res == IDYES:
        return True
    return False


def should_show_gui_dialog(allow_gui: bool) -> bool:
    """Determines whether a modal Windows MessageBox can safely be presented."""
    if not allow_gui or os.name != "nt":
        return False
    if os.environ.get("HEADLESS") == "1" or os.environ.get("CI") == "1":
        return False
    if any(arg in sys.argv for arg in ("--headless", "--no-gui", "--non-interactive")):
        return False
    if "unittest" in sys.modules or "pytest" in sys.modules:
        return False
    return True


def check_dependencies(
    requirements_path: Optional[Path] = None,
    allow_gui: bool = True,
    interactive: bool = True,
    stream: Optional[TextIO] = None,
) -> bool:
    """
    Checks if all required dependencies are installed.
    Returns:
        True: Dependencies satisfied (or successfully installed).
        False: Missing dependencies remain.
    """
    missing, outdated = check_missing_dependencies(requirements_path)
    if not missing and not outdated:
        return True

    out_stream = stream or sys.stderr
    report = format_dependency_report(missing, outdated, requirements_path)
    print(report, file=out_stream, flush=True)

    # Copy the pip install command to clipboard
    python_exe = sys.executable
    req_file = requirements_path or find_requirements_file()
    if req_file and req_file.exists():
        cmd = f'"{python_exe}" -m pip install -r "{req_file}"'
    else:
        dist_names = [r.dist_name for r in missing] + [r.dist_name for r, _ in outdated]
        cmd = f'"{python_exe}" -m pip install {" ".join(dist_names)}'
    copy_text_to_clipboard(cmd)

    # If GUI is permitted and on Windows, prompt the user with native dialog
    if should_show_gui_dialog(allow_gui):
        user_wants_install = _show_windows_message_box(
            missing, outdated, requirements_path
        )
        if user_wants_install is True:
            success = run_pip_install(requirements_path)
            if success:
                # Re-verify dependencies
                m_after, _ = check_missing_dependencies(requirements_path)
                if not m_after:
                    try:
                        import ctypes

                        ctypes.windll.user32.MessageBoxW(
                            0,
                            "Dependencies installed successfully!\n\n3GPP Tools will now start.",
                            "3GPP Tools - Installation Complete",
                            0x00000040 | 0x00040000,  # MB_ICONINFORMATION | MB_TOPMOST
                        )
                    except Exception:
                        pass
                    return True
            else:
                try:
                    import ctypes

                    ctypes.windll.user32.MessageBoxW(
                        0,
                        f"Automatic dependency installation failed.\n\nPlease run manually in your terminal:\n\n{cmd}",
                        "3GPP Tools - Installation Failed",
                        0x00000010 | 0x00040000,  # MB_ICONERROR | MB_TOPMOST
                    )
                except Exception:
                    pass
            return False

        elif user_wants_install is False:
            return False

    # If interactive in a terminal/console (stdin is a TTY)
    if interactive and sys.stdin.isatty():
        try:
            choice = input(
                "\nWould you like to install the missing dependencies now? [y/N]: "
            ).strip()
            if choice.lower() in ("y", "yes"):
                success = run_pip_install(requirements_path)
                if success:
                    m_after, _ = check_missing_dependencies(requirements_path)
                    if not m_after:
                        print("[3GPP Tools] Dependencies installed successfully! Continuing startup...")
                        return True
                else:
                    print(
                        "[3GPP Tools] Installation failed. Please run the command above manually.",
                        file=out_stream,
                    )
                    return False
        except (EOFError, KeyboardInterrupt):
            print("\nAborted.", file=out_stream)
            return False

    return False


def check_dependencies_or_exit(
    requirements_path: Optional[Path] = None,
    allow_gui: bool = True,
    interactive: bool = True,
    exit_code: int = 1,
    stream: Optional[TextIO] = None,
) -> None:
    """Checks dependencies, and if unfulfilled, exits the process cleanly."""
    if not check_dependencies(
        requirements_path=requirements_path,
        allow_gui=allow_gui,
        interactive=interactive,
        stream=stream,
    ):
        sys.exit(exit_code)
