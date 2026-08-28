# --- File: src/modules/word_tools/core/libreoffice_converter.py ---
import logging
import os
import platform
import shutil
import stat
import subprocess
import tempfile
from pathlib import Path
from typing import Optional, Union

from modules.word_tools.core.word_config import WordConfig

LIBREOFFICE_DOWNLOAD_URL = "https://portableapps.com/apps/office/libreoffice_portable"


def resolve_soffice_binary(candidate_path: Union[str, Path]) -> Optional[Path]:
    """
    Validates a file or directory path. If given LibreOfficePortable.exe or an
    extracted PortableApps folder, resolves the internal soffice.exe binary.
    """
    if not candidate_path:
        return None

    path = Path(candidate_path).resolve()

    # 1. Direct soffice executable
    if path.is_file() and path.name.lower().startswith("soffice"):
        return path

    # 2. PortableApps launcher (e.g., LibreOfficePortable.exe)
    if path.is_file() and "portable" in path.name.lower():
        portable_subpaths = [
            path.parent / "App" / "libreoffice" / "program" / "soffice.exe",
            path.parent / "App" / "LibreOffice64" / "program" / "soffice.exe",
            path.parent / "App" / "LibreOffice" / "program" / "soffice.exe",
        ]
        for sub in portable_subpaths:
            if sub.is_file():
                return sub.resolve()

    # 3. Directory selection (user selected the root install/portable folder)
    if path.is_dir():
        dir_subpaths = [
            path / "program" / "soffice.exe",
            path / "App" / "libreoffice" / "program" / "soffice.exe",
            path / "App" / "LibreOffice64" / "program" / "soffice.exe",
            path / "App" / "LibreOffice" / "program" / "soffice.exe",
            path / "soffice.exe",
        ]
        for sub in dir_subpaths:
            if sub.is_file():
                return sub.resolve()

    if path.is_file() and os.access(path, os.X_OK):
        return path

    return None


def find_libreoffice_executable() -> Optional[Path]:
    """
    Finds the soffice binary by checking:
    1. Saved JSON configuration (custom / portable path)
    2. System PATH
    3. Standard OS installation directories
    """
    # 1. Check custom path in word_config.json
    custom_path = WordConfig.get_libreoffice_path()
    if custom_path:
        resolved = resolve_soffice_binary(custom_path)
        if resolved:
            return resolved

    # 2. Check system PATH
    for binary_name in ("soffice", "soffice.exe", "libreoffice"):
        found = shutil.which(binary_name)
        if found:
            return Path(found).resolve()

    # 3. Standard Windows paths
    if platform.system() == "Windows":
        candidates = [
            Path(os.environ.get("ProgramFiles", r"C:\Program Files")) / "LibreOffice" / "program" / "soffice.exe",
            Path(os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")) / "LibreOffice" / "program" / "soffice.exe",
            Path(os.environ.get("LOCALAPPDATA", "")) / "Programs" / "LibreOffice" / "program" / "soffice.exe",
        ]
        for candidate in candidates:
            if candidate.is_file():
                return candidate.resolve()

    # 4. macOS bundle
    elif platform.system() == "Darwin":
        mac_path = Path("/Applications/LibreOffice.app/Contents/MacOS/soffice")
        if mac_path.is_file():
            return mac_path.resolve()

    # 5. Linux paths
    elif platform.system() == "Linux":
        linux_paths = [
            Path("/usr/bin/soffice"),
            Path("/usr/bin/libreoffice"),
            Path("/usr/local/bin/soffice"),
            Path("/usr/local/bin/libreoffice"),
            Path("/opt/libreoffice/program/soffice"),
        ]
        for candidate in linux_paths:
            if candidate.is_file():
                return candidate.resolve()

    return None


def is_libreoffice_available() -> bool:
    return find_libreoffice_executable() is not None


def get_libreoffice_missing_msg() -> str:
    return (
        "❌ LibreOffice is not installed or could not be found.\n\n"
        "Headless document conversion requires LibreOffice (Installed or Portable).\n"
        f"Download LibreOffice Portable from: {LIBREOFFICE_DOWNLOAD_URL}\n"
        "After downloading/extracting, use the '📂 Locate Executable' button in the Word tab to select it."
    )


def _sanitize_file_attributes(file_path: Path) -> None:
    try:
        if file_path.exists():
            os.chmod(file_path, stat.S_IWRITE | stat.S_IREAD)
    except Exception:
        pass

    try:
        zone_stream = Path(f"{file_path.resolve()}:Zone.Identifier")
        if zone_stream.exists():
            zone_stream.unlink()
    except Exception:
        pass


def convert_doc_to_docx_libreoffice(
    doc_path: Union[str, Path],
    output_path: Optional[Union[str, Path]] = None,
    logger: Optional[logging.Logger] = None,
) -> Path:
    """Synchronously converts legacy .doc to .docx using headless LibreOffice."""
    log = logger or logging.getLogger(__name__)
    source = Path(doc_path).resolve()

    if not source.exists():
        raise FileNotFoundError(f"Source document not found: {source}")

    if source.suffix.lower() == ".docx":
        return source

    target = Path(output_path).resolve() if output_path else source.with_suffix(".docx")
    if target.exists() and target.stat().st_size > 0:
        return target

    soffice_bin = find_libreoffice_executable()
    if not soffice_bin:
        err_msg = get_libreoffice_missing_msg()
        log.error(err_msg)
        raise RuntimeError(err_msg)

    _sanitize_file_attributes(source)
    target.parent.mkdir(parents=True, exist_ok=True)
    if target.exists():
        try:
            _sanitize_file_attributes(target)
            target.unlink()
        except Exception:
            pass

    temp_dir = Path(tempfile.mkdtemp(prefix="3gpp_lo_doc2docx_"))
    try:
        cmd = [
            str(soffice_bin),
            "--headless",
            "--invisible",
            "--nodefault",
            "--nofirststartwizard",
            "--nolockcheck",
            "--nologo",
            "--norestore",
            "--convert-to",
            "docx:MS Word 2007 XML",
            str(source),
            "--outdir",
            str(temp_dir),
        ]

        creation_flags = getattr(subprocess, "CREATE_NO_WINDOW", 0x08000000) if platform.system() == "Windows" else 0

        proc = subprocess.run(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            timeout=120,
            creationflags=creation_flags,
        )

        expected_file = temp_dir / f"{source.stem}.docx"
        if not expected_file.exists() or expected_file.stat().st_size == 0:
            err_details = proc.stderr.strip() or proc.stdout.strip()
            raise RuntimeError(
                f"LibreOffice failed to convert {source.name} to .docx (exit code {proc.returncode}). {err_details}"
            )

        shutil.copy2(expected_file, target)
        _sanitize_file_attributes(target)
        log.info(f"Successfully converted via LibreOffice ({soffice_bin.name}): {source.name} -> {target.name}")
        return target

    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)