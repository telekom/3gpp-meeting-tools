import logging
import os
import platform
import shutil
import stat
import subprocess
import tempfile
from pathlib import Path
from typing import Optional, Union

LIBREOFFICE_DOWNLOAD_URL = "https://portableapps.com/apps/office/libreoffice_portable"


def find_libreoffice_executable() -> Optional[Path]:
    """Scans system PATH and standard OS installation paths for the soffice binary."""
    for binary_name in ("soffice", "soffice.exe", "libreoffice"):
        found = shutil.which(binary_name)
        if found:
            return Path(found).resolve()

    if platform.system() == "Windows":
        candidates = [
            Path(os.environ.get("ProgramFiles", r"C:\Program Files")) / "LibreOffice" / "program" / "soffice.exe",
            Path(os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")) / "LibreOffice" / "program" / "soffice.exe",
            Path(os.environ.get("LOCALAPPDATA", "")) / "Programs" / "LibreOffice" / "program" / "soffice.exe",
        ]
        for candidate in candidates:
            if candidate.is_file():
                return candidate.resolve()

    elif platform.system() == "Darwin":
        mac_path = Path("/Applications/LibreOffice.app/Contents/MacOS/soffice")
        if mac_path.is_file():
            return mac_path.resolve()

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
    """Returns True if LibreOffice is installed and executable."""
    return find_libreoffice_executable() is not None


def get_libreoffice_missing_msg() -> str:
    """Returns user instructions for installing LibreOffice."""
    return (
        "❌ LibreOffice is not installed or could not be found.\n\n"
        "Headless document conversion requires LibreOffice.\n"
        f"Download and install it from: {LIBREOFFICE_DOWNLOAD_URL}\n"
        "Or install via terminal:\n"
        "  - Windows (PowerShell): winget install TheDocumentFoundation.LibreOffice\n"
        "  - Linux: sudo apt-get install libreoffice\n"
        "  - macOS: brew install --cask libreoffice"
    )


def _sanitize_file_attributes(file_path: Path) -> None:
    """Removes read-only flags and NTFS Zone.Identifier streams."""
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
        log.info(f"Successfully converted via LibreOffice: {source.name} -> {target.name}")
        return target

    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)