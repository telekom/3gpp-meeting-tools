# --- File: src/modules/word_tools/core/libreoffice_converter.py ---
import json
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

# Optimized PDF filter preserving links, document bookmarks, and navigation hierarchy
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

# Configuration to suppress printer polling and network lookups
LO_NO_PRINTER_CONFIG = """<?xml version="1.0" encoding="UTF-8"?>
<oor:items xmlns:oor="http://openoffice.org/2001/registry" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <item oor:path="/org.openoffice.Office.Common/Save/Document">
    <prop oor:name="LoadPrinterSetup" oor:type="xs:boolean"><value>false</value></prop>
  </item>
  <item oor:path="/org.openoffice.Office.Writer/Layout/Other">
    <prop oor:name="LoadPrinterSetup" oor:type="xs:boolean"><value>false</value></prop>
  </item>
  <item oor:path="/org.openoffice.Office.Common/Print/Option">
    <prop oor:name="PrinterIndependentLayout" oor:type="xs:string"><value>enabled</value></prop>
  </item>
</oor:items>
"""


def _normalize_target_format(raw_format: str) -> str:
    """Sanitizes compound format tags (e.g. 'docx_libreoffice' -> 'docx')."""
    fmt = raw_format.lower().replace(".", "").strip()
    for suffix in ("_libreoffice", "_word", "_com", "_headless"):
        if fmt.endswith(suffix):
            fmt = fmt[: -len(suffix)]
    return fmt or "docx"


def resolve_soffice_binary(candidate_path: Union[str, Path]) -> Optional[Path]:
    """Validates a file or directory path and resolves the internal soffice binary."""
    if not candidate_path:
        return None

    path = Path(candidate_path).resolve()

    if path.is_file() and path.name.lower().startswith("soffice"):
        return path

    if path.is_file() and "portable" in path.name.lower():
        portable_subpaths = [
            path.parent / "App" / "libreoffice" / "program" / "soffice.exe",
            path.parent / "App" / "LibreOffice64" / "program" / "soffice.exe",
            path.parent / "App" / "LibreOffice" / "program" / "soffice.exe",
        ]
        for sub in portable_subpaths:
            if sub.is_file():
                return sub.resolve()

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
    """Locates the soffice binary from JSON configuration, system PATH, or default OS paths."""
    custom_path = WordConfig.get_libreoffice_path()
    if custom_path:
        resolved = resolve_soffice_binary(custom_path)
        if resolved:
            return resolved

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


def _prepare_isolated_profile(profile_dir: Path) -> None:
    """Pre-seeds an isolated user profile that completely suppresses printer connections."""
    user_config_dir = profile_dir / "user"
    user_config_dir.mkdir(parents=True, exist_ok=True)
    config_file = user_config_dir / "registrymodifications.xcu"
    config_file.write_text(LO_NO_PRINTER_CONFIG, encoding="utf-8")


def convert_document_libreoffice(
    source_path: Union[str, Path],
    target_format: str = "docx",
    output_path: Optional[Union[str, Path]] = None,
    logger: Optional[logging.Logger] = None,
) -> Path:
    """
    Converts a document to the specified format using headless LibreOffice with an
    isolated profile to avoid printer connection hangs and profile locks.
    """
    log = logger or logging.getLogger(__name__)
    source = Path(source_path).resolve()
    target_ext = _normalize_target_format(target_format)

    if not source.exists():
        raise FileNotFoundError(f"Source document not found: {source}")

    if source.suffix.lower() == f".{target_ext}":
        return source

    target = Path(output_path).resolve() if output_path else source.with_suffix(f".{target_ext}")
    if target.exists() and target.stat().st_size > 0:
        return target

    soffice_bin = find_libreoffice_executable()
    if not soffice_bin:
        err_msg = get_libreoffice_missing_msg()
        log.error(err_msg)
        raise RuntimeError(err_msg)

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

    try:
        _prepare_isolated_profile(temp_profile_dir)
        profile_uri = temp_profile_dir.as_uri()

        cmd = [
            str(soffice_bin),
            f"-env:UserInstallation={profile_uri}",
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

        creation_flags = getattr(subprocess, "CREATE_NO_WINDOW", 0x08000000) if platform.system() == "Windows" else 0

        proc = subprocess.run(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            timeout=120,
            creationflags=creation_flags,
        )

        expected_file = temp_out_dir / f"{source.stem}.{target_ext}"
        if not expected_file.exists() or expected_file.stat().st_size == 0:
            err_details = proc.stderr.strip() or proc.stdout.strip()
            raise RuntimeError(
                f"LibreOffice failed to convert {source.name} to .{target_ext} (exit code {proc.returncode}). {err_details}"
            )

        shutil.copy2(expected_file, target)
        _sanitize_file_attributes(target)
        log.info(f"Successfully converted via LibreOffice ({soffice_bin.name}): {source.name} -> {target.name}")
        return target

    finally:
        shutil.rmtree(temp_work_dir, ignore_errors=True)


def convert_doc_to_docx_libreoffice(
    doc_path: Union[str, Path],
    output_path: Optional[Union[str, Path]] = None,
    logger: Optional[logging.Logger] = None,
) -> Path:
    """Backwards-compatible wrapper for .doc -> .docx conversions via LibreOffice."""
    return convert_document_libreoffice(
        source_path=doc_path,
        target_format="docx",
        output_path=output_path,
        logger=logger,
    )