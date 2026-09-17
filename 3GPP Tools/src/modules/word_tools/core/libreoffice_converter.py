# --- File: src/modules/word_tools/core/libreoffice_converter.py ---
"""Small LibreOffice fallback used by Word Tools."""
import configparser
import json
import logging
import os
import platform
import shutil
import stat
import subprocess
import tempfile
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


@dataclass(frozen=True)
class LibreOfficeRuntime:
    executable: Path
    portable: bool = False
    portable_root: Optional[Path] = None


def _normalize_target_format(value: str) -> str:
    fmt = value.lower().replace(".", "").strip()
    for suffix in ("_libreoffice", "_word", "_com", "_headless"):
        if fmt.endswith(suffix):
            fmt = fmt[:-len(suffix)]
    return fmt or "docx"


def _find_portable_root(path: Path) -> Optional[Path]:
    current = path if path.is_dir() else path.parent
    for _ in range(10):
        if (
            (current / "LibreOfficePortable.exe").is_file()
            and (current / "App" / "AppInfo").is_dir()
        ):
            return current.resolve()
        if current.parent == current:
            break
        current = current.parent
    return None


def _portable_soffice(root: Path) -> Optional[Path]:
    launcher_ini = root / "App" / "AppInfo" / "Launcher" / "LibreOfficePortable.ini"
    if launcher_ini.is_file():
        parser = configparser.ConfigParser(interpolation=None)
        try:
            parser.read(launcher_ini, encoding="utf-8")
            for key in (
                "ProgramExecutableWhenParameters64",
                "ProgramExecutable64",
                "ProgramExecutableWhenParameters",
                "ProgramExecutable",
            ):
                value = parser.get("Launch", key, fallback="").strip().strip('"')
                if value:
                    candidate = root / "App" / Path(value.replace("\\", os.sep))
                    if candidate.is_file() and candidate.name.lower().startswith("soffice"):
                        return candidate.resolve()
        except Exception:
            pass

    for candidate in (
        root / "App" / "libreoffice" / "program" / "soffice.exe",
        root / "App" / "LibreOffice64" / "program" / "soffice.exe",
        root / "App" / "LibreOffice" / "program" / "soffice.exe",
    ):
        if candidate.is_file():
            return candidate.resolve()
    return None


def _resolve_runtime(candidate: Optional[Union[str, Path]] = None) -> Optional[LibreOfficeRuntime]:
    if candidate:
        path = Path(candidate).expanduser().resolve()
        root = _find_portable_root(path)
        if root:
            soffice = _portable_soffice(root)
            return LibreOfficeRuntime(soffice, True, root) if soffice else None
        if path.is_file() and path.name.lower().startswith(("soffice", "libreoffice")):
            return LibreOfficeRuntime(path)
        if path.is_dir():
            for item in (path / "program" / "soffice.exe", path / "soffice.exe"):
                if item.is_file():
                    return LibreOfficeRuntime(item.resolve())
        return None

    configured = WordConfig.get_libreoffice_path()
    if configured:
        runtime = _resolve_runtime(configured)
        if runtime:
            return runtime

    for name in ("soffice", "soffice.exe", "libreoffice"):
        found = shutil.which(name)
        if found:
            return LibreOfficeRuntime(Path(found).resolve())

    if platform.system() == "Windows":
        for candidate_path in (
            Path(os.environ.get("ProgramFiles", r"C:\Program Files")) / "LibreOffice" / "program" / "soffice.exe",
            Path(os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")) / "LibreOffice" / "program" / "soffice.exe",
            Path(os.environ.get("LOCALAPPDATA", "")) / "Programs" / "LibreOffice" / "program" / "soffice.exe",
        ):
            if candidate_path.is_file():
                return LibreOfficeRuntime(candidate_path.resolve())
    return None


def resolve_soffice_binary(candidate_path: Union[str, Path]) -> Optional[Path]:
    """UI-compatible resolver; preserve PortableApps launcher as the saved path."""
    path = Path(candidate_path).expanduser().resolve()
    root = _find_portable_root(path)
    if root:
        return (root / "LibreOfficePortable.exe").resolve()
    runtime = _resolve_runtime(path)
    return runtime.executable if runtime else None


def find_libreoffice_executable() -> Optional[Path]:
    configured = WordConfig.get_libreoffice_path()
    if configured:
        resolved = resolve_soffice_binary(configured)
        if resolved:
            return resolved
    runtime = _resolve_runtime()
    if not runtime:
        return None
    if runtime.portable and runtime.portable_root:
        return (runtime.portable_root / "LibreOfficePortable.exe").resolve()
    return runtime.executable


def is_libreoffice_available() -> bool:
    return _resolve_runtime() is not None


def get_libreoffice_missing_msg() -> str:
    return (
        "❌ LibreOffice is not installed or could not be found.\n\n"
        "Headless conversion requires LibreOffice (Installed or Portable).\n"
        f"Download LibreOffice Portable from: {LIBREOFFICE_DOWNLOAD_URL}\n"
        "For PortableApps select LibreOfficePortable.exe."
    )


def _clean_environment(runtime: LibreOfficeRuntime) -> dict:
    env = os.environ.copy()
    for key in ("PYTHONHOME", "PYTHONPATH", "PYTHONEXECUTABLE", "__PYVENV_LAUNCHER__", "VIRTUAL_ENV"):
        env.pop(key, None)

    if runtime.portable and platform.system() == "Windows":
        system_root = Path(env.get("SystemRoot", r"C:\Windows"))
        entries = [
            runtime.executable.parent,
            system_root / "System32",
            system_root,
            system_root / "System32" / "Wbem",
        ]
        env["PATH"] = os.pathsep.join(str(p) for p in entries if p.exists())
    return env


def _sanitize(path: Path) -> None:
    try:
        if path.exists():
            os.chmod(path, stat.S_IWRITE | stat.S_IREAD)
    except Exception:
        pass


def convert_document_libreoffice(
    source_path: Union[str, Path],
    target_format: str = "docx",
    output_path: Optional[Union[str, Path]] = None,
    logger: Optional[logging.Logger] = None,
) -> Path:
    """One headless LibreOffice conversion attempt with an isolated profile."""
    log = logger or logging.getLogger(__name__)
    source = Path(source_path).resolve()
    target_ext = _normalize_target_format(target_format)

    if not source.exists():
        raise FileNotFoundError(f"Source document not found: {source}")
    if source.suffix.lower() == f".{target_ext}":
        return source

    target = Path(output_path).resolve() if output_path else source.with_suffix(f".{target_ext}")
    runtime = _resolve_runtime()
    if not runtime:
        raise RuntimeError(get_libreoffice_missing_msg())

    target.parent.mkdir(parents=True, exist_ok=True)
    if target.exists():
        _sanitize(target)
        target.unlink()

    work = Path(tempfile.mkdtemp(prefix="3gpp_lo_"))
    profile = work / "profile"
    out_dir = work / "out"
    out_dir.mkdir(parents=True)
    expected = out_dir / f"{source.stem}.{target_ext}"

    try:
        filter_spec = LO_EXPORT_FILTERS.get(target_ext, target_ext)
        cmd = [
            str(runtime.executable),
            f"-env:UserInstallation={profile.as_uri()}",
            "--headless",
            "--invisible",
            "--nodefault",
            "--nofirststartwizard",
            "--nologo",
            "--norestore",
            "--convert-to",
            filter_spec,
            str(source),
            "--outdir",
            str(out_dir),
        ]
        log.info(
            "LibreOffice fallback: %s -> .%s using %s",
            source.name, target_ext, runtime.executable,
        )
        proc = subprocess.run(
            cmd,
            stdin=subprocess.DEVNULL,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            timeout=120,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0) if platform.system() == "Windows" else 0,
            env=_clean_environment(runtime),
            cwd=str(runtime.executable.parent),
        )

        if not expected.exists() or expected.stat().st_size == 0:
            details = (proc.stderr or proc.stdout or "LibreOffice produced no output.").strip()
            raise RuntimeError(
                f"LibreOffice failed to convert {source.name} to .{target_ext} "
                f"(exit code {proc.returncode}). {details}"
            )

        shutil.copy2(expected, target)
        _sanitize(target)
        return target
    finally:
        shutil.rmtree(work, ignore_errors=True)


def convert_doc_to_docx_libreoffice(
    doc_path: Union[str, Path],
    output_path: Optional[Union[str, Path]] = None,
    logger: Optional[logging.Logger] = None,
) -> Path:
    return convert_document_libreoffice(
        doc_path, target_format="docx", output_path=output_path, logger=logger
    )
