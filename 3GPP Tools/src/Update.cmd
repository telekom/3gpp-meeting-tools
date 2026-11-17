@echo off
setlocal EnableDelayedExpansion

:: ============================================================================
:: 3GPP Delegate Tools — Standalone Self-Contained Updater (Update.cmd)
::
:: Can be placed in `3GPP Tools\src\`, `3GPP Tools\`, or its parent folder.
:: Safely downloads and applies the latest `master` update from GitHub while
:: preserving all local databases (`3gpp_data.db`), proxy/AI settings, user
:: presets, PlantUML JAR, and `output/` files.
:: ============================================================================

set "SCRIPT_DIR=%~dp0"
if "%SCRIPT_DIR:~-1%"=="\" set "SCRIPT_DIR=%SCRIPT_DIR:~0,-1%"

set "APP_ROOT="
if exist "%SCRIPT_DIR%\src\main_tools.py" (
    set "APP_ROOT=%SCRIPT_DIR%"
) else if exist "%SCRIPT_DIR%\..\src\main_tools.py" (
    for %%I in ("%SCRIPT_DIR%\..") do set "APP_ROOT=%%~fI"
) else if exist "%SCRIPT_DIR%\3GPP Tools\src\main_tools.py" (
    set "APP_ROOT=%SCRIPT_DIR%\3GPP Tools"
)

if "%APP_ROOT%"=="" (
    echo ========================================================
    echo [ERROR] Could not locate "3GPP Tools\src\main_tools.py".
    echo Please place Update.cmd inside your "3GPP Tools" or
    echo "3GPP Tools\src" folder and run it again.
    echo ========================================================
    pause
    exit /b 1
)

cd /d "%APP_ROOT%"
set "PYTHONPATH=%APP_ROOT%;%APP_ROOT%\src;%PYTHONPATH%"

:: Detect available Python interpreter (prefer py -3.14, then py -3, then python)
set "PY_CMD="
py -3.14 --version >nul 2>nul
if %ERRORLEVEL% EQU 0 (
    set "PY_CMD=py -3.14"
) else (
    py -3 --version >nul 2>nul
    if %ERRORLEVEL% EQU 0 (
        set "PY_CMD=py -3"
    ) else (
        where python >nul 2>nul
        if %ERRORLEVEL% EQU 0 (
            set "PY_CMD=python"
        )
    )
)

if "%PY_CMD%"=="" (
    echo ========================================================
    echo [ERROR] No Python interpreter found!
    echo Please install 64-bit Python 3.10+ from https://www.python.org/
    echo ========================================================
    pause
    exit /b 1
)

(
    %PY_CMD% -c "import pathlib, sys; p = pathlib.Path(sys.argv[1]); src = p.read_text(encoding='utf-8').split('# === ' + 'EMBEDDED_UPDATER_PY ===', 1)[1]; exec(compile(src, str(p), 'exec'), {'__name__': '__main__', '__file__': str(p)})" "%~f0" "%APP_ROOT%" %*
    set "EXIT_CODE=!ERRORLEVEL!"
    echo.
    if not "%~1"=="--no-pause" pause
    exit /b !EXIT_CODE!
)

# === EMBEDDED_UPDATER_PY ===
import argparse
import datetime
import fnmatch
import hashlib
import json
import os
import re
import shutil
import stat
import subprocess
import sys
import time
import urllib.request
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

GITHUB_SUBDIR = "3GPP Tools"
ATOM_URL = "https://github.com/telekom/3gpp-meeting-tools/commits/master/3GPP%20Tools.atom"
ARCHIVE_URL_TPL = "https://github.com/telekom/3gpp-meeting-tools/archive/{ref}.zip"
APP_VERSION_FILENAME = ".app_version.json"
UPDATE_STATE_FILENAME = "app_update_state.json"

PROTECTED_DIR_PREFIXES = (
    ".git/",
    ".idea/",
    "output/",
    "scratch/",
    "test_meeting_dir/",
    "src/export/",
)
PROTECTED_EXACT_PATHS = {
    ".app_version.json",
    ".editor_cache.puml",
    "3gpp_tools.log",
    "src/3gpp_tools.log",
    "src/assets/plantuml.jar",
    "src/assets/plantuml.version",
    "src/assets/3gpp_70.dotx",
    "src/config/cli_user_context.md",
}
PROTECTED_GLOB_PATTERNS = (
    "src/config/*_config.json",
    "src/config/*_user.json",
    "src/assets/template_3gpp_cr*",
    "src/assets/pseudo-cr-form*",
    "src/assets/*.zip",
    "src/assets/*.doc",
    "src/db/*.db*",
    "*.tmp",
    "*.part",
    "*.pyc",
)
SEED_IF_MISSING_PATHS = {
    "src/config/useful_links.json",
    "src/config/export_config.json",
    "src/db/3gpp_data.db",
}


def _ext_path(p) -> str:
    p_str = os.path.abspath(str(p))
    if os.name == "nt" and not p_str.startswith("\\\\?\\"):
        if p_str.startswith("\\\\"):
            return "\\\\?\\UNC\\" + p_str[2:]
        return "\\\\?\\" + p_str
    return p_str


def _clear_ro(path_str: str) -> None:
    try:
        os.chmod(path_str, stat.S_IWRITE | stat.S_IREAD)
    except OSError:
        pass


def _rmtree_safe(target: Path) -> None:
    ext = _ext_path(target)
    if not os.path.exists(ext):
        return

    def _onerror(func, path, _exc):
        _clear_ro(path)
        try:
            func(path)
        except OSError:
            pass

    try:
        shutil.rmtree(ext, onerror=_onerror)
    except OSError:
        pass


def _norm_rel(rel_path: str) -> str:
    return rel_path.replace("\\", "/").lstrip("./")


def _is_ignored_cache(rel_posix: str) -> bool:
    norm = _norm_rel(rel_posix).lower()
    parts = [p for p in norm.split("/") if p]
    if "__pycache__" in parts or ".ipynb_checkpoints" in parts:
        return True
    return norm.endswith((".pyc", ".pyo", ".tmp", ".part"))


def _is_protected(rel_posix: str) -> bool:
    low = _norm_rel(rel_posix).lower()
    if _is_ignored_cache(low):
        return True
    for prefix in PROTECTED_DIR_PREFIXES:
        if low.startswith(prefix) or low == prefix.rstrip("/"):
            return True
    if low in PROTECTED_EXACT_PATHS:
        return True
    for pat in PROTECTED_GLOB_PATTERNS:
        if fnmatch.fnmatch(low, pat):
            return True
    return False


def _is_seed_only(rel_posix: str) -> bool:
    return _norm_rel(rel_posix).lower() in SEED_IF_MISSING_PATHS


def _sha256_file(path: Path) -> str:
    ext = _ext_path(path)
    if not os.path.isfile(ext):
        return ""
    h = hashlib.sha256()
    with open(ext, "rb") as fh:
        for chunk in iter(lambda: fh.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()


def _copy_retry(src: Path, dst: Path, attempts: int = 10, delay_s: float = 0.35) -> None:
    ext_src = _ext_path(src)
    ext_dst = _ext_path(dst)
    os.makedirs(_ext_path(dst.parent), exist_ok=True)
    last_err = None
    for i in range(max(1, attempts)):
        try:
            if os.path.exists(ext_dst):
                _clear_ro(ext_dst)
            shutil.copy2(ext_src, ext_dst)
            return
        except OSError as exc:
            last_err = exc
            _clear_ro(ext_dst)
            if i < attempts - 1:
                time.sleep(delay_s)
    if last_err is not None:
        raise last_err


def _write_text_retry(target: Path, text: str, attempts: int = 8, delay_s: float = 0.25) -> None:
    ext = _ext_path(target)
    os.makedirs(_ext_path(target.parent), exist_ok=True)
    last_err = None
    for i in range(max(1, attempts)):
        try:
            if os.path.exists(ext):
                _clear_ro(ext)
            with open(ext, "w", encoding="utf-8") as fh:
                fh.write(text)
            return
        except OSError as exc:
            last_err = exc
            _clear_ro(ext)
            if i < attempts - 1:
                time.sleep(delay_s)
    if last_err is not None:
        raise last_err


def _fetch_latest_commit_meta() -> dict:
    """Fetch latest master commit metadata using NetworkSession (corporate proxy aware) or urllib."""
    xml_bytes = b""
    try:
        from core.network.session import NetworkSession

        sess = NetworkSession.get_instance()
        headers = NetworkSession.get_request_headers()
        headers["Accept"] = "application/atom+xml, application/xml, text/xml;q=0.9, */*;q=0.8"
        with sess.get(ATOM_URL, headers=headers, timeout=20) as resp:
            resp.raise_for_status()
            xml_bytes = resp.content
    except Exception:
        try:
            req = urllib.request.Request(ATOM_URL, headers={"User-Agent": "3GPP-Tools-Updater/1.0"})
            with urllib.request.urlopen(req, timeout=20) as resp:
                xml_bytes = resp.read()
        except Exception as exc:
            print(f"[WARN] Could not query commit Atom feed ({exc}); falling back to refs/heads/master.")
            return {
                "sha": "master",
                "short_sha": "master",
                "title": "Latest master snapshot",
                "updated_iso": datetime.datetime.now(datetime.timezone.utc).replace(microsecond=0).isoformat(),
            }

    try:
        ns = {"a": "http://www.w3.org/2005/Atom"}
        root = ET.fromstring(xml_bytes)
        entry = root.find("a:entry", ns)
        if entry is not None:
            entry_id = entry.findtext("a:id", "", ns) or ""
            m = re.search(r"([0-9a-fA-F]{40})", entry_id)
            sha = m.group(1).lower() if m else "master"
            title = (entry.findtext("a:title", "", ns) or "").strip() or sha[:7]
            updated_iso = (entry.findtext("a:updated", "", ns) or "").strip()
            return {
                "sha": sha,
                "short_sha": sha[:7],
                "title": title,
                "updated_iso": updated_iso,
            }
    except Exception as exc:
        print(f"[WARN] Could not parse Atom XML ({exc}); falling back to master branch archive.")

    return {
        "sha": "master",
        "short_sha": "master",
        "title": "Latest master snapshot",
        "updated_iso": datetime.datetime.now(datetime.timezone.utc).replace(microsecond=0).isoformat(),
    }


def _download_archive(url: str, dest_zip: Path) -> None:
    os.makedirs(_ext_path(dest_zip.parent), exist_ok=True)
    try:
        from core.network.session import NetworkSession

        last_pct = [-1]

        def _prog(pct: int) -> None:
            step = (int(pct) // 10) * 10
            if step != last_pct[0]:
                last_pct[0] = step
                print(f"  Downloading archive... {pct}%")

        NetworkSession.download_file(url=url, dest_path=dest_zip, timeout=90, progress_cb=_prog, atomic=True)
        return
    except Exception as exc:
        print(f"  [INFO] Direct/fallback download ({exc})...")

    req = urllib.request.Request(url, headers={"User-Agent": "3GPP-Tools-Updater/1.0"})
    tmp_zip = dest_zip.with_suffix(".zip.part")
    with urllib.request.urlopen(req, timeout=90) as resp, open(_ext_path(tmp_zip), "wb") as out_f:
        shutil.copyfileobj(resp, out_f, length=65536)
    os.replace(_ext_path(tmp_zip), _ext_path(dest_zip))


def _extract_3gpp_tools_from_zip(zip_path: Path, extract_root: Path) -> Path:
    dest_root = Path(os.path.abspath(str(extract_root)))
    os.makedirs(_ext_path(dest_root), exist_ok=True)

    with zipfile.ZipFile(_ext_path(zip_path), "r") as zf:
        members = zf.infolist()
        for member in members:
            name = member.filename.replace("\\", "/")
            parts = [p for p in name.split("/") if p]
            if not name or name.startswith("/") or ":" in name or ".." in parts:
                raise ValueError(f"Unsafe archive entry path: {member.filename}")

        has_subdir = any(
            (parts and parts[0] == GITHUB_SUBDIR) or (len(parts) >= 2 and parts[1] == GITHUB_SUBDIR)
            for parts in ([p for p in m.filename.replace("\\", "/").split("/") if p] for m in members)
        )

        for member in members:
            name = member.filename.replace("\\", "/")
            is_dir = member.is_dir() or name.endswith("/")
            parts = [p for p in name.split("/") if p]
            if not parts:
                continue

            if has_subdir:
                if parts[0] == GITHUB_SUBDIR:
                    rel_parts = parts
                elif len(parts) >= 2 and parts[1] == GITHUB_SUBDIR:
                    rel_parts = parts[1:]
                else:
                    continue
            else:
                rel_parts = parts

            target = dest_root.joinpath(*rel_parts)
            ext_target = _ext_path(target)
            if is_dir:
                os.makedirs(ext_target, exist_ok=True)
                continue
            os.makedirs(_ext_path(target.parent), exist_ok=True)
            with zf.open(member, "r") as src_f, open(ext_target, "wb") as dst_f:
                shutil.copyfileobj(src_f, dst_f, length=65536)

    staged = dest_root / GITHUB_SUBDIR
    if os.path.isfile(_ext_path(staged / "src" / "main_tools.py")):
        return staged
    if os.path.isfile(_ext_path(dest_root / "src" / "main_tools.py")):
        return dest_root
    raise FileNotFoundError("Extracted archive did not contain '3GPP Tools/src/main_tools.py'.")


def _apply_update(staged_dir: Path, app_root: Path, commit_meta: dict) -> dict:
    staged_dir = Path(os.path.abspath(str(staged_dir)))
    app_root = Path(os.path.abspath(str(app_root)))

    version_file = app_root / APP_VERSION_FILENAME
    prev_managed = []
    if os.path.isfile(_ext_path(version_file)):
        try:
            with open(_ext_path(version_file), "r", encoding="utf-8") as vf:
                prev_managed = list((json.loads(vf.read()) or {}).get("managed_files") or [])
        except Exception:
            prev_managed = []

    copied = 0
    seeded = 0
    skipped = 0
    removed = 0
    new_managed = []

    ext_staged = _ext_path(staged_dir)
    for root, dirs, files in os.walk(ext_staged):
        dirs[:] = [d for d in dirs if d not in {"__pycache__", ".git", ".idea", ".ipynb_checkpoints"}]
        for fname in files:
            src_full = os.path.join(root, fname)
            rel_posix = _norm_rel(os.path.relpath(src_full, ext_staged))
            src_file = staged_dir / rel_posix

            if _is_ignored_cache(rel_posix):
                continue

            dst_file = app_root / rel_posix
            if _is_seed_only(rel_posix):
                if os.path.exists(_ext_path(dst_file)):
                    skipped += 1
                    continue
                _copy_retry(src_file, dst_file)
                seeded += 1
                continue

            if _is_protected(rel_posix):
                skipped += 1
                continue

            _copy_retry(src_file, dst_file)
            copied += 1
            new_managed.append(rel_posix)

    new_managed_set = set(new_managed)
    for old_rel in prev_managed:
        norm_old = _norm_rel(old_rel)
        if not norm_old or norm_old in new_managed_set:
            continue
        if _is_protected(norm_old) or _is_seed_only(norm_old):
            continue
        old_target = app_root / norm_old
        ext_old = _ext_path(old_target)
        if os.path.isfile(ext_old):
            try:
                _clear_ro(ext_old)
                os.unlink(ext_old)
                removed += 1
            except OSError:
                pass

    now_iso = datetime.datetime.now(datetime.timezone.utc).replace(microsecond=0).isoformat()
    has_git = os.path.exists(_ext_path(app_root / ".git")) or os.path.exists(
        _ext_path(app_root.parent / ".git")
    )
    payload = {
        "installed_sha": str(commit_meta.get("sha") or ""),
        "short_sha": str(commit_meta.get("short_sha") or "")[:7],
        "commit_title": str(commit_meta.get("title") or ""),
        "updated_iso": str(commit_meta.get("updated_iso") or ""),
        "installed_at_iso": now_iso,
        "managed_files": sorted(new_managed),
        "zip_override": bool(has_git),
    }
    _write_text_retry(version_file, json.dumps(payload, indent=2))

    app_data_root = Path.home() / "3GPP_Delegate_Helper"
    state_path = app_data_root / UPDATE_STATE_FILENAME
    state = {}
    if os.path.isfile(_ext_path(state_path)):
        try:
            with open(_ext_path(state_path), "r", encoding="utf-8") as sf:
                state = json.loads(sf.read()) or {}
        except Exception:
            state = {}
    state["last_updated_sha"] = payload["installed_sha"]
    state["last_updated_at_iso"] = now_iso
    state["skipped_sha"] = ""
    state["last_check_ts"] = time.time()
    _write_text_retry(state_path, json.dumps(state, indent=2))

    return {
        "copied": copied,
        "seeded": seeded,
        "skipped": skipped,
        "removed": removed,
        "installed_sha": payload["installed_sha"],
    }


def main() -> int:
    parser = argparse.ArgumentParser(description="Standalone 3GPP Delegate Tools Updater")
    parser.add_argument("script_path", help="Path to Update.cmd")
    parser.add_argument("app_root", help="Resolved 3GPP Tools root directory")
    parser.add_argument("--no-pause", action="store_true", help="Do not pause on exit")
    parser.add_argument("--zip-file", default="", help="Optional local ZIP archive to apply instead of downloading")
    args, _unknown = parser.parse_known_args(sys.argv[1:])

    app_root = Path(os.path.abspath(args.app_root))
    print("=" * 60)
    print("  3GPP Delegate Tools — Standalone Updater")
    print("=" * 60)
    print(f"Target Folder : {app_root}")
    print(f"Python Engine : {sys.executable}")
    print("Note: Please close 3GPP Delegate Tools if it is currently open.\n")

    temp_root = Path.home() / "3GPP_Delegate_Helper" / "temp" / "app_update"
    os.makedirs(_ext_path(temp_root), exist_ok=True)

    if args.zip_file:
        zip_path = Path(os.path.abspath(args.zip_file))
        commit_meta = {
            "sha": "local_zip",
            "short_sha": "local",
            "title": f"Local archive ({zip_path.name})",
            "updated_iso": datetime.datetime.now(datetime.timezone.utc).replace(microsecond=0).isoformat(),
        }
        print(f"Using local archive: {zip_path}")
    else:
        print("1/4 Checking latest version on GitHub (master)...")
        commit_meta = _fetch_latest_commit_meta()
        print(f"    Latest commit: {commit_meta['short_sha']} — {commit_meta['title']}")

        zip_path = temp_root / f"manual_{commit_meta['short_sha']}.zip"
        archive_url = ARCHIVE_URL_TPL.format(ref=commit_meta["sha"])
        print(f"2/4 Downloading update from {archive_url} ...")
        _download_archive(archive_url, zip_path)

    extract_dir = temp_root / f"extract_manual_{commit_meta['short_sha']}"
    if os.path.exists(_ext_path(extract_dir)):
        _rmtree_safe(extract_dir)

    print("3/4 Extracting '3GPP Tools' package (long-path safe)...")
    staged_dir = _extract_3gpp_tools_from_zip(zip_path, extract_dir)

    before_req = _sha256_file(app_root / "requirements.txt")
    after_req = _sha256_file(staged_dir / "requirements.txt")
    req_changed = before_req != after_req

    print("4/4 Applying update (preserving 3gpp_data.db, configs, presets, and output/)...")
    summary = _apply_update(staged_dir, app_root, commit_meta)
    print(
        f"    Done: {summary['copied']} files updated, "
        f"{summary['seeded']} seeded, "
        f"{summary['skipped']} user/protected files preserved, "
        f"{summary['removed']} obsolete removed."
    )

    if req_changed and os.path.isfile(_ext_path(app_root / "requirements.txt")):
        print("\nInstalling updated dependencies from requirements.txt...")
        subprocess.run([sys.executable, "-m", "pip", "install", "-r", str(app_root / "requirements.txt")], cwd=str(app_root))

    _rmtree_safe(extract_dir)
    if not args.zip_file and os.path.isfile(_ext_path(zip_path)):
        try:
            os.unlink(_ext_path(zip_path))
        except OSError:
            pass

    print("\n" + "=" * 60)
    print("  SUCCESS: 3GPP Delegate Tools is now up to date!")
    print("=" * 60)
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(f"\n[ERROR] Update failed: {exc}")
        raise SystemExit(1)
