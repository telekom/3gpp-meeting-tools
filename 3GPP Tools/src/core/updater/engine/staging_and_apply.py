# --- File: src/core/updater/engine/staging_and_apply.py ---
"""ZIP download/ZipSlip staging, Git fast-forward apply, and detached relaunch."""

from __future__ import annotations

import hashlib
import json
import logging
import os
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path
from typing import Callable, Dict, List, Optional, Tuple

from core.config.config import GITHUB_ARCHIVE_URL_TEMPLATE, GITHUB_BRANCH, GITHUB_SUBDIR
from core.network.session import DownloadCancelledError, NetworkSession
from core.updater.apply_worker import (
    SEED_IF_MISSING_PATHS,
    _copy_with_retry,
    normalize_rel_path,
    rmtree_extended,
)
from core.utils.paths import get_app_data_root, get_project_root, get_temp_root, to_extended_path

from .git_state import inspect_local_install
from .models import CommitInfo, LocalInstallInfo, StagedUpdateInfo, _facade

logger = logging.getLogger(__name__)


def _safe_extract_zip(
    zip_path: Path,
    dest_dir: Path,
    subdir_filter: Optional[str] = GITHUB_SUBDIR,
) -> None:
    """
    Extract ``zip_path`` into ``dest_dir`` with ZipSlip path traversal protection,
    optional `3GPP Tools/` subtree filtering, top-level `<repo>-<sha>/` prefix
    stripping, and Windows extended-length (`\\\\?\\`) path support.
    """
    dest_root = Path(os.path.abspath(str(dest_dir)))
    os.makedirs(to_extended_path(dest_root), exist_ok=True)

    with zipfile.ZipFile(to_extended_path(zip_path), "r") as zf:
        members = zf.infolist()

        # Pass 1: Validate every entry in the archive against ZipSlip traversal
        for member in members:
            member_name = member.filename.replace("\\", "/")
            parts = [p for p in member_name.split("/") if p]
            if (
                not member_name
                or member_name.startswith("/")
                or ":" in member_name
                or ".." in parts
            ):
                raise ValueError(f"Unsafe archive entry path: {member.filename}")
            target = Path(os.path.abspath(str(dest_root / member_name)))
            try:
                target.relative_to(dest_root)
            except ValueError as exc:
                raise ValueError(
                    f"Archive entry escapes target directory: {member.filename}"
                ) from exc

        # Pass 2: Check whether the archive contains `subdir_filter` ("3GPP Tools")
        # so we can skip sibling projects (e.g. `3GPP Meeting Helper`) and strip the
        # redundant 60-character `<repo>-<40-char-sha>/` wrapper folder.
        filter_norm = (subdir_filter or "").strip().strip("/")
        has_subdir = False
        if filter_norm:
            for member in members:
                parts = [p for p in member.filename.replace("\\", "/").split("/") if p]
                if (parts and parts[0] == filter_norm) or (
                    len(parts) >= 2 and parts[1] == filter_norm
                ):
                    has_subdir = True
                    break

        # Pass 3: Stream-extract using Windows extended-length paths
        for member in members:
            member_name = member.filename.replace("\\", "/")
            is_dir = member.is_dir() or member_name.endswith("/")
            parts = [p for p in member_name.split("/") if p]
            if not parts:
                continue

            if has_subdir and filter_norm:
                if parts[0] == filter_norm:
                    rel_parts = parts
                elif len(parts) >= 2 and parts[1] == filter_norm:
                    # Strip top-level `<repo>-<sha>/` wrapper directory
                    rel_parts = parts[1:]
                else:
                    continue
            else:
                rel_parts = parts

            target_path = dest_root.joinpath(*rel_parts)
            ext_target = to_extended_path(target_path)
            if is_dir:
                os.makedirs(ext_target, exist_ok=True)
                continue

            os.makedirs(to_extended_path(target_path.parent), exist_ok=True)
            with zf.open(member, "r") as src_f, open(ext_target, "wb") as dst_f:
                shutil.copyfileobj(src_f, dst_f, length=65536)


def _locate_staged_app_dir(extract_root: Path) -> Path:
    """Locate `3GPP Tools/` (containing `src/main_tools.py`) inside an extracted archive."""
    direct_subdir = extract_root / GITHUB_SUBDIR
    if os.path.isfile(to_extended_path(direct_subdir / "src" / "main_tools.py")):
        return Path(os.path.abspath(str(direct_subdir)))

    for child in extract_root.iterdir():
        if not os.path.isdir(to_extended_path(child)):
            continue
        candidate = child / GITHUB_SUBDIR
        if os.path.isfile(to_extended_path(candidate / "src" / "main_tools.py")):
            return Path(os.path.abspath(str(candidate)))
        if os.path.isfile(to_extended_path(child / "src" / "main_tools.py")):
            return Path(os.path.abspath(str(child)))
    if os.path.isfile(to_extended_path(extract_root / "src" / "main_tools.py")):
        return Path(os.path.abspath(str(extract_root)))
    raise FileNotFoundError(
        "Downloaded archive does not contain the expected '3GPP Tools/src/main_tools.py' structure."
    )


def _sha256_file(path: Path) -> str:
    ext_p = to_extended_path(path)
    if not os.path.isfile(ext_p):
        return ""
    h = hashlib.sha256()
    with open(ext_p, "rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()


def stage_zip_update(
    target_commit: CommitInfo,
    local_info: Optional[LocalInstallInfo] = None,
    progress_cb: Optional[Callable[[int], None]] = None,
    cancel_cb: Optional[Callable[[], bool]] = None,
    timeout: int = 60,
) -> StagedUpdateInfo:
    """Download commit-pinned ZIP archive and extract/validate into temp staging."""
    info = local_info or inspect_local_install()
    work_root = _facade("get_temp_root", get_temp_root)() / "app_update"
    os.makedirs(to_extended_path(work_root), exist_ok=True)

    zip_path = work_root / f"{target_commit.short_sha}.zip"
    extract_root = work_root / f"extract_{target_commit.short_sha}"
    if os.path.exists(to_extended_path(extract_root)):
        rmtree_extended(extract_root)
    os.makedirs(to_extended_path(extract_root), exist_ok=True)

    archive_url = GITHUB_ARCHIVE_URL_TEMPLATE.format(ref=target_commit.sha)
    logger.info("[Updater] Downloading update archive: %s", archive_url)
    net_cls = _facade("NetworkSession", NetworkSession)
    net_cls.download_file(
        url=archive_url,
        dest_path=zip_path,
        timeout=timeout,
        progress_cb=progress_cb,
        cancel_cb=cancel_cb,
        atomic=True,
    )

    if cancel_cb and cancel_cb():
        raise DownloadCancelledError("Update cancelled before extraction.")

    _safe_extract_zip(zip_path, extract_root)
    staged_app_dir = _locate_staged_app_dir(extract_root)

    cur_req = info.app_root / "requirements.txt"
    new_req = staged_app_dir / "requirements.txt"
    requirements_changed = _sha256_file(cur_req) != _sha256_file(new_req)

    return StagedUpdateInfo(
        target_commit=target_commit,
        zip_path=zip_path,
        extract_root=extract_root,
        staged_app_dir=staged_app_dir,
        requirements_changed=requirements_changed,
        app_root=info.app_root,
        prev_managed_files=list(info.managed_files),
    )


def launch_staged_update_and_relaunch(
    staged: StagedUpdateInfo,
    parent_pid: Optional[int] = None,
    relaunch: bool = True,
) -> Path:
    """
    Copy the standalone apply worker outside `app_root`, write the manifest,
    and spawn the detached applier process.
    """
    work_root = _facade("get_temp_root", get_temp_root)() / "app_update"
    os.makedirs(to_extended_path(work_root), exist_ok=True)

    worker_src = Path(__file__).resolve().parent.parent / "apply_worker.py"
    worker_dst = work_root / "apply_worker_runner.py"
    shutil.copy2(to_extended_path(worker_src), to_extended_path(worker_dst))

    manifest_path = work_root / "update_manifest.json"
    manifest = {
        "parent_pid": int(os.getpid() if parent_pid is None else parent_pid),
        "app_root": str(staged.app_root),
        "staged_app_dir": str(staged.staged_app_dir),
        "app_data_root": str(get_app_data_root()),
        "python_executable": sys.executable,
        "requirements_changed": bool(staged.requirements_changed),
        "prev_managed_files": list(staged.prev_managed_files),
        "cleanup_dir": str(staged.extract_root),
        "cleanup_zip": str(staged.zip_path),
        "relaunch": bool(relaunch),
        "commit_meta": {
            "sha": staged.target_commit.sha,
            "short_sha": staged.target_commit.short_sha,
            "title": staged.target_commit.title,
            "updated_iso": staged.target_commit.updated_iso,
        },
    }
    manifest_path.write_text(json.dumps(manifest, indent=2), encoding="utf-8")

    sp = _facade("subprocess", subprocess)
    kwargs: Dict[str, object] = {"cwd": str(work_root)}
    if os.name == "nt":
        kwargs["creationflags"] = sp.CREATE_NEW_CONSOLE

    sp.Popen(
        [sys.executable, str(worker_dst), "--manifest", str(manifest_path)],
        **kwargs,
    )
    return manifest_path


def apply_git_update(app_root: Optional[Path] = None) -> Tuple[bool, str, bool]:
    """
    Fast-forward pull for clean developer Git checkouts while preserving local
    runtime seed files (`3gpp_data.db`, `export_config.json`, `useful_links.json`).

    Returns `(success, message, requirements_changed)`.
    """
    info = inspect_local_install(app_root=app_root)
    if info.install_mode != "git" or info.git_root is None:
        return False, "Current installation is not a Git repository.", False
    if info.is_git_dirty:
        return (
            False,
            "Uncommitted local Git changes detected. Commit or stash your changes before pulling.",
            False,
        )

    req_path = info.app_root / "requirements.txt"
    before_req_hash = _sha256_file(req_path)

    sp = _facade("subprocess", subprocess)
    kwargs: Dict[str, object] = {
        "cwd": str(info.git_root),
        "stdout": sp.PIPE,
        "stderr": sp.PIPE,
        "stdin": sp.DEVNULL,
        "timeout": 60.0,
    }
    if os.name == "nt":
        kwargs["creationflags"] = 0x08000000

    # Back up tracked seed files (`3gpp_data.db`, `export_config.json`, `useful_links.json`)
    # and reset any working-tree edits on them so `git pull --ff-only` never aborts or
    # overwrites the user's synced database and preferences.
    backup_root = _facade("get_temp_root", get_temp_root)() / "app_update" / "git_seed_backup"
    if os.path.exists(to_extended_path(backup_root)):
        rmtree_extended(backup_root)
    os.makedirs(to_extended_path(backup_root), exist_ok=True)

    backed_up_seeds: List[Tuple[Path, Path]] = []
    try:
        rel_app_to_git = "."
        try:
            rel_app = Path(os.path.abspath(str(info.app_root))).relative_to(
                Path(os.path.abspath(str(info.git_root)))
            )
            rel_app_to_git = normalize_rel_path(str(rel_app))
        except ValueError:
            rel_app_to_git = "."

        git_seed_paths: List[str] = []
        for seed_rel in sorted(SEED_IF_MISSING_PATHS):
            src_seed = info.app_root / seed_rel
            if os.path.isfile(to_extended_path(src_seed)):
                bak_seed = backup_root / seed_rel
                _copy_with_retry(src_seed, bak_seed)
                backed_up_seeds.append((bak_seed, src_seed))
                git_rel = (
                    f"{rel_app_to_git}/{seed_rel}"
                    if rel_app_to_git and rel_app_to_git != "."
                    else seed_rel
                )
                git_seed_paths.append(git_rel)

        if git_seed_paths:
            try:
                sp.run(
                    ["git", "checkout", "--", *git_seed_paths],
                    **kwargs,
                )
            except Exception:
                pass

        try:
            res = sp.run(
                ["git", "pull", "--ff-only", "origin", GITHUB_BRANCH],
                **kwargs,
            )
        except (OSError, sp.SubprocessError) as exc:
            return False, f"Could not run git pull: {exc}", False

        out = res.stdout.decode("utf-8", errors="replace").strip()
        err = res.stderr.decode("utf-8", errors="replace").strip()
        if res.returncode != 0:
            return (
                False,
                (err or out or f"git pull failed with exit code {res.returncode}"),
                False,
            )
    finally:
        for bak_seed, dst_seed in backed_up_seeds:
            try:
                if os.path.isfile(to_extended_path(bak_seed)):
                    _copy_with_retry(bak_seed, dst_seed)
            except Exception as exc:
                logger.warning("[Updater] Could not restore seed file %s: %s", dst_seed, exc)
        rmtree_extended(backup_root)

    after_req_hash = _sha256_file(req_path)
    req_changed = before_req_hash != after_req_hash
    if req_changed and os.path.isfile(to_extended_path(req_path)):
        try:
            sp.run(
                [sys.executable, "-m", "pip", "install", "-r", str(req_path)],
                cwd=str(info.app_root),
                timeout=180.0,
            )
        except Exception as exc:
            logger.warning("[Updater] pip install after git pull failed: %s", exc)

    return True, (out or "Updated via git pull --ff-only."), req_changed


def spawn_detached_relaunch(
    app_root: Optional[Path] = None,
    parent_pid: Optional[int] = None,
) -> Path:
    """
    Spawn a detached pure-stdlib watcher that waits for ``parent_pid`` to exit
    and then relaunches the GUI directly via ``sys.executable -m main_tools``
    without creating extra ``cmd.exe`` console windows.
    """
    src_root = _facade("get_project_root", get_project_root)()
    resolved_app_root = Path(app_root) if app_root is not None else src_root.parent
    pid = int(os.getpid() if parent_pid is None else parent_pid)

    work_root = _facade("get_temp_root", get_temp_root)() / "app_update"
    os.makedirs(to_extended_path(work_root), exist_ok=True)

    helper_code = (
        "import ctypes, os, subprocess, sys, time\n"
        f"pid = {pid}\n"
        f"app_root = r'''{str(resolved_app_root)}'''\n"
        f"src_root = r'''{str(src_root)}'''\n"
        f"py_exe = r'''{sys.executable}'''\n"
        "if os.name == 'nt' and pid > 0:\n"
        "    k32 = ctypes.windll.kernel32\n"
        "    k32.OpenProcess.argtypes = [ctypes.c_uint32, ctypes.c_int, ctypes.c_uint32]\n"
        "    k32.OpenProcess.restype = ctypes.c_void_p\n"
        "    k32.WaitForSingleObject.argtypes = [ctypes.c_void_p, ctypes.c_uint32]\n"
        "    k32.CloseHandle.argtypes = [ctypes.c_void_p]\n"
        "    h = k32.OpenProcess(0x00100000, 0, pid)\n"
        "    if h:\n"
        "        k32.WaitForSingleObject(h, 15000)\n"
        "        k32.CloseHandle(h)\n"
        "    time.sleep(0.25)\n"
        "else:\n"
        "    deadline = time.monotonic() + 15.0\n"
        "    while time.monotonic() < deadline:\n"
        "        try:\n"
        "            os.kill(pid, 0)\n"
        "        except OSError:\n"
        "            break\n"
        "        time.sleep(0.15)\n"
        "env = os.environ.copy()\n"
        "existing_pp = env.get('PYTHONPATH', '')\n"
        "pp_parts = [app_root, src_root] + ([existing_pp] if existing_pp else [])\n"
        "env['PYTHONPATH'] = os.pathsep.join(pp_parts)\n"
        "kwargs = {'cwd': app_root, 'env': env}\n"
        "if os.name == 'nt':\n"
        "    kwargs['creationflags'] = 0x08000000\n"
        "subprocess.Popen([py_exe, '-m', 'main_tools'], **kwargs)\n"
    )
    helper_path = work_root / "relaunch_app.py"
    with open(to_extended_path(helper_path), "w", encoding="utf-8") as fh:
        fh.write(helper_code)

    sp = _facade("subprocess", subprocess)
    kwargs: Dict[str, object] = {"cwd": str(work_root)}
    if os.name == "nt":
        kwargs["creationflags"] = 0x08000000
    sp.Popen([sys.executable, str(helper_path)], **kwargs)
    return helper_path
