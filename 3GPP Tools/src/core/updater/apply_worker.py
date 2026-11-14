# --- File: src/core/updater/apply_worker.py ---
"""
Standalone pure-stdlib update applier for 3GPP Delegate Tools.

This module intentionally uses ONLY Python standard library modules so it can
be copied into a temporary folder outside the installation directory and
executed after the main GUI process exits, avoiding Windows file-lock issues
while overwriting live application files.
"""

from __future__ import annotations

import argparse
import ctypes
import datetime
import fnmatch
import json
import os
import shutil
import subprocess
import sys
import time
from pathlib import Path
from typing import Dict, List, Optional, Set, Tuple

APP_VERSION_FILENAME = ".app_version.json"
UPDATE_STATE_FILENAME = "app_update_state.json"

# Directories / paths that belong to the user or runtime environment and must
# NEVER be overwritten or deleted during an update.
PROTECTED_DIR_PREFIXES: Tuple[str, ...] = (
    ".git/",
    ".idea/",
    "output/",
    "scratch/",
    "test_meeting_dir/",
    "src/export/",
)

# Relative files that should never be overwritten or deleted if already present,
# and never deleted by obsolete-file cleanup.
PROTECTED_EXACT_PATHS: Set[str] = {
    ".app_version.json",
    ".editor_cache.puml",
    "3gpp_tools.log",
    "src/3gpp_tools.log",
    "src/assets/plantuml.jar",
    "src/assets/plantuml.version",
    "src/assets/3gpp_70.dotx",
}

PROTECTED_GLOB_PATTERNS: Tuple[str, ...] = (
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

# Tracked seed files: copy from update archive ONLY when missing locally so
# colleagues' customized links, export preferences, and synced DBs survive.
SEED_IF_MISSING_PATHS: Set[str] = {
    "src/config/useful_links.json",
    "src/config/export_config.json",
    "src/db/3gpp_data.db",
}


def normalize_rel_path(rel_path: str) -> str:
    """Normalize a relative path to forward-slash lowercase comparison form."""
    return rel_path.replace("\\", "/").lstrip("./")


def is_ignored_cache_path(rel_posix: str) -> bool:
    """Return True for transient cache/bytecode paths excluded from updates."""
    norm = normalize_rel_path(rel_posix).lower()
    parts = [p for p in norm.split("/") if p]
    if "__pycache__" in parts or ".ipynb_checkpoints" in parts:
        return True
    if norm.endswith((".pyc", ".pyo", ".tmp", ".part")):
        return True
    return False


def is_protected_relative_path(rel_posix: str) -> bool:
    """Return True if ``rel_posix`` is runtime/user-owned and must never be overwritten."""
    norm = normalize_rel_path(rel_posix)
    low = norm.lower()
    if is_ignored_cache_path(low):
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


def is_seed_only_relative_path(rel_posix: str) -> bool:
    """Return True if ``rel_posix`` should only be copied when absent on disk."""
    low = normalize_rel_path(rel_posix).lower()
    return low in SEED_IF_MISSING_PATHS


def wait_for_pid_exit(pid: int, timeout_seconds: float = 20.0) -> bool:
    """Wait until process ``pid`` exits, returning True if gone before timeout."""
    if not pid or pid <= 0:
        return True
    deadline = time.monotonic() + max(1.0, float(timeout_seconds))

    if os.name == "nt":
        SYNCHRONIZE = 0x00100000
        WAIT_OBJECT_0 = 0x00000000
        handle = ctypes.windll.kernel32.OpenProcess(SYNCHRONIZE, False, int(pid))
        if not handle:
            return True
        try:
            remaining_ms = int(max(100, (deadline - time.monotonic()) * 1000))
            res = ctypes.windll.kernel32.WaitForSingleObject(handle, remaining_ms)
            return res == WAIT_OBJECT_0
        finally:
            ctypes.windll.kernel32.CloseHandle(handle)

    while time.monotonic() < deadline:
        try:
            os.kill(pid, 0)
        except OSError:
            return True
        time.sleep(0.2)
    return False


def _copy_with_retry(src: Path, dst: Path, attempts: int = 10, delay_s: float = 0.35) -> None:
    """Copy ``src`` to ``dst`` with retries for transient OneDrive/AV file locks."""
    dst.parent.mkdir(parents=True, exist_ok=True)
    last_err: Optional[Exception] = None
    for i in range(max(1, attempts)):
        try:
            shutil.copy2(str(src), str(dst))
            return
        except (PermissionError, OSError) as exc:
            last_err = exc
            if i < attempts - 1:
                time.sleep(delay_s)
    if last_err is not None:
        raise last_err


def _remove_with_retry(target: Path, attempts: int = 6, delay_s: float = 0.25) -> bool:
    """Remove ``target`` file safely with retries; returns True if removed."""
    for i in range(max(1, attempts)):
        try:
            if target.is_file() or target.is_symlink():
                target.unlink()
            return True
        except (PermissionError, OSError):
            if i < attempts - 1:
                time.sleep(delay_s)
    return False


def apply_staged_tree(
    staged_app_dir: Path,
    app_root: Path,
    commit_meta: Dict[str, object],
    prev_managed_files: Optional[List[str]] = None,
    app_data_root: Optional[Path] = None,
    log_fn=print,
) -> Dict[str, object]:
    """
    Overlay ``staged_app_dir`` onto ``app_root`` while preserving user configs/DBs.

    Returns a summary dict with counts of copied, seeded, skipped, and removed files.
    """
    staged_app_dir = Path(staged_app_dir).resolve()
    app_root = Path(app_root).resolve()

    if not (staged_app_dir / "src" / "main_tools.py").is_file():
        raise FileNotFoundError(
            f"Staged directory does not contain src/main_tools.py: {staged_app_dir}"
        )

    copied = 0
    seeded = 0
    skipped = 0
    removed = 0
    new_managed_files: List[str] = []

    for root, dirs, files in os.walk(staged_app_dir):
        root_path = Path(root)
        # Skip cache directories before descending
        dirs[:] = [
            d for d in dirs
            if d not in {"__pycache__", ".git", ".idea", ".ipynb_checkpoints"}
        ]
        for fname in files:
            src_file = root_path / fname
            rel_posix = normalize_rel_path(str(src_file.relative_to(staged_app_dir)))

            if is_ignored_cache_path(rel_posix):
                continue

            dst_file = app_root / rel_posix

            if is_seed_only_relative_path(rel_posix):
                if dst_file.exists():
                    skipped += 1
                    continue
                _copy_with_retry(src_file, dst_file)
                seeded += 1
                continue

            if is_protected_relative_path(rel_posix):
                skipped += 1
                continue

            _copy_with_retry(src_file, dst_file)
            copied += 1
            new_managed_files.append(rel_posix)

    new_managed_set = set(new_managed_files)
    if prev_managed_files:
        for old_rel in prev_managed_files:
            norm_old = normalize_rel_path(old_rel)
            if not norm_old or norm_old in new_managed_set:
                continue
            if is_protected_relative_path(norm_old) or is_seed_only_relative_path(norm_old):
                continue
            old_target = app_root / norm_old
            if old_target.is_file():
                if _remove_with_retry(old_target):
                    removed += 1
                    # Prune empty parent folders inside app_root
                    parent = old_target.parent
                    while parent != app_root and parent.is_dir():
                        try:
                            parent.rmdir()
                            parent = parent.parent
                        except OSError:
                            break

    now_iso = datetime.datetime.now(datetime.timezone.utc).replace(microsecond=0).isoformat()
    version_payload = {
        "installed_sha": str(commit_meta.get("sha") or ""),
        "short_sha": str(commit_meta.get("short_sha") or str(commit_meta.get("sha") or "")[:7]),
        "commit_title": str(commit_meta.get("title") or ""),
        "updated_iso": str(commit_meta.get("updated_iso") or ""),
        "installed_at_iso": now_iso,
        "managed_files": sorted(new_managed_files),
    }
    version_file = app_root / APP_VERSION_FILENAME
    version_file.write_text(json.dumps(version_payload, indent=2), encoding="utf-8")

    if app_data_root is not None:
        app_data_root = Path(app_data_root)
        app_data_root.mkdir(parents=True, exist_ok=True)
        state_path = app_data_root / UPDATE_STATE_FILENAME
        state: Dict[str, object] = {}
        if state_path.is_file():
            try:
                state = json.loads(state_path.read_text(encoding="utf-8")) or {}
            except Exception:
                state = {}
        state["last_updated_sha"] = version_payload["installed_sha"]
        state["last_updated_at_iso"] = now_iso
        state["skipped_sha"] = ""
        state_path.write_text(json.dumps(state, indent=2), encoding="utf-8")

    log_fn(
        f"[Updater] Applied update: {copied} updated, {seeded} seeded, "
        f"{skipped} user/protected preserved, {removed} obsolete removed."
    )
    return {
        "copied": copied,
        "seeded": seeded,
        "skipped": skipped,
        "removed": removed,
        "installed_sha": version_payload["installed_sha"],
    }


def _relaunch_application(app_root: Path, python_exe: str) -> None:
    """Relaunch 3GPP Delegate Tools after update completes."""
    launch_cmd = app_root / "src" / "Launch.cmd"
    if os.name == "nt" and launch_cmd.is_file():
        subprocess.Popen(
            ["cmd.exe", "/c", str(launch_cmd)],
            cwd=str(app_root),
            creationflags=subprocess.CREATE_NEW_CONSOLE,
        )
        return

    env = os.environ.copy()
    src_dir = str(app_root / "src")
    root_dir = str(app_root)
    existing_pp = env.get("PYTHONPATH", "")
    env["PYTHONPATH"] = f"{root_dir}{os.pathsep}{src_dir}" + (
        f"{os.pathsep}{existing_pp}" if existing_pp else ""
    )
    kwargs = {"cwd": root_dir, "env": env}
    if os.name == "nt":
        kwargs["creationflags"] = subprocess.CREATE_NEW_CONSOLE
    subprocess.Popen([python_exe, "-m", "main_tools"], **kwargs)


def run_manifest(manifest_path: Path) -> int:
    """Execute an update from a JSON manifest file."""
    manifest = json.loads(Path(manifest_path).read_text(encoding="utf-8"))
    parent_pid = int(manifest.get("parent_pid") or 0)
    app_root = Path(manifest["app_root"])
    staged_app_dir = Path(manifest["staged_app_dir"])
    app_data_root = Path(manifest["app_data_root"]) if manifest.get("app_data_root") else None
    python_exe = str(manifest.get("python_executable") or sys.executable)
    requirements_changed = bool(manifest.get("requirements_changed", False))
    commit_meta = dict(manifest.get("commit_meta") or {})
    prev_managed_files = list(manifest.get("prev_managed_files") or [])
    cleanup_dir = Path(manifest["cleanup_dir"]) if manifest.get("cleanup_dir") else None
    relaunch = bool(manifest.get("relaunch", True))

    print("=" * 56)
    print("3GPP Delegate Tools — Applying Update")
    print("=" * 56)
    print(f"Target version : {commit_meta.get('short_sha', '?')} — {commit_meta.get('title', '')}")
    print(f"Install folder : {app_root}")

    if parent_pid > 0:
        print(f"Waiting for running application (PID {parent_pid}) to close...")
        exited = wait_for_pid_exit(parent_pid, timeout_seconds=20.0)
        if not exited:
            print("[WARN] Application did not exit within 20s; continuing with retry-safe copy.")
        else:
            time.sleep(0.4)

    apply_staged_tree(
        staged_app_dir=staged_app_dir,
        app_root=app_root,
        commit_meta=commit_meta,
        prev_managed_files=prev_managed_files,
        app_data_root=app_data_root,
    )

    if requirements_changed:
        req_file = app_root / "requirements.txt"
        if req_file.is_file():
            print("\nInstalling updated Python dependencies from requirements.txt...")
            res = subprocess.run(
                [python_exe, "-m", "pip", "install", "-r", str(req_file)],
                cwd=str(app_root),
            )
            if res.returncode != 0:
                print(
                    f"[WARN] pip install exited with code {res.returncode}. "
                    "Startup dependency check will assist if anything is missing."
                )

    if cleanup_dir and cleanup_dir.exists():
        try:
            shutil.rmtree(str(cleanup_dir), ignore_errors=True)
        except Exception:
            pass

    if relaunch:
        print("\nRelaunching 3GPP Delegate Tools...")
        _relaunch_application(app_root, python_exe)

    print("Update complete.")
    return 0


def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(description="Apply staged 3GPP Delegate Tools update.")
    parser.add_argument("--manifest", required=True, help="Path to update manifest JSON")
    args = parser.parse_args(argv)
    try:
        return run_manifest(Path(args.manifest))
    except Exception as exc:
        print(f"\n[ERROR] Failed to apply update: {exc}")
        if os.name == "nt":
            input("\nPress Enter to close this window...")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
