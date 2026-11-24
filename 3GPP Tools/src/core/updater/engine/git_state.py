# --- File: src/core/updater/engine/git_state.py ---
"""Local install mode detection (`git`, `zip_tracked`, `zip_manual`) and Git inspection."""

from __future__ import annotations

import json
import logging
import os
import subprocess
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from core.updater.apply_worker import (
    APP_VERSION_FILENAME,
    is_ignored_cache_path,
    is_protected_relative_path,
    is_seed_only_relative_path,
    normalize_rel_path,
)
from core.utils.paths import to_extended_path

from .atom_feed import _parse_iso_timestamp
from .models import _SHA_RE, LocalInstallInfo, _facade, get_app_root

logger = logging.getLogger(__name__)


def _read_git_head_sha_from_files(git_dir: Path) -> str:
    """Fast pure-Python `.git/HEAD` SHA lookup (no `git.exe` required)."""
    head_file = git_dir / "HEAD"
    if not head_file.is_file():
        return ""
    head_text = head_file.read_text(encoding="utf-8", errors="replace").strip()
    if _SHA_RE.fullmatch(head_text):
        return head_text.lower()
    if head_text.startswith("ref:"):
        ref_rel = head_text.split(":", 1)[1].strip()
        ref_file = git_dir / ref_rel
        if ref_file.is_file():
            val = ref_file.read_text(encoding="utf-8", errors="replace").strip()
            if _SHA_RE.fullmatch(val):
                return val.lower()
        packed = git_dir / "packed-refs"
        if packed.is_file():
            for line in packed.read_text(encoding="utf-8", errors="replace").splitlines():
                line = line.strip()
                if not line or line.startswith("#") or line.startswith("^"):
                    continue
                parts = line.split()
                if len(parts) >= 2 and parts[1] == ref_rel and _SHA_RE.fullmatch(parts[0]):
                    return parts[0].lower()
    return ""


def _find_git_root(app_root: Path) -> Optional[Path]:
    for candidate in (app_root, app_root.parent):
        if (candidate / ".git").exists():
            return candidate
    return None


def _parse_porcelain_rel_paths(raw_output: str, rel_scope: str) -> List[str]:
    """
    Parse `git status --porcelain` output into paths relative to `app_root`.

    Supports both NUL-delimited (`-z`) and newline-delimited porcelain output.
    """
    if not raw_output:
        return []

    scope_norm = normalize_rel_path(rel_scope).rstrip("/")
    scope_prefix = (scope_norm + "/") if scope_norm and scope_norm != "." else ""

    rel_paths: List[str] = []
    if "\x00" in raw_output:
        tokens = [t for t in raw_output.split("\x00") if t]
        i = 0
        while i < len(tokens):
            entry = tokens[i]
            if len(entry) < 4:
                i += 1
                continue
            status_xy = entry[:2]
            raw_path = entry[3:]
            # In -z mode, rename/copy entries have an extra NUL-terminated source path token
            if ("R" in status_xy or "C" in status_xy) and (i + 1) < len(tokens):
                i += 2
            else:
                i += 1
            norm_p = normalize_rel_path(raw_path.strip('"'))
            if scope_prefix and norm_p.lower().startswith(scope_prefix.lower()):
                norm_p = norm_p[len(scope_prefix) :]
            if norm_p:
                rel_paths.append(norm_p)
    else:
        for line in raw_output.splitlines():
            if len(line) < 4:
                continue
            raw_path = line[3:].strip()
            if " -> " in raw_path and line[0] in ("R", "C"):
                raw_path = raw_path.split(" -> ", 1)[1].strip()
            norm_p = normalize_rel_path(raw_path.strip('"'))
            if scope_prefix and norm_p.lower().startswith(scope_prefix.lower()):
                norm_p = norm_p[len(scope_prefix) :]
            if norm_p:
                rel_paths.append(norm_p)
    return rel_paths


def _inspect_git_state(git_root: Path, app_root: Path) -> Tuple[str, bool]:
    """
    Return `(head_sha, is_dirty)` for a Git workspace.

    Runtime-owned and seed-only files (such as `src/db/3gpp_data.db`,
    `src/config/export_config.json`, `src/config/useful_links.json`, and
    `src/config/*_config.json`) are ignored when determining whether the source
    tree is dirty so normal application usage does not block updates.
    """
    sp = _facade("subprocess", subprocess)
    git_dir = git_root / ".git"
    sha = _read_git_head_sha_from_files(git_dir) if git_dir.is_dir() else ""
    is_dirty = False
    try:
        kwargs = {
            "cwd": str(git_root),
            "stdout": sp.PIPE,
            "stderr": sp.PIPE,
            "stdin": sp.DEVNULL,
            "timeout": 4.0,
        }
        if os.name == "nt":
            kwargs["creationflags"] = 0x08000000
        if not sha:
            res_sha = sp.run(["git", "rev-parse", "HEAD"], **kwargs)
            if res_sha.returncode == 0:
                candidate = res_sha.stdout.decode("utf-8", errors="replace").strip()
                if _SHA_RE.fullmatch(candidate):
                    sha = candidate.lower()

        rel_scope = "."
        try:
            rel = Path(os.path.abspath(str(app_root))).relative_to(
                Path(os.path.abspath(str(git_root)))
            )
            rel_scope = str(rel) if str(rel) != "." else "."
        except ValueError:
            rel_scope = "."

        res_status = sp.run(
            ["git", "status", "--porcelain", "--untracked-files=no", "-z", "--", rel_scope],
            **kwargs,
        )
        if res_status.returncode == 0:
            raw_status = res_status.stdout.decode("utf-8", errors="replace")
            changed_paths = _parse_porcelain_rel_paths(raw_status, rel_scope)
            meaningful_changes = [
                p
                for p in changed_paths
                if not (is_protected_relative_path(p) or is_seed_only_relative_path(p))
            ]
            is_dirty = bool(meaningful_changes)
    except Exception:
        pass
    return sha, is_dirty


def _git_is_ancestor(git_root: Optional[Path], commit_sha: str) -> bool:
    """Return True if ``commit_sha`` is already in local Git HEAD history (ahead of remote)."""
    if git_root is None or not commit_sha:
        return False
    sp = _facade("subprocess", subprocess)
    try:
        kwargs: Dict[str, object] = {
            "cwd": str(git_root),
            "stdout": sp.DEVNULL,
            "stderr": sp.DEVNULL,
            "stdin": sp.DEVNULL,
            "timeout": 4.0,
        }
        if os.name == "nt":
            kwargs["creationflags"] = 0x08000000
        res = sp.run(
            ["git", "merge-base", "--is-ancestor", commit_sha, "HEAD"],
            **kwargs,
        )
        return res.returncode == 0
    except Exception:
        return False


def _estimate_source_mtime(app_root: Path) -> float:
    """Estimate download/commit timestamp from newest tracked-like source file mtime."""
    src_dir = app_root / "src"
    scan_root = src_dir if src_dir.is_dir() else app_root
    newest = 0.0
    for root, dirs, files in os.walk(scan_root):
        dirs[:] = [d for d in dirs if d not in {"__pycache__", ".git", ".idea", "db", "output"}]
        root_path = Path(root)
        for fname in files:
            fpath = root_path / fname
            try:
                rel = normalize_rel_path(str(fpath.relative_to(app_root)))
            except ValueError:
                continue
            if is_ignored_cache_path(rel) or is_protected_relative_path(rel):
                continue
            if not rel.endswith((".py", ".cmd", ".txt", ".md", ".json")):
                continue
            try:
                mtime = fpath.stat().st_mtime
                if mtime > newest:
                    newest = mtime
            except OSError:
                continue
    return newest


def inspect_local_install(app_root: Optional[Path] = None) -> LocalInstallInfo:
    """Inspect whether ``app_root`` is a Git checkout, tracked ZIP, or manual ZIP."""
    root = Path(os.path.abspath(str(app_root))) if app_root is not None else get_app_root()

    version_file = root / APP_VERSION_FILENAME
    tracked_info: Optional[LocalInstallInfo] = None
    zip_override = False
    if os.path.isfile(to_extended_path(version_file)):
        try:
            with open(to_extended_path(version_file), "r", encoding="utf-8") as vf:
                data = json.loads(vf.read()) or {}
            sha = str(data.get("installed_sha") or "").strip().lower()
            zip_override = bool(data.get("zip_override", False))
            if sha:
                tracked_info = LocalInstallInfo(
                    install_mode="zip_tracked",
                    app_root=root,
                    installed_sha=sha,
                    short_sha=str(data.get("short_sha") or sha[:7]),
                    installed_at_iso=str(data.get("installed_at_iso") or data.get("updated_iso") or ""),
                    estimated_ts=_parse_iso_timestamp(str(data.get("updated_iso") or "")),
                    managed_files=list(data.get("managed_files") or []),
                )
        except Exception as exc:
            logger.warning("[Updater] Could not read %s: %s", version_file, exc)

    if tracked_info is not None and zip_override:
        return tracked_info

    git_root = _find_git_root(root)
    if git_root is not None:
        sha, dirty = _inspect_git_state(git_root, root)
        if sha:
            return LocalInstallInfo(
                install_mode="git",
                app_root=root,
                installed_sha=sha,
                short_sha=sha[:7],
                is_git_dirty=dirty,
                git_root=git_root,
                managed_files=tracked_info.managed_files if tracked_info is not None else [],
            )

    if tracked_info is not None:
        return tracked_info

    estimated_ts = _estimate_source_mtime(root)
    return LocalInstallInfo(
        install_mode="zip_manual",
        app_root=root,
        estimated_ts=estimated_ts,
    )
