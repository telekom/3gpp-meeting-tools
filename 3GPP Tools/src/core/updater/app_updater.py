# --- File: src/core/updater/app_updater.py ---
"""
Continuous GitHub `master` self-updater for 3GPP Delegate Tools.

Supports both:
- End-user ZIP installations (no `.git` folder required) via GitHub's Atom
  commit feed (rate-limit free) and commit-pinned ZIP archive downloads.
- Developer Git checkouts via non-destructive `git pull --ff-only`.
"""

from __future__ import annotations

import datetime
import hashlib
import html
import json
import logging
import os
import re
import shutil
import subprocess
import sys
import time
import xml.etree.ElementTree as ET
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Dict, List, Optional, Tuple

import requests

from core.config.config import (
    GITHUB_ARCHIVE_URL_TEMPLATE,
    GITHUB_BRANCH,
    GITHUB_COMMITS_ATOM_URL,
    GITHUB_SUBDIR,
)
from core.network.session import (
    DownloadCancelledError,
    HttpError,
    NetworkError,
    NetworkSession,
    NetworkTimeoutError,
)
from core.updater.apply_worker import (
    APP_VERSION_FILENAME,
    UPDATE_STATE_FILENAME,
    is_ignored_cache_path,
    is_protected_relative_path,
    normalize_rel_path,
)
from core.utils.paths import get_app_data_root, get_project_root, get_temp_root

logger = logging.getLogger(__name__)

_ATOM_NS = {"a": "http://www.w3.org/2005/Atom"}
_SHA_RE = re.compile(r"([0-9a-fA-F]{40})")
_TAG_RE = re.compile(r"<[^>]+>")

DEFAULT_CHECK_INTERVAL_HOURS = 12


@dataclass(frozen=True)
class CommitInfo:
    """Single commit entry parsed from GitHub's commits Atom feed."""

    sha: str
    short_sha: str
    title: str
    message: str
    author: str
    updated_iso: str
    updated_ts: float
    url: str


@dataclass
class LocalInstallInfo:
    """Detected local installation mode and baseline version information."""

    install_mode: str  # "git" | "zip_tracked" | "zip_manual"
    app_root: Path
    installed_sha: str = ""
    short_sha: str = ""
    installed_at_iso: str = ""
    estimated_ts: float = 0.0
    is_git_dirty: bool = False
    git_root: Optional[Path] = None
    managed_files: List[str] = field(default_factory=list)


@dataclass
class UpdateCheckResult:
    """Outcome of comparing local installation against remote `master`."""

    status: str  # "up_to_date" | "update_available" | "error"
    local_info: LocalInstallInfo
    latest_commit: Optional[CommitInfo] = None
    pending_commits: List[CommitInfo] = field(default_factory=list)
    all_commits: List[CommitInfo] = field(default_factory=list)
    history_truncated: bool = False
    error: str = ""


@dataclass
class StagedUpdateInfo:
    """Verified staged ZIP update ready to be applied after process exit."""

    target_commit: CommitInfo
    zip_path: Path
    extract_root: Path
    staged_app_dir: Path
    requirements_changed: bool
    app_root: Path
    prev_managed_files: List[str] = field(default_factory=list)


def get_app_root() -> Path:
    """Return the `3GPP Tools` installation root directory."""
    return get_project_root().parent.resolve()


def _parse_iso_timestamp(iso_str: str) -> float:
    raw = (iso_str or "").strip()
    if not raw:
        return 0.0
    if raw.endswith("Z"):
        raw = raw[:-1] + "+00:00"
    try:
        dt = datetime.datetime.fromisoformat(raw)
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=datetime.timezone.utc)
        return dt.timestamp()
    except ValueError:
        return 0.0


def _strip_html_content(raw_html: str) -> str:
    if not raw_html:
        return ""
    text = re.sub(r"<br\s*/?>", "\n", raw_html, flags=re.IGNORECASE)
    text = _TAG_RE.sub("", text)
    text = html.unescape(text)
    lines = [line.rstrip() for line in text.strip().splitlines()]
    return "\n".join(lines).strip()


def parse_commits_atom(xml_bytes: bytes) -> List[CommitInfo]:
    """Parse GitHub's `commits/*.atom` XML feed into chronological CommitInfo items."""
    if not xml_bytes:
        return []
    try:
        root = ET.fromstring(xml_bytes)
    except ET.ParseError as exc:
        raise ValueError(f"Invalid GitHub commits Atom XML: {exc}") from exc

    commits: List[CommitInfo] = []
    for entry in root.findall("a:entry", _ATOM_NS):
        entry_id = (entry.findtext("a:id", "", _ATOM_NS) or "").strip()
        sha_match = _SHA_RE.search(entry_id)
        if not sha_match:
            continue
        sha = sha_match.group(1).lower()

        raw_title = (entry.findtext("a:title", "", _ATOM_NS) or "").strip()
        raw_content = entry.findtext("a:content", "", _ATOM_NS) or ""
        full_message = _strip_html_content(raw_content) or raw_title
        first_line = full_message.splitlines()[0].strip() if full_message else raw_title
        title = first_line or raw_title or sha[:7]

        author_name = (entry.findtext("a:author/a:name", "", _ATOM_NS) or "").strip()
        author_email = (entry.findtext("a:author/a:email", "", _ATOM_NS) or "").strip()
        author = author_name or author_email or "Maintainer"

        updated_iso = (entry.findtext("a:updated", "", _ATOM_NS) or "").strip()
        updated_ts = _parse_iso_timestamp(updated_iso)

        url = ""
        for link in entry.findall("a:link", _ATOM_NS):
            if (link.attrib.get("rel") or "alternate") == "alternate":
                url = (link.attrib.get("href") or "").strip()
                break
        if not url:
            url = f"https://github.com/telekom/3gpp-meeting-tools/commit/{sha}"

        commits.append(
            CommitInfo(
                sha=sha,
                short_sha=sha[:7],
                title=title,
                message=full_message,
                author=author,
                updated_iso=updated_iso,
                updated_ts=updated_ts,
                url=url,
            )
        )
    return commits


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


def _inspect_git_state(git_root: Path, app_root: Path) -> Tuple[str, bool]:
    """Return `(head_sha, is_dirty)` for a Git workspace."""
    git_dir = git_root / ".git"
    sha = _read_git_head_sha_from_files(git_dir) if git_dir.is_dir() else ""
    is_dirty = False
    try:
        kwargs = {
            "cwd": str(git_root),
            "stdout": subprocess.PIPE,
            "stderr": subprocess.PIPE,
            "stdin": subprocess.DEVNULL,
            "timeout": 4.0,
        }
        if os.name == "nt":
            kwargs["creationflags"] = 0x08000000
        if not sha:
            res_sha = subprocess.run(["git", "rev-parse", "HEAD"], **kwargs)
            if res_sha.returncode == 0:
                candidate = res_sha.stdout.decode("utf-8", errors="replace").strip()
                if _SHA_RE.fullmatch(candidate):
                    sha = candidate.lower()

        rel_scope = "."
        try:
            rel = app_root.resolve().relative_to(git_root.resolve())
            rel_scope = str(rel) if str(rel) != "." else "."
        except ValueError:
            rel_scope = "."

        res_status = subprocess.run(
            ["git", "status", "--porcelain", "--untracked-files=no", "--", rel_scope],
            **kwargs,
        )
        if res_status.returncode == 0:
            is_dirty = bool(res_status.stdout.decode("utf-8", errors="replace").strip())
    except Exception:
        pass
    return sha, is_dirty


def _git_is_ancestor(git_root: Optional[Path], commit_sha: str) -> bool:
    """Return True if ``commit_sha`` is already in local Git HEAD history (ahead of remote)."""
    if git_root is None or not commit_sha:
        return False
    try:
        kwargs: Dict[str, object] = {
            "cwd": str(git_root),
            "stdout": subprocess.DEVNULL,
            "stderr": subprocess.DEVNULL,
            "stdin": subprocess.DEVNULL,
            "timeout": 4.0,
        }
        if os.name == "nt":
            kwargs["creationflags"] = 0x08000000
        res = subprocess.run(
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
    root = Path(app_root).resolve() if app_root is not None else get_app_root()

    git_root = _find_git_root(root)
    if git_root is not None:
        sha, dirty = _inspect_git_state(git_root, root)
        return LocalInstallInfo(
            install_mode="git",
            app_root=root,
            installed_sha=sha,
            short_sha=sha[:7] if sha else "",
            is_git_dirty=dirty,
            git_root=git_root,
        )

    version_file = root / APP_VERSION_FILENAME
    if version_file.is_file():
        try:
            data = json.loads(version_file.read_text(encoding="utf-8")) or {}
            sha = str(data.get("installed_sha") or "").strip().lower()
            if sha:
                return LocalInstallInfo(
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

    estimated_ts = _estimate_source_mtime(root)
    return LocalInstallInfo(
        install_mode="zip_manual",
        app_root=root,
        estimated_ts=estimated_ts,
    )


def evaluate_update_status(
    local_info: LocalInstallInfo,
    commits: List[CommitInfo],
) -> UpdateCheckResult:
    """Compare local installation baseline against parsed remote commits."""
    if not commits:
        return UpdateCheckResult(
            status="error",
            local_info=local_info,
            error="No commits were found in the remote GitHub feed.",
        )

    latest = commits[0]

    if local_info.installed_sha:
        if latest.sha == local_info.installed_sha or (
            local_info.install_mode == "git"
            and _git_is_ancestor(local_info.git_root, latest.sha)
        ):
            return UpdateCheckResult(
                status="up_to_date",
                local_info=local_info,
                latest_commit=latest,
                pending_commits=[],
                all_commits=commits,
            )

        pending: List[CommitInfo] = []
        found_local = False
        for commit in commits:
            if commit.sha == local_info.installed_sha:
                found_local = True
                break
            pending.append(commit)

        return UpdateCheckResult(
            status="update_available",
            local_info=local_info,
            latest_commit=latest,
            pending_commits=pending,
            all_commits=commits,
            history_truncated=(not found_local and len(pending) == len(commits)),
        )

    # Manual ZIP install without `.app_version.json`: compare commit timestamps
    # with a 120-second tolerance so a freshly downloaded ZIP isn't flagged.
    cutoff = (local_info.estimated_ts or 0.0) + 120.0
    pending = [c for c in commits if c.updated_ts > cutoff]
    if pending:
        return UpdateCheckResult(
            status="update_available",
            local_info=local_info,
            latest_commit=latest,
            pending_commits=pending,
            all_commits=commits,
            history_truncated=(len(pending) == len(commits)),
        )

    return UpdateCheckResult(
        status="up_to_date",
        local_info=local_info,
        latest_commit=latest,
        pending_commits=[],
        all_commits=commits,
    )


def get_update_state_path() -> Path:
    """Return the path to persisted updater preferences in user app data."""
    return get_app_data_root() / UPDATE_STATE_FILENAME


def load_update_state() -> Dict[str, object]:
    """Load persisted updater settings and throttling timestamps."""
    path = get_update_state_path()
    defaults: Dict[str, object] = {
        "auto_check_enabled": False,
        "check_interval_hours": DEFAULT_CHECK_INTERVAL_HOURS,
        "last_check_ts": 0.0,
        "skipped_sha": "",
    }
    if not path.is_file():
        return defaults
    try:
        loaded = json.loads(path.read_text(encoding="utf-8")) or {}
        if isinstance(loaded, dict):
            defaults.update(loaded)
    except Exception as exc:
        logger.warning("[Updater] Failed to read update state: %s", exc)
    return defaults


def save_update_state(updates: Dict[str, object]) -> Dict[str, object]:
    """Merge ``updates`` into persisted updater state and save atomically."""
    state = load_update_state()
    state.update(updates)
    path = get_update_state_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp_path = path.with_suffix(".json.tmp")
    tmp_path.write_text(json.dumps(state, indent=2), encoding="utf-8")
    os.replace(str(tmp_path), str(path))
    return state


def should_auto_check(state: Dict[str, object], now_ts: Optional[float] = None) -> bool:
    """Return True when background startup update check is enabled and due."""
    if not bool(state.get("auto_check_enabled", True)):
        return False
    now = time.time() if now_ts is None else float(now_ts)
    last_ts = float(state.get("last_check_ts") or 0.0)
    interval_h = max(1.0, float(state.get("check_interval_hours") or DEFAULT_CHECK_INTERVAL_HOURS))
    return (now - last_ts) >= (interval_h * 3600.0)


def fetch_remote_commits(
    timeout: int = 15,
    feed_url: str = GITHUB_COMMITS_ATOM_URL,
) -> List[CommitInfo]:
    """Fetch and parse recent commits from GitHub using NetworkSession proxies."""
    session = NetworkSession.get_instance()
    headers = NetworkSession.get_request_headers()
    headers["Accept"] = "application/atom+xml, application/xml, text/xml;q=0.9, */*;q=0.8"

    try:
        with session.get(feed_url, headers=headers, timeout=timeout) as resp:
            resp.raise_for_status()
            return parse_commits_atom(resp.content)
    except requests.exceptions.Timeout as exc:
        raise NetworkTimeoutError(f"Timed out checking GitHub updates after {timeout}s.") from exc
    except requests.exceptions.HTTPError as exc:
        code = exc.response.status_code if exc.response is not None else None
        raise HttpError(f"GitHub returned HTTP {code} while checking updates.", status_code=code, url=feed_url) from exc
    except requests.exceptions.RequestException as exc:
        raise NetworkError(f"Network error checking GitHub updates: {exc}") from exc


def check_for_updates(
    app_root: Optional[Path] = None,
    timeout: int = 15,
    record_check_time: bool = True,
) -> UpdateCheckResult:
    """Check GitHub `master` for new commits and compare against the local install."""
    local_info = inspect_local_install(app_root=app_root)
    try:
        commits = fetch_remote_commits(timeout=timeout)
    except Exception as exc:
        return UpdateCheckResult(
            status="error",
            local_info=local_info,
            error=str(exc),
        )

    if record_check_time:
        try:
            save_update_state({"last_check_ts": time.time()})
        except Exception:
            pass

    return evaluate_update_status(local_info, commits)


def _safe_extract_zip(zip_path: Path, dest_dir: Path) -> None:
    """Extract ``zip_path`` into ``dest_dir`` with ZipSlip path traversal protection."""
    dest_root = dest_dir.resolve()
    with zipfile.ZipFile(str(zip_path), "r") as zf:
        for member in zf.infolist():
            member_name = member.filename.replace("\\", "/")
            if not member_name or member_name.startswith("/"):
                raise ValueError(f"Unsafe archive entry path: {member.filename}")
            target = (dest_root / member_name).resolve()
            try:
                target.relative_to(dest_root)
            except ValueError as exc:
                raise ValueError(f"Archive entry escapes target directory: {member.filename}") from exc
        zf.extractall(str(dest_root))


def _locate_staged_app_dir(extract_root: Path) -> Path:
    """Locate `3GPP Tools/` (containing `src/main_tools.py`) inside an extracted archive."""
    for child in extract_root.iterdir():
        if not child.is_dir():
            continue
        candidate = child / GITHUB_SUBDIR
        if (candidate / "src" / "main_tools.py").is_file():
            return candidate.resolve()
        if (child / "src" / "main_tools.py").is_file():
            return child.resolve()
    if (extract_root / "src" / "main_tools.py").is_file():
        return extract_root.resolve()
    raise FileNotFoundError(
        "Downloaded archive does not contain the expected '3GPP Tools/src/main_tools.py' structure."
    )


def _sha256_file(path: Path) -> str:
    if not path.is_file():
        return ""
    h = hashlib.sha256()
    with open(path, "rb") as f:
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
    work_root = get_temp_root() / "app_update"
    work_root.mkdir(parents=True, exist_ok=True)

    zip_path = work_root / f"{target_commit.short_sha}.zip"
    extract_root = work_root / f"extract_{target_commit.short_sha}"
    if extract_root.exists():
        shutil.rmtree(str(extract_root), ignore_errors=True)
    extract_root.mkdir(parents=True, exist_ok=True)

    archive_url = GITHUB_ARCHIVE_URL_TEMPLATE.format(ref=target_commit.sha)
    logger.info("[Updater] Downloading update archive: %s", archive_url)
    NetworkSession.download_file(
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
    work_root = get_temp_root() / "app_update"
    work_root.mkdir(parents=True, exist_ok=True)

    worker_src = Path(__file__).resolve().parent / "apply_worker.py"
    worker_dst = work_root / "apply_worker_runner.py"
    shutil.copy2(str(worker_src), str(worker_dst))

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
        "relaunch": bool(relaunch),
        "commit_meta": {
            "sha": staged.target_commit.sha,
            "short_sha": staged.target_commit.short_sha,
            "title": staged.target_commit.title,
            "updated_iso": staged.target_commit.updated_iso,
        },
    }
    manifest_path.write_text(json.dumps(manifest, indent=2), encoding="utf-8")

    kwargs: Dict[str, object] = {"cwd": str(work_root)}
    if os.name == "nt":
        kwargs["creationflags"] = subprocess.CREATE_NEW_CONSOLE

    subprocess.Popen(
        [sys.executable, str(worker_dst), "--manifest", str(manifest_path)],
        **kwargs,
    )
    return manifest_path


def apply_git_update(app_root: Optional[Path] = None) -> Tuple[bool, str, bool]:
    """
    Fast-forward pull for clean developer Git checkouts.

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

    kwargs: Dict[str, object] = {
        "cwd": str(info.git_root),
        "stdout": subprocess.PIPE,
        "stderr": subprocess.PIPE,
        "stdin": subprocess.DEVNULL,
        "timeout": 60.0,
    }
    if os.name == "nt":
        kwargs["creationflags"] = 0x08000000

    res = subprocess.run(["git", "pull", "--ff-only", "origin", GITHUB_BRANCH], **kwargs)
    out = res.stdout.decode("utf-8", errors="replace").strip()
    err = res.stderr.decode("utf-8", errors="replace").strip()
    if res.returncode != 0:
        return False, (err or out or f"git pull failed with exit code {res.returncode}"), False

    after_req_hash = _sha256_file(req_path)
    req_changed = before_req_hash != after_req_hash
    if req_changed and req_path.is_file():
        subprocess.run(
            [sys.executable, "-m", "pip", "install", "-r", str(req_path)],
            cwd=str(info.app_root),
            timeout=180.0,
        )

    return True, (out or "Updated via git pull --ff-only."), req_changed
