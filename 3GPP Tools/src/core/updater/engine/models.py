# --- File: src/core/updater/engine/models.py ---
"""Data models, constants, and facade patch resolution for the self-updater."""

from __future__ import annotations

import re
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import List, Optional

from core.utils.paths import get_project_root

_ATOM_NS = {"a": "http://www.w3.org/2005/Atom"}
_SHA_RE = re.compile(r"([0-9a-fA-F]{40})")
_TAG_RE = re.compile(r"<[^>]+>")

DEFAULT_CHECK_INTERVAL_HOURS = 12


def _facade(name: str, default):
    """Resolve monkeypatchable symbols dynamically through `core.updater.app_updater`."""
    mod = sys.modules.get("core.updater.app_updater")
    return getattr(mod, name, default) if mod is not None else default


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
    return _facade("get_project_root", get_project_root)().parent.resolve()
