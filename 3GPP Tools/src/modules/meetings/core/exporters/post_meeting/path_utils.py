# --- File: src/modules/meetings/core/exporters/post_meeting/path_utils.py ---
"""
Path sanitization, relative Markdown link, and anchor slug helpers for Post/Pre-Meeting exporters.
"""
import logging
import os
import re
import urllib.parse
from pathlib import Path


logger = logging.getLogger("modules.meetings.core.exporters.post_meeting_exporter")


def sanitize_folder_name(name: str) -> str:
    """Sanitizes directory names, removing illegal Windows/filesystem and URL-breaking characters."""
    clean = re.sub(r'[\\/*?:"<>|#%]', "_", str(name or "")).strip()
    return re.sub(r'_+', '_', clean).strip("._ ") or "General"


def safe_relpath(target: Path, base: Path) -> str:
    """
    Computes a clean, forward-slashed relative path between two paths,
    resilient to Windows extended path prefixes (\\?\\) and drive casing differences.
    """
    try:
        t_str = str(target).replace("\\\\?\\", "").replace("//?/", "")
        b_str = str(base).replace("\\\\?\\", "").replace("//?/", "")
        t_abs = Path(t_str).resolve()
        b_abs = Path(b_str).resolve()
        return os.path.relpath(t_abs, b_abs).replace("\\", "/")
    except Exception:
        return str(target).replace("\\", "/")


def quote_markdown_path(rel_path: str) -> str:
    """
    Percent-encodes relative paths for standard CommonMark / GitHub-flavored Markdown links,
    preserving slashes and fragment hashes while encoding spaces (%20) and other special characters.
    """
    if not rel_path:
        return ""
    clean_path = str(rel_path).replace("\\", "/")
    if "#" in clean_path:
        path_part, fragment = clean_path.split("#", 1)
        quoted_path = urllib.parse.quote(path_part, safe="/")
        quoted_frag = urllib.parse.quote(fragment, safe="")
        return f"{quoted_path}#{quoted_frag}"
    return urllib.parse.quote(clean_path, safe="/")


def safe_markdown_rel_link(target: Path, base: Path) -> str:
    """Computes a URL-encoded relative path for Markdown links."""
    rel = safe_relpath(target, base)
    return quote_markdown_path(rel)


def slugify_anchor(name: str) -> str:
    """Generates a clean, URL-safe anchor slug for Markdown and HTML heading targets."""
    slug = re.sub(r'[^a-zA-Z0-9_-]+', '-', str(name or '').lower()).strip('-')
    return slug or "section"
