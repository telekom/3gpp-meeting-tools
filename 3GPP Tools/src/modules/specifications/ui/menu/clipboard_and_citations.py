"""Clipboard, configuration preference, metadata resolution, and citation formatting."""

from __future__ import annotations

import html
import json
import logging
import re

from PyQt5.QtCore import QMimeData
from PyQt5.QtGui import QCursor
from PyQt5.QtWidgets import QApplication, QToolTip, QWidget

from core.utils.paths import get_project_root
from modules.specifications.utils.utils import build_dynareport_url

logger = logging.getLogger(__name__)

CONFIG_FILE_NAME = "specs_config.json"


def get_saved_copy_mode() -> str:
    """Returns the user's preferred copy link mode ('both', 'web', 'zip'). Defaults to 'both'."""
    try:
        cfg_path = get_project_root() / "config" / CONFIG_FILE_NAME
        if cfg_path.exists():
            with open(cfg_path, "r", encoding="utf-8") as f:
                data = json.load(f)
                mode = data.get("copy_link_mode", "both")
                if mode in ("both", "web", "zip"):
                    return mode
    except Exception:
        pass
    return "both"


def save_copy_mode(mode: str):
    """Persists the preferred copy link mode ('both', 'web', 'zip') to specs_config.json."""
    if mode not in ("both", "web", "zip"):
        return
    try:
        cfg_path = get_project_root() / "config" / CONFIG_FILE_NAME
        data = {}
        if cfg_path.exists():
            with open(cfg_path, "r", encoding="utf-8") as f:
                data = json.load(f)
        data["copy_link_mode"] = mode
        with open(cfg_path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=4)
    except Exception as e:
        logger.error(f"Failed to persist copy_link_mode to specs_config.json: {e}")


def _copy_to_clipboard(
    plain_text: str,
    html_text: str = None,
    label: str = "text",
    parent: QWidget = None,
):
    """
    Copies both plain text and optional rich HTML text to the system clipboard.
    Rich-text applications (Outlook, Word, Teams, OneNote) paste the formatted HTML version,
    while plain text editors (Notepad, code editor, terminal) seamlessly fall back to plain text.
    """
    mime = QMimeData()
    mime.setText(plain_text)
    if html_text:
        mime.setHtml(html_text)
    QApplication.clipboard().setMimeData(mime)
    QToolTip.showText(QCursor.pos(), f"📋 Copied {label} to clipboard!", parent)


def format_canonical_spec_number(spec_num: str, spec_type: str = "TS") -> str:
    """Normalizes and prefixes the specification number with its type (e.g. 'TS 23.501' or 'TR 38.801')."""
    clean = str(spec_num or "").strip()
    clean = re.sub(r"^(?:3GPP\s+)?(?:TS|TR)\s*", "", clean, flags=re.IGNORECASE).strip()
    t = (spec_type or "TS").strip().upper()
    if t not in ("TS", "TR"):
        t = "TS"
    return f"{t} {clean}" if clean else str(spec_num or "").strip()


def resolve_spec_metadata(
    spec_num: str,
    db,
    details: dict = None,
    version_info: dict = None,
) -> dict:
    """Resolves specification metadata and archive file links from DB or existing objects."""
    if details:
        spec_data = dict(details)
    elif db:
        spec_data = db.get_spec_details(spec_num)
        if not spec_data and "-" in spec_num:
            base_num = spec_num.split("-")[0]
            spec_data = db.get_spec_details(base_num)
        if not spec_data:
            spec_data = {}
    else:
        spec_data = {}

    s_type = spec_data.get("type") or "TS"
    canonical_num = format_canonical_spec_number(spec_num, s_type)
    title = (spec_data.get("title") or "").strip()
    wg = (spec_data.get("primary_group") or "").strip()
    if wg in ("-", "None"):
        wg = ""
    dynareport_url = build_dynareport_url(spec_num)
    ftp_url = (spec_data.get("url") or "").strip()

    # Resolve target version and direct ZIP URL
    version = ""
    zip_url = ""
    zip_filename = ""
    if version_info and isinstance(version_info, dict):
        version = version_info.get("version") or ""
        zip_url = version_info.get("url") or ""
        zip_filename = version_info.get("fname") or version_info.get("filename") or ""
    elif db and hasattr(db, "get_spec_files"):
        files = db.get_spec_files(spec_num)
        if files:
            latest = files[0]
            version = latest.get("version") or ""
            zip_url = latest.get("url") or ""
            zip_filename = latest.get("filename") or ""

    clean_ver = ""
    if version:
        v_str = str(version).strip()
        clean_ver = v_str if v_str.startswith("v") else f"v{v_str}"

    return {
        "spec_num": spec_num,
        "type": s_type,
        "canonical_num": canonical_num,
        "title": title,
        "wg": wg,
        "dynareport_url": dynareport_url,
        "ftp_url": ftp_url,
        "version": version,
        "clean_ver": clean_ver,
        "zip_url": zip_url,
        "zip_filename": zip_filename,
        "details": spec_data,
    }


def build_citation_data(meta: dict) -> dict:
    """
    Computes plain text, rich HTML, and Markdown representations for all three link modes:
    - both: Spec number links to DynaReport, version badge links to ZIP download.
    - web:  Spec number links to DynaReport.
    - zip:  Spec number (+ version) links directly to ZIP download.
    """
    canonical = meta["canonical_num"]
    title = meta["title"]
    wg = meta["wg"]
    dynareport_url = meta["dynareport_url"]
    zip_url = meta["zip_url"]
    clean_ver = meta["clean_ver"]

    safe_canonical = html.escape(canonical)
    safe_title = html.escape(title)
    safe_wg = html.escape(wg)
    safe_ver = html.escape(clean_ver)

    # 1. HTML Anchors
    web_anchor = (
        f'<a href="{dynareport_url}">{safe_canonical}</a>'
        if dynareport_url
        else safe_canonical
    )

    if zip_url and safe_ver:
        zip_anchor = f'<a href="{zip_url}">{safe_canonical} ({safe_ver})</a>'
    elif zip_url:
        zip_anchor = f'<a href="{zip_url}">{safe_canonical}</a>'
    else:
        zip_anchor = safe_canonical

    if dynareport_url and zip_url and safe_ver:
        both_anchor = (
            f'<a href="{dynareport_url}">{safe_canonical}</a> '
            f'(<a href="{zip_url}">{safe_ver}</a>)'
        )
    elif dynareport_url and zip_url:
        both_anchor = (
            f'<a href="{dynareport_url}">{safe_canonical}</a> '
            f'(<a href="{zip_url}">ZIP</a>)'
        )
    elif zip_url:
        both_anchor = zip_anchor
    else:
        both_anchor = web_anchor

    # 2. Markdown Hyperlinks
    web_md = f"[{canonical}]({dynareport_url})" if dynareport_url else canonical
    if zip_url and clean_ver:
        zip_md = f"[{canonical} ({clean_ver})]({zip_url})"
    elif zip_url:
        zip_md = f"[{canonical}]({zip_url})"
    else:
        zip_md = canonical

    if dynareport_url and zip_url and clean_ver:
        both_md = f"[{canonical}]({dynareport_url}) ([{clean_ver}]({zip_url}))"
    elif dynareport_url and zip_url:
        both_md = f"[{canonical}]({dynareport_url}) ([ZIP]({zip_url}))"
    elif zip_url:
        both_md = zip_md
    else:
        both_md = web_md

    # 3. Plain text numbers
    num_with_ver = f"{canonical} ({clean_ver})" if clean_ver else canonical
    num_web = canonical

    # 4. Full Citations
    if title and wg:
        cit_with_ver_plain = f"{num_with_ver}: {title} ({wg})"
        cit_web_plain = f"{num_web}: {title} ({wg})"
    elif title:
        cit_with_ver_plain = f"{num_with_ver}: {title}"
        cit_web_plain = f"{num_web}: {title}"
    else:
        cit_with_ver_plain = num_with_ver
        cit_web_plain = num_web

    # Full Citation HTML
    if title and wg:
        cit_both_html = f"{both_anchor}: {safe_title} ({safe_wg})"
        cit_web_html = f"{web_anchor}: {safe_title} ({safe_wg})"
        cit_zip_html = f"{zip_anchor}: {safe_title} ({safe_wg})"
    elif title:
        cit_both_html = f"{both_anchor}: {safe_title}"
        cit_web_html = f"{web_anchor}: {safe_title}"
        cit_zip_html = f"{zip_anchor}: {safe_title}"
    else:
        cit_both_html = both_anchor
        cit_web_html = web_anchor
        cit_zip_html = zip_anchor

    # Full Citation Markdown
    if title and wg:
        cit_both_md = f"{both_md}: {title} ({wg})"
        cit_web_md = f"{web_md}: {title} ({wg})"
        cit_zip_md = f"{zip_md}: {title} ({wg})"
    elif title:
        cit_both_md = f"{both_md}: {title}"
        cit_web_md = f"{web_md}: {title}"
        cit_zip_md = f"{zip_md}: {title}"
    else:
        cit_both_md = both_md
        cit_web_md = web_md
        cit_zip_md = zip_md

    # Number & Title
    num_title_plain = f"{num_with_ver}: {title}" if title else num_with_ver
    num_title_web_plain = f"{num_web}: {title}" if title else num_web

    num_title_both_html = f"{both_anchor}: {safe_title}" if title else both_anchor
    num_title_web_html = f"{web_anchor}: {safe_title}" if title else web_anchor
    num_title_zip_html = f"{zip_anchor}: {safe_title}" if title else zip_anchor

    return {
        "both_anchor": both_anchor,
        "web_anchor": web_anchor,
        "zip_anchor": zip_anchor,
        "both_md": both_md,
        "web_md": web_md,
        "zip_md": zip_md,
        "num_with_ver": num_with_ver,
        "num_web": num_web,
        "cit_with_ver_plain": cit_with_ver_plain,
        "cit_web_plain": cit_web_plain,
        "cit_both_html": cit_both_html,
        "cit_web_html": cit_web_html,
        "cit_zip_html": cit_zip_html,
        "cit_both_md": cit_both_md,
        "cit_web_md": cit_web_md,
        "cit_zip_md": cit_zip_md,
        "num_title_plain": num_title_plain,
        "num_title_web_plain": num_title_web_plain,
        "num_title_both_html": num_title_both_html,
        "num_title_web_html": num_title_web_html,
        "num_title_zip_html": num_title_zip_html,
    }
