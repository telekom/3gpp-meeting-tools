# --- File: src/modules/specifications/ui/spec_menu.py ---
"""
Context menu builder and clipboard copy utilities for the Specifications table.
Provides multi-select actions for toggling favorites, editing annotations,
granular clipboard copying (Web link, direct ZIP link, both links, markdown),
metadata synchronization, and deletion.
"""
import html
import json
import logging
from pathlib import Path
import re
from typing import Dict, List, Optional, Tuple, Union

from PyQt5.QtCore import QMimeData, Qt
from PyQt5.QtGui import QCursor
from PyQt5.QtWidgets import (
    QAction,
    QActionGroup,
    QApplication,
    QComboBox,
    QLabel,
    QMenu,
    QToolTip,
    QWidget,
)

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
            f'(<a href="{zip_url}">ZIP {safe_ver}</a>)'
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
        both_md = f"[{canonical}]({dynareport_url}) ([ZIP {clean_ver}]({zip_url}))"
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


def populate_spec_copy_menu(
    copy_sub: QMenu,
    parent_widget: QWidget,
    spec_num: str,
    details: dict = None,
    version_info: dict = None,
    db=None,
):
    """
    Populates the Quick Copy submenu with standard copy actions for a single specification.
    Includes Dual Links (Web + ZIP), Web Link (DynaReport), Direct ZIP Link,
    Markdown links, direct URLs, and Ctrl+C preference setting.
    """
    meta = resolve_spec_metadata(spec_num, db, details, version_info)
    cdata = build_citation_data(meta)

    canonical = meta["canonical_num"]
    title = meta["title"]
    wg = meta["wg"]
    dynareport_url = meta["dynareport_url"]
    zip_url = meta["zip_url"]
    clean_ver = meta["clean_ver"]
    ftp_url = meta["ftp_url"]

    safe_title = html.escape(title)
    safe_wg = html.escape(wg)

    ver_label = f" ({clean_ver})" if clean_ver else ""

    # =========================================================================
    # 1. TOP-LEVEL 1-CLICK ACTIONS
    # =========================================================================
    if zip_url:
        act_top_both = copy_sub.addAction("🔗 Copy Citation (Both Links: Web + ZIP)")
        act_top_both.setToolTip(
            "Copy citation with both links: specification number links to DynaReport, "
            "and version badge links directly to the ZIP download."
        )
        act_top_both.triggered.connect(
            lambda _, p=cdata["cit_with_ver_plain"], h=cdata["cit_both_html"]: _copy_to_clipboard(
                p, h, "Citation (Both Links)", parent_widget
            )
        )

    act_top_web = copy_sub.addAction("🌐 Copy Citation (Web Link - DynaReport)")
    act_top_web.setToolTip("Copy citation with specification number linked to the 3GPP DynaReport status page.")
    act_top_web.triggered.connect(
        lambda _, p=cdata["cit_web_plain"], h=cdata["cit_web_html"]: _copy_to_clipboard(
            p, h, "Citation (Web Link)", parent_widget
        )
    )

    if zip_url:
        act_top_zip = copy_sub.addAction(f"📦 Copy Citation (Direct ZIP Link{ver_label})")
        act_top_zip.setToolTip("Copy citation with specification number linked directly to the ZIP archive.")
        act_top_zip.triggered.connect(
            lambda _, p=cdata["cit_with_ver_plain"], h=cdata["cit_zip_html"]: _copy_to_clipboard(
                p, h, "Citation (ZIP Link)", parent_widget
            )
        )

    copy_sub.addSeparator()

    # =========================================================================
    # 2. DEDICATED SUBMENUS FOR GRANULAR FLAVORS
    # =========================================================================
    if zip_url and dynareport_url:
        sub_both = copy_sub.addMenu("🔗 Dual Links (Web + ZIP)")
        sub_both.setToolTipsVisible(True)
        sub_both.setToolTip("Options including both DynaReport status link and direct ZIP download link.")

        sub_both.addAction("📋 Copy Full Citation (Both Links)").triggered.connect(
            lambda _, p=cdata["cit_with_ver_plain"], h=cdata["cit_both_html"]: _copy_to_clipboard(
                p, h, "Full Citation (Both Links)", parent_widget
            )
        )
        if title:
            sub_both.addAction("📋 Copy Number & Title (Both Links)").triggered.connect(
                lambda _, p=cdata["num_title_plain"], h=cdata["num_title_both_html"]: _copy_to_clipboard(
                    p, h, "Number & Title (Both Links)", parent_widget
                )
            )
        sub_both.addAction(f"📋 Copy Spec Number (Both Links)").triggered.connect(
            lambda _, p=cdata["num_with_ver"], h=cdata["both_anchor"]: _copy_to_clipboard(
                p, h, "Spec Number (Both Links)", parent_widget
            )
        )
        sub_both.addSeparator()
        sub_both.addAction("📋 Copy as Markdown (Both Links)").triggered.connect(
            lambda _, p=cdata["both_md"], h=cdata["both_anchor"]: _copy_to_clipboard(
                p, h, "Markdown (Both Links)", parent_widget
            )
        )
        if title:
            sub_both.addAction("📋 Copy as Markdown with Title (Both Links)").triggered.connect(
                lambda _, p=cdata["cit_both_md"], h=cdata["cit_both_html"]: _copy_to_clipboard(
                    p, h, "Markdown with Title (Both Links)", parent_widget
                )
            )

    sub_web = copy_sub.addMenu("🌐 3GPP Web Links (DynaReport)")
    sub_web.setToolTipsVisible(True)
    sub_web.setToolTip("Options linking exclusively to the official 3GPP DynaReport specification status page.")

    sub_web.addAction("📋 Copy Full Citation (Web Link)").triggered.connect(
        lambda _, p=cdata["cit_web_plain"], h=cdata["cit_web_html"]: _copy_to_clipboard(
            p, h, "Full Citation (Web Link)", parent_widget
        )
    )
    if title:
        sub_web.addAction("📋 Copy Number & Title (Web Link)").triggered.connect(
            lambda _, p=cdata["num_title_web_plain"], h=cdata["num_title_web_html"]: _copy_to_clipboard(
                p, h, "Number & Title (Web Link)", parent_widget
            )
        )
    sub_web.addAction(f"📋 Copy Spec Number ({canonical})").triggered.connect(
        lambda _, p=canonical, h=cdata["web_anchor"]: _copy_to_clipboard(
            p, h, "Spec Number", parent_widget
        )
    )
    if wg:
        sub_web.addAction(f"📋 Copy Number & WG ({canonical} ({wg}))").triggered.connect(
            lambda _, p=f"{canonical} ({wg})", h=f"{cdata['web_anchor']} ({safe_wg})": _copy_to_clipboard(
                p, h, "Number & WG", parent_widget
            )
        )
    if title:
        sub_web.addAction("📋 Copy Title Only").triggered.connect(
            lambda _, p=title, h=safe_title: _copy_to_clipboard(
                p, h, "Title", parent_widget
            )
        )
    sub_web.addSeparator()
    sub_web.addAction(f"📋 Copy as Markdown Link ([{canonical}](url))").triggered.connect(
        lambda _, p=cdata["web_md"], h=cdata["web_anchor"]: _copy_to_clipboard(
            p, h, "Markdown Link", parent_widget
        )
    )
    if title:
        sub_web.addAction("📋 Copy as Markdown with Title").triggered.connect(
            lambda _, p=cdata["cit_web_md"], h=cdata["cit_web_html"]: _copy_to_clipboard(
                p, h, "Markdown with Title", parent_widget
            )
        )

    if zip_url:
        sub_zip = copy_sub.addMenu(f"📦 Direct ZIP Download Links{ver_label}")
        sub_zip.setToolTipsVisible(True)
        sub_zip.setToolTip("Options linking directly to the specification ZIP archive file on 3GPP FTP.")

        sub_zip.addAction(f"📋 Copy Full Citation (ZIP Link{ver_label})").triggered.connect(
            lambda _, p=cdata["cit_with_ver_plain"], h=cdata["cit_zip_html"]: _copy_to_clipboard(
                p, h, f"Full Citation (ZIP{ver_label})", parent_widget
            )
        )
        if title:
            sub_zip.addAction(f"📋 Copy Number & Title (ZIP Link{ver_label})").triggered.connect(
                lambda _, p=cdata["num_title_plain"], h=cdata["num_title_zip_html"]: _copy_to_clipboard(
                    p, h, f"Number & Title (ZIP{ver_label})", parent_widget
                )
            )
        sub_zip.addAction(f"📋 Copy Spec Number (with ZIP Link{ver_label})").triggered.connect(
            lambda _, p=cdata["num_with_ver"], h=cdata["zip_anchor"]: _copy_to_clipboard(
                p, h, f"Spec Number (ZIP{ver_label})", parent_widget
            )
        )
        if wg:
            sub_zip.addAction(f"📋 Copy Number & WG (with ZIP Link)").triggered.connect(
                lambda _, p=f"{canonical} ({clean_ver}) ({wg})", h=f"{cdata['zip_anchor']} ({safe_wg})": _copy_to_clipboard(
                    p, h, f"Number & WG (ZIP{ver_label})", parent_widget
                )
            )
        sub_zip.addSeparator()
        sub_zip.addAction(f"📋 Copy as Markdown (ZIP Link)").triggered.connect(
            lambda _, p=cdata["zip_md"], h=cdata["zip_anchor"]: _copy_to_clipboard(
                p, h, f"Markdown (ZIP{ver_label})", parent_widget
            )
        )
        if title:
            sub_zip.addAction(f"📋 Copy as Markdown with Title (ZIP Link)").triggered.connect(
                lambda _, p=cdata["cit_zip_md"], h=cdata["cit_zip_html"]: _copy_to_clipboard(
                    p, h, f"Markdown with Title (ZIP{ver_label})", parent_widget
                )
            )

    copy_sub.addSeparator()

    # =========================================================================
    # 3. DIRECT RAW URLS
    # =========================================================================
    if dynareport_url:
        act_dyna = copy_sub.addAction("🌐 Copy 3GPP Web Report URL (DynaReport)")
        act_dyna.setToolTip("Copy direct URL to the 3GPP DynaReport specification status page.")
        act_dyna.triggered.connect(
            lambda _, p=dynareport_url, h=f'<a href="{dynareport_url}">{dynareport_url}</a>': _copy_to_clipboard(
                p, h, "Web Report URL", parent_widget
            )
        )

    if zip_url:
        act_zip_url = copy_sub.addAction(f"📦 Copy ZIP Download URL{ver_label}")
        act_zip_url.setToolTip("Copy direct URL to download the specification ZIP archive file.")
        act_zip_url.triggered.connect(
            lambda _, p=zip_url, h=f'<a href="{zip_url}">{zip_url}</a>': _copy_to_clipboard(
                p, h, f"ZIP Download URL{ver_label}", parent_widget
            )
        )

    if ftp_url:
        act_ftp = copy_sub.addAction("📂 Copy FTP Archive Folder URL")
        act_ftp.setToolTip("Copy direct URL to the 3GPP FTP specification directory.")
        act_ftp.triggered.connect(
            lambda _, p=ftp_url, h=f'<a href="{ftp_url}">{ftp_url}</a>': _copy_to_clipboard(
                p, h, "FTP Archive Folder URL", parent_widget
            )
        )

    # =========================================================================
    # 4. PREFERENCES (DEFAULT MODE FOR Ctrl+C)
    # =========================================================================
    copy_sub.addSeparator()
    sub_cfg = copy_sub.addMenu("⚙️ Default Link Mode (Ctrl+C)")
    sub_cfg.setToolTipsVisible(True)
    sub_cfg.setToolTip("Select the default link style when using the Ctrl+C keyboard shortcut on specifications.")

    saved_mode = get_saved_copy_mode()
    mode_group = QActionGroup(sub_cfg)
    mode_group.setExclusive(True)

    mode_options = [
        ("both", "🔗 Both Links (Web + ZIP)", "Spec number links to DynaReport, version badge links to ZIP download."),
        ("web", "🌐 3GPP Web Report (DynaReport)", "Specification number links to DynaReport page."),
        ("zip", "📦 Direct ZIP Download", "Specification number links directly to the ZIP download."),
    ]

    for mode_key, mode_label, mode_tip in mode_options:
        act_mode = sub_cfg.addAction(mode_label)
        act_mode.setCheckable(True)
        act_mode.setChecked(mode_key == saved_mode)
        act_mode.setToolTip(mode_tip)
        mode_group.addAction(act_mode)

        def _make_handler(k=mode_key, l=mode_label):
            return lambda: (save_copy_mode(k), QToolTip.showText(QCursor.pos(), f"📋 Default link mode set to: {l}!", parent_widget))

        act_mode.triggered.connect(_make_handler())


def populate_multi_spec_copy_menu(
    copy_sub: QMenu,
    parent_widget: QWidget,
    target_items: list,
    db,
):
    """
    Populates copy options for multiple selected specifications.
    target_items can be a list of spec number strings or (spec_num, version_info) tuples.
    """
    if not target_items:
        return

    items_processed = []
    for item in target_items:
        if isinstance(item, tuple):
            s_num, v_info = item
            items_processed.append(resolve_spec_metadata(s_num, db, version_info=v_info))
        else:
            items_processed.append(resolve_spec_metadata(str(item), db))

    count = len(items_processed)
    cdatas = [build_citation_data(m) for m in items_processed]

    # 1. Full Citations - Both Links
    both_plain = "\n".join(cd["cit_with_ver_plain"] for cd in cdatas)
    both_html = "<p>" + "<br>".join(cd["cit_both_html"] for cd in cdatas) + "</p>"
    act_both = copy_sub.addAction(f"🔗 Copy Full Citations (Both Links: Web + ZIP) ({count})")
    act_both.setToolTip("Copy complete citations for all specifications with DynaReport and ZIP download links.")
    act_both.triggered.connect(
        lambda _, p=both_plain, h=both_html, c=count: _copy_to_clipboard(
            p, h, f"{c} Citations (Both Links)", parent_widget
        )
    )

    # 2. Full Citations - Web Links
    web_plain = "\n".join(cd["cit_web_plain"] for cd in cdatas)
    web_html = "<p>" + "<br>".join(cd["cit_web_html"] for cd in cdatas) + "</p>"
    act_web = copy_sub.addAction(f"🌐 Copy Full Citations (Web Links) ({count})")
    act_web.setToolTip("Copy complete citations for all specifications with DynaReport web links.")
    act_web.triggered.connect(
        lambda _, p=web_plain, h=web_html, c=count: _copy_to_clipboard(
            p, h, f"{c} Citations (Web Links)", parent_widget
        )
    )

    # 3. Full Citations - Direct ZIP Links
    zip_plain = "\n".join(cd["cit_with_ver_plain"] for cd in cdatas)
    zip_html = "<p>" + "<br>".join(cd["cit_zip_html"] for cd in cdatas) + "</p>"
    act_zip = copy_sub.addAction(f"📦 Copy Full Citations (Direct ZIP Links) ({count})")
    act_zip.setToolTip("Copy complete citations for all specifications with direct ZIP download links.")
    act_zip.triggered.connect(
        lambda _, p=zip_plain, h=zip_html, c=count: _copy_to_clipboard(
            p, h, f"{c} Citations (ZIP Links)", parent_widget
        )
    )

    copy_sub.addSeparator()

    # 4. Spec Numbers Only
    nums_plain = ", ".join(m["canonical_num"] for m in items_processed)
    nums_html = ", ".join(cd["both_anchor"] for cd in cdatas)
    act_nums = copy_sub.addAction(f"📋 Copy Spec Numbers ({count})")
    act_nums.setToolTip("Copy all selected specification numbers as a comma-separated list.")
    act_nums.triggered.connect(
        lambda _, p=nums_plain, h=nums_html, c=count: _copy_to_clipboard(
            p, h, f"{c} Spec Numbers", parent_widget
        )
    )

    # 5. Numbers & Titles
    nt_plain = "\n".join(cd["num_title_plain"] for cd in cdatas)
    nt_html = "<p>" + "<br>".join(cd["num_title_both_html"] for cd in cdatas) + "</p>"
    act_nt = copy_sub.addAction(f"📋 Copy Numbers & Titles ({count})")
    act_nt.setToolTip("Copy specification numbers and titles (one per line).")
    act_nt.triggered.connect(
        lambda _, p=nt_plain, h=nt_html, c=count: _copy_to_clipboard(
            p, h, f"{c} Numbers & Titles", parent_widget
        )
    )

    copy_sub.addSeparator()

    # 6. Markdown Lists
    md_both = "\n".join(f"- {cd['cit_both_md']}" for cd in cdatas)
    act_md_both = copy_sub.addAction(f"🔗 Copy as Markdown List (Both Links) ({count})")
    act_md_both.setToolTip("Copy as Markdown bullet list with both DynaReport and ZIP links.")
    act_md_both.triggered.connect(
        lambda _, p=md_both, h=both_html, c=count: _copy_to_clipboard(
            p, h, f"{c} Markdown Items (Both Links)", parent_widget
        )
    )

    md_web = "\n".join(f"- {cd['cit_web_md']}" for cd in cdatas)
    act_md_web = copy_sub.addAction(f"🌐 Copy as Markdown List (Web Links) ({count})")
    act_md_web.setToolTip("Copy as Markdown bullet list with DynaReport hyperlinks.")
    act_md_web.triggered.connect(
        lambda _, p=md_web, h=web_html, c=count: _copy_to_clipboard(
            p, h, f"{c} Markdown Items (Web Links)", parent_widget
        )
    )

    md_zip = "\n".join(f"- {cd['cit_zip_md']}" for cd in cdatas)
    act_md_zip = copy_sub.addAction(f"📦 Copy as Markdown List (ZIP Links) ({count})")
    act_md_zip.setToolTip("Copy as Markdown bullet list with direct ZIP download hyperlinks.")
    act_md_zip.triggered.connect(
        lambda _, p=md_zip, h=zip_html, c=count: _copy_to_clipboard(
            p, h, f"{c} Markdown Items (ZIP Links)", parent_widget
        )
    )

    copy_sub.addSeparator()

    # 7. URLs Lists
    web_urls = [m["dynareport_url"] for m in items_processed if m["dynareport_url"]]
    if web_urls:
        act_w_urls = copy_sub.addAction(f"🌐 Copy 3GPP Web Report URLs ({len(web_urls)})")
        act_w_urls.setToolTip("Copy direct URLs to 3GPP DynaReport specification pages.")
        act_w_urls.triggered.connect(
            lambda _, p="\n".join(web_urls), h="<br>".join(f'<a href="{u}">{u}</a>' for u in web_urls), c=len(web_urls): _copy_to_clipboard(
                p, h, f"{c} Web URLs", parent_widget
            )
        )

    zip_urls = [m["zip_url"] for m in items_processed if m["zip_url"]]
    if zip_urls:
        act_z_urls = copy_sub.addAction(f"📦 Copy ZIP Download URLs ({len(zip_urls)})")
        act_z_urls.setToolTip("Copy direct URLs to download specification ZIP archives.")
        act_z_urls.triggered.connect(
            lambda _, p="\n".join(zip_urls), h="<br>".join(f'<a href="{u}">{u}</a>' for u in zip_urls), c=len(zip_urls): _copy_to_clipboard(
                p, h, f"{c} ZIP URLs", parent_widget
            )
        )


def copy_specs_default(parent_widget: QWidget, target_items: list, db):
    """
    Default copy behavior triggered by Ctrl+C on the specifications table.
    Uses the user's saved preference ('both', 'web', 'zip').
    """
    if not target_items:
        return

    mode = get_saved_copy_mode()

    items_processed = []
    for item in target_items:
        if isinstance(item, tuple):
            s_num, v_info = item
            items_processed.append(resolve_spec_metadata(s_num, db, version_info=v_info))
        else:
            items_processed.append(resolve_spec_metadata(str(item), db))

    count = len(items_processed)
    cdatas = [build_citation_data(m) for m in items_processed]

    mode_label = "Both Links" if mode == "both" else ("ZIP Link" if mode == "zip" else "Web Link")

    if count == 1:
        cd = cdatas[0]
        m = items_processed[0]
        canonical = m["canonical_num"]

        if mode == "both":
            plain = cd["cit_with_ver_plain"]
            html_t = cd["cit_both_html"]
        elif mode == "zip":
            plain = cd["cit_with_ver_plain"]
            html_t = cd["cit_zip_html"]
        else:
            plain = cd["cit_web_plain"]
            html_t = cd["cit_web_html"]

        _copy_to_clipboard(plain, html_t, f"Specification {canonical} ({mode_label})", parent_widget)
    else:
        if mode == "both":
            plain = "\n".join(cd["cit_with_ver_plain"] for cd in cdatas)
            html_t = "<p>" + "<br>".join(cd["cit_both_html"] for cd in cdatas) + "</p>"
        elif mode == "zip":
            plain = "\n".join(cd["cit_with_ver_plain"] for cd in cdatas)
            html_t = "<p>" + "<br>".join(cd["cit_zip_html"] for cd in cdatas) + "</p>"
        else:
            plain = "\n".join(cd["cit_web_plain"] for cd in cdatas)
            html_t = "<p>" + "<br>".join(cd["cit_web_html"] for cd in cdatas) + "</p>"

        _copy_to_clipboard(plain, html_t, f"{count} Specifications ({mode_label})", parent_widget)


def show_spec_context_menu(tab, position):
    """
    Constructs and executes the context menu for selected specifications rows in the table.

    :param tab: The SpecificationsTab parent widget instance.
    :param position: QPoint position of the context menu request from table viewport.
    """
    selected_rows = tab.table.selectionModel().selectedRows()
    if not selected_rows:
        return

    target_items = []
    for index in selected_rows:
        row = index.row()
        widget = tab.table.cellWidget(row, 0)
        if widget:
            num_label = widget.findChild(QLabel, "specNumberLabel")
            if num_label:
                canonical = num_label.property("specNumber")
                spec_num = ""
                if canonical:
                    spec_num = str(canonical).strip()
                else:
                    display_text = num_label.text().strip()
                    match = re.match(r"^(\d{2}\.\d{2,3}(?:-[A-Za-z0-9]+)?)", display_text)
                    if match:
                        spec_num = match.group(1)

                if spec_num:
                    # Resolve currently selected version from Column 2 combo
                    cell_widget = tab.table.cellWidget(row, 2)
                    v_info = None
                    if cell_widget:
                        combo = cell_widget.findChild(QComboBox)
                        if combo:
                            v_info = combo.currentData()
                    target_items.append((spec_num, v_info))

    if not target_items:
        return

    menu = QMenu(tab)
    menu.setStyleSheet("""
        QMenu { background-color: #FAFAFA; border: 1px solid #CCC; }
        QMenu::item { padding: 5px 20px 5px 15px; color: #333333; }
        QMenu::item:selected { background-color: #E1F0FF; color: #0078D7; }
        QMenu::item:disabled { color: #AAAAAA; }
    """)

    if len(target_items) == 1:
        spec_num, v_info = target_items[0]
        is_fav = spec_num in tab.favorites
        fav_label = "⭐ Remove from Favorites" if is_fav else "⭐ Add to Favorites"
        menu.addAction(fav_label).triggered.connect(lambda _, s=spec_num: tab._toggle_favorite(s))

        edit_ann_action = menu.addAction("🏷️ Edit Tags & Notes...")
        edit_ann_action.triggered.connect(lambda _, s=spec_num: tab._open_edit_annotations(s))

        menu.addSeparator()

        # Quick Copy Submenu
        ver_badge = f" (v{v_info['version']})" if v_info and v_info.get("version") else ""
        copy_sub = menu.addMenu(f"📋 Quick Copy{ver_badge}")
        copy_sub.setToolTipsVisible(True)
        copy_sub.setToolTip("Copy specification metadata in plain text and rich HTML form.")
        populate_spec_copy_menu(copy_sub, tab, spec_num, version_info=v_info, db=tab.db)

        menu.addSeparator()

        info_action = menu.addAction("ℹ️ View Details")
        info_action.triggered.connect(lambda _, s=spec_num: tab._show_spec_info(s))

        web_action = menu.addAction("🌐 Open 3GPP Web Report")
        web_action.triggered.connect(lambda _, s=spec_num: tab._open_web_report(s))

        search_action = menu.addAction("🔎 Track Evolution in Spec Search")
        search_action.triggered.connect(
            lambda _, s=spec_num: tab.search_in_spec_search_requested.emit(s, "")
        )

        menu.addSeparator()

        spec_target_dir = Path(tab.download_dir) / spec_num
        folder_action = menu.addAction("📂 Open Local Folder")
        folder_action.triggered.connect(lambda _, s=spec_num: tab._open_spec_folder(s))
        if not (spec_target_dir.exists() and any(spec_target_dir.iterdir())):
            folder_action.setText("📁 Folder Not Created")
            folder_action.setEnabled(False)

        menu.addSeparator()

        update_action = menu.addAction(f"🔄 Sync metadata for TS {spec_num}")
        update_action.triggered.connect(
            lambda: tab.update_specific_requested.emit([spec_num], tab.force_meta_action.isChecked())
        )

        menu.addSeparator()

        del_action = menu.addAction(f"🗑️ Delete TS {spec_num}")
        del_action.triggered.connect(lambda _, s=spec_num: tab._delete_single_spec(s))
    else:
        target_specs = [item[0] for item in target_items]
        menu.addAction(f"⭐ Add selected ({len(target_specs)}) to Favorites").triggered.connect(
            lambda: tab._batch_set_favorites(target_specs, True)
        )
        menu.addAction(f"🗑️ Remove selected ({len(target_specs)}) from Favorites").triggered.connect(
            lambda: tab._batch_set_favorites(target_specs, False)
        )
        menu.addSeparator()

        batch_tag_action = menu.addAction(f"🏷️ Add Tag to selected ({len(target_specs)})...")
        batch_tag_action.triggered.connect(lambda: tab._batch_add_tag(target_specs))

        menu.addSeparator()

        copy_sub = menu.addMenu(f"📋 Quick Copy ({len(target_items)} selected)")
        copy_sub.setToolTipsVisible(True)
        copy_sub.setToolTip("Copy metadata and download links for all selected specifications.")
        populate_multi_spec_copy_menu(copy_sub, tab, target_items, db=tab.db)

        menu.addSeparator()

        update_action = menu.addAction(f"🔄 Sync metadata for selected ({len(target_specs)})")
        update_action.triggered.connect(
            lambda: tab.update_specific_requested.emit(target_specs, tab.force_meta_action.isChecked())
        )

        menu.addSeparator()

        del_action = menu.addAction(
            f"🗑️ Delete selected ({len(target_specs)}) specifications"
        )
        del_action.triggered.connect(lambda: tab._delete_batch_specs(target_specs))

    menu.exec_(tab.table.viewport().mapToGlobal(position))
