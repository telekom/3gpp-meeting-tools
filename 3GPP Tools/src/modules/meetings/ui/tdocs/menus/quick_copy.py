"""Quick-copy clipboard helpers, cache checks, and portal/folder launchers for TDocs."""

from __future__ import annotations

import html
import logging
import os
import re
import sys
import webbrowser
from pathlib import Path

from PyQt5.QtCore import QMimeData
from PyQt5.QtGui import QCursor
from PyQt5.QtWidgets import QApplication, QMenu, QToolTip, QWidget

from core.utils.portal_utils import get_portal_tdoc_url
from modules.meetings.core.tdocs.tdoc_url_resolver import is_tdoc_revision, resolve_tdoc_download_url

logger = logging.getLogger(__name__)


def is_tdoc_cached(meeting_dir: Path, target_name: str) -> bool:
    """Checks if a TDoc or revision file exists locally on disk."""
    if not meeting_dir:
        return False

    name = target_name.strip()
    base_match = re.match(r"^([A-Za-z0-9]+-\d+)", name)
    base_tdoc = base_match.group(1).upper() if base_match else name.upper()

    mtg_path = Path(meeting_dir)
    candidate_roots = []
    if (mtg_path / "TDocs").is_dir():
        candidate_roots.append(mtg_path / "TDocs")
    candidate_roots.append(mtg_path)

    for root in candidate_roots:
        candidates = [
            root / f"{name}.zip",
            root / f"{name}.docx",
            root / f"{name}.doc",
            root / base_tdoc / f"{name}.zip",
            root / base_tdoc / f"{name}.docx",
            root / base_tdoc / f"{name}.doc",
        ]
        if any(c.exists() for c in candidates):
            return True

        sub_dir = root / base_tdoc
        if sub_dir.is_dir():
            for f in sub_dir.glob(f"*{name}*"):
                if f.is_file():
                    return True
    return False


def open_tdoc_local_folder(meeting_dir: Path, base_tdoc: str):
    """Opens the TDoc directory or meeting TDocs folder in Windows Explorer."""
    tdocs_dir = Path(meeting_dir)
    target_sub = tdocs_dir / base_tdoc
    logger.info(f"📂 Opening local TDoc folder: {target_sub}")
    target = target_sub if target_sub.exists() else tdocs_dir
    if not target.exists():
        target.mkdir(parents=True, exist_ok=True)
    if hasattr(os, "startfile"):
        logger.debug(f"Opening resolved local folder: {target}")
        os.startfile(str(target))
    else:
        webbrowser.open(f"file:///{target}")


def open_3gu_portal(base_tdoc: str):
    """Opens the official 3GU portal view for the target TDoc."""
    url = get_portal_tdoc_url(base_tdoc)
    webbrowser.open(url)


def _copy_to_clipboard(plain_text: str, html_text: str = None, label: str = "text", parent: QWidget = None):
    """
    Copies both plain text and optional rich HTML text to the clipboard.
    Rich text apps (Outlook, Word) paste the HTML version, while plain text
    editors (Notepad, terminal) seamlessly fall back to the plain text version.
    """
    mime = QMimeData()
    mime.setText(plain_text)
    if html_text:
        mime.setHtml(html_text)
    QApplication.clipboard().setMimeData(mime)
    QToolTip.showText(QCursor.pos(), f"📋 Copied {label} to clipboard!", parent)


def _resolve_copy_helper():
    """Preserve unittest.mock.patch hooks targeting ``tdocs_menus._copy_to_clipboard``."""
    facade = sys.modules.get("modules.meetings.ui.tdocs_menus")
    if facade is not None and hasattr(facade, "_copy_to_clipboard"):
        return getattr(facade, "_copy_to_clipboard")
    return _copy_to_clipboard


def populate_quick_copy_menu(
    menu: QMenu,
    tdoc_id: str,
    row_data: dict = None,
    docs_ftp_url: str = "",
    revisions_url: str = "",
    main_url: str = "",
    parent: QWidget = None,
) -> QMenu:
    """
    Populates a QMenu with standardized 3GPP clipboard copy actions:
    - TDoc identifier (plain text + rich HTML hyperlink)
    - Document title (if available)
    - Full citation (TDoc: Title (Source) with rich link)
    - Direct archive download URL (.zip on 3GPP FTP/HTTPS)
    - Baseline document download URL (if different from revision URL)
    - 3GU Portal URL (official web link)
    - Markdown hyperlink syntax ([TDoc](url))
    """
    menu.setToolTipsVisible(True)
    row_data = row_data or {}
    clean_tdoc = str(tdoc_id or row_data.get("TDoc", "")).strip()
    if not clean_tdoc:
        return menu

    base_match = re.match(r"^([A-Za-z0-9]+-\d+)", clean_tdoc)
    base_tdoc = base_match.group(1).upper() if base_match else clean_tdoc.upper()
    is_rev = is_tdoc_revision(clean_tdoc)

    # Resolve direct download URLs via central resolver
    download_url = resolve_tdoc_download_url(
        clean_tdoc,
        docs_url=docs_ftp_url,
        revisions_url=revisions_url,
        main_url=main_url,
        meta=row_data,
    )

    if is_rev:
        baseline_url = resolve_tdoc_download_url(
            base_tdoc,
            docs_url=docs_ftp_url,
            revisions_url=revisions_url,
            main_url=main_url,
            meta=row_data,
        )
    else:
        baseline_url = download_url

    portal_url = get_portal_tdoc_url(base_tdoc)
    best_url = download_url or baseline_url or portal_url

    # HTML formatted anchor tag for the TDoc
    safe_tdoc = html.escape(clean_tdoc)
    tdoc_html_link = f'<a href="{best_url}">{safe_tdoc}</a>' if best_url else safe_tdoc

    def copy_action(plain_text: str, html_text: str = None, label: str = "text", target_parent: QWidget = None):
        _resolve_copy_helper()(
            plain_text=plain_text,
            html_text=html_text,
            label=label,
            parent=target_parent,
        )

    # Action 1: Copy TDoc Number
    act_copy_id = menu.addAction(f"📋 Copy TDoc Number ({clean_tdoc})")
    act_copy_id.setToolTip("Copy the TDoc number; rich-text targets also receive a clickable link.")
    act_copy_id.triggered.connect(
        lambda: copy_action(
            plain_text=clean_tdoc,
            html_text=tdoc_html_link,
            label="TDoc Number",
            target_parent=parent or menu,
        )
    )

    # Action 2: Copy Title (if present)
    title = str(row_data.get("Title", "")).strip()
    safe_title = html.escape(title)
    if title:
        act_copy_title = menu.addAction("📋 Copy Title")
        act_copy_title.setToolTip("Copy the TDoc title.")
        act_copy_title.triggered.connect(
            lambda: copy_action(
                plain_text=title,
                html_text=safe_title,
                label="Title",
                target_parent=parent or menu,
            )
        )

    # Action 3: Copy Full Citation (if title present)
    source = str(row_data.get("Source", "")).strip()
    if title:
        citation_plain = f"{clean_tdoc}: {title} ({source})" if source else f"{clean_tdoc}: {title}"
        safe_source = html.escape(source)
        citation_html = f"{tdoc_html_link}: {safe_title} ({safe_source})" if source else f"{tdoc_html_link}: {safe_title}"

        act_copy_cit = menu.addAction("📋 Copy Full Citation")
        act_copy_cit.setToolTip("Copy a compact TDoc citation including title and source.")
        act_copy_cit.triggered.connect(
            lambda: copy_action(
                plain_text=citation_plain,
                html_text=citation_html,
                label="Citation",
                target_parent=parent or menu,
            )
        )

    # Action 4: Copy Document Download URL (Direct ZIP archive link)
    if download_url:
        url_label = "📋 Copy Revision Download URL" if is_rev else "📋 Copy Document Download URL"
        act_copy_dl = menu.addAction(url_label)
        act_copy_dl.setToolTip(f"Copy the direct download URL for the document archive ({download_url}).")
        act_copy_dl.triggered.connect(
            lambda: copy_action(
                plain_text=download_url,
                html_text=f'<a href="{download_url}">{download_url}</a>',
                label="Document Download URL",
                target_parent=parent or menu,
            )
        )

    # Action 5: Copy Baseline Document URL (if different from revision URL)
    if is_rev and baseline_url and baseline_url != download_url:
        act_copy_base_url = menu.addAction("📋 Copy Baseline Document URL")
        act_copy_base_url.setToolTip(f"Copy the baseline document archive download link ({baseline_url}).")
        act_copy_base_url.triggered.connect(
            lambda: copy_action(
                plain_text=baseline_url,
                html_text=f'<a href="{baseline_url}">{baseline_url}</a>',
                label="Baseline Document URL",
                target_parent=parent or menu,
            )
        )

    # Action 6: Copy 3GU Portal Web URL
    act_copy_portal = menu.addAction("🌐 Copy 3GU Portal URL")
    act_copy_portal.setToolTip(f"Copy official 3GPP 3GU Portal web link ({portal_url}).")
    act_copy_portal.triggered.connect(
        lambda: copy_action(
            plain_text=portal_url,
            html_text=f'<a href="{portal_url}">{portal_url}</a>',
            label="3GU Portal URL",
            target_parent=parent or menu,
        )
    )

    # Action 7: Copy Markdown Link
    if best_url:
        md_text = f"[{clean_tdoc}]({best_url})"
        act_copy_md = menu.addAction("🔗 Copy Markdown Link")
        act_copy_md.setToolTip(f"Copy as Markdown hyperlink: {md_text}")
        act_copy_md.triggered.connect(
            lambda: copy_action(
                plain_text=md_text,
                html_text=tdoc_html_link,
                label="Markdown Link",
                target_parent=parent or menu,
            )
        )

    return menu


def build_quick_copy_menu(
    parent: QWidget = None,
    tdoc_id: str = "",
    row_data: dict = None,
    docs_ftp_url: str = "",
    revisions_url: str = "",
    main_url: str = "",
) -> QMenu:
    """Builds and returns a standalone QMenu pre-populated with Quick Copy actions."""
    menu = QMenu(parent)
    menu.setToolTipsVisible(True)
    populate_quick_copy_menu(
        menu=menu,
        tdoc_id=tdoc_id,
        row_data=row_data,
        docs_ftp_url=docs_ftp_url,
        revisions_url=revisions_url,
        main_url=main_url,
        parent=parent,
    )
    return menu
