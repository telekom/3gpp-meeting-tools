# --- File: src/modules/meetings/ui/tdocs_menus.py ---
import html
import os
import re
import webbrowser
import logging
from pathlib import Path

from PyQt5.QtCore import QPoint, QMimeData
from PyQt5.QtGui import QCursor
from PyQt5.QtWidgets import QApplication, QMenu, QToolTip, QWidget

from core.utils.portal_utils import get_portal_tdoc_url, open_portal_tdoc
from modules.meetings.core.tdocs.tdoc_url_resolver import is_tdoc_revision, resolve_tdoc_download_url

logger = logging.getLogger(__name__)


def is_tdoc_cached(meeting_dir: Path, target_name: str) -> bool:
    """Checks if a TDoc or revision file exists locally on disk."""
    if not meeting_dir:
        return False
    tdocs_dir = Path(meeting_dir) / "TDocs"
    if not tdocs_dir.exists():
        return False

    name = target_name.strip()
    candidates = [
        tdocs_dir / f"{name}.zip",
        tdocs_dir / f"{name}.docx",
        tdocs_dir / f"{name}.doc",
    ]
    if any(c.exists() for c in candidates):
        return True

    # Check extracted subdirectories: TDocs/{base_tdoc}/*{name}*
    base_match = re.match(r"^([A-Za-z0-9]+-\d+)", name)
    base_tdoc = base_match.group(1).upper() if base_match else name.upper()
    sub_dir = tdocs_dir / base_tdoc
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

    baseline_url = ""
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
    if best_url:
        tdoc_html_link = f'<a href="{best_url}">{safe_tdoc}</a>'
    else:
        tdoc_html_link = safe_tdoc

    # Action 1: Copy TDoc Number
    act_copy_id = menu.addAction(f"📋 Copy TDoc Number ({clean_tdoc})")
    act_copy_id.setToolTip('Copy the TDoc number; rich-text targets also receive a clickable link.')
    act_copy_id.triggered.connect(
        lambda: _copy_to_clipboard(
            plain_text=clean_tdoc,
            html_text=tdoc_html_link,
            label="TDoc Number",
            parent=parent or menu,
        )
    )

    # Action 2: Copy Title (if present)
    title = str(row_data.get("Title", "")).strip()
    safe_title = html.escape(title)
    if title:
        act_copy_title = menu.addAction("📋 Copy Title")
        act_copy_title.setToolTip('Copy the TDoc title.')
        act_copy_title.triggered.connect(
            lambda: _copy_to_clipboard(
                plain_text=title,
                html_text=safe_title,
                label="Title",
                parent=parent or menu,
            )
        )

    # Action 3: Copy Full Citation (if title present)
    source = str(row_data.get("Source", "")).strip()
    if title:
        citation_plain = f"{clean_tdoc}: {title} ({source})" if source else f"{clean_tdoc}: {title}"
        safe_source = html.escape(source)
        if source:
            citation_html = f"{tdoc_html_link}: {safe_title} ({safe_source})"
        else:
            citation_html = f"{tdoc_html_link}: {safe_title}"

        act_copy_cit = menu.addAction("📋 Copy Full Citation")
        act_copy_cit.setToolTip('Copy a compact TDoc citation including title and source.')
        act_copy_cit.triggered.connect(
            lambda: _copy_to_clipboard(
                plain_text=citation_plain,
                html_text=citation_html,
                label="Citation",
                parent=parent or menu,
            )
        )

    # Action 4: Copy Document Download URL (Direct ZIP archive link)
    if download_url:
        url_label = "📋 Copy Revision Download URL" if is_rev else "📋 Copy Document Download URL"
        act_copy_dl = menu.addAction(url_label)
        act_copy_dl.setToolTip(f'Copy the direct download URL for the document archive ({download_url}).')
        act_copy_dl.triggered.connect(
            lambda: _copy_to_clipboard(
                plain_text=download_url,
                html_text=f'<a href="{download_url}">{download_url}</a>',
                label="Document Download URL",
                parent=parent or menu,
            )
        )

    # Action 5: Copy Baseline Document URL (if different from revision URL)
    if is_rev and baseline_url and baseline_url != download_url:
        act_copy_base_url = menu.addAction("📋 Copy Baseline Document URL")
        act_copy_base_url.setToolTip(f'Copy the baseline document archive download link ({baseline_url}).')
        act_copy_base_url.triggered.connect(
            lambda: _copy_to_clipboard(
                plain_text=baseline_url,
                html_text=f'<a href="{baseline_url}">{baseline_url}</a>',
                label="Baseline Document URL",
                parent=parent or menu,
            )
        )

    # Action 6: Copy 3GU Portal Web URL
    act_copy_portal = menu.addAction("🌐 Copy 3GU Portal URL")
    act_copy_portal.setToolTip(f'Copy official 3GPP 3GU Portal web link ({portal_url}).')
    act_copy_portal.triggered.connect(
        lambda: _copy_to_clipboard(
            plain_text=portal_url,
            html_text=f'<a href="{portal_url}">{portal_url}</a>',
            label="3GU Portal URL",
            parent=parent or menu,
        )
    )

    # Action 7: Copy Markdown Link
    if best_url:
        md_text = f"[{clean_tdoc}]({best_url})"
        act_copy_md = menu.addAction("🔗 Copy Markdown Link")
        act_copy_md.setToolTip(f'Copy as Markdown hyperlink: {md_text}')
        act_copy_md.triggered.connect(
            lambda: _copy_to_clipboard(
                plain_text=md_text,
                html_text=tdoc_html_link,
                label="Markdown Link",
                parent=parent or menu,
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


# =========================================================================
# 1. STREAMLINED ACTION MENU (COLUMN 0 "OPEN" BUTTON)
# =========================================================================
def build_action_menu(
    parent: QWidget,
    base_tdoc: str,
    docs_ftp_url: str,
    revisions_url: str,
    revisions: list,
    meeting_dir: Path,
    download_callback,
    export_llm_callback=None,
    compose_email_callback=None,
    pos: QPoint = None,
    compare_active_word_callback=None,
):
    """Fast, dedicated document opener for the Column 0 button."""
    menu = QMenu(parent)
    menu.setToolTipsVisible(True)
    revisions = revisions or []
    base_cached = is_tdoc_cached(meeting_dir, base_tdoc)

    if revisions:
        latest_rev = revisions[-1]
        latest_cached = is_tdoc_cached(meeting_dir, latest_rev)
        act_latest = menu.addAction(
            f"🚀 Open Latest: {latest_rev}" + (" (Local)" if latest_cached else "")
        )
        act_latest.setToolTip('Open the newest known revision of this TDoc.')
        act_latest.triggered.connect(
            lambda: download_callback(
                base_tdoc, latest_rev, is_silent_compare=False
            )
        )
        menu.addSeparator()

        act_base = menu.addAction(
            f"📄 Open Base: {base_tdoc}" + (" (Local)" if base_cached else "")
        )
        act_base.setToolTip('Open the baseline TDoc.')
        act_base.triggered.connect(
            lambda: download_callback(
                base_tdoc, base_tdoc, is_silent_compare=False
            )
        )

        for rev in revisions:
            rev_cached = is_tdoc_cached(meeting_dir, rev)
            act_rev = menu.addAction(
                f"📄 Open Rev: {rev}" + (" (Local)" if rev_cached else "")
            )
            act_rev.triggered.connect(
                lambda _, r=rev: download_callback(
                    base_tdoc, r, is_silent_compare=False
                )
            )
    else:
        act_base = menu.addAction(
            f"📄 Open Base: {base_tdoc}" + (" (Local)" if base_cached else "")
        )
        act_base.setToolTip('Open the baseline TDoc.')
        act_base.triggered.connect(
            lambda: download_callback(
                base_tdoc, base_tdoc, is_silent_compare=False
            )
        )

    if compare_active_word_callback:
        menu.addSeparator()
        if revisions:
            latest_rev = revisions[-1]
            act_cmp_lat = menu.addAction(f"⚖️ Compare Latest ({latest_rev}) vs Active Word")
            act_cmp_lat.setToolTip(
                "Compare the newest revision of this TDoc against the document currently active in Microsoft Word."
            )
            act_cmp_lat.triggered.connect(
                lambda: compare_active_word_callback(base_tdoc, latest_rev)
            )
            act_cmp_base = menu.addAction(f"⚖️ Compare Base ({base_tdoc}) vs Active Word")
            act_cmp_base.setToolTip(
                "Compare the baseline TDoc against the document currently active in Microsoft Word."
            )
            act_cmp_base.triggered.connect(
                lambda: compare_active_word_callback(base_tdoc, base_tdoc)
            )
        else:
            act_cmp_base = menu.addAction(f"⚖️ Compare {base_tdoc} vs Active Word")
            act_cmp_base.setToolTip(
                "Compare this TDoc against the document currently active in Microsoft Word."
            )
            act_cmp_base.triggered.connect(
                lambda: compare_active_word_callback(base_tdoc, base_tdoc)
            )

    menu.addSeparator()
    act_folder = menu.addAction("📂 Open Local Folder")
    act_folder.setToolTip('Open the local folder used to cache this TDoc and its extracted files.')
    act_folder.triggered.connect(
        lambda: open_tdoc_local_folder(meeting_dir, base_tdoc)
    )

    act_portal = menu.addAction("🌐 View TDoc on 3GU Portal")
    act_portal.setToolTip('Open this TDoc in the official 3GPP 3GU portal.')
    act_portal.triggered.connect(lambda: open_3gu_portal(base_tdoc))

    menu.exec_(pos or QCursor.pos())


# =========================================================================
# 2. UNIFIED ROW CONTEXT MENU (RIGHT-CLICK COMMAND CENTER)
# =========================================================================
def build_row_context_menu(
    parent: QWidget,
    row_data: dict,
    revisions: list,
    family: list,
    docs_ftp_url: str,
    revisions_url: str,
    meeting_dir: Path,
    download_callback,
    export_llm_callback,
    compose_email_callback,
    open_emails_callback,
    mark_read_callback,
    mark_unread_callback,
    open_details_callback,
    unread_emails_count: int = 0,
    pos: QPoint = None,
    triage_callback=None,  # <--- Added parameter
    draft_cr_callback=None,
    compare_active_word_callback=None,
    export_dossier_callback=None,
    load_session_plan_callback=None,
):
    """Complete command center menu for row-level operations."""
    menu = QMenu(parent)
    menu.setToolTipsVisible(True)
    revisions = revisions or []
    tdoc_id = str(row_data.get("TDoc", "")).strip()

    # --- 1. INSPECT FULL DETAILS & AI TRIAGE ---
    act_details = menu.addAction(f"ℹ️ View Full Details for {tdoc_id}...")
    act_details.setToolTip('Show all stored metadata and links for this TDoc.')
    act_details.triggered.connect(open_details_callback)

    if triage_callback:
        act_triage = menu.addAction(f"🤖 AI Technical Triage for {tdoc_id}...")
        act_triage.setToolTip('Run the local-LLM technical triage for this TDoc.')
        act_triage.triggered.connect(lambda: triage_callback(tdoc_id))

    if draft_cr_callback:
        act_draft_cr = menu.addAction(f"📝 Draft CR / pCR from Spec...")
        act_draft_cr.setToolTip("Compose a Change Request (CR) or Pseudo-CR (pCR) pre-populated from this TDoc.")
        act_draft_cr.triggered.connect(lambda: draft_cr_callback(row_data))

    # --- 1b. SESSION PLAN INGESTION SHORTCUT ---
    title_lower = str(row_data.get("Title", "")).lower()
    if "session plan" in title_lower or "time allocation" in title_lower:
        act_sp = menu.addAction(f"🗓️ Load {tdoc_id} into Session Plan...")
        act_sp.setToolTip("Parse and load this TDoc into the Meeting Session Plan viewer.")
        if load_session_plan_callback:
            act_sp.triggered.connect(lambda: load_session_plan_callback(tdoc_id))

    menu.addSeparator()

    # --- 2. OPEN DOCUMENT SUBMENU ---
    open_sub = menu.addMenu("📂 Open Document")
    open_sub.setToolTipsVisible(True)
    open_sub.setToolTip("Open the base TDoc or one of its revisions, using the local cache when available.")
    base_cached = is_tdoc_cached(meeting_dir, tdoc_id)

    if revisions:
        latest_rev = revisions[-1]
        latest_cached = is_tdoc_cached(meeting_dir, latest_rev)
        act_open_lat = open_sub.addAction(
            f"🚀 Open Latest: {latest_rev}"
            + (" (Local)" if latest_cached else "")
        )
        act_open_lat.setToolTip('Open the newest known revision of this TDoc.')
        act_open_lat.triggered.connect(
            lambda: download_callback(
                tdoc_id, latest_rev, is_silent_compare=False
            )
        )
        open_sub.addSeparator()

        act_open_base = open_sub.addAction(
            f"📄 Open Base: {tdoc_id}" + (" (Local)" if base_cached else "")
        )
        act_open_base.setToolTip('Open the baseline TDoc.')
        act_open_base.triggered.connect(
            lambda: download_callback(tdoc_id, tdoc_id, is_silent_compare=False)
        )

        for rev in revisions:
            rev_cached = is_tdoc_cached(meeting_dir, rev)
            act_open_r = open_sub.addAction(
                f"📄 Open Rev: {rev}" + (" (Local)" if rev_cached else "")
            )
            act_open_r.triggered.connect(
                lambda _, r=rev: download_callback(
                    tdoc_id, r, is_silent_compare=False
                )
            )
    else:
        act_open_base = open_sub.addAction(
            f"📄 Open Base: {tdoc_id}" + (" (Local)" if base_cached else "")
        )
        act_open_base.setToolTip('Open the baseline TDoc.')
        act_open_base.triggered.connect(
            lambda: download_callback(tdoc_id, tdoc_id, is_silent_compare=False)
        )

    open_sub.addSeparator()
    act_folder = open_sub.addAction("📂 Open Local Folder")
    act_folder.setToolTip('Open the local folder used to cache this TDoc and its extracted files.')
    act_folder.triggered.connect(
        lambda: open_tdoc_local_folder(meeting_dir, tdoc_id)
    )

    act_portal = open_sub.addAction("🌐 View TDoc on 3GU Portal")
    act_portal.setToolTip('Open this TDoc in the official 3GPP 3GU portal.')
    act_portal.triggered.connect(lambda: open_3gu_portal(tdoc_id))

    # --- 3. COMPARE WITH ACTIVE WORD & COMPARISON CART SUBMENU ---
    if compare_active_word_callback:
        if revisions:
            latest_rev = revisions[-1]
            act_cmp_lat = menu.addAction(f"⚖️ Compare Latest ({latest_rev}) vs Active Word")
            act_cmp_lat.setToolTip(
                "Immediately compare the latest revision of this TDoc against the document currently active in Microsoft Word."
            )
            act_cmp_lat.triggered.connect(
                lambda: compare_active_word_callback(tdoc_id, latest_rev)
            )

            cmp_rev_sub = menu.addMenu(f"⚖️ Compare {tdoc_id} (Base / Revs) vs Active Word")
            cmp_rev_sub.setToolTipsVisible(True)
            act_cmp_base = cmp_rev_sub.addAction(f"⚖️ Compare Base: {tdoc_id} vs Active Word")
            act_cmp_base.setToolTip("Compare baseline TDoc against active Word document.")
            act_cmp_base.triggered.connect(
                lambda: compare_active_word_callback(tdoc_id, tdoc_id)
            )
            cmp_rev_sub.addSeparator()
            for rev in revisions:
                act_cmp_r = cmp_rev_sub.addAction(f"⚖️ Compare Rev: {rev} vs Active Word")
                act_cmp_r.triggered.connect(
                    lambda _, r=rev: compare_active_word_callback(tdoc_id, r)
                )
        else:
            act_cmp_active = menu.addAction(f"⚖️ Compare {tdoc_id} vs Active Word")
            act_cmp_active.setToolTip(
                "Immediately compare this TDoc against the document currently active in Microsoft Word."
            )
            act_cmp_active.triggered.connect(
                lambda: compare_active_word_callback(tdoc_id, tdoc_id)
            )

    cart_sub = menu.addMenu("⚖️ Comparison Cart")
    cart_sub.setToolTipsVisible(True)
    cart_sub.setToolTip("Add the base TDoc, a revision, or an open Word document to the comparison cart.")
    act_cart_base = cart_sub.addAction(f"➕ Add Base: {tdoc_id}")
    act_cart_base.setToolTip('Download/cache the baseline as needed and add it to the comparison cart.')
    act_cart_base.triggered.connect(
        lambda: download_callback(tdoc_id, tdoc_id, is_silent_compare=True)
    )

    if revisions:
        for rev in revisions:
            act_cart_rev = cart_sub.addAction(f"➕ Add Rev: {rev}")
            act_cart_rev.setToolTip('Download/cache this revision as needed and add it to the comparison cart.')
            act_cart_rev.triggered.connect(
                lambda _, r=rev: download_callback(
                    tdoc_id, r, is_silent_compare=True
                )
            )

    cart_sub.addSeparator()
    from modules.meetings.ui.tdocs_components import populate_slot_word_menu

    word_slot_a_sub = cart_sub.addMenu("🪟 Stage Open Word Doc in Slot A (Original)")
    word_slot_a_sub.setToolTipsVisible(True)
    populate_slot_word_menu(word_slot_a_sub, "A", parent)

    word_slot_b_sub = cart_sub.addMenu("🪟 Stage Open Word Doc in Slot B (Revised)")
    word_slot_b_sub.setToolTipsVisible(True)
    populate_slot_word_menu(word_slot_b_sub, "B", parent)

    # --- 4. MARKDOWN & DOSSIER EXPORT ---
    if export_dossier_callback or export_llm_callback:
        if export_dossier_callback:
            act_dossier = menu.addAction("📑 Export to Navigable Dossier (HTML && PDF)...")
            act_dossier.setToolTip("Compile this TDoc into an interactive HTML dossier and bookmarked vector PDF with companion AI citation prompt.")
            act_dossier.triggered.connect(lambda: export_dossier_callback([row_data]))
        if export_llm_callback:
            act_md = menu.addAction("📄 Export to Markdown (.md)")
            act_md.setToolTip("Export this TDoc as an individual Markdown document with metadata and extracted figures.")
            act_md.triggered.connect(lambda: export_llm_callback(tdoc_id))
        menu.addSeparator()

    # --- 5. EMAILS SUBMENU ---
    email_sub = menu.addMenu("📧 Emails")
    email_sub.setToolTipsVisible(True)
    email_sub.setToolTip("Inspect, draft, or update read state for emails associated with this TDoc family.")
    email_label = (
        f"👁️ View Related Emails ({unread_emails_count} unread)..."
        if unread_emails_count > 0
        else f"👁️ View Related Emails for {tdoc_id}..."
    )
    act_view_emails = email_sub.addAction(email_label)
    act_view_emails.setToolTip('Open locally indexed emails associated with this TDoc family.')
    act_view_emails.triggered.connect(lambda: open_emails_callback(tdoc_id))

    act_draft_email = email_sub.addAction("✉️ Draft Email with Subject...")
    act_draft_email.setToolTip('Create a new email draft with meeting, AI and TDoc information in the subject.')
    act_draft_email.triggered.connect(lambda: compose_email_callback(tdoc_id))

    email_sub.addSeparator()
    act_mark_read = email_sub.addAction(
        f"✔️ Mark all emails as read for {tdoc_id} family"
    )
    act_mark_read.setToolTip('Mark indexed emails for this TDoc family as read.')
    act_mark_read.triggered.connect(mark_read_callback)

    act_mark_unread = email_sub.addAction(
        f"✉️ Mark all emails as unread for {tdoc_id} family"
    )
    act_mark_unread.setToolTip('Mark indexed emails for this TDoc family as unread.')
    act_mark_unread.triggered.connect(mark_unread_callback)
    menu.addSeparator()

    # --- 6. QUICK COPY SUBMENU ---
    copy_sub = menu.addMenu("📋 Quick Copy")
    copy_sub.setToolTipsVisible(True)
    copy_sub.setToolTip("Copy commonly used TDoc metadata in plain text and, where useful, rich HTML form.")
    populate_quick_copy_menu(
        menu=copy_sub,
        tdoc_id=tdoc_id,
        row_data=row_data,
        docs_ftp_url=docs_ftp_url,
        revisions_url=revisions_url,
        parent=parent,
    )

    menu.exec_(pos or QCursor.pos())


# =========================================================================
# 3. RELATED LINKS CONTEXT MENU (SECRETARY REMARKS & RELATED TDOCS)
# =========================================================================
def build_related_menu(
    parent: QWidget,
    target_tdoc: str,
    valid_tdocs: set,
    docs_ftp_url: str,
    revisions_url: str,
    scroll_callback,
    download_callback,
    export_llm_callback,
    global_search_callback,
    compose_email_callback,
    pos: QPoint,
    triage_callback=None,  # <--- Added parameter
    compare_active_word_callback=None,
):
    """Context menu triggered when right-clicking hyperlinked TDoc tags."""
    menu = QMenu(parent)
    menu.setToolTipsVisible(True)
    clean_tdoc = target_tdoc.strip().upper()

    match = re.search(
        r"^(.*?)-?(?:r|rev)\d{1,2}[a-zA-Z]?$", clean_tdoc, re.IGNORECASE
    )
    base_tdoc = match.group(1).upper() if match else clean_tdoc
    is_internal = base_tdoc in valid_tdocs

    if is_internal:
        act_jump = menu.addAction(f"🎯 Jump to {clean_tdoc} in Table")
        act_jump.setToolTip('Scroll the table to this related TDoc.')
        act_jump.triggered.connect(lambda: scroll_callback(clean_tdoc))

        act_open = menu.addAction(f"📄 Open {clean_tdoc}")
        act_open.setToolTip('Open/download this related TDoc.')
        act_open.triggered.connect(
            lambda: download_callback(
                base_tdoc, clean_tdoc, is_silent_compare=False
            )
        )

        if compare_active_word_callback:
            act_cmp_active = menu.addAction(f"⚖️ Compare {clean_tdoc} vs Active Word")
            act_cmp_active.setToolTip("Compare this related TDoc against your currently active Microsoft Word document.")
            act_cmp_active.triggered.connect(
                lambda: compare_active_word_callback(base_tdoc, clean_tdoc)
            )

        act_cart = menu.addAction(f"⚖️ Add {clean_tdoc} to Comparison Cart")
        act_cart.setToolTip('Add this related TDoc to the comparison cart.')
        act_cart.triggered.connect(
            lambda: download_callback(
                base_tdoc, clean_tdoc, is_silent_compare=True
            )
        )

        menu.addSeparator()

        if triage_callback:
            act_triage = menu.addAction(f"🤖 AI Technical Triage ({clean_tdoc})")
            act_triage.setToolTip('Run the local-LLM technical triage for this TDoc.')
            act_triage.triggered.connect(lambda: triage_callback(clean_tdoc))

        if compose_email_callback:
            act_email = menu.addAction(f"✉️ Draft Email ({clean_tdoc})")
            act_email.setToolTip('Draft an email referring to this TDoc.')
            act_email.triggered.connect(
                lambda: compose_email_callback(clean_tdoc)
            )

        if export_llm_callback:
            act_md = menu.addAction("📄 Export to Markdown (.md)")
            act_md.setToolTip("Export this TDoc as an individual Markdown document with metadata and extracted figures.")
            act_md.triggered.connect(lambda: export_llm_callback(clean_tdoc))
    else:
        act_global = menu.addAction(
            f"🌐 Search {clean_tdoc} Across All Meetings"
        )
        act_global.setToolTip('Find this external TDoc in the global meeting database.')
        act_global.triggered.connect(
            lambda: global_search_callback(clean_tdoc, "open_meeting")
        )

        act_open_ext = menu.addAction(f"📄 Try Downloading {clean_tdoc}")
        act_open_ext.setToolTip('Attempt to download this TDoc even though it is not in the current meeting list.')
        act_open_ext.triggered.connect(
            lambda: download_callback(
                base_tdoc, clean_tdoc, is_silent_compare=False
            )
        )

        if compare_active_word_callback:
            act_cmp_ext = menu.addAction(f"⚖️ Compare {clean_tdoc} vs Active Word")
            act_cmp_ext.setToolTip("Compare this TDoc against your currently active Microsoft Word document.")
            act_cmp_ext.triggered.connect(
                lambda: compare_active_word_callback(base_tdoc, clean_tdoc)
            )

    menu.addSeparator()
    act_copy = menu.addAction(f"📋 Copy TDoc ID ({clean_tdoc})")
    act_copy.setToolTip('Copy this TDoc identifier to the clipboard.')
    act_copy.triggered.connect(
        lambda: [
            QApplication.clipboard().setText(clean_tdoc),
            QToolTip.showText(QCursor.pos(), "📋 Copied!", parent),
        ]
    )

    menu.exec_(pos or QCursor.pos())
