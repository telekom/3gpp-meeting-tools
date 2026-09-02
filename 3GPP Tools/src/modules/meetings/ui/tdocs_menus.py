# --- File: src/modules/meetings/ui/tdocs_menus.py ---
import os
import re
import webbrowser
from pathlib import Path

from PyQt5.QtCore import QPoint, Qt
from PyQt5.QtGui import QCursor
from PyQt5.QtWidgets import QApplication, QMenu, QMessageBox, QToolTip, QWidget


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
    tdocs_dir = Path(meeting_dir) / "TDocs"
    target_sub = tdocs_dir / base_tdoc
    target = target_sub if target_sub.exists() else tdocs_dir
    if not target.exists():
        target.mkdir(parents=True, exist_ok=True)
    if hasattr(os, "startfile"):
        os.startfile(str(target))
    else:
        webbrowser.open(f"file:///{target}")


def open_3gu_portal(base_tdoc: str):
    """Opens the official 3GU portal view for the target TDoc."""
    url = f"https://portal.3gpp.org/ngppapp/CreateTdoc.aspx?mode=view&tdocId={base_tdoc}"
    webbrowser.open(url)


def _copy_text(text: str, label: str, parent: QWidget):
    """Copies text to clipboard and flashes a non-blocking tooltip."""
    QApplication.clipboard().setText(text)
    QToolTip.showText(QCursor.pos(), f"📋 Copied {label} to clipboard!", parent)


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
):
    """Fast, dedicated document opener for the Column 0 button."""
    menu = QMenu(parent)
    revisions = revisions or []
    base_cached = is_tdoc_cached(meeting_dir, base_tdoc)

    if revisions:
        latest_rev = revisions[-1]
        latest_cached = is_tdoc_cached(meeting_dir, latest_rev)
        act_latest = menu.addAction(
            f"🚀 Open Latest: {latest_rev}" + (" (Local)" if latest_cached else "")
        )
        act_latest.triggered.connect(
            lambda: download_callback(
                base_tdoc, latest_rev, is_silent_compare=False
            )
        )
        menu.addSeparator()

        act_base = menu.addAction(
            f"📄 Open Base: {base_tdoc}" + (" (Local)" if base_cached else "")
        )
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
        act_base.triggered.connect(
            lambda: download_callback(
                base_tdoc, base_tdoc, is_silent_compare=False
            )
        )

    menu.addSeparator()
    act_folder = menu.addAction("📂 Open Local Folder")
    act_folder.triggered.connect(
        lambda: open_tdoc_local_folder(meeting_dir, base_tdoc)
    )

    act_portal = menu.addAction("🌐 View TDoc on 3GU Portal")
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
):
    """Complete command center menu for row-level operations."""
    menu = QMenu(parent)
    revisions = revisions or []
    tdoc_id = str(row_data.get("TDoc", "")).strip()

    # --- 1. INSPECT FULL DETAILS ---
    act_details = menu.addAction(f"ℹ️ View Full Details for {tdoc_id}...")
    act_details.triggered.connect(open_details_callback)
    menu.addSeparator()

    # --- 2. OPEN DOCUMENT SUBMENU ---
    open_sub = menu.addMenu("📂 Open Document")
    base_cached = is_tdoc_cached(meeting_dir, tdoc_id)

    if revisions:
        latest_rev = revisions[-1]
        latest_cached = is_tdoc_cached(meeting_dir, latest_rev)
        act_open_lat = open_sub.addAction(
            f"🚀 Open Latest: {latest_rev}"
            + (" (Local)" if latest_cached else "")
        )
        act_open_lat.triggered.connect(
            lambda: download_callback(
                tdoc_id, latest_rev, is_silent_compare=False
            )
        )
        open_sub.addSeparator()

        act_open_base = open_sub.addAction(
            f"📄 Open Base: {tdoc_id}" + (" (Local)" if base_cached else "")
        )
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
        act_open_base.triggered.connect(
            lambda: download_callback(tdoc_id, tdoc_id, is_silent_compare=False)
        )

    open_sub.addSeparator()
    act_folder = open_sub.addAction("📂 Open Local Folder")
    act_folder.triggered.connect(
        lambda: open_tdoc_local_folder(meeting_dir, tdoc_id)
    )

    act_portal = open_sub.addAction("🌐 View TDoc on 3GU Portal")
    act_portal.triggered.connect(lambda: open_3gu_portal(tdoc_id))

    # --- 3. COMPARISON CART SUBMENU ---
    cart_sub = menu.addMenu("⚖️ Comparison Cart")
    act_cart_base = cart_sub.addAction(f"➕ Add Base: {tdoc_id}")
    act_cart_base.triggered.connect(
        lambda: download_callback(tdoc_id, tdoc_id, is_silent_compare=True)
    )

    if revisions:
        for rev in revisions:
            act_cart_rev = cart_sub.addAction(f"➕ Add Rev: {rev}")
            act_cart_rev.triggered.connect(
                lambda _, r=rev: download_callback(
                    tdoc_id, r, is_silent_compare=True
                )
            )

    # --- 4. LLM CORPUS EXPORT ---
    act_llm = menu.addAction("🤖 Export for LLM Analysis")
    act_llm.triggered.connect(lambda: export_llm_callback(tdoc_id))
    menu.addSeparator()

    # --- 5. EMAILS SUBMENU ---
    email_sub = menu.addMenu("📧 Emails")
    email_label = (
        f"👁️ View Related Emails ({unread_emails_count} unread)..."
        if unread_emails_count > 0
        else f"👁️ View Related Emails for {tdoc_id}..."
    )
    act_view_emails = email_sub.addAction(email_label)
    act_view_emails.triggered.connect(lambda: open_emails_callback(tdoc_id))

    act_draft_email = email_sub.addAction("✉️ Draft Email with Subject...")
    act_draft_email.triggered.connect(lambda: compose_email_callback(tdoc_id))

    email_sub.addSeparator()
    act_mark_read = email_sub.addAction(
        f"✔️ Mark all emails as read for {tdoc_id} family"
    )
    act_mark_read.triggered.connect(mark_read_callback)

    act_mark_unread = email_sub.addAction(
        f"✉️ Mark all emails as unread for {tdoc_id} family"
    )
    act_mark_unread.triggered.connect(mark_unread_callback)
    menu.addSeparator()

    # --- 6. QUICK COPY SUBMENU ---
    copy_sub = menu.addMenu("📋 Quick Copy")
    act_copy_id = copy_sub.addAction(f"📋 Copy TDoc Number ({tdoc_id})")
    act_copy_id.triggered.connect(
        lambda: _copy_text(tdoc_id, "TDoc Number", parent)
    )

    title = str(row_data.get("Title", "")).strip()
    if title:
        act_copy_title = copy_sub.addAction("📋 Copy Title")
        act_copy_title.triggered.connect(
            lambda: _copy_text(title, "Title", parent)
        )

    source = str(row_data.get("Source", "")).strip()
    citation = (
        f"{tdoc_id}: {title} ({source})" if source else f"{tdoc_id}: {title}"
    )
    act_copy_cit = copy_sub.addAction("📋 Copy Full Citation")
    act_copy_cit.triggered.connect(
        lambda: _copy_text(citation, "Citation", parent)
    )

    if docs_ftp_url and tdoc_id and tdoc_id.upper() != "UNKNOWN":
        clean_docs_url = docs_ftp_url.rstrip('/')
        if clean_docs_url.startswith("ftp://"):
            clean_docs_url = "https://" + clean_docs_url[6:]
        elif not clean_docs_url.startswith("http"):
            clean_docs_url = "https://www.3gpp.org/ftp/" + clean_docs_url.lstrip('/')

        clean_docs_url = clean_docs_url.replace("https://ftp.3gpp.org/", "https://www.3gpp.org/ftp/")

        if not clean_docs_url.endswith("/Docs") and not clean_docs_url.endswith("/docs"):
            if "/Docs" not in clean_docs_url and "/docs" not in clean_docs_url:
                clean_docs_url = f"{clean_docs_url}/Docs"

        base_match = re.match(r"^([A-Za-z0-9]+-\d+)", tdoc_id)
        base_tdoc = base_match.group(1).upper() if base_match else tdoc_id.upper()
        baseline_url = f"{clean_docs_url}/{base_tdoc}.zip"

        act_copy_url = copy_sub.addAction("📋 Copy Baseline Document URL")
        act_copy_url.triggered.connect(
            lambda: _copy_text(baseline_url, "Baseline Document URL", parent)
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
):
    """Context menu triggered when right-clicking hyperlinked TDoc tags."""
    menu = QMenu(parent)
    clean_tdoc = target_tdoc.strip().upper()

    match = re.search(
        r"^(.*?)-?(?:r|rev)\d{1,2}[a-zA-Z]?$", clean_tdoc, re.IGNORECASE
    )
    base_tdoc = match.group(1).upper() if match else clean_tdoc
    is_internal = base_tdoc in valid_tdocs

    if is_internal:
        act_jump = menu.addAction(f"🎯 Jump to {clean_tdoc} in Table")
        act_jump.triggered.connect(lambda: scroll_callback(clean_tdoc))

        act_open = menu.addAction(f"📄 Open {clean_tdoc}")
        act_open.triggered.connect(
            lambda: download_callback(
                base_tdoc, clean_tdoc, is_silent_compare=False
            )
        )

        act_cart = menu.addAction(f"⚖️ Add {clean_tdoc} to Comparison Cart")
        act_cart.triggered.connect(
            lambda: download_callback(
                base_tdoc, clean_tdoc, is_silent_compare=True
            )
        )

        menu.addSeparator()
        if compose_email_callback:
            act_email = menu.addAction(f"✉️ Draft Email ({clean_tdoc})")
            act_email.triggered.connect(
                lambda: compose_email_callback(clean_tdoc)
            )

        if export_llm_callback:
            act_llm = menu.addAction("🤖 Export for LLM Analysis")
            act_llm.triggered.connect(lambda: export_llm_callback(clean_tdoc))
    else:
        act_global = menu.addAction(
            f"🌐 Search {clean_tdoc} Across All Meetings"
        )
        act_global.triggered.connect(
            lambda: global_search_callback(clean_tdoc, "open_meeting")
        )

        act_open_ext = menu.addAction(f"📄 Try Downloading {clean_tdoc}")
        act_open_ext.triggered.connect(
            lambda: download_callback(
                base_tdoc, clean_tdoc, is_silent_compare=False
            )
        )

    menu.addSeparator()
    act_copy = menu.addAction(f"📋 Copy TDoc ID ({clean_tdoc})")
    act_copy.triggered.connect(
        lambda: [
            QApplication.clipboard().setText(clean_tdoc),
            QToolTip.showText(QCursor.pos(), "📋 Copied!", parent),
        ]
    )

    menu.exec_(pos or QCursor.pos())