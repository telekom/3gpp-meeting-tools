"""Streamlined Column 0 action menu and related-TDoc hyperlink context menu."""

from __future__ import annotations

import re
from pathlib import Path

from PyQt5.QtCore import QPoint
from PyQt5.QtGui import QCursor
from PyQt5.QtWidgets import QApplication, QMenu, QToolTip, QWidget

from core.ai.cli_launcher import populate_cli_agent_menu
from modules.meetings.ui.tdocs.menus.quick_copy import (
    is_tdoc_cached,
    open_3gu_portal,
    open_tdoc_local_folder,
)


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
        act_latest.setToolTip("Open the newest known revision of this TDoc.")
        act_latest.triggered.connect(
            lambda: download_callback(base_tdoc, latest_rev, is_silent_compare=False)
        )
        menu.addSeparator()

        act_base = menu.addAction(
            f"📄 Open Base: {base_tdoc}" + (" (Local)" if base_cached else "")
        )
        act_base.setToolTip("Open the baseline TDoc.")
        act_base.triggered.connect(
            lambda: download_callback(base_tdoc, base_tdoc, is_silent_compare=False)
        )

        for rev in revisions:
            rev_cached = is_tdoc_cached(meeting_dir, rev)
            act_rev = menu.addAction(
                f"📄 Open Rev: {rev}" + (" (Local)" if rev_cached else "")
            )
            act_rev.triggered.connect(
                lambda _, r=rev: download_callback(base_tdoc, r, is_silent_compare=False)
            )
    else:
        act_base = menu.addAction(
            f"📄 Open Base: {base_tdoc}" + (" (Local)" if base_cached else "")
        )
        act_base.setToolTip("Open the baseline TDoc.")
        act_base.triggered.connect(
            lambda: download_callback(base_tdoc, base_tdoc, is_silent_compare=False)
        )

    if compare_active_word_callback:
        menu.addSeparator()
        if revisions:
            latest_rev = revisions[-1]
            act_cmp_lat = menu.addAction(f"⚖️ Compare Latest ({latest_rev}) vs Active Word")
            act_cmp_lat.setToolTip(
                "Compare the newest revision of this TDoc against the document currently active in Microsoft Word."
            )
            act_cmp_lat.triggered.connect(lambda: compare_active_word_callback(base_tdoc, latest_rev))
            act_cmp_base = menu.addAction(f"⚖️ Compare Base ({base_tdoc}) vs Active Word")
            act_cmp_base.setToolTip(
                "Compare the baseline TDoc against the document currently active in Microsoft Word."
            )
            act_cmp_base.triggered.connect(lambda: compare_active_word_callback(base_tdoc, base_tdoc))
        else:
            act_cmp_base = menu.addAction(f"⚖️ Compare {base_tdoc} vs Active Word")
            act_cmp_base.setToolTip(
                "Compare this TDoc against the document currently active in Microsoft Word."
            )
            act_cmp_base.triggered.connect(lambda: compare_active_word_callback(base_tdoc, base_tdoc))

    menu.addSeparator()
    act_folder = menu.addAction("📂 Open Local Folder")
    act_folder.setToolTip("Open the local folder used to cache this TDoc and its extracted files.")
    act_folder.triggered.connect(lambda: open_tdoc_local_folder(meeting_dir, base_tdoc))

    act_portal = menu.addAction("🌐 View TDoc on 3GU Portal")
    act_portal.setToolTip("Open this TDoc in the official 3GPP 3GU portal.")
    act_portal.triggered.connect(lambda: open_3gu_portal(base_tdoc))

    menu.exec_(pos or QCursor.pos())


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
    triage_callback=None,
    compare_active_word_callback=None,
    launch_cli_callback=None,
    send_tdoc_email_callback=None,
):
    """Context menu triggered when right-clicking hyperlinked TDoc tags."""
    menu = QMenu(parent)
    menu.setToolTipsVisible(True)
    clean_tdoc = target_tdoc.strip().upper()

    match = re.search(r"^(.*?)-?(?:r|rev)\d{1,2}[a-zA-Z]?$", clean_tdoc, re.IGNORECASE)
    base_tdoc = match.group(1).upper() if match else clean_tdoc
    is_internal = base_tdoc in valid_tdocs

    if is_internal:
        act_jump = menu.addAction(f"🎯 Jump to {clean_tdoc} in Table")
        act_jump.setToolTip("Scroll the table to this related TDoc.")
        act_jump.triggered.connect(lambda: scroll_callback(clean_tdoc))

        act_open = menu.addAction(f"📄 Open {clean_tdoc}")
        act_open.setToolTip("Open/download this related TDoc.")
        act_open.triggered.connect(
            lambda: download_callback(base_tdoc, clean_tdoc, is_silent_compare=False)
        )

        if compare_active_word_callback:
            act_cmp_active = menu.addAction(f"⚖️ Compare {clean_tdoc} vs Active Word")
            act_cmp_active.setToolTip("Compare this related TDoc against your currently active Microsoft Word document.")
            act_cmp_active.triggered.connect(lambda: compare_active_word_callback(base_tdoc, clean_tdoc))

        act_cart = menu.addAction(f"⚖️ Add {clean_tdoc} to Comparison Cart")
        act_cart.setToolTip("Add this related TDoc to the comparison cart.")
        act_cart.triggered.connect(
            lambda: download_callback(base_tdoc, clean_tdoc, is_silent_compare=True)
        )

        menu.addSeparator()

        if triage_callback:
            act_triage = menu.addAction(f"🤖 AI Technical Triage ({clean_tdoc})")
            act_triage.setToolTip("Run the local-LLM technical triage for this TDoc.")
            act_triage.triggered.connect(lambda: triage_callback(clean_tdoc))

        if launch_cli_callback:
            cli_sub = menu.addMenu(f"💬 Ask AI Agent (CLI) ({clean_tdoc})")
            populate_cli_agent_menu(
                cli_sub,
                on_launch_callback=lambda aid, mode, t=clean_tdoc: launch_cli_callback([{"TDoc": t}], aid, mode),
                has_multiple_or_revs=False,
                parent=parent,
            )

        if compose_email_callback:
            act_email = menu.addAction(f"✉️ Draft Email ({clean_tdoc})")
            act_email.setToolTip("Draft an email referring to this TDoc.")
            act_email.triggered.connect(lambda: compose_email_callback(clean_tdoc))

        if send_tdoc_email_callback:
            act_send_rel = menu.addAction(f"📎 Send {clean_tdoc} via Email...")
            act_send_rel.setToolTip("Create an Outlook email draft with this TDoc's subject and attach the document.")
            act_send_rel.triggered.connect(lambda: send_tdoc_email_callback(clean_tdoc))

        if export_llm_callback:
            act_md = menu.addAction("📄 Export to Markdown (.md)")
            act_md.setToolTip("Export this TDoc as an individual Markdown document with metadata and extracted figures.")
            act_md.triggered.connect(lambda: export_llm_callback(clean_tdoc))
    else:
        act_global = menu.addAction(f"🌐 Search {clean_tdoc} Across All Meetings")
        act_global.setToolTip("Find this external TDoc in the global meeting database.")
        act_global.triggered.connect(lambda: global_search_callback(clean_tdoc, "open_meeting"))

        act_open_ext = menu.addAction(f"📄 Try Downloading {clean_tdoc}")
        act_open_ext.setToolTip("Attempt to download this TDoc even though it is not in the current meeting list.")
        act_open_ext.triggered.connect(
            lambda: download_callback(base_tdoc, clean_tdoc, is_silent_compare=False)
        )

        if compare_active_word_callback:
            act_cmp_ext = menu.addAction(f"⚖️ Compare {clean_tdoc} vs Active Word")
            act_cmp_ext.setToolTip("Compare this TDoc against your currently active Microsoft Word document.")
            act_cmp_ext.triggered.connect(lambda: compare_active_word_callback(base_tdoc, clean_tdoc))

    menu.addSeparator()
    act_copy = menu.addAction(f"📋 Copy TDoc ID ({clean_tdoc})")
    act_copy.setToolTip("Copy this TDoc identifier to the clipboard.")
    act_copy.triggered.connect(
        lambda: [
            QApplication.clipboard().setText(clean_tdoc),
            QToolTip.showText(QCursor.pos(), "📋 Copied!", parent),
        ]
    )

    menu.exec_(pos or QCursor.pos())
