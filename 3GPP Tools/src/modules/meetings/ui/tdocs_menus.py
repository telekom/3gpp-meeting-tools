# --- File: src/modules/meetings/ui/tdocs_menus.py ---
from PyQt5.QtWidgets import QMenu, QApplication, QToolTip
from PyQt5.QtCore import Qt
from pathlib import Path
import re


class TDocActionMenu(QMenu):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.RightButton:
            action = self.actionAt(event.pos())
            if action and action.data():
                url = action.data()
                QApplication.clipboard().setText(url)
                QToolTip.showText(event.globalPos(), "📋 URL Copied to Clipboard!", self)
                self.close()
                return
        super().mouseReleaseEvent(event)


def build_action_menu(parent_widget, base_tdoc, docs_ftp_url, revisions_url, revisions_list, meeting_dir,
                      download_callback, llm_export_callback, global_pos):
    menu = TDocActionMenu(parent_widget)
    menu.setStyleSheet("QMenu { font-size: 13px; }")
    menu.setToolTipsVisible(True)

    docs_url = docs_ftp_url if docs_ftp_url.startswith("http") else "https://www.3gpp.org/ftp/" + docs_ftp_url.lstrip(
        '/')
    base_zip = meeting_dir / base_tdoc / f"{base_tdoc}.zip"

    act_base = menu.addAction(f"🗎 Open Base: {base_tdoc}" + ("  (Local)" if base_zip.exists() else ""))
    act_base.setData(docs_url.rstrip('/') + f"/{base_tdoc}.zip")
    act_base.setToolTip("Left-click to open. Right-click to copy FTP link.")
    act_base.triggered.connect(lambda _, t=base_tdoc: download_callback(base_tdoc, t, docs_url, False))

    if revisions_list:
        menu.addSeparator()
        for rev in revisions_list:
            target_filename = f"{base_tdoc}{rev}"
            rev_zip = meeting_dir / base_tdoc / f"{target_filename}.zip"
            act_rev = menu.addAction(f"📝 Open Revision: {target_filename}" + ("  (Local)" if rev_zip.exists() else ""))
            act_rev.setData(revisions_url.rstrip('/') + f"/{target_filename}.zip")
            act_rev.setToolTip("Left-click to open. Right-click to copy FTP link.")
            act_rev.triggered.connect(
                lambda _, t=target_filename: download_callback(base_tdoc, t, revisions_url, False))

    menu.addSeparator()
    act_folder = menu.addAction("📂 Open Local Folder")
    act_folder.triggered.connect(lambda _, d=(meeting_dir / base_tdoc): __open_folder(d))

    menu.addSeparator()
    compare_menu = TDocActionMenu("⚖️ Add to Comparison Cart...", parent_widget)
    compare_menu.setToolTipsVisible(True)
    menu.addMenu(compare_menu)

    act_cmp_base = compare_menu.addAction(f"🗎 Base Version: {base_tdoc}" + ("  (Local)" if base_zip.exists() else ""))
    act_cmp_base.setData(docs_url.rstrip('/') + f"/{base_tdoc}.zip")
    act_cmp_base.setToolTip("Right-click to copy FTP link.")
    act_cmp_base.triggered.connect(lambda _, t=base_tdoc: download_callback(base_tdoc, t, docs_url, True))

    for rev in revisions_list:
        target_filename = f"{base_tdoc}{rev}"
        rev_zip = meeting_dir / base_tdoc / f"{target_filename}.zip"
        act_cmp_rev = compare_menu.addAction(
            f"📝 Revision: {target_filename}" + ("  (Local)" if rev_zip.exists() else ""))
        act_cmp_rev.setData(revisions_url.rstrip('/') + f"/{target_filename}.zip")
        act_cmp_rev.setToolTip("Right-click to copy FTP link.")
        act_cmp_rev.triggered.connect(lambda _, t=target_filename: download_callback(base_tdoc, t, revisions_url, True))

    menu.addSeparator()
    act_llm = menu.addAction("🤖 Export for LLM Analysis")
    act_llm.setToolTip("Extract Track Changes and prepare Markdown for Gemini")
    act_llm.triggered.connect(lambda _, t=base_tdoc: llm_export_callback(t))

    menu.exec_(global_pos)


def build_related_menu(parent_widget, target_tdoc, valid_tdocs, docs_ftp_url, revisions_url, scroll_callback,
                       download_callback, llm_export_callback, global_req_callback, global_pos):
    menu = QMenu(parent_widget)
    menu.setStyleSheet("QMenu { font-size: 13px; }")

    match = re.search(r'^(.*?)-?(?:r|rev)\d{1,2}[a-zA-Z]?$', target_tdoc, re.IGNORECASE)
    base_tdoc = match.group(1).upper() if match else target_tdoc.upper()

    is_local = base_tdoc in valid_tdocs
    dl_url = revisions_url if (match and revisions_url) else docs_ftp_url

    if is_local:
        menu.addAction("⬇️ Go to Row").triggered.connect(lambda: scroll_callback(target_tdoc))
        menu.addAction(f"📄 Open Document: {target_tdoc}").triggered.connect(
            lambda: download_callback(base_tdoc, target_tdoc, dl_url, False))
        menu.addAction(f"⚖️ Add to Comparison Cart: {target_tdoc}").triggered.connect(
            lambda: download_callback(base_tdoc, target_tdoc, dl_url, True))

        menu.addSeparator()
        menu.addAction("🤖 Export for LLM Analysis").triggered.connect(
            lambda: llm_export_callback(base_tdoc))
    else:
        menu.addAction("🌐 Search && Open Meeting").triggered.connect(
            lambda: global_req_callback(base_tdoc, 'open_meeting'))
        menu.addAction(f"📄 Search && Open Document: {target_tdoc}").triggered.connect(
            lambda: global_req_callback(target_tdoc, 'open_doc'))
        menu.addAction(f"⚖️ Search && Add to Comparison Cart: {target_tdoc}").triggered.connect(
            lambda: global_req_callback(target_tdoc, 'add_to_cart'))

    menu.exec_(global_pos)


def __open_folder(folder_path: Path):
    import os, webbrowser
    if folder_path.exists():
        if hasattr(os, 'startfile'):
            os.startfile(str(folder_path))
        else:
            webbrowser.open(f"file:///{folder_path}")