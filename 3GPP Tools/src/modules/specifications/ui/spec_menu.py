# --- File: src/modules/specifications/ui/spec_menu.py ---
"""
Context menu builder and clipboard copy utilities for the Specifications table.
Provides multi-select actions for toggling favorites, editing annotations,
granular clipboard copying (spec number, title, WG, links, markdown),
metadata synchronization, and deletion.
"""
import html
from pathlib import Path
import re

from PyQt5.QtCore import QMimeData, Qt
from PyQt5.QtGui import QCursor
from PyQt5.QtWidgets import QApplication, QLabel, QMenu, QToolTip, QWidget

from modules.specifications.utils.utils import build_dynareport_url


def _copy_to_clipboard(plain_text: str, html_text: str = None, label: str = "text", parent: QWidget = None):
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


def resolve_spec_metadata(spec_num: str, db, details: dict = None) -> dict:
    """Resolves specification metadata from DB or existing dictionary."""
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

    return {
        "spec_num": spec_num,
        "type": s_type,
        "canonical_num": canonical_num,
        "title": title,
        "wg": wg,
        "dynareport_url": dynareport_url,
        "ftp_url": ftp_url,
        "details": spec_data,
    }


def populate_spec_copy_menu(
    copy_sub: QMenu,
    parent_widget: QWidget,
    spec_num: str,
    details: dict = None,
    db=None,
):
    """
    Populates the Quick Copy submenu with standard copy actions for a single specification.
    Includes spec number (plain + rich link), number & title, full citation with WG,
    number & WG, title only, markdown link, and direct web / FTP URLs.
    """
    meta = resolve_spec_metadata(spec_num, db, details)
    canonical = meta["canonical_num"]
    title = meta["title"]
    wg = meta["wg"]
    dynareport_url = meta["dynareport_url"]
    ftp_url = meta["ftp_url"]

    safe_canonical = html.escape(canonical)
    safe_title = html.escape(title)
    safe_wg = html.escape(wg)

    # HTML anchor hyperlink
    if dynareport_url:
        spec_html_link = f'<a href="{dynareport_url}">{safe_canonical}</a>'
    else:
        spec_html_link = safe_canonical

    # 1. Copy Spec Number (TS XX.XXX)
    act_num = copy_sub.addAction(f"📋 Copy Spec Number ({canonical})")
    act_num.setToolTip("Copy the specification number (embeds clickable 3GPP link for rich-text apps).")
    act_num.triggered.connect(
        lambda _, p=canonical, h=spec_html_link: _copy_to_clipboard(
            plain_text=p,
            html_text=h,
            label="Spec Number",
            parent=parent_widget,
        )
    )

    # 2. Copy Number & Title
    if title:
        plain_num_title = f"{canonical}: {title}"
        html_num_title = f"{spec_html_link}: {safe_title}"
        act_num_title = copy_sub.addAction("📋 Copy Number & Title")
        act_num_title.setToolTip(
            "Copy specification number and title (embeds clickable link for rich-text apps)."
        )
        act_num_title.triggered.connect(
            lambda _, p=plain_num_title, h=html_num_title: _copy_to_clipboard(
                plain_text=p,
                html_text=h,
                label="Number & Title",
                parent=parent_widget,
            )
        )

    # 3. Copy Full Citation (with WG)
    if wg:
        citation_plain = f"{canonical}: {title} ({wg})" if title else f"{canonical} ({wg})"
        citation_html = (
            f"{spec_html_link}: {safe_title} ({safe_wg})"
            if title
            else f"{spec_html_link} ({safe_wg})"
        )
        act_cit = copy_sub.addAction(f"📋 Copy Full Citation (with {wg})")
        act_cit.setToolTip(
            "Copy complete citation including specification number, title, and Working Group."
        )
        act_cit.triggered.connect(
            lambda _, p=citation_plain, h=citation_html, w=wg: _copy_to_clipboard(
                plain_text=p,
                html_text=h,
                label=f"Full Citation ({w})",
                parent=parent_widget,
            )
        )

        # 4. Copy Number & WG
        num_wg_plain = f"{canonical} ({wg})"
        num_wg_html = f"{spec_html_link} ({safe_wg})"
        act_num_wg = copy_sub.addAction(f"📋 Copy Number & WG ({canonical} ({wg}))")
        act_num_wg.setToolTip("Copy specification number and responsible Working Group.")
        act_num_wg.triggered.connect(
            lambda _, p=num_wg_plain, h=num_wg_html: _copy_to_clipboard(
                plain_text=p,
                html_text=h,
                label="Number & WG",
                parent=parent_widget,
            )
        )
    elif title:
        citation_plain = f"{canonical}: {title}"
        citation_html = f"{spec_html_link}: {safe_title}"
        act_cit = copy_sub.addAction("📋 Copy Full Citation")
        act_cit.setToolTip("Copy complete specification citation.")
        act_cit.triggered.connect(
            lambda _, p=citation_plain, h=citation_html: _copy_to_clipboard(
                plain_text=p,
                html_text=h,
                label="Full Citation",
                parent=parent_widget,
            )
        )

    # 5. Copy Title
    if title:
        act_title = copy_sub.addAction("📋 Copy Title")
        act_title.setToolTip("Copy the specification title.")
        act_title.triggered.connect(
            lambda _, p=title, h=safe_title: _copy_to_clipboard(
                plain_text=p,
                html_text=h,
                label="Title",
                parent=parent_widget,
            )
        )

    # Separator before external links
    copy_sub.addSeparator()

    # 6. Copy as Markdown Link
    if dynareport_url:
        md_plain = f"[{canonical}]({dynareport_url})"
        act_md = copy_sub.addAction(f"📋 Copy as Markdown Link ([{canonical}](url))")
        act_md.setToolTip("Copy as Markdown hyperlink syntax.")
        act_md.triggered.connect(
            lambda _, p=md_plain, h=spec_html_link: _copy_to_clipboard(
                plain_text=p,
                html_text=h,
                label="Markdown Link",
                parent=parent_widget,
            )
        )

        if title:
            md_with_title = f"[{canonical}]({dynareport_url}) - {title}"
            md_with_title_html = f"{spec_html_link} - {safe_title}"
            act_md_title = copy_sub.addAction("📋 Copy as Markdown with Title")
            act_md_title.setToolTip("Copy as Markdown link followed by specification title.")
            act_md_title.triggered.connect(
                lambda _, p=md_with_title, h=md_with_title_html: _copy_to_clipboard(
                    plain_text=p,
                    html_text=h,
                    label="Markdown with Title",
                    parent=parent_widget,
                )
            )

    # 7. Copy 3GPP Web Report URL (DynaReport)
    if dynareport_url:
        act_dyna = copy_sub.addAction("📋 Copy 3GPP Web Report URL")
        act_dyna.setToolTip("Copy direct URL to the 3GPP DynaReport specification status page.")
        act_dyna.triggered.connect(
            lambda _, p=dynareport_url, h=f'<a href="{dynareport_url}">{dynareport_url}</a>': _copy_to_clipboard(
                plain_text=p,
                html_text=h,
                label="Web Report URL",
                parent=parent_widget,
            )
        )

    # 8. Copy FTP Archive URL
    if ftp_url:
        act_ftp = copy_sub.addAction("📋 Copy FTP Archive URL")
        act_ftp.setToolTip("Copy direct URL to the 3GPP FTP spec archive directory.")
        act_ftp.triggered.connect(
            lambda _, p=ftp_url, h=f'<a href="{ftp_url}">{ftp_url}</a>': _copy_to_clipboard(
                plain_text=p,
                html_text=h,
                label="FTP Archive URL",
                parent=parent_widget,
            )
        )


def populate_multi_spec_copy_menu(
    copy_sub: QMenu,
    parent_widget: QWidget,
    target_specs: list,
    db,
):
    """Populates copy options for multiple selected specifications."""
    if not target_specs:
        return

    specs_meta = [resolve_spec_metadata(s, db) for s in target_specs]
    count = len(specs_meta)

    # 1. Copy Spec Numbers
    nums_plain = ", ".join(m["canonical_num"] for m in specs_meta)
    nums_html_parts = []
    for m in specs_meta:
        if m["dynareport_url"]:
            nums_html_parts.append(
                f'<a href="{m["dynareport_url"]}">{html.escape(m["canonical_num"])}</a>'
            )
        else:
            nums_html_parts.append(html.escape(m["canonical_num"]))
    nums_html = ", ".join(nums_html_parts)

    act_nums = copy_sub.addAction(f"📋 Copy Spec Numbers ({count})")
    act_nums.setToolTip("Copy all selected specification numbers as a comma-separated list.")
    act_nums.triggered.connect(
        lambda _, p=nums_plain, h=nums_html, c=count: _copy_to_clipboard(
            p, h, f"{c} Spec Numbers", parent_widget
        )
    )

    # 2. Copy Numbers & Titles
    lines_plain = []
    lines_html = []
    for m in specs_meta:
        c_num = m["canonical_num"]
        c_title = m["title"]
        url = m["dynareport_url"]
        link_h = f'<a href="{url}">{html.escape(c_num)}</a>' if url else html.escape(c_num)
        if c_title:
            lines_plain.append(f"{c_num}: {c_title}")
            lines_html.append(f"{link_h}: {html.escape(c_title)}")
        else:
            lines_plain.append(c_num)
            lines_html.append(link_h)

    act_titles = copy_sub.addAction(f"📋 Copy Numbers & Titles ({count})")
    act_titles.setToolTip("Copy specification numbers and titles (one per line).")
    act_titles.triggered.connect(
        lambda _, p="\n".join(lines_plain), h="<p>" + "<br>".join(lines_html) + "</p>", c=count: _copy_to_clipboard(
            p, h, f"{c} Numbers & Titles", parent_widget
        )
    )

    # 3. Copy Full Citations (with WG)
    cit_plain = []
    cit_html = []
    for m in specs_meta:
        c_num = m["canonical_num"]
        c_title = m["title"]
        wg = m["wg"]
        url = m["dynareport_url"]
        link_h = f'<a href="{url}">{html.escape(c_num)}</a>' if url else html.escape(c_num)
        safe_t = html.escape(c_title)
        safe_wg = html.escape(wg)

        if c_title and wg:
            cit_plain.append(f"{c_num}: {c_title} ({wg})")
            cit_html.append(f"{link_h}: {safe_t} ({safe_wg})")
        elif c_title:
            cit_plain.append(f"{c_num}: {c_title}")
            cit_html.append(f"{link_h}: {safe_t}")
        elif wg:
            cit_plain.append(f"{c_num} ({wg})")
            cit_html.append(f"{link_h} ({safe_wg})")
        else:
            cit_plain.append(c_num)
            cit_html.append(link_h)

    act_cit = copy_sub.addAction(f"📋 Copy Full Citations (with WG) ({count})")
    act_cit.setToolTip(
        "Copy complete citations including specification numbers, titles, and Working Groups."
    )
    act_cit.triggered.connect(
        lambda _, p="\n".join(cit_plain), h="<p>" + "<br>".join(cit_html) + "</p>", c=count: _copy_to_clipboard(
            p, h, f"{c} Citations", parent_widget
        )
    )

    # Separator
    copy_sub.addSeparator()

    # 4. Copy as Markdown List
    md_lines = []
    for m in specs_meta:
        c_num = m["canonical_num"]
        c_title = m["title"]
        wg = m["wg"]
        url = m["dynareport_url"]
        wg_str = f" ({wg})" if wg else ""
        if url and c_title:
            md_lines.append(f"- [{c_num}]({url}): {c_title}{wg_str}")
        elif url:
            md_lines.append(f"- [{c_num}]({url}){wg_str}")
        elif c_title:
            md_lines.append(f"- {c_num}: {c_title}{wg_str}")
        else:
            md_lines.append(f"- {c_num}")

    act_md = copy_sub.addAction(f"📋 Copy as Markdown List ({count})")
    act_md.setToolTip("Copy as Markdown bullet list with hyperlinks.")
    act_md.triggered.connect(
        lambda _, p="\n".join(md_lines), h="<p>" + "<br>".join(cit_html) + "</p>", c=count: _copy_to_clipboard(
            p, h, f"{c} Markdown Items", parent_widget
        )
    )

    # 5. Copy Web Report URLs
    urls = [m["dynareport_url"] for m in specs_meta if m["dynareport_url"]]
    if urls:
        act_urls = copy_sub.addAction(f"📋 Copy 3GPP Web Report URLs ({len(urls)})")
        act_urls.setToolTip("Copy direct URLs to 3GPP DynaReport specification pages.")
        act_urls.triggered.connect(
            lambda _, p="\n".join(urls), h="<br>".join(f'<a href="{u}">{u}</a>' for u in urls), c=len(urls): _copy_to_clipboard(
                p, h, f"{c} URLs", parent_widget
            )
        )


def copy_specs_default(parent_widget: QWidget, target_specs: list, db):
    """Default copy behavior triggered by Ctrl+C on the specifications table."""
    if not target_specs:
        return
    specs_meta = [resolve_spec_metadata(s, db) for s in target_specs]
    count = len(specs_meta)

    if count == 1:
        m = specs_meta[0]
        canonical = m["canonical_num"]
        title = m["title"]
        wg = m["wg"]
        url = m["dynareport_url"]
        link_h = f'<a href="{url}">{html.escape(canonical)}</a>' if url else html.escape(canonical)

        if title and wg:
            plain = f"{canonical}: {title} ({wg})"
            html_t = f"{link_h}: {html.escape(title)} ({html.escape(wg)})"
        elif title:
            plain = f"{canonical}: {title}"
            html_t = f"{link_h}: {html.escape(title)}"
        else:
            plain = canonical
            html_t = link_h

        _copy_to_clipboard(plain, html_t, f"Specification {canonical}", parent_widget)
    else:
        cit_plain = []
        cit_html = []
        for m in specs_meta:
            c_num = m["canonical_num"]
            c_title = m["title"]
            wg = m["wg"]
            url = m["dynareport_url"]
            link_h = f'<a href="{url}">{html.escape(c_num)}</a>' if url else html.escape(c_num)
            safe_t = html.escape(c_title)
            safe_wg = html.escape(wg)

            if c_title and wg:
                cit_plain.append(f"{c_num}: {c_title} ({wg})")
                cit_html.append(f"{link_h}: {safe_t} ({safe_wg})")
            elif c_title:
                cit_plain.append(f"{c_num}: {c_title}")
                cit_html.append(f"{link_h}: {safe_t}")
            else:
                cit_plain.append(c_num)
                cit_html.append(link_h)

        _copy_to_clipboard(
            "\n".join(cit_plain),
            "<p>" + "<br>".join(cit_html) + "</p>",
            f"{count} Specifications",
            parent_widget,
        )


def show_spec_context_menu(tab, position):
    """
    Constructs and executes the context menu for selected specifications rows in the table.

    :param tab: The SpecificationsTab parent widget instance.
    :param position: QPoint position of the context menu request from table viewport.
    """
    selected_rows = tab.table.selectionModel().selectedRows()
    if not selected_rows:
        return

    target_specs = []
    for index in selected_rows:
        widget = tab.table.cellWidget(index.row(), 0)
        if widget:
            num_label = widget.findChild(QLabel, "specNumberLabel")
            if num_label:
                # Never use decorated UI text (e.g. "33.801-01 ⭐") as an identifier.
                canonical = num_label.property("specNumber")
                if canonical:
                    target_specs.append(str(canonical).strip())
                else:
                    # Compatibility fallback for widgets created by older code.
                    display_text = num_label.text().strip()
                    match = re.match(r"^(\d{2}\.\d{2,3}(?:-[A-Za-z0-9]+)?)", display_text)
                    if match:
                        target_specs.append(match.group(1))

    if not target_specs:
        return

    menu = QMenu(tab)
    menu.setStyleSheet("""
        QMenu { background-color: #FAFAFA; border: 1px solid #CCC; }
        QMenu::item { padding: 5px 20px 5px 15px; color: #333333; }
        QMenu::item:selected { background-color: #E1F0FF; color: #0078D7; }
        QMenu::item:disabled { color: #AAAAAA; }
    """)

    if len(target_specs) == 1:
        spec_num = target_specs[0]
        is_fav = spec_num in tab.favorites
        fav_label = "⭐ Remove from Favorites" if is_fav else "⭐ Add to Favorites"
        menu.addAction(fav_label).triggered.connect(lambda _, s=spec_num: tab._toggle_favorite(s))

        edit_ann_action = menu.addAction("🏷️ Edit Tags & Notes...")
        edit_ann_action.triggered.connect(lambda _, s=spec_num: tab._open_edit_annotations(s))

        menu.addSeparator()

        # Quick Copy Submenu
        copy_sub = menu.addMenu("📋 Quick Copy")
        copy_sub.setToolTipsVisible(True)
        copy_sub.setToolTip("Copy specification metadata in plain text and rich HTML form.")
        populate_spec_copy_menu(copy_sub, tab, spec_num, db=tab.db)

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

        copy_sub = menu.addMenu(f"📋 Quick Copy ({len(target_specs)} selected)")
        copy_sub.setToolTipsVisible(True)
        copy_sub.setToolTip("Copy metadata for all selected specifications.")
        populate_multi_spec_copy_menu(copy_sub, tab, target_specs, db=tab.db)

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
