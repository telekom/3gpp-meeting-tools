"""Quick-copy submenu builders and Ctrl+C default copy handlers."""

from __future__ import annotations

import html

from PyQt5.QtGui import QCursor
from PyQt5.QtWidgets import QActionGroup, QMenu, QToolTip, QWidget

from modules.specifications.ui.menu.clipboard_and_citations import (
    _copy_to_clipboard,
    build_citation_data,
    get_saved_copy_mode,
    resolve_spec_metadata,
    save_copy_mode,
)


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
            sub_both.addAction("📋 Copy Number && Title (Both Links)").triggered.connect(
                lambda _, p=cdata["num_title_plain"], h=cdata["num_title_both_html"]: _copy_to_clipboard(
                    p, h, "Number & Title (Both Links)", parent_widget
                )
            )
        sub_both.addAction("📋 Copy Spec Number (Both Links)").triggered.connect(
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
        sub_web.addAction("📋 Copy Number && Title (Web Link)").triggered.connect(
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
        sub_web.addAction(f"📋 Copy Number && WG ({canonical} ({wg}))").triggered.connect(
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
            sub_zip.addAction(f"📋 Copy Number && Title (ZIP Link{ver_label})").triggered.connect(
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
            sub_zip.addAction("📋 Copy Number && WG (with ZIP Link)").triggered.connect(
                lambda _, p=f"{canonical} ({clean_ver}) ({wg})", h=f"{cdata['zip_anchor']} ({safe_wg})": _copy_to_clipboard(
                    p, h, f"Number & WG (ZIP{ver_label})", parent_widget
                )
            )
        sub_zip.addSeparator()
        sub_zip.addAction("📋 Copy as Markdown (ZIP Link)").triggered.connect(
            lambda _, p=cdata["zip_md"], h=cdata["zip_anchor"]: _copy_to_clipboard(
                p, h, f"Markdown (ZIP{ver_label})", parent_widget
            )
        )
        if title:
            sub_zip.addAction("📋 Copy as Markdown with Title (ZIP Link)").triggered.connect(
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
            return lambda: (
                save_copy_mode(k),
                QToolTip.showText(QCursor.pos(), f"📋 Default link mode set to: {l}!", parent_widget),
            )

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
    act_nt = copy_sub.addAction(f"📋 Copy Numbers && Titles ({count})")
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
        web_plain = "\n".join(web_urls)
        web_html = "<br>".join(f'<a href="{url}">{url}</a>' for url in web_urls)
        web_count = len(web_urls)
        act_w_urls = copy_sub.addAction(f"🌐 Copy 3GPP Web Report URLs ({web_count})")
        act_w_urls.setToolTip("Copy direct URLs to 3GPP DynaReport specification pages.")
        act_w_urls.triggered.connect(
            lambda _, p=web_plain, h=web_html, c=web_count: _copy_to_clipboard(
                p, h, f"{c} Web URLs", parent_widget
            )
        )

    zip_urls = [m["zip_url"] for m in items_processed if m["zip_url"]]
    if zip_urls:
        zip_plain = "\n".join(zip_urls)
        zip_html = "<br>".join(f'<a href="{url}">{url}</a>' for url in zip_urls)
        zip_count = len(zip_urls)
        act_z_urls = copy_sub.addAction(f"📦 Copy ZIP Download URLs ({zip_count})")
        act_z_urls.setToolTip("Copy direct URLs to download specification ZIP archives.")
        act_z_urls.triggered.connect(
            lambda _, p=zip_plain, h=zip_html, c=zip_count: _copy_to_clipboard(
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
