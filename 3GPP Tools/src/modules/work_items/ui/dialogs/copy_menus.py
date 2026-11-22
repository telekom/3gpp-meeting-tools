"""Clipboard payload builders and QMenu helpers for 3GPP Work Items."""

from __future__ import annotations

import html
import re
from typing import List, Optional

from PyQt5.QtWidgets import QMenu, QWidget

from core.utils.portal_utils import get_portal_tdoc_url, get_portal_work_item_url
from modules.meetings.ui.tdocs_menus import _copy_to_clipboard


def build_wi_copy_data(row_data: dict) -> dict:
    """
    Builds plain-text, Markdown, and rich-HTML (Microsoft Word / Outlook compatible)
    clipboard payloads for a 3GPP Work Item row.
    """
    wi_code = str(row_data.get("code") or "").strip()
    raw_acronym = str(row_data.get("acronym") or "").strip()
    acronym = raw_acronym or f"WI #{wi_code}"
    name = str(row_data.get("name") or "").strip()
    wg = str(row_data.get("wg_names") or "").strip()
    rel = str(row_data.get("release") or "").strip()
    latest_wid = str(row_data.get("latest_wid") or "").strip()

    wi_url = get_portal_work_item_url(wi_code) if wi_code else ""

    wid_url = ""
    if latest_wid:
        base_wid = re.split(r"(?:r|rev)\d+", latest_wid, maxsplit=1, flags=re.IGNORECASE)[0].rstrip("-")
        if not base_wid:
            base_wid = latest_wid.split("r")[0].split("R")[0]
        wid_url = get_portal_tdoc_url(base_wid)

    safe_code = html.escape(wi_code)
    safe_acr = html.escape(acronym)
    safe_name = html.escape(name)
    safe_wid = html.escape(latest_wid)

    meta_parts = [p for p in (rel, wg) if p]
    meta_str = f" ({', '.join(meta_parts)})" if meta_parts else ""
    safe_meta = html.escape(meta_str)

    # 1. Acronym + URL
    acr_link_plain = f"{acronym} ({wi_url})" if wi_url else acronym
    acr_link_html = f'<a href="{wi_url}">{safe_acr}</a>' if wi_url else safe_acr
    acr_vis_url_html = f'<b>{safe_acr}</b> (<a href="{wi_url}">{wi_url}</a>)' if wi_url else safe_acr
    acr_md = f"[{acronym}]({wi_url})" if wi_url else acronym

    # 2. Latest WID + URL
    wid_link_plain = f"{latest_wid} ({wid_url})" if (latest_wid and wid_url) else latest_wid
    wid_link_html = f'<a href="{wid_url}">{safe_wid}</a>' if (latest_wid and wid_url) else safe_wid
    wid_vis_url_html = (
        f'<b>{safe_wid}</b> (<a href="{wid_url}">{wid_url}</a>)'
        if (latest_wid and wid_url)
        else safe_wid
    )
    wid_md = f"[{latest_wid}]({wid_url})" if (latest_wid and wid_url) else latest_wid

    # 3. Full Citation
    title_part = f" — {name}" if name else ""
    safe_title_part = f" — {safe_name}" if name else ""
    cit_plain = f"WI {wi_code}: {acronym}{title_part}{meta_str}" + (f" [{wi_url}]" if wi_url else "")
    if wi_url:
        cit_html = f'<a href="{wi_url}"><b>WI {safe_code}: {safe_acr}</b></a>{safe_title_part}{safe_meta}'
    else:
        cit_html = f"<b>WI {safe_code}: {safe_acr}</b>{safe_title_part}{safe_meta}"

    if latest_wid and wid_url:
        cit_with_wid_plain = (
            f"WI {wi_code}: {acronym}{title_part}{meta_str} "
            f"[WI: {wi_url} | WID {latest_wid}: {wid_url}]"
        )
        cit_with_wid_html = f'{cit_html} [WID: <a href="{wid_url}">{safe_wid}</a>]'
    else:
        cit_with_wid_plain = cit_plain
        cit_with_wid_html = cit_html

    cit_md = (
        f"[WI {wi_code}: {acronym}]({wi_url}){title_part}{meta_str}"
        if wi_url
        else f"WI {wi_code}: {acronym}{title_part}{meta_str}"
    )

    return {
        "code": wi_code,
        "acronym": acronym,
        "name": name,
        "wg": wg,
        "rel": rel,
        "latest_wid": latest_wid,
        "wi_url": wi_url,
        "wid_url": wid_url,
        "acr_link_plain": acr_link_plain,
        "acr_link_html": acr_link_html,
        "acr_vis_url_html": acr_vis_url_html,
        "acr_md": acr_md,
        "wid_link_plain": wid_link_plain,
        "wid_link_html": wid_link_html,
        "wid_vis_url_html": wid_vis_url_html,
        "wid_md": wid_md,
        "cit_plain": cit_plain,
        "cit_html": cit_html,
        "cit_with_wid_plain": cit_with_wid_plain,
        "cit_with_wid_html": cit_with_wid_html,
        "cit_md": cit_md,
    }


def populate_wi_copy_menu(copy_sub: QMenu, row_data: dict, parent: Optional[QWidget] = None):
    """Populates a QMenu with rich single-Work-Item copy options."""
    copy_sub.setToolTipsVisible(True)
    d = build_wi_copy_data(row_data)
    acronym = d["acronym"]
    wi_code = d["code"]
    name = d["name"]
    latest_wid = d["latest_wid"]
    wi_url = d["wi_url"]
    wid_url = d["wid_url"]

    # 1. WI Acronym including URL (Hyperlink for Word/Outlook + Plain text with URL)
    act_acr_link = copy_sub.addAction(f"📋 Copy WI Acronym with Link ({acronym})")
    act_acr_link.setToolTip(
        "Pastes as a clickable hyperlink in Word/Outlook, or 'Acronym (URL)' in plain text."
    )
    act_acr_link.triggered.connect(
        lambda _, p=d["acr_link_plain"], h=d["acr_link_html"]: _copy_to_clipboard(
            p, h, f"WI Acronym ({acronym})", parent
        )
    )

    act_acr_vis = copy_sub.addAction("📋 Copy WI Acronym + Visible URL")
    act_acr_vis.setToolTip("Pastes 'Acronym (https://...)' with both the acronym and visible URL.")
    act_acr_vis.triggered.connect(
        lambda _, p=d["acr_link_plain"], h=d["acr_vis_url_html"]: _copy_to_clipboard(
            p, h, "WI Acronym + URL", parent
        )
    )

    # 2. Only URL (WI Portal URL)
    if wi_url:
        act_wi_url = copy_sub.addAction("🌐 Copy WI Portal URL Only")
        act_wi_url.setToolTip("Copy only the 3GPP Portal URL for this Work Item.")
        act_wi_url.triggered.connect(
            lambda _, u=wi_url: _copy_to_clipboard(
                u, f'<a href="{u}">{u}</a>', "WI Portal URL", parent
            )
        )

    # 3. Latest WID including URL
    if latest_wid:
        copy_sub.addSeparator()
        act_wid_link = copy_sub.addAction(f"📄 Copy Latest WID with Link ({latest_wid})")
        act_wid_link.setToolTip(
            "Pastes the Latest WID TDoc ID as a clickable 3GU Portal link in Word/Outlook, or 'WID (URL)' in plain text."
        )
        act_wid_link.triggered.connect(
            lambda _, p=d["wid_link_plain"], h=d["wid_link_html"]: _copy_to_clipboard(
                p, h, f"Latest WID ({latest_wid})", parent
            )
        )

        act_wid_vis = copy_sub.addAction(f"📄 Copy Latest WID + Visible URL ({latest_wid})")
        act_wid_vis.setToolTip("Pastes 'WID (https://...)' with visible 3GU Portal URL.")
        act_wid_vis.triggered.connect(
            lambda _, p=d["wid_link_plain"], h=d["wid_vis_url_html"]: _copy_to_clipboard(
                p, h, f"Latest WID + URL ({latest_wid})", parent
            )
        )

        if wid_url:
            act_wid_url = copy_sub.addAction("🌐 Copy Latest WID URL Only")
            act_wid_url.setToolTip("Copy only the 3GU Portal URL for the Latest WID TDoc.")
            act_wid_url.triggered.connect(
                lambda _, u=wid_url: _copy_to_clipboard(
                    u, f'<a href="{u}">{u}</a>', "Latest WID URL", parent
                )
            )

    copy_sub.addSeparator()

    # 4. Full Citations & Markdown
    act_cit = copy_sub.addAction("📋 Copy Full WI Citation (with Link)")
    act_cit.setToolTip("Copy complete Work Item citation (Code, Acronym, Title, Release, WG) with hyperlink.")
    act_cit.triggered.connect(
        lambda _, p=d["cit_plain"], h=d["cit_html"]: _copy_to_clipboard(
            p, h, "WI Citation", parent
        )
    )

    if latest_wid:
        act_cit_wid = copy_sub.addAction("📋 Copy Full Citation + Latest WID Link")
        act_cit_wid.setToolTip("Copy complete Work Item citation including both WI and Latest WID links.")
        act_cit_wid.triggered.connect(
            lambda _, p=d["cit_with_wid_plain"], h=d["cit_with_wid_html"]: _copy_to_clipboard(
                p, h, "WI & WID Citation", parent
            )
        )

    act_md = copy_sub.addAction(f"🔗 Copy as Markdown Link ([{acronym}](url))")
    act_md.setToolTip("Copy Markdown link syntax for notes, Obsidian, or GitHub.")
    act_md.triggered.connect(
        lambda _, p=d["acr_md"], h=d["acr_link_html"]: _copy_to_clipboard(
            p, h, "Markdown Link", parent
        )
    )

    copy_sub.addSeparator()

    # 5. Plain Text Identifiers
    act_plain_acr = copy_sub.addAction(f"📋 Copy Acronym Only ({acronym})")
    act_plain_acr.setToolTip("Copy only the plain-text Work Item acronym without any URL.")
    act_plain_acr.triggered.connect(
        lambda _, a=acronym: _copy_to_clipboard(a, html.escape(a), "Acronym", parent)
    )

    if wi_code:
        act_plain_code = copy_sub.addAction(f"📋 Copy WI Code Only ({wi_code})")
        act_plain_code.setToolTip("Copy only the numeric Work Item UID/code.")
        act_plain_code.triggered.connect(
            lambda _, c=wi_code: _copy_to_clipboard(c, html.escape(c), "WI Code", parent)
        )

    if name:
        act_plain_name = copy_sub.addAction("📋 Copy Title / Name Only")
        act_plain_name.setToolTip("Copy only the Work Item title.")
        act_plain_name.triggered.connect(
            lambda _, n=name: _copy_to_clipboard(n, html.escape(n), "WI Title", parent)
        )


def populate_wi_batch_copy_menu(copy_sub: QMenu, rows_data: List[dict], parent: Optional[QWidget] = None):
    """Populates a QMenu with batch copy options for multiple selected Work Items."""
    copy_sub.setToolTipsVisible(True)
    items = [build_wi_copy_data(r) for r in rows_data if r and r.get("code")]
    count = len(items)
    if not count:
        return

    # 1. Acronyms with Links
    acr_links_plain = ", ".join(d["acr_link_plain"] for d in items)
    acr_links_html = ", ".join(d["acr_link_html"] for d in items)
    act_acr_links = copy_sub.addAction(f"📋 Copy Acronyms with Links ({count})")
    act_acr_links.setToolTip("Copy comma-separated Work Item acronyms as clickable links.")
    act_acr_links.triggered.connect(
        lambda _, p=acr_links_plain, h=acr_links_html, c=count: _copy_to_clipboard(
            p, h, f"{c} WI Acronyms (with Links)", parent
        )
    )

    # 2. Portal URLs List
    urls = [d["wi_url"] for d in items if d["wi_url"]]
    if urls:
        urls_plain = "\n".join(urls)
        urls_html = "<br>".join(f'<a href="{u}">{u}</a>' for u in urls)
        act_urls = copy_sub.addAction(f"🌐 Copy WI Portal URLs List ({len(urls)})")
        act_urls.setToolTip("Copy direct 3GPP Portal URLs (one per line).")
        act_urls.triggered.connect(
            lambda _, p=urls_plain, h=urls_html, c=len(urls): _copy_to_clipboard(
                p, h, f"{c} WI Portal URLs", parent
            )
        )

    # 3. Latest WIDs with Links
    wid_items = [d for d in items if d["latest_wid"]]
    if wid_items:
        wids_plain = "\n".join(f"{d['acronym']}: {d['wid_link_plain']}" for d in wid_items)
        wids_html = "<br>".join(
            f"<b>{html.escape(d['acronym'])}</b>: {d['wid_link_html']}" for d in wid_items
        )
        act_wids = copy_sub.addAction(f"📄 Copy Latest WIDs with Links ({len(wid_items)})")
        act_wids.setToolTip("Copy Latest WID TDoc links for selected Work Items.")
        act_wids.triggered.connect(
            lambda _, p=wids_plain, h=wids_html, c=len(wid_items): _copy_to_clipboard(
                p, h, f"{c} Latest WID Links", parent
            )
        )

    copy_sub.addSeparator()

    # 4. Full Citations List
    cits_plain = "\n".join(d["cit_plain"] for d in items)
    cits_html = "<p>" + "<br>".join(d["cit_html"] for d in items) + "</p>"
    act_cits = copy_sub.addAction(f"📋 Copy Full Citations List ({count})")
    act_cits.setToolTip("Copy full citations for all selected Work Items (one per line).")
    act_cits.triggered.connect(
        lambda _, p=cits_plain, h=cits_html, c=count: _copy_to_clipboard(
            p, h, f"{c} WI Citations", parent
        )
    )

    # 5. Markdown Bullet List
    md_plain = "\n".join(f"- {d['cit_md']}" for d in items)
    md_html = "<ul>" + "".join(f"<li>{d['cit_html']}</li>" for d in items) + "</ul>"
    act_md = copy_sub.addAction(f"🔗 Copy as Markdown List ({count})")
    act_md.setToolTip("Copy as a formatted bullet list in Markdown and rich HTML.")
    act_md.triggered.connect(
        lambda _, p=md_plain, h=md_html, c=count: _copy_to_clipboard(
            p, h, f"{c} Markdown Items", parent
        )
    )

    copy_sub.addSeparator()

    # 6. Plain Acronyms Only
    acrs_plain = ", ".join(d["acronym"] for d in items)
    acrs_html = ", ".join(html.escape(d["acronym"]) for d in items)
    act_acrs = copy_sub.addAction(f"📋 Copy Acronyms Only (Comma-separated) ({count})")
    act_acrs.setToolTip("Copy plain-text Work Item acronyms separated by commas.")
    act_acrs.triggered.connect(
        lambda _, p=acrs_plain, h=acrs_html, c=count: _copy_to_clipboard(
            p, h, f"{c} WI Acronyms", parent
        )
    )


def copy_wis_default(rows_data: List[dict], parent: Optional[QWidget] = None):
    """Default Ctrl+C copy handler for selected rows in the Work Items table."""
    items = [build_wi_copy_data(r) for r in rows_data if r and r.get("code")]
    if not items:
        return
    if len(items) == 1:
        d = items[0]
        _copy_to_clipboard(
            d["acr_link_plain"], d["acr_link_html"], f"WI {d['acronym']}", parent
        )
    else:
        plain = ", ".join(d["acr_link_plain"] for d in items)
        html_t = ", ".join(d["acr_link_html"] for d in items)
        _copy_to_clipboard(plain, html_t, f"{len(items)} Work Items", parent)
