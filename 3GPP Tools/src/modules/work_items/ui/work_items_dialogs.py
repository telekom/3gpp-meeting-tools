# --- File: src/modules/work_items/ui/work_items_dialogs.py ---
import html
import logging
import re
import webbrowser
from datetime import datetime
from typing import List, Optional

from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QDialog,
    QFormLayout,
    QFrame,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMenu,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QVBoxLayout,
    QWidget,
)

from core.ui.ui_components import BUTTON_STYLE_TOOLBAR_SECONDARY
from modules.meetings.ui.tdocs_menus import _copy_to_clipboard
from modules.work_items.core.wi_database import WorkItemsDatabase
from modules.work_items.core.wi_scraper import extract_wi_code, fetch_wi_details_from_portal

logger = logging.getLogger(__name__)


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

    wi_url = (
        f"https://portal.3gpp.org/desktopmodules/WorkItem/WorkItemDetails.aspx?workitemId={wi_code}"
        if wi_code
        else ""
    )

    wid_url = ""
    if latest_wid:
        base_wid = re.split(r"(?:r|rev)\d+", latest_wid, maxsplit=1, flags=re.IGNORECASE)[0].rstrip("-")
        if not base_wid:
            base_wid = latest_wid.split("r")[0].split("R")[0]
        wid_url = f"https://portal.3gpp.org/ngppapp/CreateTdoc.aspx?mode=view&tdocId={base_wid}"

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
        cit_with_wid_html = (
            f'{cit_html} [WID: <a href="{wid_url}">{safe_wid}</a>]'
        )
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
        _copy_to_clipboard(
            plain, html_t, f"{len(items)} Work Items", parent
        )



class WorkItemInfoDialog(QDialog):
    """Inspector Dialog displaying full Work Item details, linked specs, and remarks history."""

    def __init__(
        self,
        details: dict,
        parent=None,
        is_starred: bool = False,
        is_following: bool = False,
        on_toggle_star=None,
        on_toggle_follow=None,
    ):
        super().__init__(parent)
        self.details = details
        self.is_starred = is_starred
        self.is_following = is_following
        self.on_toggle_star = on_toggle_star
        self.on_toggle_follow = on_toggle_follow

        wi_code = details.get('code', 'Unknown')
        acronym = details.get('acronym', 'Unknown')
        name = details.get('name', 'No Title Available')
        release = details.get('release', '')
        end_date = details.get('end_date', '')

        self.setWindowTitle(f"Work Item Details: {acronym} (#{wi_code})")
        self.setMinimumWidth(620)
        self.setMinimumHeight(480)
        self.setStyleSheet("""
            QFrame#cardFrame {
                background-color: #FFFFFF;
                border: 1px solid #CBD5E1;
                border-radius: 8px;
            }
            QLabel { font-size: 13px; color: #1E293B; }
            QPushButton#specChipBtn {
                background-color: #EBF3FC;
                border: 1px solid #BFDBFE;
                border-radius: 12px;
                padding: 3px 10px;
                font-size: 11px;
                color: #1E5C99;
                font-weight: bold;
            }
            QPushButton#specChipBtn:hover {
                background-color: #DBEAFE;
                border-color: #1E5C99;
            }
            QScrollArea { border: none; background-color: transparent; }
        """)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)

        # --- 1. HEADER CARD ---
        header_card = QFrame()
        header_card.setObjectName("cardFrame")
        header_layout = QVBoxLayout(header_card)
        header_layout.setContentsMargins(14, 12, 14, 12)
        header_layout.setSpacing(8)

        title_row = QHBoxLayout()

        code_badge = QLabel(f"<b>WI #{wi_code}</b>")
        code_badge.setStyleSheet("""
            background-color: #EBF3FC; color: #1E5C99; border: 1px solid #BFDBFE;
            border-radius: 4px; padding: 2px 6px; font-size: 12px; font-weight: bold;
        """)
        code_badge.setTextInteractionFlags(Qt.TextSelectableByMouse)

        if release:
            rel_badge = QLabel(f"<b>{release}</b>")
            rel_badge.setStyleSheet("""
                background-color: #ECFDF5; color: #065F46; border: 1px solid #A7F3D0;
                border-radius: 4px; padding: 2px 6px; font-size: 12px; font-weight: bold;
            """)
            title_row.addWidget(rel_badge)

        is_finished = bool(end_date and end_date.strip() and end_date < datetime.now().strftime("%Y-%m-%d"))
        status_badge = QLabel("🏁 Finished" if is_finished else "🟢 Active / Ongoing")
        status_badge.setStyleSheet(
            "background-color: #F1F5F9; color: #475569; border-radius: 4px; padding: 2px 6px; font-size: 11px; font-weight: bold;"
            if is_finished else
            "background-color: #ECFDF5; color: #065F46; border: 1px solid #A7F3D0; border-radius: 4px; padding: 2px 6px; font-size: 11px; font-weight: bold;"
        )

        acronym_lbl = QLabel(f"<b>{acronym}</b>")
        acronym_lbl.setStyleSheet("font-size: 16px; color: #1E293B; font-weight: bold;")
        acronym_lbl.setTextInteractionFlags(Qt.TextSelectableByMouse)

        title_row.addWidget(code_badge)
        title_row.addWidget(acronym_lbl)
        title_row.addWidget(status_badge)
        title_row.addStretch()

        # Starred Toggle Button
        self.star_btn = QPushButton()
        self.star_btn.setCursor(Qt.PointingHandCursor)
        self.star_btn.setToolTip("Toggle Starred status (Personal Interest / Watchlist)")
        self._update_star_btn_style()
        self.star_btn.clicked.connect(self._handle_toggle_star)
        title_row.addWidget(self.star_btn)

        # Following Toggle Button
        self.follow_btn = QPushButton()
        self.follow_btn.setCursor(Qt.PointingHandCursor)
        self.follow_btn.setToolTip("Toggle Following status (Assigned Delegation Responsibility)")
        self._update_follow_btn_style()
        self.follow_btn.clicked.connect(self._handle_toggle_follow)
        title_row.addWidget(self.follow_btn)

        header_layout.addLayout(title_row)

        desc_label = QLabel(name)
        desc_label.setWordWrap(True)
        desc_label.setStyleSheet("color: #475569; font-size: 13px; line-height: 1.4;")
        desc_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        header_layout.addWidget(desc_label)

        layout.addWidget(header_card)

        # --- 2. DETAILS & TIMELINE CARD ---
        details_card = QFrame()
        details_card.setObjectName("cardFrame")
        form = QFormLayout(details_card)
        form.setContentsMargins(14, 12, 14, 12)
        form.setSpacing(8)
        form.setLabelAlignment(Qt.AlignRight)

        self._add_row(form, "Working Groups", details.get('wg_names') or '-')
        self._add_row(form, "Start Date", details.get('start_date') or '-')
        self._add_row(form, "Target End Date", details.get('end_date') or '-')
        self._add_row(form, "Latest WID", details.get('latest_wid') or '-')

        linked_specs = details.get('linked_specs', [])
        if linked_specs:
            specs_container = QWidget()
            specs_layout = QHBoxLayout(specs_container)
            specs_layout.setContentsMargins(0, 0, 0, 0)
            specs_layout.setSpacing(6)

            for sp in linked_specs:
                num = sp.get('number', '')
                is_pri = sp.get('is_primary', False)
                clean_num = num.split("-")[0].replace('.', '').strip()
                label = f"⭐ {num}" if is_pri else num

                btn = QPushButton(label)
                btn.setObjectName("specChipBtn")
                btn.setCursor(Qt.PointingHandCursor)
                btn.setToolTip(f"{sp.get('title', '')}\nClick to open 3GPP Portal Report")
                btn.clicked.connect(lambda _, n=clean_num: webbrowser.open(f"https://www.3gpp.org/DynaReport/{n}.htm"))
                specs_layout.addWidget(btn)

            specs_layout.addStretch()
            form.addRow(self._make_key_label("Impacted Specs:"), specs_container)
        else:
            self._add_row(form, "Impacted Specs", "-")

        layout.addWidget(details_card)

        # --- 3. HISTORICAL REMARKS FEED ---
        remarks = details.get('remarks', [])
        if remarks:
            remarks_lbl = QLabel(f"<b>Secretary Remarks ({len(remarks)}):</b>")
            remarks_lbl.setStyleSheet("color: #475569; font-size: 12px; margin-top: 4px;")
            layout.addWidget(remarks_lbl)

            remarks_scroll = QScrollArea()
            remarks_scroll.setWidgetResizable(True)
            remarks_scroll.setMaximumHeight(160)

            remarks_content = QWidget()
            remarks_vbox = QVBoxLayout(remarks_content)
            remarks_vbox.setContentsMargins(0, 0, 0, 0)
            remarks_vbox.setSpacing(6)

            for rm in remarks:
                r_card = QFrame()
                r_card.setStyleSheet("background-color: #FFFFFF; border: 1px solid #E2E8F0; border-radius: 6px; padding: 6px;")
                r_box = QVBoxLayout(r_card)
                r_box.setContentsMargins(6, 4, 6, 4)
                r_box.setSpacing(3)

                if rm.get('date'):
                    date_lbl = QLabel(f"📅 {rm['date']}")
                    date_lbl.setStyleSheet("color: #64748B; font-size: 11px; font-weight: bold;")
                    r_box.addWidget(date_lbl)

                text_lbl = QLabel(rm.get('text', ''))
                text_lbl.setWordWrap(True)
                text_lbl.setStyleSheet("color: #1E293B; font-size: 12px;")
                text_lbl.setTextInteractionFlags(Qt.TextSelectableByMouse)
                r_box.addWidget(text_lbl)

                remarks_vbox.addWidget(r_card)

            remarks_vbox.addStretch()
            remarks_scroll.setWidget(remarks_content)
            layout.addWidget(remarks_scroll)

        # --- 4. ACTION BUTTONS ---
        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(8)

        portal_btn = QPushButton("🌐 Open 3GPP Portal")
        portal_btn.setObjectName("primaryBtn")
        portal_btn.setCursor(Qt.PointingHandCursor)
        portal_btn.clicked.connect(lambda: webbrowser.open(
            f"https://portal.3gpp.org/desktopmodules/WorkItem/WorkItemDetails.aspx?workitemId={wi_code}"
        ))
        btn_layout.addWidget(portal_btn)

        copy_btn = QPushButton("📋 Quick Copy ▾")
        copy_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        copy_btn.setCursor(Qt.PointingHandCursor)
        copy_menu = QMenu(copy_btn)
        populate_wi_copy_menu(copy_menu, self.details, self)
        copy_btn.setMenu(copy_menu)
        btn_layout.addWidget(copy_btn)

        btn_layout.addStretch()

        close_btn = QPushButton("Close")
        close_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        close_btn.setCursor(Qt.PointingHandCursor)
        close_btn.clicked.connect(self.accept)
        btn_layout.addWidget(close_btn)

        layout.addLayout(btn_layout)

    def _update_star_btn_style(self):
        if self.is_starred:
            self.star_btn.setText("⭐ Starred")
            self.star_btn.setStyleSheet("""
                QPushButton {
                    background-color: #FEF3C7; color: #92400E; border: 1px solid #F59E0B;
                    border-radius: 4px; padding: 3px 8px; font-size: 11px; font-weight: bold;
                }
                QPushButton:hover { background-color: #FDE68A; }
            """)
        else:
            self.star_btn.setText("☆ Star")
            self.star_btn.setStyleSheet("""
                QPushButton {
                    background-color: #F8FAFC; color: #64748B; border: 1px solid #CBD5E1;
                    border-radius: 4px; padding: 3px 8px; font-size: 11px;
                }
                QPushButton:hover { background-color: #F1F5F9; color: #1E293B; }
            """)

    def _update_follow_btn_style(self):
        if self.is_following:
            self.follow_btn.setText("🎯 Following")
            self.follow_btn.setStyleSheet("""
                QPushButton {
                    background-color: #DBEAFE; color: #1E40AF; border: 1px solid #3B82F6;
                    border-radius: 4px; padding: 3px 8px; font-size: 11px; font-weight: bold;
                }
                QPushButton:hover { background-color: #BFDBFE; }
            """)
        else:
            self.follow_btn.setText("🎯 Follow")
            self.follow_btn.setStyleSheet("""
                QPushButton {
                    background-color: #F8FAFC; color: #64748B; border: 1px solid #CBD5E1;
                    border-radius: 4px; padding: 3px 8px; font-size: 11px;
                }
                QPushButton:hover { background-color: #F1F5F9; color: #1E293B; }
            """)

    def _handle_toggle_star(self):
        self.is_starred = not self.is_starred
        self._update_star_btn_style()
        if self.on_toggle_star:
            self.on_toggle_star(str(self.details.get('code', '')).strip())

    def _handle_toggle_follow(self):
        self.is_following = not self.is_following
        self._update_follow_btn_style()
        if self.on_toggle_follow:
            self.on_toggle_follow(str(self.details.get('code', '')).strip())

    def _copy_citation(self):
        d = build_wi_copy_data(self.details)
        _copy_to_clipboard(d["cit_plain"], d["cit_html"], "WI Citation", self)

    def _make_key_label(self, text: str) -> QLabel:
        lbl = QLabel(f"<b>{text}</b>")
        lbl.setStyleSheet("color: #64748B; font-size: 12px;")
        return lbl

    def _add_row(self, form: QFormLayout, label_text: str, value_text: str):
        val_label = QLabel(value_text)
        val_label.setWordWrap(True)
        val_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        form.addRow(self._make_key_label(f"{label_text}:"), val_label)


class LinkedSpecsDialog(QDialog):
    """Dialog displaying all specifications linked to a given Work Item."""

    def __init__(self, wi_code: str, acronym: str, specs: list, parent=None):
        super().__init__(parent)
        title_str = f"Linked Specifications for WI {acronym} ({wi_code})" if acronym else f"Linked Specifications for WI #{wi_code}"
        self.setWindowTitle(title_str)
        self.setMinimumWidth(580)
        self.setMinimumHeight(380)
        self.setStyleSheet("""
            QFrame#specCard { background-color: #FFFFFF; border: 1px solid #CBD5E1; border-radius: 6px; }
            QFrame#specCard:hover { border-color: #94A3B8; background-color: #FAFCFF; }
            QLabel { color: #1E293B; }
            QScrollArea { border: 1px solid #CBD5E1; border-radius: 6px; background-color: #FFFFFF; }
        """)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)

        header_text = (
            f"Specifications Impacted by Work Item <b>{acronym}</b> ({wi_code}):"
            if acronym else f"Specifications Impacted by Work Item <b>#{wi_code}</b>:"
        )
        header = QLabel(header_text)
        header.setStyleSheet("font-size: 14px; color: #1E293B;")
        header.setTextInteractionFlags(Qt.TextSelectableByMouse)
        layout.addWidget(header)

        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)

        scroll_content = QWidget()
        cards_layout = QVBoxLayout(scroll_content)
        cards_layout.setContentsMargins(8, 8, 8, 8)
        cards_layout.setSpacing(8)

        if specs:
            for sp in specs:
                num = sp.get('number', 'Unknown')
                title = sp.get('title', 'No Title Available')
                spec_type = sp.get('type', 'TS')
                series_name = sp.get('series_name', num.split('.')[0] if '.' in num else '')
                is_pri = bool(sp.get('is_primary', False))

                card = self._create_spec_card(num, title, spec_type, series_name, is_pri)
                cards_layout.addWidget(card)
        else:
            no_data = QLabel("No linked specifications recorded in the local database for this Work Item.")
            no_data.setStyleSheet("color: #64748B; font-style: italic; padding: 25px;")
            no_data.setAlignment(Qt.AlignCenter)
            cards_layout.addWidget(no_data)

        cards_layout.addStretch()
        scroll_area.setWidget(scroll_content)
        layout.addWidget(scroll_area)

        footer_layout = QHBoxLayout()
        count_lbl = QLabel(f"<b>Total:</b> {len(specs)} specification(s)")
        count_lbl.setStyleSheet("color: #64748B; font-size: 12px;")
        footer_layout.addWidget(count_lbl)
        footer_layout.addStretch()

        close_btn = QPushButton("Close")
        close_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        close_btn.setCursor(Qt.PointingHandCursor)
        close_btn.clicked.connect(self.accept)
        footer_layout.addWidget(close_btn)

        layout.addLayout(footer_layout)

    def _create_spec_card(self, num: str, title: str, spec_type: str, series_name: str, is_primary: bool) -> QFrame:
        card = QFrame()
        card.setObjectName("specCard")
        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(12, 10, 12, 10)
        card_layout.setSpacing(6)

        top_row = QHBoxLayout()
        top_row.setSpacing(8)

        type_badge = QLabel(f"<b>{spec_type}</b>")
        type_badge.setStyleSheet("""
            background-color: #EBF3FC; color: #1E5C99; border: 1px solid #BFDBFE;
            border-radius: 4px; padding: 2px 6px; font-size: 11px; font-weight: bold;
        """)
        top_row.addWidget(type_badge)

        num_lbl = QLabel(f"<b>{num}</b>")
        num_lbl.setStyleSheet("font-size: 14px; font-weight: bold; color: #1E293B;")
        num_lbl.setTextInteractionFlags(Qt.TextSelectableByMouse)
        top_row.addWidget(num_lbl)

        if is_primary:
            pri_badge = QLabel("⭐ Primary Spec")
            pri_badge.setStyleSheet("""
                background-color: #FEF3C7; color: #92400E; border: 1px solid #FDE68A;
                border-radius: 4px; padding: 2px 6px; font-size: 11px; font-weight: bold;
            """)
            top_row.addWidget(pri_badge)

        top_row.addStretch()

        clean_num = num.split("-")[0].replace('.', '').strip()
        dyna_url = f"https://www.3gpp.org/DynaReport/{clean_num}.htm" if clean_num else ""
        ftp_url = f"https://www.3gpp.org/ftp/Specs/archive/{series_name}_series/{num}/" if series_name else ""

        if dyna_url:
            dyna_btn = QPushButton("🌐 DynaReport")
            dyna_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
            dyna_btn.setCursor(Qt.PointingHandCursor)
            dyna_btn.clicked.connect(lambda: webbrowser.open(dyna_url))
            top_row.addWidget(dyna_btn)

        if ftp_url:
            ftp_btn = QPushButton("📂 FTP Archive")
            ftp_btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
            ftp_btn.setCursor(Qt.PointingHandCursor)
            ftp_btn.clicked.connect(lambda: webbrowser.open(ftp_url))
            top_row.addWidget(ftp_btn)

        card_layout.addLayout(top_row)

        title_lbl = QLabel(title)
        title_lbl.setWordWrap(True)
        title_lbl.setStyleSheet("font-size: 12px; color: #475569; line-height: 1.4;")
        title_lbl.setTextInteractionFlags(Qt.TextSelectableByMouse)
        card_layout.addWidget(title_lbl)

        return card


class ManualWIFetcherThread(QThread):
    """
    Background worker that fetches and parses Work Item details directly from
    the 3GPP Portal without blocking the GUI thread.
    """
    fetch_finished = pyqtSignal(bool, dict, str)
    log_msg = pyqtSignal(str, int)

    def __init__(self, query: str):
        super().__init__()
        self.query = query.strip()

    def run(self):
        try:
            if self.isInterruptionRequested():
                return

            code = extract_wi_code(self.query)
            if not code:
                self.fetch_finished.emit(False, {}, "Could not determine a valid Work Item ID from the input.")
                return

            self.log_msg.emit(f"Connecting to 3GPP Portal for WI #{code}...", logging.INFO)
            metadata = fetch_wi_details_from_portal(code, log_cb=self.log_msg.emit)

            if self.isInterruptionRequested():
                return

            if metadata and (metadata.get("name") or metadata.get("acronym")):
                acr = metadata.get("acronym") or metadata.get("name")
                remarks_cnt = len(metadata.get("remarks", []))
                msg = f"Fetched metadata for WI #{code} ({acr}) successfully ({remarks_cnt} remarks)."
                self.fetch_finished.emit(True, metadata, msg)
            else:
                err_cause = metadata.get("error") if metadata else ""
                reason = f": {err_cause}" if err_cause else " (Check WI ID or network connection)"
                fail_msg = f"Could not extract details for WI #{code}{reason}"
                self.fetch_finished.emit(False, metadata or {}, fail_msg)
        except Exception as e:
            self.fetch_finished.emit(False, {}, f"Error fetching details: {e}")


class AddWorkItemDialog(QDialog):
    """Dialog allowing users to query, preview, manually edit, and register an individual Work Item."""

    def __init__(self, db: WorkItemsDatabase, initial_query: str = "", parent=None):
        super().__init__(parent)
        self.db = db
        self.fetch_thread: Optional[ManualWIFetcherThread] = None
        self.saved_wi_code = ""
        self.is_starred_requested = False
        self.is_following_requested = False
        self._fetched_remarks = []

        self.setWindowTitle("➕ Add / Fetch Work Item")
        self.setMinimumWidth(580)
        self.setStyleSheet("""
            QDialog { background-color: #FAFAFA; }
            QGroupBox {
                font-weight: bold;
                border: 1px solid #D0D0D0;
                border-radius: 6px;
                margin-top: 10px;
                padding-top: 10px;
                background-color: white;
            }
            QGroupBox::title { subcontrol-origin: margin; left: 10px; padding: 0 4px; }
            QLineEdit, QComboBox { padding: 5px; border: 1px solid #CCC; border-radius: 4px; }
            QLineEdit:focus, QComboBox:focus { border: 1px solid #0078D7; }
        """)

        self._setup_ui()

        if initial_query:
            self.query_input.setText(initial_query)
            self._start_fetch()

    def _setup_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setSpacing(12)

        # 1. Query Section
        query_group = QGroupBox("1. Query 3GPP Work Item")
        query_layout = QVBoxLayout(query_group)

        row_layout = QHBoxLayout()
        self.query_input = QLineEdit()
        self.query_input.setPlaceholderText("e.g. 1121123, #1121123, or https://portal.3gpp.org/...workitemId=1121123")
        self.query_input.setToolTip("Enter a 3GPP Work Item ID or Portal URL")
        self.query_input.returnPressed.connect(self._start_fetch)

        self.btn_fetch = QPushButton("🔍 Fetch Details")
        self.btn_fetch.setCursor(Qt.PointingHandCursor)
        self.btn_fetch.setStyleSheet("""
            QPushButton {
                background-color: #0078D7;
                color: white;
                font-weight: bold;
                padding: 6px 14px;
                border-radius: 4px;
            }
            QPushButton:hover { background-color: #005A9E; }
            QPushButton:disabled { background-color: #B0D0F0; }
        """)
        self.btn_fetch.clicked.connect(self._start_fetch)

        row_layout.addWidget(self.query_input, 1)
        row_layout.addWidget(self.btn_fetch)
        query_layout.addLayout(row_layout)

        self.lbl_status = QLabel("Enter a Work Item ID or URL and click 'Fetch Details'.")
        self.lbl_status.setStyleSheet("color: #64748B; font-size: 11px; margin-top: 2px;")
        self.lbl_status.setWordWrap(True)
        query_layout.addWidget(self.lbl_status)
        main_layout.addWidget(query_group)

        # 2. Form Preview (Editable)
        self.preview_group = QGroupBox("2. Work Item Details (Editable)")
        form = QFormLayout(self.preview_group)
        form.setLabelAlignment(Qt.AlignRight)

        self.edit_code = QLineEdit()
        self.edit_code.setPlaceholderText("e.g. 1121123")
        form.addRow("WI Code *:", self.edit_code)

        self.edit_acronym = QLineEdit()
        self.edit_acronym.setPlaceholderText("e.g. NR_Sensing_bis-Core")
        form.addRow("Acronym *:", self.edit_acronym)

        self.edit_name = QLineEdit()
        self.edit_name.setPlaceholderText("e.g. Core part: Integrated Sensing and communication (ISAC) for NR")
        form.addRow("Title / Name *:", self.edit_name)

        self.edit_release = QLineEdit()
        self.edit_release.setPlaceholderText("e.g. Rel-20, Rel-19, Rel-18...")
        form.addRow("Release:", self.edit_release)

        self.edit_groups = QLineEdit()
        self.edit_groups.setPlaceholderText("e.g. RAN3 (comma-separated for multiple)")
        form.addRow("Responsible WG(s):", self.edit_groups)

        self.edit_status = QLineEdit()
        self.edit_status.setPlaceholderText("e.g. PCG approved, In progress, Finished...")
        form.addRow("Status:", self.edit_status)

        self.edit_start_date = QLineEdit()
        self.edit_start_date.setPlaceholderText("YYYY-MM-DD")
        form.addRow("Start Date:", self.edit_start_date)

        self.edit_end_date = QLineEdit()
        self.edit_end_date.setPlaceholderText("YYYY-MM-DD")
        form.addRow("End Date:", self.edit_end_date)

        self.edit_latest_wid = QLineEdit()
        self.edit_latest_wid.setPlaceholderText("e.g. RP-261901")
        form.addRow("Latest WID:", self.edit_latest_wid)

        tracking_layout = QVBoxLayout()
        tracking_layout.setSpacing(4)
        self.chk_starred = QCheckBox("⭐ Add to Starred (Personal Watchlist)")
        self.chk_following = QCheckBox("🎯 Add to Following (Delegation Responsibilities)")
        tracking_layout.addWidget(self.chk_starred)
        tracking_layout.addWidget(self.chk_following)
        form.addRow("Tracking:", tracking_layout)

        main_layout.addWidget(self.preview_group)

        # 3. Actions
        btn_layout = QHBoxLayout()
        self.btn_save = QPushButton("💾 Save Work Item")
        self.btn_save.setCursor(Qt.PointingHandCursor)
        self.btn_save.setStyleSheet("""
            QPushButton {
                background-color: #107C41;
                color: white;
                font-weight: bold;
                padding: 7px 18px;
                border-radius: 4px;
            }
            QPushButton:hover { background-color: #0B5A30; }
        """)
        self.btn_save.clicked.connect(self._save_wi)

        self.btn_cancel = QPushButton("Cancel")
        self.btn_cancel.setCursor(Qt.PointingHandCursor)
        self.btn_cancel.clicked.connect(self.reject)

        btn_layout.addStretch()
        btn_layout.addWidget(self.btn_save)
        btn_layout.addWidget(self.btn_cancel)
        main_layout.addLayout(btn_layout)

    def _start_fetch(self):
        if self.fetch_thread and self.fetch_thread.isRunning():
            return

        query = self.query_input.text().strip()
        code = extract_wi_code(query)
        if not code:
            QMessageBox.warning(self, "Input Required", "Please enter a valid Work Item ID or 3GPP Portal URL.")
            return

        self.btn_fetch.setEnabled(False)
        self.btn_fetch.setText("⏳ Fetching...")
        self.lbl_status.setText(f"⏳ Connecting to 3GPP Portal for WI #{code}...")
        self.lbl_status.setStyleSheet("color: #0078D7; font-weight: bold;")

        self.fetch_thread = ManualWIFetcherThread(query)
        self.fetch_thread.log_msg.connect(self._on_fetch_log)
        self.fetch_thread.fetch_finished.connect(self._on_fetch_finished)
        self.fetch_thread.start()

    def _on_fetch_log(self, msg: str, level: int):
        self.lbl_status.setText(msg)
        if level >= logging.ERROR:
            self.lbl_status.setStyleSheet("color: #DC2626; font-weight: bold;")
        elif level >= logging.WARNING:
            self.lbl_status.setStyleSheet("color: #D97706; font-weight: bold;")
        else:
            self.lbl_status.setStyleSheet("color: #0078D7;")

    def _on_fetch_finished(self, success: bool, data: dict, msg: str):
        self.btn_fetch.setEnabled(True)
        self.btn_fetch.setText("🔍 Fetch Details")

        if success:
            self.lbl_status.setText(f"✅ {msg}")
            self.lbl_status.setStyleSheet("color: #107C41; font-weight: bold;")

            self.edit_code.setText(data.get("code", ""))
            self.edit_acronym.setText(data.get("acronym", ""))
            self.edit_name.setText(data.get("name", ""))
            self.edit_release.setText(data.get("release", ""))
            groups = data.get("responsible_groups", [])
            self.edit_groups.setText(", ".join(groups) if isinstance(groups, list) else str(groups))
            self.edit_status.setText(data.get("status", ""))
            self.edit_start_date.setText(data.get("start_date", ""))
            self.edit_end_date.setText(data.get("end_date", ""))
            self.edit_latest_wid.setText(data.get("latest_wid", ""))
            self._fetched_remarks = data.get("remarks", [])
        else:
            self.lbl_status.setText(f"⚠️ {msg}")
            self.lbl_status.setStyleSheet("color: #DC2626; font-weight: bold;")

    def _save_wi(self):
        code = self.edit_code.text().strip()
        acronym = self.edit_acronym.text().strip()
        name = self.edit_name.text().strip()

        if not code:
            QMessageBox.warning(self, "Validation Error", "Please provide a valid Work Item Code.")
            self.edit_code.setFocus()
            return

        if not acronym and not name:
            QMessageBox.warning(self, "Validation Error", "Please provide at least a Work Item Acronym or Title.")
            self.edit_acronym.setFocus()
            return

        groups_str = self.edit_groups.text().strip()
        groups = [g.strip() for g in re.split(r'[,;/|\s]+', groups_str) if g.strip()]

        wi_data = {
            "code": code,
            "acronym": acronym,
            "name": name,
            "release": self.edit_release.text().strip(),
            "latest_wid": self.edit_latest_wid.text().strip(),
            "start_date": self.edit_start_date.text().strip(),
            "end_date": self.edit_end_date.text().strip(),
            "responsible_groups": groups,
            "remarks": self._fetched_remarks
        }

        if self.db.upsert_single_work_item(wi_data):
            self.saved_wi_code = code
            self.is_starred_requested = self.chk_starred.isChecked()
            self.is_following_requested = self.chk_following.isChecked()
            self.accept()
        else:
            QMessageBox.critical(self, "Database Error", f"Failed to save Work Item #{code} to the database.")

    def closeEvent(self, event):
        if self.fetch_thread and self.fetch_thread.isRunning():
            self.fetch_thread.requestInterruption()
            self.fetch_thread.wait(1000)
        super().closeEvent(event)
