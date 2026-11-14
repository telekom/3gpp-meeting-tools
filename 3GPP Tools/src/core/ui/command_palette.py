# --- File: src/core/ui/command_palette.py ---
"""
Global Ctrl+K Command Palette & Quick Switcher for 3GPP Delegate Tools (Phase 3.4).

Provides zero-latency fuzzy/prefix switching across Tools, Word Workbenches,
Side Panels & Shell Actions, cached Specifications & Protocol Messages, and
indexed SQLite Meetings / TDocs / Work Items.
"""

from __future__ import annotations

import logging
import re
import sqlite3
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Dict, List, Optional, Sequence, Tuple

from PyQt5.QtCore import QEvent, QRect, QSize, Qt, pyqtSignal
from PyQt5.QtGui import QColor, QFont, QIcon, QKeySequence, QPen
from PyQt5.QtWidgets import (
    QApplication,
    QDialog,
    QFrame,
    QGraphicsDropShadowEffect,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QSizePolicy,
    QStyle,
    QStyledItemDelegate,
    QVBoxLayout,
    QWidget,
)

from core.ui.icons import lucide_icon
from core.ui.right_side_panel import MODE_ASSISTANT, MODE_SPEC_SEARCH
from core.ui.theme import C, R, S, T

logger = logging.getLogger(__name__)

_ROLE_KIND = Qt.UserRole + 20
_ROLE_CMD = Qt.UserRole + 21
_KIND_HEADER = "header"
_KIND_ITEM = "item"

_TDOC_PATTERN = re.compile(r"^[A-Z0-9]{2,4}-\d{6,8}$", re.IGNORECASE)


@dataclass
class CommandItem:
    """Single actionable entry rendered in the Ctrl+K Command Palette."""

    id: str
    category: str
    title: str
    subtitle: str = ""
    shortcut: str = ""
    badge: str = ""
    icon_name: str = "compass"
    callback: Optional[Callable[[], None]] = None
    keywords: str = ""


class _PaletteRowDelegate(QStyledItemDelegate):
    """Crisp custom delegate for group headers (26px) and action rows (36px)."""

    def __init__(self, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self._icon_cache: Dict[Tuple[str, bool], QIcon] = {}

    def _get_icon(self, name: str, active: bool) -> QIcon:
        key = (name or "compass", active)
        if key not in self._icon_cache:
            color = C.ACCENT_PRIMARY if active else C.TEXT_MUTED
            self._icon_cache[key] = lucide_icon(key[0], color=color, size=15)
        return self._icon_cache[key]

    def paint(self, painter, option, index) -> None:
        painter.save()
        kind = index.data(_ROLE_KIND)
        rect = option.rect

        if kind == _KIND_HEADER:
            painter.setPen(QColor(C.TEXT_FAINT))
            f = QFont(T.FONT_FAMILY_UI)
            f.setPixelSize(10)
            f.setWeight(QFont.Bold)
            painter.setFont(f)
            text = str(index.data(Qt.DisplayRole) or "").upper()
            painter.drawText(rect.adjusted(14, 4, -12, -2), Qt.AlignVCenter | Qt.AlignLeft, text)
            painter.restore()
            return

        cmd: Optional[CommandItem] = index.data(_ROLE_CMD)
        if cmd is None:
            painter.restore()
            return

        selected = bool(option.state & QStyle.State_Selected)
        hovered = bool(option.state & QStyle.State_MouseOver)

        row_rect = rect.adjusted(6, 2, -6, -2)
        if selected:
            painter.setBrush(QColor(C.ACCENT_SOFT))
            painter.setPen(Qt.NoPen)
            painter.drawRoundedRect(row_rect, R.MD, R.MD)
        elif hovered:
            painter.setBrush(QColor(C.BG_SURFACE))
            painter.setPen(Qt.NoPen)
            painter.drawRoundedRect(row_rect, R.MD, R.MD)

        # Icon on left
        icon = self._get_icon(cmd.icon_name, selected)
        icon_rect = QRect(row_rect.left() + 8, row_rect.top() + (row_rect.height() - 16) // 2, 16, 16)
        icon.paint(painter, icon_rect, Qt.AlignCenter)

        # Right-side shortcut / badge pill width calculation
        right_x = row_rect.right() - 8
        pill_text = cmd.shortcut or cmd.badge or ""
        if pill_text:
            pf = QFont(T.FONT_FAMILY_MONO if cmd.shortcut else T.FONT_FAMILY_UI)
            pf.setPixelSize(11)
            pf.setWeight(QFont.DemiBold)
            painter.setFont(pf)
            fm_p = painter.fontMetrics()
            pw = fm_p.horizontalAdvance(pill_text) + 12
            ph = 20
            pill_rect = QRect(right_x - pw, row_rect.top() + (row_rect.height() - ph) // 2, pw, ph)
            if cmd.shortcut:
                painter.setBrush(QColor(C.BG_SURFACE))
                painter.setPen(QPen(QColor(C.BORDER_SUBTLE), 1))
            else:
                painter.setBrush(QColor(C.ACCENT_SOFT if selected else C.BG_SURFACE))
                painter.setPen(QPen(QColor(C.BORDER_SUBTLE), 1))
            painter.drawRoundedRect(pill_rect, R.SM, R.SM)
            painter.setPen(QColor(C.ACCENT_PRIMARY if selected and not cmd.shortcut else C.TEXT_MUTED))
            painter.drawText(pill_rect, Qt.AlignCenter, pill_text)
            right_x = pill_rect.left() - 8

        text_left = icon_rect.right() + 10
        avail_w = max(80, right_x - text_left)

        # Title + inline muted subtitle
        tf = QFont(T.FONT_FAMILY_UI)
        tf.setPixelSize(T.SIZE_BODY)
        tf.setWeight(QFont.DemiBold if selected else QFont.Medium)
        painter.setFont(tf)
        fm_t = painter.fontMetrics()
        title_elided = fm_t.elidedText(cmd.title, Qt.ElideRight, avail_w)
        painter.setPen(QColor(C.ACCENT_PRIMARY if selected else C.TEXT_PRIMARY))
        painter.drawText(
            QRect(text_left, row_rect.top(), avail_w, row_rect.height()),
            Qt.AlignVCenter | Qt.AlignLeft,
            title_elided,
        )

        title_w = fm_t.horizontalAdvance(title_elided)
        rem_w = avail_w - title_w - 10
        if cmd.subtitle and rem_w > 48:
            sf = QFont(T.FONT_FAMILY_UI)
            sf.setPixelSize(T.SIZE_CAPTION)
            painter.setFont(sf)
            fm_s = painter.fontMetrics()
            sub_elided = fm_s.elidedText(cmd.subtitle, Qt.ElideRight, rem_w)
            painter.setPen(QColor(C.TEXT_MUTED))
            painter.drawText(
                QRect(text_left + title_w + 10, row_rect.top(), rem_w, row_rect.height()),
                Qt.AlignVCenter | Qt.AlignLeft,
                sub_elided,
            )

        painter.restore()

    def sizeHint(self, option, index) -> QSize:
        kind = index.data(_ROLE_KIND)
        if kind == _KIND_HEADER:
            return QSize(option.rect.width(), 24)
        return QSize(option.rect.width(), 36)


def _score_command(item: CommandItem, query: str) -> Optional[Tuple[int, str]]:
    """Return a sort rank (lower is better) if item matches query tokens."""
    tokens = [t for t in query.lower().split() if t]
    if not tokens:
        return (0, item.title.lower())

    title_l = item.title.lower()
    sub_l = item.subtitle.lower()
    kw_l = item.keywords.lower()
    cat_l = item.category.lower()
    haystack = f"{title_l} {sub_l} {kw_l} {cat_l}"

    if not all(tok in haystack for tok in tokens):
        return None

    first = tokens[0]
    if title_l == query.lower():
        tier = 0
    elif title_l.startswith(first):
        tier = 1
    elif first in title_l:
        tier = 2
    elif first in kw_l:
        tier = 3
    else:
        tier = 4
    return (tier, title_l)


class DynamicPaletteSearcher:
    """Fast local SQLite + cached metadata provider for CommandPaletteDialog."""

    def __init__(
        self,
        *,
        db_path: Optional[Path] = None,
        protocol_db_path: Optional[Path] = None,
        on_open_spec: Optional[Callable[[str], None]] = None,
        on_open_meeting: Optional[Callable[[str], None]] = None,
        on_open_work_item: Optional[Callable[[str], None]] = None,
        on_open_protocol: Optional[Callable[[str, str], None]] = None,
    ) -> None:
        self.db_path = Path(db_path) if db_path else None
        self.protocol_db_path = Path(protocol_db_path) if protocol_db_path else None
        self.on_open_spec = on_open_spec
        self.on_open_meeting = on_open_meeting
        self.on_open_work_item = on_open_work_item
        self.on_open_protocol = on_open_protocol

        self._spec_titles_cache: Optional[List[Tuple[str, str, str]]] = None
        self._protocol_cache: Optional[List[Tuple[str, str, str]]] = None

    def invalidate_cache(self) -> None:
        self._spec_titles_cache = None
        self._protocol_cache = None

    def _ensure_spec_cache(self) -> List[Tuple[str, str, str]]:
        if self._spec_titles_cache is not None:
            return self._spec_titles_cache
        items: List[Tuple[str, str, str]] = []
        if self.db_path and self.db_path.exists():
            try:
                from modules.specifications.core.database import SpecsDatabase

                db = SpecsDatabase(self.db_path)
                titles = db.get_all_spec_titles() or {}
                # Fetch optional WG metadata cleanly in one lightweight read
                wg_map: Dict[str, str] = {}
                try:
                    with db._get_connection() as conn:
                        cur = conn.cursor()
                        cur.execute(
                            """
                            SELECT sp.number, COALESCE(wg.name, '')
                            FROM specifications sp
                            LEFT JOIN working_groups wg ON sp.primary_group_id = wg.id
                            """
                        )
                        for row in cur.fetchall():
                            s_num = str(row[0] or "").strip()
                            wg = str(row[1] or "").strip()
                            if s_num and wg:
                                wg_map[s_num] = wg
                except Exception:
                    pass
                for spec_num, title in titles.items():
                    items.append((str(spec_num), str(title or ""), wg_map.get(str(spec_num), "")))
            except Exception as exc:
                logger.debug("Palette failed loading spec title cache: %s", exc)
        self._spec_titles_cache = items
        return items

    def _ensure_protocol_cache(self) -> List[Tuple[str, str, str]]:
        if self._protocol_cache is not None:
            return self._protocol_cache
        items: List[Tuple[str, str, str]] = []
        if self.protocol_db_path and self.protocol_db_path.exists():
            try:
                conn = sqlite3.connect(str(self.protocol_db_path), timeout=1.5)
                try:
                    cur = conn.cursor()
                    cur.execute(
                        "SELECT name FROM sqlite_master WHERE type='table' AND name='nas_messages'"
                    )
                    if cur.fetchone():
                        cur.execute(
                            """
                            SELECT DISTINCT m.message_name, COALESCE(v.spec_type, 'NAS'), COALESCE(v.spec_number, '')
                            FROM nas_messages m
                            JOIN spec_versions v ON m.version_id = v.id
                            ORDER BY m.message_name ASC
                            """
                        )
                        for row in cur.fetchall():
                            name = str(row[0] or "").strip()
                            if name:
                                items.append((name, str(row[1] or "NAS"), str(row[2] or "")))
                finally:
                    conn.close()
            except Exception as exc:
                logger.debug("Palette failed loading protocol cache: %s", exc)
        self._protocol_cache = items
        return items

    def search(self, query: str) -> List[CommandItem]:
        q = (query or "").strip()
        if len(q) < 2:
            return []

        results: List[CommandItem] = []
        q_low = q.lower()
        clean_spec_q = re.sub(r"^(?:ts|tr)\s*", "", q_low).strip()

        # 1. Cached Specifications (up to 8 matches)
        spec_matches: List[Tuple[int, str, str, str]] = []
        for spec_num, title, wg in self._ensure_spec_cache():
            s_low = spec_num.lower()
            t_low = title.lower()
            if clean_spec_q and clean_spec_q in s_low:
                rank = 0 if s_low.startswith(clean_spec_q) else 1
            elif q_low in t_low or (wg and q_low in wg.lower()):
                rank = 2
            else:
                continue
            spec_matches.append((rank, spec_num, title, wg))

        spec_matches.sort(key=lambda x: (x[0], x[1]))
        for _rk, spec_num, title, wg in spec_matches[:8]:
            results.append(
                CommandItem(
                    id=f"spec:{spec_num}",
                    category="Specifications",
                    title=f"TS {spec_num}",
                    subtitle=title,
                    badge=wg,
                    icon_name="book-open",
                    callback=(lambda s=spec_num: self.on_open_spec(s)) if self.on_open_spec else None,
                    keywords=f"{spec_num} {title} {wg}",
                )
            )

        # 2. Live SQLite Meetings & TDoc jump (up to 6 matches)
        if self.db_path and self.db_path.exists():
            try:
                from modules.meetings.core.meetings_db import MeetingsDatabase

                mdb = MeetingsDatabase(self.db_path)
                if _TDOC_PATTERN.match(q):
                    tdoc_up = q.upper()
                    owning = mdb.find_meeting_by_tdoc(tdoc_up)
                    sub = (
                        f"{owning.get('wg_name', '')} · {owning.get('meeting_number', '')} ({owning.get('location', '')})"
                        if owning
                        else "Jump to TDoc in Meetings tab"
                    )
                    results.append(
                        CommandItem(
                            id=f"tdoc:{tdoc_up}",
                            category="Meetings & TDocs",
                            title=f"TDoc {tdoc_up}",
                            subtitle=sub,
                            badge="TDoc",
                            icon_name="file-text",
                            callback=(lambda t=tdoc_up: self.on_open_meeting(t)) if self.on_open_meeting else None,
                            keywords=tdoc_up,
                        )
                    )

                meetings = mdb.search_meetings(search_term=q) or []
                for m in meetings[:6]:
                    m_num = str(m.get("meeting_number") or m.get("name") or "").strip()
                    wg = str(m.get("wg_name") or "").strip()
                    loc = str(m.get("location") or "").strip()
                    s_date = str(m.get("start_date") or "").strip()
                    label = f"{wg} {m_num}".strip() if wg and wg not in m_num else m_num
                    if not label:
                        continue
                    sub_parts = [p for p in (loc, s_date) if p]
                    results.append(
                        CommandItem(
                            id=f"mtg:{label}",
                            category="Meetings & TDocs",
                            title=label,
                            subtitle=" · ".join(sub_parts),
                            badge=wg,
                            icon_name="calendar",
                            callback=(lambda key=m_num or label: self.on_open_meeting(key))
                            if self.on_open_meeting
                            else None,
                            keywords=f"{label} {loc} {s_date}",
                        )
                    )
            except Exception as exc:
                logger.debug("Palette failed querying meetings: %s", exc)

            # 3. Live SQLite Work Items (up to 6 matches)
            try:
                from modules.work_items.core.wi_database import WorkItemsDatabase

                wdb = WorkItemsDatabase(self.db_path)
                wis = wdb.search_work_item_suggestions(q, limit=6) or []
                for wi in wis[:6]:
                    acronym = str(wi.get("acronym") or "").strip()
                    code = str(wi.get("code") or "").strip()
                    name = str(wi.get("name") or "").strip()
                    results.append(
                        CommandItem(
                            id=f"wi:{code or acronym}",
                            category="Work Items",
                            title=f"{acronym} ({code})" if code else acronym,
                            subtitle=name,
                            badge="WI",
                            icon_name="clipboard-list",
                            callback=(lambda key=acronym or code: self.on_open_work_item(key))
                            if self.on_open_work_item
                            else None,
                            keywords=f"{acronym} {code} {name}",
                        )
                    )
            except Exception as exc:
                logger.debug("Palette failed querying work items: %s", exc)

        # 4. Cached Protocol Messages (up to 6 matches)
        proto_hits = [
            row
            for row in self._ensure_protocol_cache()
            if q_low in row[0].lower() or q_low in row[1].lower()
        ]
        for msg_name, ptype, spec_num in proto_hits[:6]:
            sub = f"{ptype} · TS {spec_num}" if spec_num else ptype
            results.append(
                CommandItem(
                    id=f"proto:{ptype}:{msg_name}",
                    category="Protocols",
                    title=msg_name,
                    subtitle=sub,
                    badge=ptype,
                    icon_name="network",
                    callback=(lambda p=ptype, m=msg_name: self.on_open_protocol(p, m))
                    if self.on_open_protocol
                    else None,
                    keywords=f"{msg_name} {ptype} {spec_num}",
                )
            )

        return results


def build_static_command_items(window: object) -> List[CommandItem]:
    """Assemble Tools, Word Workbenches, and Shell/Action commands for ``window``."""
    items: List[CommandItem] = []

    def _switch_tab(idx: int) -> None:
        tabs = getattr(window, "tabs", None)
        if tabs is not None and 0 <= idx < tabs.count():
            tabs.setCurrentIndex(idx)

    # 1. Core Tools (Ctrl+1..Ctrl+9)
    tools = [
        ("tool:plantuml", "PlantUML Studio", "Sequence diagrams, SVG live preview & AI generator", "Ctrl+1", "code", 0),
        ("tool:visio", "Visio & Office Studio", "Batch .puml conversion, EMF bridge & 3GPP stencils", "Ctrl+2", "layers", 1),
        ("tool:word", "Word Tools", "Compare docs, Markdown studio, CR/pCR & Sanitizer", "Ctrl+3", "file-text", 2),
        ("tool:specs", "Specifications Archive", "3GPP TS/TR browser, release downloads & change history", "Ctrl+4", "book-open", 3),
        ("tool:spec_search", "Spec Search", "Full-text FTS5 clause search & version comparison", "Ctrl+5", "search", 4),
        ("tool:protocols", "Protocol Explorer", "NAS, RRC, NGAP & PFCP evolution matrix", "Ctrl+6", "network", 5),
        ("tool:meetings", "Meetings & TDocs", "Agenda browser, TDocs, agreements & Session Plan", "Ctrl+7", "calendar", 6),
        ("tool:work_items", "Work Items Tracker", "Feature & WI search, tracking & linked specs", "Ctrl+8", "clipboard-list", 7),
        ("tool:assistant", "3GPP Research Assistant", "Local evidence-grounded Ollama AI assistant", "Ctrl+9", "sparkles", 8),
    ]
    for cid, title, sub, sc, icon, idx in tools:
        items.append(
            CommandItem(
                id=cid,
                category="Tools",
                title=title,
                subtitle=sub,
                shortcut=sc,
                icon_name=icon,
                callback=lambda i=idx: _switch_tab(i),
                keywords=f"tab switch goto {title} {sub}",
            )
        )

    # 2. Side Panels & Shell Actions
    def _dock(mode: str) -> None:
        if hasattr(window, "open_right_panel"):
            window.open_right_panel(mode)

    def _open_word_workbench(wb_key: str) -> None:
        _switch_tab(2)
        wt = getattr(window, "word_tab", None)
        if wt is not None and hasattr(wt, "open_workbench"):
            wt.open_workbench(wb_key)

    items.extend([
        CommandItem(
            id="action:dock_assistant",
            category="Side Panels & Actions",
            title="Dock Assistant Side Panel",
            subtitle="Open compact Assistant beside current workspace",
            shortcut="Ctrl+Shift+A",
            icon_name="sparkles",
            callback=lambda: _dock(MODE_ASSISTANT),
            keywords="side panel right dock ai assistant",
        ),
        CommandItem(
            id="action:dock_search",
            category="Side Panels & Actions",
            title="Dock Quick Spec Search",
            subtitle="Open compact clause search beside current workspace",
            shortcut="Ctrl+Shift+F",
            icon_name="search",
            callback=lambda: _dock(MODE_SPEC_SEARCH),
            keywords="side panel right dock quick spec search",
        ),
        CommandItem(
            id="action:toggle_drawer",
            category="Side Panels & Actions",
            title="Toggle Console & Queue Drawer",
            subtitle="Show or collapse bottom system output and task queue",
            shortcut="Ctrl+J",
            icon_name="terminal",
            callback=lambda: getattr(window, "toggle_bottom_drawer", lambda: None)(),
            keywords="logs output queue drawer console",
        ),
        CommandItem(
            id="action:toggle_rail",
            category="Side Panels & Actions",
            title="Toggle Navigation Sidebar",
            subtitle="Expand or collapse left icon navigation rail",
            shortcut="Ctrl+\\",
            icon_name="panel-left-close",
            callback=lambda: getattr(getattr(window, "nav_rail", None), "toggle_collapsed", lambda: None)(),
            keywords="sidebar nav rail collapse expand",
        ),
    ])

    # 3. Word Workbenches (7)
    workbenches = [
        ("wb:compare", "Word: Compare Documents", "Native Track Changes Word COM diff", "git-compare", "compare"),
        ("wb:markdown", "Word: Markdown <-> Word", "Bidirectional docx/pptx/pdf to Markdown", "file-code", "markdown"),
        ("wb:cr_pcr", "Word: 3GPP CR & pCR Workbench", "Draft Change Request / Pseudo-CR contributions", "file-plus", "cr_pcr"),
        ("wb:sanitizer", "Word: Sanitizer & Revisions", "Proofing language & Track Changes author cleanup", "sliders", "sanitizer"),
        ("wb:excerpt", "Word: Slice & Excerpt", "Clause slicing & heading tree extraction", "scissors", "excerpt"),
        ("wb:converter", "Word: Format Converter", "DOC / DOCX / PDF / HTML conversion", "refresh-cw", "converter"),
        ("wb:visio", "Word: Visio & Media Extractor", "Extract embedded Visio & image assets", "image", "visio"),
    ]
    for cid, title, sub, icon, wb_key in workbenches:
        items.append(
            CommandItem(
                id=cid,
                category="Word Workbenches",
                title=title,
                subtitle=sub,
                icon_name=icon,
                callback=lambda k=wb_key: _open_word_workbench(k),
                keywords=f"word workbench {title} {sub}",
            )
        )

    # 4. Sync & System Dialogs
    def _sync_meetings() -> None:
        _switch_tab(6)
        mt = getattr(window, "meetings_tab", None)
        if mt is not None and hasattr(mt, "update_db_requested"):
            mt.update_db_requested.emit(True, True, False)

    def _sync_work_items() -> None:
        _switch_tab(7)
        wt = getattr(window, "work_items_tab", None)
        if wt is not None and hasattr(wt, "sync_btn"):
            wt.sync_btn.click()

    sys_actions = [
        ("sys:sync_meetings", "Sync 3GPP Meetings Database", "Fetch latest WG meetings and TDocs metadata", "refresh-cw", _sync_meetings),
        ("sys:sync_wi", "Sync 3GPP Work Items", "Refresh Work Items catalog from portal", "refresh-cw", _sync_work_items),
        ("sys:ollama", "Ollama AI Settings", "Configure local model host, timeout & model selection", "sparkles", getattr(window, "open_ollama_settings", None)),
        ("sys:proxy", "Proxy Settings", "Configure HTTP/HTTPS proxy & corporate bypass", "shield", getattr(window, "open_proxy_settings", None)),
        ("sys:db_maint", "Database Maintenance", "Vacuum, inspect & repair SQLite databases", "database", getattr(window, "open_db_maintenance", None)),
        ("sys:task_mgr", "COM Task Manager", "Inspect Word/Visio/PPT background processes", "activity", getattr(window, "open_task_manager", None)),
        ("sys:java", "Java & PlantUML Engine", "Verify Java runtime & PlantUML JAR status", "cpu", getattr(window, "open_java_maintenance", None)),
        ("sys:mcp", "MCP Server & AI Cockpit", "Inspect stdio MCP tools and agent activity", "terminal", getattr(window, "open_mcp_monitor", None)),
        ("sys:check_updates", "Check for App Updates", "Check GitHub master for new changes and update", "download", getattr(window, "open_app_updater", None)),
    ]
    for cid, title, sub, icon, fn in sys_actions:
        if fn is None:
            continue
        items.append(
            CommandItem(
                id=cid,
                category="Side Panels & Actions",
                title=title,
                subtitle=sub,
                icon_name=icon,
                callback=fn,
                keywords=f"system settings dialog {title} {sub}",
            )
        )

    return items


class CommandPaletteDialog(QDialog):
    """Modal Ctrl+K command palette with instant filter + arrow-key execution."""

    item_executed = pyqtSignal(object)  # CommandItem

    def __init__(
        self,
        static_items: Sequence[CommandItem],
        dynamic_searcher: Optional[Callable[[str], Sequence[CommandItem]]] = None,
        parent: Optional[QWidget] = None,
    ) -> None:
        super().__init__(parent)
        self.setObjectName("commandPaletteDialog")
        self.setWindowFlags(Qt.Dialog | Qt.FramelessWindowHint)
        self.setAttribute(Qt.WA_TranslucentBackground, True)
        self.setModal(True)

        self._static_items = list(static_items)
        self._dynamic_searcher = dynamic_searcher

        self._setup_ui()
        self._apply_styles()
        self.populate_results("")
        self._center_on_parent()

    def _setup_ui(self) -> None:
        outer = QVBoxLayout(self)
        outer.setContentsMargins(16, 16, 16, 20)

        self.card = QFrame(self)
        self.card.setObjectName("cmdPaletteCard")
        self.card.setFixedWidth(580)
        shadow = QGraphicsDropShadowEffect(self.card)
        shadow.setBlurRadius(28)
        shadow.setOffset(0, 8)
        shadow.setColor(QColor(0, 0, 0, 55))
        self.card.setGraphicsEffect(shadow)

        card_layout = QVBoxLayout(self.card)
        card_layout.setContentsMargins(0, 0, 0, 0)
        card_layout.setSpacing(0)

        # --- Top Search Bar Row (50px) ---
        search_bar = QFrame(self.card)
        search_bar.setObjectName("cmdPaletteSearchRow")
        search_bar.setFixedHeight(50)
        sb_layout = QHBoxLayout(search_bar)
        sb_layout.setContentsMargins(14, 0, 14, 0)
        sb_layout.setSpacing(10)

        icon_lbl = QLabel(search_bar)
        icon_lbl.setPixmap(lucide_icon("search", color=C.ACCENT_PRIMARY, size=17).pixmap(17, 17))
        sb_layout.addWidget(icon_lbl)

        self.search_input = QLineEdit(search_bar)
        self.search_input.setObjectName("cmdPaletteInput")
        self.search_input.setPlaceholderText(
            "Jump to tool, spec (23.501), meeting, TDoc (S2-…), WI, or action…"
        )
        self.search_input.setClearButtonEnabled(True)
        self.search_input.installEventFilter(self)
        self.search_input.textChanged.connect(self.populate_results)
        sb_layout.addWidget(self.search_input, 1)

        esc_badge = QLabel("Esc", search_bar)
        esc_badge.setObjectName("cmdPaletteEscBadge")
        sb_layout.addWidget(esc_badge)
        card_layout.addWidget(search_bar)

        # --- Results List ---
        self.results_list = QListWidget(self.card)
        self.results_list.setObjectName("cmdPaletteList")
        self.results_list.setItemDelegate(_PaletteRowDelegate(self.results_list))
        self.results_list.setVerticalScrollMode(QListWidget.ScrollPerPixel)
        self.results_list.setMaximumHeight(380)
        self.results_list.setMinimumHeight(180)
        self.results_list.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.results_list.itemActivated.connect(self._on_item_activated)
        self.results_list.itemClicked.connect(self._on_item_activated)
        card_layout.addWidget(self.results_list, 1)

        # --- Footer Hint Strip ---
        footer = QFrame(self.card)
        footer.setObjectName("cmdPaletteFooter")
        footer.setFixedHeight(32)
        f_layout = QHBoxLayout(footer)
        f_layout.setContentsMargins(14, 0, 14, 0)
        f_layout.setSpacing(12)

        hint_lbl = QLabel("↑↓ Navigate   ·   ↵ Open   ·   Esc Dismiss", footer)
        hint_lbl.setObjectName("cmdPaletteHintLabel")
        f_layout.addWidget(hint_lbl)
        f_layout.addStretch(1)

        self.count_label = QLabel("", footer)
        self.count_label.setObjectName("cmdPaletteHintLabel")
        f_layout.addWidget(self.count_label)
        card_layout.addWidget(footer)

        outer.addWidget(self.card)

    def _apply_styles(self) -> None:
        self.setStyleSheet(f"""
            QFrame#cmdPaletteCard {{
                background-color: {C.BG_CARD};
                border: 1px solid {C.BORDER_STRONG};
                border-radius: {R.XL}px;
            }}
            QFrame#cmdPaletteSearchRow {{
                background-color: {C.BG_CARD};
                border-bottom: 1px solid {C.BORDER_SUBTLE};
                border-top-left-radius: {R.XL}px;
                border-top-right-radius: {R.XL}px;
            }}
            QLineEdit#cmdPaletteInput {{
                background-color: transparent;
                border: none;
                color: {C.TEXT_PRIMARY};
                font-family: {T.FONT_FAMILY_UI};
                font-size: {T.SIZE_SECTION_TITLE}px;
                padding: 4px 0px;
            }}
            QLabel#cmdPaletteEscBadge {{
                background-color: {C.BG_SURFACE};
                color: {C.TEXT_MUTED};
                border: 1px solid {C.BORDER_SUBTLE};
                border-radius: {R.SM}px;
                padding: 2px 6px;
                font-family: {T.FONT_FAMILY_MONO};
                font-size: 11px;
                font-weight: 600;
            }}
            QListWidget#cmdPaletteList {{
                background-color: {C.BG_CARD};
                border: none;
                outline: none;
                padding: 4px 0px;
            }}
            QFrame#cmdPaletteFooter {{
                background-color: {C.BG_SURFACE};
                border-top: 1px solid {C.BORDER_SUBTLE};
                border-bottom-left-radius: {R.XL}px;
                border-bottom-right-radius: {R.XL}px;
            }}
            QLabel#cmdPaletteHintLabel {{
                color: {C.TEXT_MUTED};
                font-family: {T.FONT_FAMILY_UI};
                font-size: {T.SIZE_CAPTION}px;
            }}
        """)

    def _center_on_parent(self) -> None:
        parent = self.parentWidget()
        self.adjustSize()
        if parent is not None and parent.isVisible():
            pg = parent.frameGeometry()
            x = pg.center().x() - self.width() // 2
            y = pg.top() + max(48, int(pg.height() * 0.14))
            self.move(x, y)

    def showEvent(self, event) -> None:
        super().showEvent(event)
        self._center_on_parent()
        self.search_input.setFocus()
        self.search_input.selectAll()

    def populate_results(self, query: str) -> None:
        q = (query or "").strip()
        self.results_list.clear()

        matched_static: List[Tuple[Tuple[int, str], CommandItem]] = []
        for item in self._static_items:
            rank = _score_command(item, q)
            if rank is not None:
                matched_static.append((rank, item))

        if q:
            matched_static.sort(key=lambda pair: pair[0])

        grouped: Dict[str, List[CommandItem]] = {}
        for _rank, item in matched_static:
            grouped.setdefault(item.category, []).append(item)

        if self._dynamic_searcher is not None and len(q) >= 2:
            try:
                dyn = self._dynamic_searcher(q) or []
            except Exception as exc:
                logger.debug("Command palette dynamic provider error: %s", exc)
                dyn = []
            for item in dyn:
                grouped.setdefault(item.category, []).append(item)

        # Render groups in clean predictable order
        preferred_order = [
            "Tools",
            "Side Panels & Actions",
            "Specifications",
            "Meetings & TDocs",
            "Work Items",
            "Protocols",
            "Word Workbenches",
        ]
        ordered_cats = [c for c in preferred_order if c in grouped] + [
            c for c in grouped if c not in preferred_order
        ]

        total_items = 0
        first_selectable_row = -1
        for cat in ordered_cats:
            cat_items = grouped[cat]
            if not q and cat == "Word Workbenches":
                cat_items = cat_items[:4]
            if not cat_items:
                continue

            hdr = QListWidgetItem(cat)
            hdr.setFlags(Qt.NoItemFlags)
            hdr.setData(_ROLE_KIND, _KIND_HEADER)
            self.results_list.addItem(hdr)

            for item in cat_items:
                row_item = QListWidgetItem(item.title)
                row_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
                row_item.setData(_ROLE_KIND, _KIND_ITEM)
                row_item.setData(_ROLE_CMD, item)
                self.results_list.addItem(row_item)
                if first_selectable_row < 0:
                    first_selectable_row = self.results_list.count() - 1
                total_items += 1

        if first_selectable_row >= 0:
            self.results_list.setCurrentRow(first_selectable_row)
            self.count_label.setText(f"{total_items} command{'s' if total_items != 1 else ''}")
        else:
            empty_row = QListWidgetItem("No matching commands, specs, meetings, or work items")
            empty_row.setFlags(Qt.NoItemFlags)
            empty_row.setData(_ROLE_KIND, _KIND_HEADER)
            self.results_list.addItem(empty_row)
            self.count_label.setText("0 matches")

    def _move_selection(self, delta: int) -> None:
        count = self.results_list.count()
        if count == 0:
            return
        cur = self.results_list.currentRow()
        idx = cur
        for _ in range(count):
            idx = (idx + delta) % count
            it = self.results_list.item(idx)
            if it is not None and bool(it.flags() & Qt.ItemIsSelectable):
                self.results_list.setCurrentRow(idx)
                return

    def _execute_current(self) -> bool:
        it = self.results_list.currentItem()
        if it is None:
            return False
        return self._on_item_activated(it)

    def _on_item_activated(self, item: QListWidgetItem) -> bool:
        if item is None or item.data(_ROLE_KIND) != _KIND_ITEM:
            return False
        cmd: Optional[CommandItem] = item.data(_ROLE_CMD)
        if cmd is None:
            return False
        self.accept()
        self.item_executed.emit(cmd)
        if callable(cmd.callback):
            cmd.callback()
        return True

    def eventFilter(self, watched, event) -> bool:
        if watched is self.search_input and event.type() == QEvent.KeyPress:
            key = event.key()
            if key == Qt.Key_Down:
                self._move_selection(1)
                return True
            if key == Qt.Key_Up:
                self._move_selection(-1)
                return True
            if key in (Qt.Key_Return, Qt.Key_Enter):
                self._execute_current()
                return True
            if key == Qt.Key_Escape:
                self.reject()
                return True
        return super().eventFilter(watched, event)
