# --- File: src/core/ui/ui_components.py ---
import urllib.parse
from pathlib import Path
from typing import Optional, Tuple

from PyQt5.QtCore import QEvent, QObject, QSize, Qt, QThread, pyqtSignal
from PyQt5.QtGui import QColor, QIcon, QPainter, QPen, QPixmap
from PyQt5.QtWidgets import (
    QCheckBox,
    QDialog,
    QFormLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from core.network.session import NetworkSession, ProxyProfileManager
from core.ui.theme import C, R, T
from core.utils.dpapi import SecureCredentialStore, redact_sensitive_urls

# ==========================================
# --- GLOBAL STYLESHEET (ALL-BLUE THEME) ---
# ==========================================
GLOBAL_STYLE = f"""
    QWidget {{
        font-family: {T.FONT_FAMILY_UI};
        font-size: {T.SIZE_BODY}px;
        color: {C.TEXT_BODY};
    }}
    QToolTip {{
        color: {C.TEXT_BODY};
        background-color: {C.BG_SURFACE};
        border: 1px solid {C.BORDER_STRONG};
        border-radius: {R.SM}px;
        padding: 4px;
    }}
    QTabWidget::pane {{
        border: 1px solid {C.BORDER_STRONG};
        border-radius: {R.LG}px;
        background: {C.BG_CARD};
        top: -1px;
    }}
    QTabBar::tab {{
        background: {C.BG_CANVAS};
        border: 1px solid {C.BORDER_STRONG};
        padding: 8px 16px;
        margin-right: 2px;
        border-top-left-radius: {R.MD}px;
        border-top-right-radius: {R.MD}px;
    }}
    QTabBar::tab:selected {{
        background: {C.BG_CARD};
        border-bottom-color: {C.BG_CARD};
        font-weight: bold;
        color: {C.ACCENT_PRIMARY};
    }}
    QTabBar::tab:hover:!selected {{
        background: {C.BG_SURFACE};
    }}
    QPushButton {{
        padding: 8px 16px;
        border-radius: {R.MD}px;
        border: 1px solid {C.BORDER_STRONG};
        background-color: {C.BG_SURFACE};
        font-weight: bold;
    }}
    QPushButton:hover {{
        background-color: {C.BG_CANVAS};
    }}
    QPushButton:disabled {{
        background-color: {C.BG_SURFACE};
        color: {C.TEXT_FAINT};
        border: 1px solid {C.BORDER_SUBTLE};
    }}
    QPushButton:checked {{
        background-color: {C.ACCENT_SOFT};
        border: 2px solid {C.ACCENT_PRIMARY};
        color: {C.ACCENT_PRIMARY};
    }}
    QPushButton#primaryBtn, QPushButton#pptBtn, QPushButton#svgBtn {{
        background-color: {C.ACCENT_PRIMARY};
        color: {C.TEXT_ON_ACCENT};
        border: none;
    }}
    QPushButton#primaryBtn:hover, QPushButton#pptBtn:hover, QPushButton#svgBtn:hover {{
        background-color: {C.ACCENT_HOVER};
    }}
    QPushButton#primaryBtn:disabled, QPushButton#pptBtn:disabled, QPushButton#svgBtn:disabled {{
        background-color: {C.BG_MUTED};
        color: {C.TEXT_FAINT};
        border: 1px solid {C.BORDER_STRONG};
    }}
    QMenu {{
        background-color: {C.BG_CARD};
        border: 1px solid {C.BORDER_STRONG};
        border-radius: {R.MD}px;
        padding: 4px;
    }}
    QMenu::item {{
        padding: 8px 24px 8px 12px;
        border-radius: {R.SM}px;
        font-size: {T.SIZE_BODY}px;
        color: {C.TEXT_BODY};
    }}
    QMenu::item:selected {{
        background-color: {C.ACCENT_SOFT};
        color: {C.ACCENT_PRIMARY};
        font-weight: bold;
    }}
    QPushButton::menu-indicator {{
        width: 0px;
    }}
    QSplitter::handle {{
        background-color: {C.BORDER_SUBTLE};
        height: 2px;
        margin: 4px 0px;
    }}
    QSplitter::handle:hover {{
        background-color: {C.ACCENT_PRIMARY};
    }}
    QStatusBar {{
        background-color: {C.BG_SURFACE};
        border-top: 1px solid {C.BORDER_STRONG};
        color: {C.TEXT_BODY};
    }}
    QComboBox {{
        combobox-popup: 0;
        padding: 4px 8px;
        border-radius: {R.SM}px;
        border: 1px solid {C.BORDER_STRONG};
        background-color: {C.BG_CARD};
        font-weight: bold;
        color: {C.TEXT_BODY};
    }}
    QComboBox::drop-down {{
        border: none;
        width: 20px;
    }}
    QTextEdit#console, QListWidget#queueList {{
        background-color: {C.CONSOLE_LEGACY_BG};
        color: {C.CONSOLE_TEXT};
        font-family: {T.FONT_FAMILY_MONO};
        font-size: {T.SIZE_BODY}px;
        border-radius: {R.LG}px;
        padding: 8px;
        border: 1px solid {C.CONSOLE_BORDER};
    }}
    QListWidget#queueList::item {{
        padding: 4px;
        border-bottom: 1px solid {C.CONSOLE_ROW_DIVIDER};
    }}
    QListWidget#queueList::item:selected {{
        background-color: {C.CONSOLE_SEL_BG};
        color: {C.TEXT_ON_ACCENT};
        border-radius: {R.SM}px;
    }}
"""

# ==========================================
# --- COMMON REUSABLE UI STYLES ---
# ==========================================

CARD_FRAME_STYLE = f"""
QFrame#cardFrame {{
    background-color: {C.BG_CARD};
    border: 1px solid {C.BORDER_SUBTLE};
    border-radius: {R.LG}px;
}}
"""

TABLE_STYLE_CLEAN = f"""
QTableWidget {{
    background-color: {C.BG_CARD};
    border: 1px solid {C.BORDER_SUBTLE};
    border-radius: {R.MD}px;
    gridline-color: {C.BORDER_ROW};
    font-size: {T.SIZE_TABLE}px;
}}
QTableWidget::item {{
    padding: 4px 8px;
    border-bottom: 1px solid {C.BORDER_ROW};
}}
QTableWidget::item:selected {{
    background-color: {C.ACCENT_SOFT};
    color: {C.TEXT_STRONG};
}}
QHeaderView::section {{
    background-color: {C.BG_SURFACE};
    color: {C.TEXT_SECONDARY};
    font-weight: bold;
    font-size: {T.SIZE_CAPTION}px;
    border: none;
    border-bottom: 1px solid {C.BORDER_STRONG};
    padding: 6px 8px;
}}
"""

TREE_STYLE_CLEAN = f"""
QTreeWidget {{
    background-color: {C.BG_CARD};
    border: 1px solid {C.BORDER_SUBTLE};
    border-radius: {R.MD}px;
    alternate-background-color: {C.BG_SURFACE};
    font-size: {T.SIZE_TABLE}px;
    outline: none;
}}
QTreeWidget::item {{
    min-height: 28px;
    padding: 2px 6px;
    border-bottom: 1px solid {C.BORDER_ROW};
}}
QTreeWidget::item:selected {{
    background-color: {C.ACCENT_SOFT};
    color: {C.TEXT_STRONG};
}}
QTreeWidget::item:hover {{
    background-color: {C.BG_SURFACE};
}}
QHeaderView::section {{
    background-color: {C.BG_SURFACE};
    color: {C.TEXT_SECONDARY};
    font-weight: bold;
    font-size: {T.SIZE_CAPTION}px;
    border: none;
    border-bottom: 1px solid {C.BORDER_STRONG};
    padding: 6px 8px;
}}
"""

SEARCH_INPUT_STYLE = f"""
QLineEdit {{
    padding: 4px 8px;
    border: 1px solid {C.BORDER_STRONG};
    border-radius: {R.SM}px;
    font-size: {T.SIZE_TABLE}px;
    background-color: {C.BG_CARD};
}}
QLineEdit:focus {{
    border: 1px solid {C.ACCENT_PRIMARY};
}}
"""

BADGE_STYLE_PRIMARY = f"""
background-color: {C.SUCCESS_SOFT_MINT};
color: {C.SUCCESS_ALT};
font-weight: bold;
font-size: {T.SIZE_CAPTION}px;
border: 1px solid {C.SUCCESS_BORDER_MINT};
border-radius: {R.SM}px;
padding: 2px 8px;
"""

BADGE_STYLE_MUTED = f"""
background-color: {C.BG_CANVAS};
color: {C.TEXT_SECONDARY};
font-size: {T.SIZE_CAPTION}px;
border: 1px solid {C.BORDER_SUBTLE};
border-radius: {R.SM}px;
padding: 2px 8px;
"""

BADGE_STYLE_INFO = f"""
background-color: {C.INFO_SOFT_ALT};
color: {C.INFO_ALT};
border: 1px solid {C.INFO_BORDER_ALT};
border-radius: {R.SM}px;
padding: 2px 8px;
font-size: {T.SIZE_CAPTION}px;
font-weight: bold;
"""

BUTTON_STYLE_TOOLBAR_SECONDARY = f"""
QPushButton {{
    background-color: {C.BG_SURFACE};
    color: {C.TEXT_STRONG};
    border: 1px solid {C.BORDER_STRONG};
    border-radius: {R.SM}px;
    padding: 5px 12px;
    font-size: {T.SIZE_TABLE}px;
    font-weight: 500;
}}
QPushButton:hover {{
    background-color: {C.BG_CANVAS};
    border-color: {C.BORDER_HOVER};
}}
QPushButton:pressed {{
    background-color: {C.BG_MUTED};
}}
QPushButton:checked {{
    background-color: {C.ACCENT_SOFT};
    color: {C.ACCENT_PRIMARY};
    border: 1.5px solid {C.ACCENT_PRIMARY};
    font-weight: 600;
}}
QPushButton:checked:hover {{
    background-color: {C.ACCENT_SOFT_HOVER};
    border-color: {C.ACCENT_HOVER};
}}
QPushButton:disabled {{
    background-color: {C.BG_SURFACE};
    color: {C.TEXT_FAINT};
    border-color: {C.BORDER_SUBTLE};
}}
"""

BUTTON_STYLE_TOOLBAR_DANGER = f"""
QPushButton {{
    background-color: {C.DANGER_SOFT};
    color: {C.DANGER_ALT};
    border: 1px solid {C.DANGER_BORDER};
    border-radius: {R.SM}px;
    padding: 5px 12px;
    font-size: {T.SIZE_TABLE}px;
    font-weight: bold;
}}
QPushButton:hover {{
    background-color: {C.DANGER_SOFT_HOVER};
    border-color: {C.DANGER_BORDER_HOVER};
    color: {C.DANGER};
}}
QPushButton:pressed {{
    background-color: {C.DANGER_BORDER};
    border-color: {C.DANGER_ACCENT};
}}
QPushButton:checked {{
    background-color: {C.DANGER_SOFT_HOVER};
    color: {C.DANGER_DARK};
    border: 1.5px solid {C.DANGER_ACCENT};
    font-weight: bold;
}}
QPushButton:checked:hover {{
    background-color: {C.DANGER_BORDER};
    border-color: {C.DANGER_ALT};
}}
QPushButton:disabled {{
    background-color: {C.BG_SURFACE};
    color: {C.TEXT_FAINT};
    border-color: {C.BORDER_SUBTLE};
}}
"""

BUTTON_STYLE_TOOLBAR_WARNING = f"""
QPushButton {{
    background-color: {C.WARNING_SOFT};
    color: {C.WARNING};
    border: 1px solid {C.WARNING_BORDER};
    border-radius: {R.SM}px;
    padding: 5px 12px;
    font-size: {T.SIZE_TABLE}px;
    font-weight: 500;
}}
QPushButton:hover {{
    background-color: {C.WARNING_SOFT_ALT};
    border-color: {C.WARNING_ACCENT};
}}
QPushButton:pressed {{
    background-color: {C.WARNING_BORDER};
}}
QPushButton:checked {{
    background-color: {C.WARNING_SOFT_ALT};
    color: {C.WARNING_DARK};
    border: 1.5px solid {C.WARNING_ACCENT};
    font-weight: 600;
}}
QPushButton:checked:hover {{
    background-color: {C.WARNING_BORDER};
    border-color: {C.WARNING};
}}
QPushButton:disabled {{
    background-color: {C.BG_SURFACE};
    color: {C.TEXT_FAINT};
    border-color: {C.BORDER_SUBTLE};
}}
"""

COMBOBOX_STYLE_TOOLBAR = f"""
QComboBox {{
    combobox-popup: 0;
    font-size: {T.SIZE_CAPTION}px;
    font-weight: 600;
    color: {C.TEXT_STRONG};
    background-color: {C.BG_SURFACE};
    border: 1px solid {C.BORDER_STRONG};
    border-radius: {R.SM}px;
    padding: 3px 22px 3px 10px;
    min-height: 22px;
}}
QComboBox:hover {{
    background-color: {C.BG_CANVAS};
    border-color: {C.BORDER_HOVER};
}}
QComboBox:focus, QComboBox:on {{
    background-color: {C.BG_CARD};
    border-color: {C.ACCENT_PRIMARY};
}}
QComboBox QAbstractItemView {{
    border: 1px solid {C.BORDER_STRONG};
    border-radius: {R.MD}px;
    background-color: {C.BG_CARD};
    color: {C.TEXT_STRONG};
    selection-background-color: {C.ACCENT_SOFT};
    selection-color: {C.ACCENT_PRIMARY};
    padding: 4px;
    outline: none;
}}
QComboBox QAbstractItemView::item {{
    min-height: 22px;
    padding: 2px 8px;
    border-radius: 3px;
}}
QComboBox QAbstractItemView::item:hover {{
    background-color: {C.BG_CANVAS};
    color: {C.TEXT_PRIMARY};
}}
QComboBox QAbstractItemView::item:selected {{
    background-color: {C.ACCENT_SOFT};
    color: {C.ACCENT_PRIMARY};
    font-weight: bold;
}}
QScrollBar:vertical {{
    border: none;
    background: {C.BG_SURFACE};
    width: 8px;
    margin: 4px 0;
    border-radius: {R.SM}px;
}}
QScrollBar::handle:vertical {{
    background: {C.BORDER_STRONG};
    min-height: 20px;
    border-radius: {R.SM}px;
}}
QScrollBar::handle:vertical:hover {{
    background: {C.BORDER_HOVER};
}}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
    height: 0px;
}}
"""

SEGMENT_BTN_STYLE = f"""
QPushButton {{
    padding: 3px 9px;
    font-size: {T.SIZE_CAPTION}px;
    font-weight: 500;
    background-color: {C.BG_SURFACE};
    color: {C.TEXT_BODY};
    border: 1px solid {C.BORDER_STRONG};
    border-radius: {R.SM}px;
}}
QPushButton:hover {{
    background-color: {C.BG_CANVAS};
    border-color: {C.BORDER_HOVER};
}}
QPushButton:checked {{
    background-color: {C.ACCENT_BRIGHT};
    color: {C.TEXT_ON_ACCENT};
    font-weight: bold;
    border: 1px solid {C.ACCENT_BRIGHT_BORDER};
}}
"""

TOGGLE_PILL_STYLE = f"""
QPushButton {{
    padding: 3px 12px;
    font-size: {T.SIZE_CAPTION}px;
    font-weight: 500;
    background-color: {C.BG_SURFACE};
    color: {C.TEXT_BODY};
    border: 1px solid {C.BORDER_STRONG};
    border-radius: {R.PILL}px;
    min-height: 22px;
}}
QPushButton:hover {{
    background-color: {C.BG_CANVAS};
    border-color: {C.BORDER_HOVER};
}}
QPushButton:checked {{
    background-color: {C.ACCENT_SOFT};
    color: {C.ACCENT_PRIMARY};
    font-weight: bold;
    border: 1.5px solid {C.ACCENT_PRIMARY};
    border-radius: {R.PILL}px;
}}
QPushButton:checked:hover {{
    background-color: {C.ACCENT_SOFT_HOVER};
}}
"""

FAV_PILL_STYLE = f"""
QPushButton {{
    padding: 3px 10px;
    font-size: {T.SIZE_CAPTION}px;
    font-weight: 500;
    background-color: {C.BG_SURFACE};
    color: {C.TEXT_BODY};
    border: 1px solid {C.BORDER_STRONG};
    border-radius: {R.PILL}px;
}}
QPushButton:hover {{
    background-color: {C.BG_CANVAS};
    border-color: {C.BORDER_HOVER};
}}
QPushButton:checked {{
    background-color: {C.WARNING_SOFT_ALT};
    color: {C.WARNING_DARK};
    font-weight: bold;
    border: 1.5px solid {C.WARNING_ACCENT};
}}
QPushButton:checked:hover {{
    background-color: {C.WARNING_BORDER};
}}
"""

DISMISSIBLE_TAG_STYLE = f"""
QPushButton {{
    padding: 2px 8px;
    font-size: {T.SIZE_CAPTION}px;
    font-weight: 500;
    background-color: {C.INFO_SOFT};
    color: {C.INFO_DARK};
    border: 1px solid {C.INFO_BORDER};
    border-radius: {R.XL}px;
}}
QPushButton:hover {{
    background-color: {C.DANGER_SOFT_HOVER};
    color: {C.DANGER_DARK};
    border-color: {C.DANGER_BORDER_ALT};
}}
"""

CLEAR_ALL_BTN_STYLE = f"""
QPushButton {{
    padding: 2px 8px;
    font-size: {T.SIZE_CAPTION}px;
    font-weight: bold;
    background-color: {C.DANGER_SOFT};
    color: {C.DANGER};
    border: 1px solid {C.DANGER_BORDER};
    border-radius: {R.SM}px;
}}
QPushButton:hover {{
    background-color: {C.DANGER_SOFT_HOVER};
    border-color: {C.DANGER_BORDER_HOVER};
}}
"""

SPEC_FILTER_COMBO_STYLE = f"""
QComboBox {{
    combobox-popup: 0;
    padding: 3px 8px;
    border: 1px solid {C.BORDER_STRONG};
    border-radius: {R.SM}px;
    font-size: {T.SIZE_TABLE}px;
    background-color: {C.BG_CARD};
    color: {C.TEXT_STRONG};
}}
QComboBox:hover {{
    border-color: {C.BORDER_HOVER};
}}
QComboBox:focus {{
    border-color: {C.ACCENT_PRIMARY};
}}
"""

SPEC_FILTER_INPUT_STYLE = f"""
QLineEdit {{
    padding: 4px 8px;
    border: 1px solid {C.BORDER_STRONG};
    border-radius: {R.SM}px;
    font-size: {T.SIZE_TABLE}px;
    background-color: {C.BG_CARD};
    color: {C.TEXT_STRONG};
}}
QLineEdit:focus {{
    border-color: {C.ACCENT_PRIMARY};
}}
"""

# ==========================================
# --- 3GPP WORKING GROUP STYLING & PALETTES ---
# ==========================================

WG_DESCRIPTIONS = {
    # SA
    "SA": "Service and System Aspects (Plenary)",
    "SA1": "Services",
    "SA2": "Architecture",
    "SA3": "Security",
    "SA4": "Codec",
    "SA5": "Telecom Management",
    "SA6": "Mission Critical & Vertical Applications",
    # RAN
    "RAN": "Radio Access Network (Plenary)",
    "RAN1": "Radio Layer 1 (Physical layer)",
    "RAN2": "Radio Layer 2 & 3",
    "RAN3": "Architecture & Network Interfaces",
    "RAN4": "Radio Performance & Protocol Aspects",
    "RAN5": "Mobile Terminal Conformance Testing",
    "RAN6": "GSM/EDGE Radio Access Network (Closed/Legacy)",
    # CT
    "CT": "Core Network and Terminals (Plenary)",
    "CT1": "User Equipment - Core Network Protocols",
    "CT3": "Interworking with External Networks",
    "CT4": "Core Network Protocols",
    "CT6": "Smart Card Application Aspects",
    # Legacy / Other
    "AHG": "Ad Hoc Group (Legacy)",
    "GERAN": "GSM EDGE Radio Access Network (Closed/Legacy)",
    "GERAN1": "GERAN Radio Aspects (Closed/Legacy)",
    "GERAN2": "GERAN Protocol Aspects (Closed/Legacy)",
    "GERAN3": "GERAN Base Station Testing (Closed/Legacy)",
    "CN1": "Core Network 1 (Legacy)",
    "CN3": "Core Network 3 (Legacy)",
    "CN4": "Core Network 4 (Legacy)",
    "CN5": "Open Service Access (Legacy)",
    "ITU": "ITU Interworking Ad Hoc (Legacy)",
    "T2": "Mobile Terminal Services (Legacy)",
    "T3": "USIM / Smart Cards (Legacy)",
    "SMG01": "Special Mobile Group 1 (Legacy)",
    "SMG03": "Special Mobile Group 3 (Legacy)",
}


def get_wg_colors(wg: str) -> dict:
    """Returns standardized palette hex colors for Working Group badges."""
    wg_upper = str(wg or "").strip().upper()
    desc = WG_DESCRIPTIONS.get(wg_upper, "Working Group")

    # SA (Soft Indigo)
    if wg_upper.startswith("SA"):
        return {
            "bg": C.WG_SA_BG,
            "text": C.WG_SA_TEXT,
            "border": C.WG_SA_BORDER,
            "desc": desc,
            "tsg": "SA",
        }
    # RAN (Soft Emerald)
    elif wg_upper.startswith("RAN") or wg_upper.startswith("GERAN"):
        return {
            "bg": C.WG_RAN_BG,
            "text": C.WG_RAN_TEXT,
            "border": C.WG_RAN_BORDER,
            "desc": desc,
            "tsg": "RAN",
        }
    # CT (Soft Amber)
    elif wg_upper.startswith("CT") or wg_upper.startswith("CN"):
        return {
            "bg": C.WG_CT_BG,
            "text": C.WG_CT_TEXT,
            "border": C.WG_CT_BORDER,
            "desc": desc,
            "tsg": "CT",
        }
    # Default (Slate)
    else:
        return {
            "bg": C.WG_DEFAULT_BG,
            "text": C.WG_DEFAULT_TEXT,
            "border": C.WG_DEFAULT_BORDER,
            "desc": desc,
            "tsg": "OTHER",
        }


def get_wg_badge_style_and_desc(wg: str) -> Tuple[str, str]:
    """Returns CSS stylesheet and description for QWidget/HTML rendering."""
    c = get_wg_colors(wg)
    style = f"""
        background-color: {c['bg']};
        color: {c['text']};
        border: 1px solid {c['border']};
        border-radius: 3px;
        padding: 2px 4px;
        font-size: 11px;
        font-weight: bold;
    """
    return style, c["desc"]


def escape_qt_mnemonic(text: Optional[str]) -> str:
    """
    Escape literal ampersands for Qt controls, actions, and menus (& -> &&).

    In Qt/PyQt, widgets and actions that support keyboard shortcuts and mnemonics
    (QAction, QMenu, QPushButton, QCheckBox, QRadioButton, QGroupBox, QTabBar,
    QTabWidget, etc.) interpret a single '&' as an accelerator key prefix,
    stripping the ampersand and underlining the subsequent character.

    IMPORTANT:
    This function is strictly for mnemonic-aware controls. Do NOT use it for
    non-mnemonic widgets such as QLabel, QLineEdit, QTextEdit, QToolTip,
    QStatusBar, setWindowTitle, or HTML rich text. Those controls do not parse
    mnemonics, and escaping '&' to '&&' will cause Qt to render literal '&&'
    in the UI!

    This function safely converts single '&' to '&&' while preserving already-escaped
    '&&' sequences, preventing double-escaping bugs when processing dynamic data
    (e.g., from JSON files, databases, or network feeds).

    :param text: Text string potentially containing ampersands, or None.
    :return: Escaped string safe for Qt mnemonic-aware controls.
    """
    if not text:
        return ""
    if not isinstance(text, str):
        text = str(text)
    # Protect existing '&&' by substituting a temporary marker, replacing single '&', then restoring
    return text.replace("&&", "\x00").replace("&", "&&").replace("\x00", "&&")


APP_USER_MODEL_ID = "3GPP Delegate Tools.1.0"
APP_DISPLAY_NAME = "3GPP Delegate Tools"

_CACHED_APP_ICON: Optional[QIcon] = None
_CACHED_HICON_BIG: Optional[int] = None
_CACHED_HICON_SMALL: Optional[int] = None


def _get_assets_dir() -> Path:
    try:
        from core.utils.paths import get_project_root
        return get_project_root() / "assets"
    except ImportError:
        return Path("assets")


def _hydrate_file_if_needed(path: Path) -> None:
    """Force OneDrive Files On-Demand hydration by reading the file bytes."""
    try:
        if path.is_file():
            _ = path.read_bytes()
    except Exception:
        pass


def _render_fallback_pixmap(size: int = 256) -> QPixmap:
    """Render the geometric 3GPP network emblem in memory if disk assets are unhydrated or unavailable."""
    pixmap = QPixmap(size, size)
    pixmap.fill(Qt.transparent)

    painter = QPainter(pixmap)
    painter.setRenderHint(QPainter.Antialiasing)

    painter.setBrush(QColor(C.TEXT_PRIMARY))
    painter.setPen(Qt.NoPen)
    corner_radius = max(2, size // 5)
    margin = max(1, size // 128)
    painter.drawRoundedRect(margin, margin, size - 2 * margin, size - 2 * margin, corner_radius, corner_radius)

    pen = QPen(QColor(C.INFO_ACCENT))
    pen.setWidth(max(2, size // 12))
    pen.setJoinStyle(Qt.RoundJoin)
    pen.setCapStyle(Qt.RoundCap)
    painter.setPen(pen)

    center_x = size * 0.50
    top_y = size * 0.28
    bl_x = size * 0.25
    bl_y = size * 0.72
    br_x = size * 0.75
    br_y = size * 0.72

    painter.drawLine(int(center_x), int(top_y), int(bl_x), int(bl_y))
    painter.drawLine(int(center_x), int(top_y), int(br_x), int(br_y))
    painter.drawLine(int(bl_x), int(bl_y), int(br_x), int(br_y))

    painter.setBrush(QColor(C.BG_CARD))
    painter.setPen(Qt.NoPen)
    node_radius = max(2, size // 10)
    painter.drawEllipse(int(center_x - node_radius), int(top_y - node_radius), node_radius * 2, node_radius * 2)
    painter.drawEllipse(int(bl_x - node_radius), int(bl_y - node_radius), node_radius * 2, node_radius * 2)
    painter.drawEllipse(int(br_x - node_radius), int(br_y - node_radius), node_radius * 2, node_radius * 2)
    painter.end()
    return pixmap


def get_app_icon_path() -> Optional[Path]:
    """Return the path to `3gpp_tools.ico` after ensuring OneDrive hydration, or a cached fallback."""
    assets_dir = _get_assets_dir()
    ico_path = assets_dir / "3gpp_tools.ico"
    _hydrate_file_if_needed(ico_path)
    if ico_path.is_file() and ico_path.stat().st_size > 0:
        return ico_path

    try:
        from core.utils.paths import get_cache_root
        fallback_ico = get_cache_root() / "3gpp_tools_fallback.ico"
        if not fallback_ico.is_file():
            fallback_ico.parent.mkdir(parents=True, exist_ok=True)
            pm = _render_fallback_pixmap(256)
            pm.save(str(fallback_ico), "ICO")
        if fallback_ico.is_file():
            return fallback_ico
    except Exception:
        pass
    return None


def create_app_icon() -> QIcon:
    """
    Build a multi-resolution application icon across all standard Windows DPI sizes
    (16, 20, 24, 32, 40, 48, 64, 128, 256 px) with in-memory fallback.
    """
    global _CACHED_APP_ICON
    if _CACHED_APP_ICON is not None and not _CACHED_APP_ICON.isNull():
        return _CACHED_APP_ICON

    assets_dir = _get_assets_dir()
    ico_path = assets_dir / "3gpp_tools.ico"
    png_path = assets_dir / "3gpp_tools.png"

    _hydrate_file_if_needed(ico_path)
    _hydrate_file_if_needed(png_path)

    icon = QIcon()
    if ico_path.is_file():
        icon.addFile(str(ico_path))

    master_pixmap = QPixmap(str(png_path)) if png_path.is_file() else QPixmap()
    if master_pixmap.isNull():
        master_pixmap = _render_fallback_pixmap(256)

    for sz in (16, 20, 24, 32, 40, 48, 64, 128, 256):
        scaled = master_pixmap.scaled(
            sz, sz, Qt.KeepAspectRatio, Qt.SmoothTransformation
        )
        if not scaled.isNull():
            icon.addPixmap(scaled, QIcon.Normal, QIcon.Off)
            icon.addPixmap(scaled, QIcon.Active, QIcon.Off)
            icon.addPixmap(scaled, QIcon.Selected, QIcon.Off)

    _CACHED_APP_ICON = icon
    return icon


def _get_win32_hicons() -> Tuple[int, int]:
    """Load and cache native Win32 big and small HICON handles from `3gpp_tools.ico`."""
    global _CACHED_HICON_BIG, _CACHED_HICON_SMALL
    if _CACHED_HICON_BIG and _CACHED_HICON_SMALL:
        return _CACHED_HICON_BIG, _CACHED_HICON_SMALL

    import os
    if os.name != "nt":
        return 0, 0

    ico_path = get_app_icon_path()
    if not ico_path:
        return 0, 0

    try:
        import ctypes
        import ctypes.wintypes as wt

        user32 = ctypes.windll.user32
        user32.LoadImageW.argtypes = [
            wt.HINSTANCE,
            wt.LPCWSTR,
            wt.UINT,
            ctypes.c_int,
            ctypes.c_int,
            wt.UINT,
        ]
        user32.LoadImageW.restype = wt.HANDLE
        user32.GetSystemMetrics.argtypes = [ctypes.c_int]
        user32.GetSystemMetrics.restype = ctypes.c_int

        IMAGE_ICON = 1
        LR_LOADFROMFILE = 0x00000010
        SM_CXICON = 11
        SM_CYICON = 12
        SM_CXSMICON = 49
        SM_CYSMICON = 50

        big_w = user32.GetSystemMetrics(SM_CXICON) or 32
        big_h = user32.GetSystemMetrics(SM_CYICON) or 32
        sm_w = user32.GetSystemMetrics(SM_CXSMICON) or 16
        sm_h = user32.GetSystemMetrics(SM_CYSMICON) or 16

        h_big = int(user32.LoadImageW(None, str(ico_path), IMAGE_ICON, big_w, big_h, LR_LOADFROMFILE) or 0)
        h_sm = int(user32.LoadImageW(None, str(ico_path), IMAGE_ICON, sm_w, sm_h, LR_LOADFROMFILE) or 0)
        if not h_big:
            h_big = int(user32.LoadImageW(None, str(ico_path), IMAGE_ICON, 0, 0, LR_LOADFROMFILE | 0x00000040) or 0)
        if not h_sm:
            h_sm = h_big

        _CACHED_HICON_BIG = h_big
        _CACHED_HICON_SMALL = h_sm
        return h_big, h_sm
    except Exception:
        return 0, 0


def _apply_win32_shell_window_properties(hwnd: int) -> None:
    """
    Bind explicit `IPropertyStore` shell metadata (`PKEY_AppUserModel_ID`,
    `PKEY_AppUserModel_RelaunchIconResource`, `PKEY_AppUserModel_RelaunchDisplayNameResource`)
    directly to `hwnd` so the Windows 10/11 Taskbar never falls back to a blank
    generic window icon when running as an unpinned `python.exe` process.
    """
    import os
    if os.name != "nt" or not hwnd:
        return

    ico_path = get_app_icon_path()
    if not ico_path:
        return

    try:
        import ctypes
        import ctypes.wintypes as wt
        import uuid

        class PROPERTYKEY(ctypes.Structure):
            _fields_ = [("fmtid", ctypes.c_byte * 16), ("pid", wt.DWORD)]

        class PROPVARIANT_UNION(ctypes.Union):
            _fields_ = [("pwszVal", wt.LPWSTR), ("ullVal", ctypes.c_ulonglong)]

        class PROPVARIANT(ctypes.Structure):
            _fields_ = [
                ("vt", wt.USHORT),
                ("wReserved1", wt.USHORT),
                ("wReserved2", wt.USHORT),
                ("wReserved3", wt.USHORT),
                ("u", PROPVARIANT_UNION),
            ]

        fmtid = (ctypes.c_byte * 16)(*uuid.UUID("{9F4C2855-9F79-4B39-A8D0-E1D42DE1D5F3}").bytes_le)
        pkey_aumid = PROPERTYKEY(fmtid, 5)
        pkey_relaunch_icon = PROPERTYKEY(fmtid, 3)
        pkey_relaunch_name = PROPERTYKEY(fmtid, 4)
        iid_prop_store = (ctypes.c_byte * 16)(*uuid.UUID("{886D8EEB-8CF2-4446-8D02-CDBA1DBDCF99}").bytes_le)

        shell32 = ctypes.windll.shell32
        shell32.SHGetPropertyStoreForWindow.argtypes = [
            wt.HWND,
            ctypes.c_void_p,
            ctypes.POINTER(ctypes.c_void_p),
        ]
        shell32.SHGetPropertyStoreForWindow.restype = ctypes.HRESULT

        store = ctypes.c_void_p()
        hr = shell32.SHGetPropertyStoreForWindow(hwnd, iid_prop_store, ctypes.byref(store))
        if hr != 0 or not store.value:
            return

        vt = ctypes.cast(store, ctypes.POINTER(ctypes.POINTER(ctypes.c_void_p))).contents
        set_value = ctypes.WINFUNCTYPE(
            ctypes.HRESULT,
            ctypes.c_void_p,
            ctypes.POINTER(PROPERTYKEY),
            ctypes.POINTER(PROPVARIANT),
        )(vt[6])
        commit = ctypes.WINFUNCTYPE(ctypes.HRESULT, ctypes.c_void_p)(vt[7])
        release = ctypes.WINFUNCTYPE(wt.ULONG, ctypes.c_void_p)(vt[2])

        try:
            pv = PROPVARIANT()
            pv.vt = 31  # VT_LPWSTR

            pv.u.pwszVal = APP_USER_MODEL_ID
            set_value(store, ctypes.byref(pkey_aumid), ctypes.byref(pv))

            pv.u.pwszVal = f"{ico_path},0"
            set_value(store, ctypes.byref(pkey_relaunch_icon), ctypes.byref(pv))

            pv.u.pwszVal = APP_DISPLAY_NAME
            set_value(store, ctypes.byref(pkey_relaunch_name), ctypes.byref(pv))

            commit(store)
        finally:
            release(store)
    except Exception:
        pass


def apply_native_window_icon(widget: Optional[QWidget], set_class_icon: bool = False) -> None:
    """
    Apply both Qt `QIcon` and native Win32 `WM_SETICON` / `IPropertyStore` icons
    to a top-level window widget.
    """
    if widget is None:
        return

    icon = create_app_icon()
    if not icon.isNull():
        widget.setWindowIcon(icon)

    import os
    if os.name != "nt":
        return

    try:
        hwnd = int(widget.winId())
    except Exception:
        return
    if not hwnd:
        return

    h_big, h_sm = _get_win32_hicons()
    if h_big or h_sm:
        try:
            import ctypes
            import ctypes.wintypes as wt

            user32 = ctypes.windll.user32
            user32.SendMessageW.argtypes = [wt.HWND, wt.UINT, wt.WPARAM, wt.LPARAM]
            user32.SendMessageW.restype = wt.LPARAM

            WM_SETICON = 0x0080
            ICON_SMALL = 0
            ICON_BIG = 1

            if h_big:
                user32.SendMessageW(hwnd, WM_SETICON, ICON_BIG, h_big)
            if h_sm:
                user32.SendMessageW(hwnd, WM_SETICON, ICON_SMALL, h_sm)

            if set_class_icon and hasattr(user32, "SetClassLongPtrW"):
                user32.SetClassLongPtrW.argtypes = [wt.HWND, ctypes.c_int, ctypes.c_void_p]
                user32.SetClassLongPtrW.restype = ctypes.c_void_p
                GCLP_HICON = -14
                GCLP_HICONSM = -34
                if h_big:
                    user32.SetClassLongPtrW(hwnd, GCLP_HICON, h_big)
                if h_sm:
                    user32.SetClassLongPtrW(hwnd, GCLP_HICONSM, h_sm)
        except Exception:
            pass

    _apply_win32_shell_window_properties(hwnd)


def apply_console_window_icon() -> None:
    """Set the 3GPP Tools icon on the attached Windows Console/Terminal window if present."""
    import os
    if os.name != "nt":
        return

    h_big, h_sm = _get_win32_hicons()
    if not (h_big or h_sm):
        return

    try:
        import ctypes
        import ctypes.wintypes as wt

        kernel32 = ctypes.windll.kernel32
        user32 = ctypes.windll.user32
        kernel32.GetConsoleWindow.restype = wt.HWND
        user32.SendMessageW.argtypes = [wt.HWND, wt.UINT, wt.WPARAM, wt.LPARAM]
        user32.SendMessageW.restype = wt.LPARAM

        hwnd = int(kernel32.GetConsoleWindow() or 0)
        if hwnd:
            WM_SETICON = 0x0080
            if h_big:
                user32.SendMessageW(hwnd, WM_SETICON, 1, h_big)
            if h_sm:
                user32.SendMessageW(hwnd, WM_SETICON, 0, h_sm)
    except Exception:
        pass


class _WindowIconEventFilter(QObject):
    """
    Application-wide event filter that automatically stamps every top-level
    `Qt.Window` or `Qt.Dialog` with both the Qt `QIcon` and native Win32 taskbar
    icon handles (`WM_SETICON` + `IPropertyStore`) whenever its native `HWND`
    is shown or recreated (e.g. after `setWindowFlags(Qt.Window)`).
    """

    def eventFilter(self, obj: QObject, event: QEvent) -> bool:
        et = event.type()
        if et in (QEvent.Show, QEvent.WinIdChange) and isinstance(obj, QWidget) and obj.isWindow():
            wtype = int(obj.windowFlags() & Qt.WindowType_Mask)
            if wtype in (int(Qt.Window), int(Qt.Dialog)):
                apply_native_window_icon(obj, set_class_icon=False)
        return False


def install_window_icon_manager(app) -> None:
    """Install the global top-level window icon event filter on `QApplication`."""
    if app is None:
        return
    if getattr(app, "_window_icon_filter", None) is None:
        flt = _WindowIconEventFilter(app)
        app.installEventFilter(flt)
        app._window_icon_filter = flt
    apply_console_window_icon()


class ProxyTestWorker(QThread):
    """Executes proxy connection testing in a separate thread to keep the UI fluid."""
    finished_signal = pyqtSignal(bool, str)

    def __init__(self, proxies: dict, parent=None):
        super().__init__(parent)
        self.proxies = proxies

    def run(self):
        try:
            success, message = NetworkSession.test_connection(self.proxies)
            self.finished_signal.emit(success, redact_sensitive_urls(message))
        except Exception as e:
            self.finished_signal.emit(False, f"Test error: {redact_sensitive_urls(str(e))}")


class ProxyDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Network Proxy Configuration")
        self.setModal(True)
        self.resize(540, 370)

        self._test_worker: Optional[ProxyTestWorker] = None

        self._setup_ui()
        self._load_existing_settings()

    def _setup_ui(self):
        layout = QVBoxLayout(self)

        title = QLabel("🌐 Proxy Configuration & Profiles")
        title.setStyleSheet("font-size: 15px; font-weight: bold; margin-bottom: 2px;")
        layout.addWidget(title)

        desc = QLabel(
            "Configure corporate HTTP/HTTPS proxies. All proxy parameters (hosts, ports, username, "
            "and password) are encrypted locally via Windows DPAPI and are never stored in plain text."
        )
        desc.setWordWrap(True)
        desc.setStyleSheet(f"color: {C.TEXT_MUTED}; font-size: {T.SIZE_CAPTION}px; margin-bottom: 8px;")
        layout.addWidget(desc)

        self.enable_checkbox = QCheckBox("Activate this proxy profile")
        self.enable_checkbox.setStyleSheet(f"font-weight: bold; color: {C.TEXT_STRONG}; margin-bottom: 6px;")
        self.enable_checkbox.toggled.connect(self._on_enable_toggled)
        layout.addWidget(self.enable_checkbox)

        self.form_container = QWidget()
        form = QFormLayout(self.form_container)
        form.setContentsMargins(0, 0, 0, 0)
        form.setLabelAlignment(Qt.AlignRight)

        input_style = f"padding: 4px; border: 1px solid {C.BORDER_STRONG}; border-radius: {R.SM}px;"
        self.http_input = QLineEdit()
        self.http_input.setPlaceholderText("e.g., proxy.company.com:8080")
        self.http_input.setStyleSheet(input_style)

        self.https_input = QLineEdit()
        self.https_input.setPlaceholderText("e.g., proxy.company.com:8080")
        self.https_input.setStyleSheet(input_style)

        self.sync_checkbox = QCheckBox("Use the same proxy address for HTTPS")
        self.sync_checkbox.setChecked(True)
        self.sync_checkbox.stateChanged.connect(self._on_sync_changed)
        self.http_input.textChanged.connect(self._on_http_changed)

        self.user_input = QLineEdit()
        self.user_input.setPlaceholderText("Domain\\Username or Username (Optional)")
        self.user_input.setStyleSheet(input_style)

        self.pass_input = QLineEdit()
        self.pass_input.setEchoMode(QLineEdit.Password)
        self.pass_input.setPlaceholderText("Password (Optional)")
        self.pass_input.setStyleSheet(input_style)

        self.remember_cred_checkbox = QCheckBox("Save proxy profile securely (Windows DPAPI)")
        self.remember_cred_checkbox.setChecked(True)
        self.remember_cred_checkbox.setToolTip(
            "Encrypts all proxy configuration (host, port, username, password) using OS-level DPAPI."
        )

        form.addRow("HTTP Proxy:", self.http_input)
        form.addRow("", self.sync_checkbox)
        form.addRow("HTTPS Proxy:", self.https_input)
        form.addRow("Username:", self.user_input)
        form.addRow("Password:", self.pass_input)
        form.addRow("", self.remember_cred_checkbox)
        layout.addWidget(self.form_container)

        self.status_lbl = QLabel("")
        self.status_lbl.setWordWrap(True)
        self.status_lbl.setAlignment(Qt.AlignCenter)
        self.status_lbl.setStyleSheet(f"font-size: {T.SIZE_CAPTION}px; color: {C.TEXT_SECONDARY}; margin: 4px 0;")
        layout.addWidget(self.status_lbl)

        btn_layout = QHBoxLayout()

        self.deactivate_btn = QPushButton("Deactivate (Direct)")
        self.deactivate_btn.setToolTip("Turn off proxy usage while retaining encrypted settings.")
        self.deactivate_btn.clicked.connect(self.deactivate_profile)

        self.cancel_btn = QPushButton("Cancel")
        self.cancel_btn.clicked.connect(self.reject)

        self.test_btn = QPushButton("⚡ Test Connection")
        self.test_btn.clicked.connect(self.test_proxy)

        self.save_btn = QPushButton("Save && Apply")
        self.save_btn.setObjectName("primaryBtn")
        self.save_btn.setStyleSheet(f"background-color: {C.ACCENT_PRIMARY}; color: {C.TEXT_ON_ACCENT}; font-weight: bold;")
        self.save_btn.clicked.connect(self.accept)

        btn_layout.addWidget(self.deactivate_btn)
        btn_layout.addWidget(self.cancel_btn)
        btn_layout.addStretch()
        btn_layout.addWidget(self.test_btn)
        btn_layout.addWidget(self.save_btn)

        layout.addLayout(btn_layout)

    def _on_enable_toggled(self, checked: bool):
        self.form_container.setEnabled(checked)
        self.test_btn.setEnabled(checked)

    def _on_sync_changed(self, state):
        is_synced = (state == Qt.Checked)
        self.https_input.setEnabled(not is_synced)
        if is_synced:
            self.https_input.setText(self.http_input.text())

    def _on_http_changed(self, text):
        if self.sync_checkbox.isChecked():
            self.https_input.setText(text)

    def _load_existing_settings(self):
        profile = ProxyProfileManager.load_profile()
        self.enable_checkbox.setChecked(profile.get("enabled", False))
        self._on_enable_toggled(self.enable_checkbox.isChecked())

        self.http_input.setText(profile.get("http_host", ""))
        self.https_input.setText(profile.get("https_host", ""))
        self.sync_checkbox.setChecked(profile.get("sync_https", True))
        self.user_input.setText(profile.get("username", ""))
        self.pass_input.setText(profile.get("password", ""))

        has_data = bool(
            profile.get("http_host") or profile.get("https_host") or profile.get("username")
        )
        self.remember_cred_checkbox.setChecked(has_data or True)

    def _build_proxy_uri(self, raw_address: str) -> str:
        raw_address = raw_address.strip()
        if not raw_address:
            return ""

        scheme, netloc = (
            raw_address.split("://", 1)
            if "://" in raw_address
            else ("http", raw_address)
        )
        username = self.user_input.text().strip()
        password = self.pass_input.text().strip()

        if username:
            safe_user = urllib.parse.quote(username, safe="")
            if password:
                safe_pass = urllib.parse.quote(password, safe="")
                auth = f"{safe_user}:{safe_pass}@"
            else:
                auth = f"{safe_user}@"
            return f"{scheme}://{auth}{netloc}"

        return f"{scheme}://{netloc}"

    def build_proxies_dict(self) -> dict:
        proxies = {}
        http_raw = self.http_input.text().strip()
        https_raw = self.https_input.text().strip()

        if http_raw:
            proxies["http"] = self._build_proxy_uri(http_raw)
        if https_raw:
            proxies["https"] = self._build_proxy_uri(https_raw)

        return proxies

    def test_proxy(self):
        proxies = self.build_proxies_dict()

        self.test_btn.setEnabled(False)
        self.test_btn.setText("Testing...")
        self.save_btn.setEnabled(False)
        self.status_lbl.setText("⏳ Connecting to www.3gpp.org through proxy...")
        self.status_lbl.setStyleSheet(f"color: {C.ACCENT_SKY}; font-size: {T.SIZE_CAPTION}px;")

        self._test_worker = ProxyTestWorker(proxies, parent=self)
        self._test_worker.finished_signal.connect(self._on_test_finished)
        self._test_worker.start()

    def _on_test_finished(self, success: bool, message: str):
        self.test_btn.setEnabled(True)
        self.test_btn.setText("⚡ Test Connection")
        self.save_btn.setEnabled(True)

        if success:
            self.status_lbl.setText("✅ Connection verified successfully!")
            self.status_lbl.setStyleSheet(f"color: {C.SUCCESS_DARK}; font-size: {T.SIZE_CAPTION}px;")
            QMessageBox.information(self, "Proxy Success", message)
        else:
            self.status_lbl.setText("❌ Connection failed.")
            self.status_lbl.setStyleSheet(f"color: {C.DANGER_ALT}; font-size: {T.SIZE_CAPTION}px;")
            QMessageBox.warning(self, "Proxy Test Failed", message)

    def deactivate_profile(self):
        """Turns off proxy routing while preserving the encrypted configuration on disk."""
        ProxyProfileManager.set_enabled(False)
        NetworkSession.update_proxies({})
        super().accept()

    def accept(self):
        """
        Validates encryption before persistence.
        Guarantees that unencrypted proxy information (hosts, ports, usernames, passwords)
        is NEVER written to disk, updates NetworkSession, and dismisses the dialog cleanly.
        """
        is_enabled = self.enable_checkbox.isChecked()
        http_host = self.http_input.text().strip()
        https_host = self.https_input.text().strip()
        sync_https = self.sync_checkbox.isChecked()
        username = self.user_input.text().strip()
        password = self.pass_input.text()

        encrypted_payload = ""

        if self.remember_cred_checkbox.isChecked() and (http_host or https_host):
            payload_dict = {
                "http_host": http_host,
                "https_host": https_host,
                "sync_https": sync_https,
                "username": username,
                "password": password,
            }
            success, enc_result = SecureCredentialStore.encrypt_dict(payload_dict)
            if not success:
                QMessageBox.warning(
                    self,
                    "Security Notice",
                    "Windows DPAPI encryption is unavailable or encountered an error.\n\n"
                    "Your proxy settings were NOT written to disk to prevent exposing corporate network data. "
                    "They will remain active in memory for the current session only."
                )
                encrypted_payload = ""
            else:
                encrypted_payload = enc_result

        profile_data = {
            "enabled": is_enabled,
            "encrypted_payload": encrypted_payload,
        }
        ProxyProfileManager.save_profile(profile_data)

        if is_enabled:
            NetworkSession.update_proxies(self.build_proxies_dict())
        else:
            NetworkSession.update_proxies({})

        super().accept()

    def closeEvent(self, event):
        if self._test_worker and self._test_worker.isRunning():
            self._test_worker.requestInterruption()
            self._test_worker.wait(1000)
        super().closeEvent(event)

    def get_proxies(self) -> Tuple[str, str]:
        proxies = self.build_proxies_dict()
        return proxies.get("http", ""), proxies.get("https", "")


# ==========================================
# --- DRAG & DROP INTERACTION WIDGET ---
# ==========================================
from core.ui.drop_zone import (
    ModernDropZone,
    InteractiveDropLabel,
    ClickableDropLabel,
    format_file_size,
)
