# --- File: src/core/ui/ui_components.py ---
import urllib.parse
from pathlib import Path
from typing import Optional, Tuple

from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QIcon
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


def create_app_icon():
    """Load the stable application icon used by Qt and the Windows shell."""
    try:
        from core.utils.paths import get_project_root
        icon_path = get_project_root() / "assets" / "3gpp_tools.ico"
    except ImportError:
        icon_path = Path("assets") / "3gpp_tools.ico"

    if icon_path.is_file():
        return QIcon(str(icon_path))

    # Keep startup resilient if an installation is missing the optional asset.
    # QIcon() lets Qt/Windows fall back cleanly instead of creating or rewriting
    # files in the project directory during application startup.
    import logging
    logging.warning("Application icon not found: %s", icon_path)
    return QIcon()


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
