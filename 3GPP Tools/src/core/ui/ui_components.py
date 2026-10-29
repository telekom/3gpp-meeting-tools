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
from core.utils.dpapi import SecureCredentialStore, redact_sensitive_urls

# ==========================================
# --- GLOBAL STYLESHEET (ALL-BLUE THEME) ---
# ==========================================
GLOBAL_STYLE = """
    QWidget {
        font-family: "Segoe UI", Arial, sans-serif;
        font-size: 13px;
        color: #333333;
    }
    QToolTip {
        color: #333333;
        background-color: #F8F8F8;
        border: 1px solid #D0D0D0;
        border-radius: 4px;
        padding: 4px;
    }
    QTabWidget::pane {
        border: 1px solid #D0D0D0;
        border-radius: 8px;
        background: #FFFFFF;
        top: -1px;
    }
    QTabBar::tab {
        background: #EAEAEA;
        border: 1px solid #D0D0D0;
        padding: 8px 16px;
        margin-right: 2px;
        border-top-left-radius: 6px;
        border-top-right-radius: 6px;
    }
    QTabBar::tab:selected {
        background: #FFFFFF;
        border-bottom-color: #FFFFFF;
        font-weight: bold;
        color: #395396;
    }
    QTabBar::tab:hover:!selected {
        background: #F0F0F0;
    }
    QPushButton {
        padding: 8px 16px;
        border-radius: 6px;
        border: 1px solid #CCCCCC;
        background-color: #F8F8F8;
        font-weight: bold;
    }
    QPushButton:hover {
        background-color: #EAEAEA;
    }
    QPushButton:disabled {
        background-color: #F0F0F0;
        color: #A0A0A0;
        border: 1px solid #DFDFDF;
    }
    QPushButton:checked {
        background-color: #EBF3FC;
        border: 2px solid #395396;
        color: #395396;
    }
    QPushButton#primaryBtn, QPushButton#pptBtn, QPushButton#svgBtn {
        background-color: #1E5C99; 
        color: white; 
        border: none;
    }
    QPushButton#primaryBtn:hover, QPushButton#pptBtn:hover, QPushButton#svgBtn:hover {
        background-color: #15426E;
    }
    QPushButton#primaryBtn:disabled, QPushButton#pptBtn:disabled, QPushButton#svgBtn:disabled {
        background-color: #E2E8F0;
        color: #94A3B8;
        border: 1px solid #CBD5E1;
    }
    QMenu {
        background-color: #FFFFFF;
        border: 1px solid #CCCCCC;
        border-radius: 6px;
        padding: 4px;
    }
    QMenu::item {
        padding: 8px 24px 8px 12px;
        border-radius: 4px;
        font-size: 13px;
        color: #333333;
    }
    QMenu::item:selected {
        background-color: #EBF3FC;
        color: #395396;
        font-weight: bold;
    }
    QPushButton::menu-indicator {
        width: 0px; 
    }
    QSplitter::handle {
        background-color: #E0E0E0;
        height: 2px;
        margin: 4px 0px;
    }
    QSplitter::handle:hover {
        background-color: #395396;
    }
    QStatusBar {
        background-color: #F0F0F0;
        border-top: 1px solid #D0D0D0;
        color: #333333;
    }
    QComboBox {
        padding: 4px 8px;
        border-radius: 4px;
        border: 1px solid #CCCCCC;
        background-color: #FFFFFF;
        font-weight: bold;
        color: #333333;
    }
    QComboBox::drop-down {
        border: none;
        width: 20px;
    }
    QTextEdit#console, QListWidget#queueList {
        background-color: #1E1E1E; 
        color: #D4D4D4; 
        font-family: Consolas, 'Courier New', monospace; 
        font-size: 13px; 
        border-radius: 8px; 
        padding: 8px;
        border: 1px solid #444444;
    }
    QListWidget#queueList::item {
        padding: 4px;
        border-bottom: 1px solid #333333;
    }
    QListWidget#queueList::item:selected {
        background-color: #264F78;
        color: #FFFFFF;
        border-radius: 4px;
    }
"""

# ==========================================
# --- COMMON REUSABLE UI STYLES ---
# ==========================================

CARD_FRAME_STYLE = """
QFrame#cardFrame {
    background-color: #FFFFFF;
    border: 1px solid #E2E8F0;
    border-radius: 8px;
}
"""

TABLE_STYLE_CLEAN = """
QTableWidget {
    background-color: #FFFFFF;
    border: 1px solid #E2E8F0;
    border-radius: 6px;
    gridline-color: #F1F5F9;
    font-size: 12px;
}
QTableWidget::item {
    padding: 4px 8px;
    border-bottom: 1px solid #F1F5F9;
}
QTableWidget::item:selected {
    background-color: #EBF3FC;
    color: #1E293B;
}
QHeaderView::section {
    background-color: #F8FAFC;
    color: #475569;
    font-weight: bold;
    font-size: 11px;
    border: none;
    border-bottom: 1px solid #CBD5E1;
    padding: 6px 8px;
}
"""



TREE_STYLE_CLEAN = """
QTreeWidget {
    background-color: #FFFFFF;
    border: 1px solid #E2E8F0;
    border-radius: 6px;
    alternate-background-color: #F8FAFC;
    font-size: 12px;
    outline: none;
}
QTreeWidget::item {
    min-height: 28px;
    padding: 2px 6px;
    border-bottom: 1px solid #F1F5F9;
}
QTreeWidget::item:selected {
    background-color: #EBF3FC;
    color: #1E293B;
}
QTreeWidget::item:hover {
    background-color: #F8FAFC;
}
QHeaderView::section {
    background-color: #F8FAFC;
    color: #475569;
    font-weight: bold;
    font-size: 11px;
    border: none;
    border-bottom: 1px solid #CBD5E1;
    padding: 6px 8px;
}
"""

SEARCH_INPUT_STYLE = """
QLineEdit {
    padding: 4px 8px;
    border: 1px solid #CBD5E1;
    border-radius: 4px;
    font-size: 12px;
    background-color: #FFFFFF;
}
QLineEdit:focus {
    border: 1px solid #1E5C99;
}
"""

BADGE_STYLE_PRIMARY = """
background-color: #E6F4EA;
color: #137333;
font-weight: bold;
font-size: 11px;
border: 1px solid #CEEAD6;
border-radius: 4px;
padding: 2px 8px;
"""

BADGE_STYLE_MUTED = """
background-color: #F1F5F9;
color: #475569;
font-size: 11px;
border: 1px solid #E2E8F0;
border-radius: 4px;
padding: 2px 8px;
"""

BADGE_STYLE_INFO = """
background-color: #EBF8FF;
color: #2B6CB0;
border: 1px solid #BEE3F8;
border-radius: 4px;
padding: 2px 8px;
font-size: 11px;
font-weight: bold;
"""

BUTTON_STYLE_TOOLBAR_SECONDARY = """
QPushButton {
    background-color: #F8FAFC;
    color: #1E293B;
    border: 1px solid #CBD5E1;
    border-radius: 4px;
    padding: 5px 12px;
    font-size: 12px;
    font-weight: 500;
}
QPushButton:hover {
    background-color: #F1F5F9;
    border-color: #94A3B8;
}
QPushButton:pressed {
    background-color: #E2E8F0;
}
QPushButton:checked {
    background-color: #EBF3FC;
    color: #1E5C99;
    border: 1.5px solid #1E5C99;
    font-weight: 600;
}
QPushButton:checked:hover {
    background-color: #DFEDFA;
    border-color: #15426E;
}
QPushButton:disabled {
    background-color: #F8FAFC;
    color: #94A3B8;
    border-color: #E2E8F0;
}
"""

BUTTON_STYLE_TOOLBAR_DANGER = """
QPushButton {
    background-color: #FEF2F2;
    color: #DC2626;
    border: 1px solid #FECACA;
    border-radius: 4px;
    padding: 5px 12px;
    font-size: 12px;
    font-weight: bold;
}
QPushButton:hover {
    background-color: #FEE2E2;
    border-color: #F87171;
    color: #B91C1C;
}
QPushButton:pressed {
    background-color: #FECACA;
    border-color: #EF4444;
}
QPushButton:checked {
    background-color: #FEE2E2;
    color: #991B1B;
    border: 1.5px solid #EF4444;
    font-weight: bold;
}
QPushButton:checked:hover {
    background-color: #FECACA;
    border-color: #DC2626;
}
QPushButton:disabled {
    background-color: #F8FAFC;
    color: #94A3B8;
    border-color: #E2E8F0;
}
"""

BUTTON_STYLE_TOOLBAR_WARNING = """
QPushButton {
    background-color: #FFFBEB;
    color: #B45309;
    border: 1px solid #FDE68A;
    border-radius: 4px;
    padding: 5px 12px;
    font-size: 12px;
    font-weight: 500;
}
QPushButton:hover {
    background-color: #FEF3C7;
    border-color: #F59E0B;
}
QPushButton:pressed {
    background-color: #FDE68A;
}
QPushButton:checked {
    background-color: #FEF3C7;
    color: #92400E;
    border: 1.5px solid #F59E0B;
    font-weight: 600;
}
QPushButton:checked:hover {
    background-color: #FDE68A;
    border-color: #D97706;
}
QPushButton:disabled {
    background-color: #F8FAFC;
    color: #94A3B8;
    border-color: #E2E8F0;
}
"""

COMBOBOX_STYLE_TOOLBAR = """
QComboBox {
    combobox-popup: 0;
    font-size: 11px;
    font-weight: 600;
    color: #1E293B;
    background-color: #F8FAFC;
    border: 1px solid #CBD5E1;
    border-radius: 4px;
    padding: 3px 22px 3px 10px;
    min-height: 22px;
}
QComboBox:hover {
    background-color: #F1F5F9;
    border-color: #94A3B8;
}
QComboBox:focus, QComboBox:on {
    background-color: #FFFFFF;
    border-color: #1E5C99;
}
QComboBox QAbstractItemView {
    border: 1px solid #CBD5E1;
    border-radius: 6px;
    background-color: #FFFFFF;
    color: #1E293B;
    selection-background-color: #EBF3FC;
    selection-color: #1E5C99;
    padding: 4px;
    outline: none;
}
QComboBox QAbstractItemView::item {
    min-height: 22px;
    padding: 2px 8px;
    border-radius: 3px;
}
QComboBox QAbstractItemView::item:hover {
    background-color: #F1F5F9;
    color: #0F172A;
}
QComboBox QAbstractItemView::item:selected {
    background-color: #EBF3FC;
    color: #1E5C99;
    font-weight: bold;
}
QScrollBar:vertical {
    border: none;
    background: #F8FAFC;
    width: 8px;
    margin: 4px 0;
    border-radius: 4px;
}
QScrollBar::handle:vertical {
    background: #CBD5E1;
    min-height: 20px;
    border-radius: 4px;
}
QScrollBar::handle:vertical:hover {
    background: #94A3B8;
}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
    height: 0px;
}
"""

SEGMENT_BTN_STYLE = """
QPushButton {
    padding: 3px 9px;
    font-size: 11px;
    font-weight: 500;
    background-color: #F8FAFC;
    color: #334155;
    border: 1px solid #CBD5E1;
    border-radius: 4px;
}
QPushButton:hover {
    background-color: #F1F5F9;
    border-color: #94A3B8;
}
QPushButton:checked {
    background-color: #0066CC;
    color: #FFFFFF;
    font-weight: bold;
    border: 1px solid #0055AA;
}
"""

TOGGLE_PILL_STYLE = """
QPushButton {
    padding: 3px 12px;
    font-size: 11px;
    font-weight: 500;
    background-color: #F8FAFC;
    color: #334155;
    border: 1px solid #CBD5E1;
    border-radius: 12px;
    min-height: 22px;
}
QPushButton:hover {
    background-color: #F1F5F9;
    border-color: #94A3B8;
}
QPushButton:checked {
    background-color: #EBF3FC;
    color: #1E5C99;
    font-weight: bold;
    border: 1.5px solid #1E5C99;
    border-radius: 12px;
}
QPushButton:checked:hover {
    background-color: #DFEDFA;
}
"""

FAV_PILL_STYLE = """
QPushButton {
    padding: 3px 10px;
    font-size: 11px;
    font-weight: 500;
    background-color: #F8FAFC;
    color: #334155;
    border: 1px solid #CBD5E1;
    border-radius: 12px;
}
QPushButton:hover {
    background-color: #F1F5F9;
    border-color: #94A3B8;
}
QPushButton:checked {
    background-color: #FEF3C7;
    color: #92400E;
    font-weight: bold;
    border: 1.5px solid #F59E0B;
}
QPushButton:checked:hover {
    background-color: #FDE68A;
}
"""

DISMISSIBLE_TAG_STYLE = """
QPushButton {
    padding: 2px 8px;
    font-size: 11px;
    font-weight: 500;
    background-color: #EFF6FF;
    color: #1E40AF;
    border: 1px solid #BFDBFE;
    border-radius: 10px;
}
QPushButton:hover {
    background-color: #FEE2E2;
    color: #991B1B;
    border-color: #FCA5A5;
}
"""

CLEAR_ALL_BTN_STYLE = """
QPushButton {
    padding: 2px 8px;
    font-size: 11px;
    font-weight: bold;
    background-color: #FEF2F2;
    color: #B91C1C;
    border: 1px solid #FECACA;
    border-radius: 4px;
}
QPushButton:hover {
    background-color: #FEE2E2;
    border-color: #F87171;
}
"""

SPEC_FILTER_COMBO_STYLE = """
QComboBox {
    padding: 3px 8px;
    border: 1px solid #CBD5E1;
    border-radius: 4px;
    font-size: 12px;
    background-color: #FFFFFF;
    color: #1E293B;
}
QComboBox:hover {
    border-color: #94A3B8;
}
QComboBox:focus {
    border-color: #1E5C99;
}
"""

SPEC_FILTER_INPUT_STYLE = """
QLineEdit {
    padding: 4px 8px;
    border: 1px solid #CBD5E1;
    border-radius: 4px;
    font-size: 12px;
    background-color: #FFFFFF;
    color: #1E293B;
}
QLineEdit:focus {
    border: 1px solid #1E5C99;
}
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
    "RAN6": "GSM/EDGE Radio Access Network",
    # CT
    "CT": "Core Network and Terminals (Plenary)",
    "CT1": "User Equipment - Core Network Protocols",
    "CT3": "Interworking with External Networks",
    "CT4": "Core Network Protocols",
    "CT6": "Smart Card Application Aspects",
    # Legacy / Other
    "GERAN": "GSM EDGE Radio Access Network",
    "GERAN1": "GERAN Radio Aspects",
    "GERAN2": "GERAN Protocol Aspects",
    "GERAN3": "GERAN Base Station Testing",
    "CN1": "Core Network 1 (Legacy)",
    "CN3": "Core Network 3 (Legacy)",
    "CN4": "Core Network 4 (Legacy)",
    "CN5": "Open Service Access (Legacy)",
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
            "bg": "#EEF2FF",
            "text": "#3730A3",
            "border": "#C7D2FE",
            "desc": desc,
            "tsg": "SA",
        }
    # RAN (Soft Emerald)
    elif wg_upper.startswith("RAN") or wg_upper.startswith("GERAN"):
        return {
            "bg": "#ECFDF5",
            "text": "#065F46",
            "border": "#A7F3D0",
            "desc": desc,
            "tsg": "RAN",
        }
    # CT (Soft Amber)
    elif wg_upper.startswith("CT") or wg_upper.startswith("CN"):
        return {
            "bg": "#FFF7ED",
            "text": "#9A3412",
            "border": "#FED7AA",
            "desc": desc,
            "tsg": "CT",
        }
    # Default (Slate)
    else:
        return {
            "bg": "#F1F5F9",
            "text": "#475569",
            "border": "#CBD5E1",
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
        padding: 1px 0px;
        font-size: 10px;
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
        desc.setStyleSheet("color: #64748B; font-size: 11px; margin-bottom: 8px;")
        layout.addWidget(desc)

        self.enable_checkbox = QCheckBox("Activate this proxy profile")
        self.enable_checkbox.setStyleSheet("font-weight: bold; color: #1E293B; margin-bottom: 6px;")
        self.enable_checkbox.toggled.connect(self._on_enable_toggled)
        layout.addWidget(self.enable_checkbox)

        self.form_container = QWidget()
        form = QFormLayout(self.form_container)
        form.setContentsMargins(0, 0, 0, 0)
        form.setLabelAlignment(Qt.AlignRight)

        self.http_input = QLineEdit()
        self.http_input.setPlaceholderText("e.g., proxy.company.com:8080")
        self.http_input.setStyleSheet("padding: 4px; border: 1px solid #CBD5E1; border-radius: 4px;")

        self.https_input = QLineEdit()
        self.https_input.setPlaceholderText("e.g., proxy.company.com:8080")
        self.https_input.setStyleSheet("padding: 4px; border: 1px solid #CBD5E1; border-radius: 4px;")

        self.sync_checkbox = QCheckBox("Use the same proxy address for HTTPS")
        self.sync_checkbox.setChecked(True)
        self.sync_checkbox.stateChanged.connect(self._on_sync_changed)
        self.http_input.textChanged.connect(self._on_http_changed)

        self.user_input = QLineEdit()
        self.user_input.setPlaceholderText("Domain\\Username or Username (Optional)")
        self.user_input.setStyleSheet("padding: 4px; border: 1px solid #CBD5E1; border-radius: 4px;")

        self.pass_input = QLineEdit()
        self.pass_input.setEchoMode(QLineEdit.Password)
        self.pass_input.setPlaceholderText("Password (Optional)")
        self.pass_input.setStyleSheet("padding: 4px; border: 1px solid #CBD5E1; border-radius: 4px;")

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
        self.status_lbl.setStyleSheet("font-size: 11px; color: #475569; margin: 4px 0;")
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
        self.save_btn.setStyleSheet("background-color: #1E5C99; color: white; font-weight: bold;")
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
        self.status_lbl.setStyleSheet("color: #0284C7; font-size: 11px;")

        self._test_worker = ProxyTestWorker(proxies, parent=self)
        self._test_worker.finished_signal.connect(self._on_test_finished)
        self._test_worker.start()

    def _on_test_finished(self, success: bool, message: str):
        self.test_btn.setEnabled(True)
        self.test_btn.setText("⚡ Test Connection")
        self.save_btn.setEnabled(True)

        if success:
            self.status_lbl.setText("✅ Connection verified successfully!")
            self.status_lbl.setStyleSheet("color: #166534; font-size: 11px;")
            QMessageBox.information(self, "Proxy Success", message)
        else:
            self.status_lbl.setText("❌ Connection failed.")
            self.status_lbl.setStyleSheet("color: #DC2626; font-size: 11px;")
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
