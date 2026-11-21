"""3GPP Working Group palette/descriptions and Qt mnemonic ampersand escaping."""

from typing import Optional, Tuple

from core.ui.theme import C


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

