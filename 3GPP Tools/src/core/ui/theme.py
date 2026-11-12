# --- File: src/core/ui/theme.py ---
"""
Centralized Design System for 3GPP Delegate Tools.

Defines semantic color tokens, typography scale, spacing/radii constants,
responsive layout breakpoints, and font loading helpers. All new UI code and
migrated modules should reference these tokens instead of raw hex literals.
"""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Dict, List, Optional

logger = logging.getLogger(__name__)


class ThemeColors:
    """Semantic light-theme color palette (Slate + 3GPP Blue)."""

    # --- Surfaces & Backgrounds ---
    BG_CANVAS = "#F1F5F9"        # Main application canvas / secondary hover fill
    BG_SURFACE = "#F8FAFC"       # Elevated subtle surface (toolbars, table headers, stripes)
    BG_CARD = "#FFFFFF"          # Primary card, input, and table background
    BG_MUTED = "#E2E8F0"         # Pressed surfaces, disabled primary fills, track backgrounds
    BG_HOVER = "#EEF2F6"         # Subtle button hover surface

    # --- Borders & Dividers ---
    BORDER_SUBTLE = "#E2E8F0"    # Card outlines, subtle dividers
    BORDER_STRONG = "#CBD5E1"    # Input fields, secondary buttons, headers
    BORDER_HOVER = "#94A3B8"     # Hover state border
    BORDER_ROW = "#F1F5F9"       # Table/tree row separators

    # --- Typography Colors ---
    TEXT_PRIMARY = "#0F172A"     # Headings, active strong text
    TEXT_STRONG = "#1E293B"      # Primary labels, input text
    TEXT_BODY = "#334155"        # Default body and control text
    TEXT_SECONDARY = "#475569"   # Table headers, descriptions
    TEXT_MUTED = "#64748B"       # Captions, placeholders, status bar
    TEXT_FAINT = "#94A3B8"       # Disabled text, group labels, subtle icons
    TEXT_ON_ACCENT = "#FFFFFF"   # Text rendered on primary accent buttons

    # --- Brand / Accent (3GPP Blue) ---
    ACCENT_PRIMARY = "#1E5C99"
    ACCENT_HOVER = "#15426E"
    ACCENT_SOFT = "#EBF3FC"
    ACCENT_SOFT_HOVER = "#DFEDFA"
    ACCENT_RING = "#BFD7F0"
    ACCENT_ALT = "#2563EB"
    ACCENT_BRIGHT = "#0066CC"
    ACCENT_BRIGHT_BORDER = "#0055AA"
    ACCENT_SKY = "#0284C7"
    ACCENT_SKY_DARK = "#0369A1"

    # --- Semantic Status ---
    SUCCESS = "#15803D"
    SUCCESS_DARK = "#166534"
    SUCCESS_ALT = "#137333"
    SUCCESS_SOFT = "#F0FDF4"
    SUCCESS_SOFT_ALT = "#DCFCE7"
    SUCCESS_SOFT_MINT = "#E6F4EA"
    SUCCESS_BORDER = "#BBF7D0"
    SUCCESS_BORDER_MINT = "#CEEAD6"

    WARNING = "#B45309"
    WARNING_DARK = "#92400E"
    WARNING_SOFT = "#FFFBEB"
    WARNING_SOFT_ALT = "#FEF3C7"
    WARNING_SOFT_LIGHT = "#FEFCE8"
    WARNING_BORDER = "#FDE68A"
    WARNING_ACCENT = "#F59E0B"

    DANGER = "#B91C1C"
    DANGER_ALT = "#DC2626"
    DANGER_DARK = "#991B1B"
    DANGER_SOFT = "#FEF2F2"
    DANGER_SOFT_HOVER = "#FEE2E2"
    DANGER_BORDER = "#FECACA"
    DANGER_BORDER_HOVER = "#F87171"
    DANGER_BORDER_ALT = "#FCA5A5"
    DANGER_ACCENT = "#EF4444"

    INFO = "#1D4ED8"
    INFO_DARK = "#1E40AF"
    INFO_ALT = "#2B6CB0"
    INFO_SOFT = "#EFF6FF"
    INFO_SOFT_ALT = "#EBF8FF"
    INFO_BORDER = "#BFDBFE"
    INFO_BORDER_ALT = "#BEE3F8"
    INFO_BORDER_SKY = "#90CDF4"

    # --- Console / Dark Terminal Drawer ---
    CONSOLE_BG = "#0F172A"
    CONSOLE_LEGACY_BG = "#1E1E1E"
    CONSOLE_SURFACE = "#1E293B"
    CONSOLE_TEXT = "#D4D4D4"
    CONSOLE_MUTED = "#94A3B8"
    CONSOLE_BORDER = "#444444"
    CONSOLE_ROW_DIVIDER = "#333333"
    CONSOLE_SEL_BG = "#264F78"
    CONSOLE_LOG_OK = "#6A9955"
    CONSOLE_LOG_WARN = "#D7BA7D"
    CONSOLE_LOG_ERR = "#F44747"

    # --- 3GPP Working Group Badges ---
    WG_SA_BG = "#EEF2FF"
    WG_SA_TEXT = "#3730A3"
    WG_SA_BORDER = "#C7D2FE"

    WG_RAN_BG = "#ECFDF5"
    WG_RAN_TEXT = "#065F46"
    WG_RAN_BORDER = "#A7F3D0"

    WG_CT_BG = "#FFF7ED"
    WG_CT_TEXT = "#9A3412"
    WG_CT_BORDER = "#FED7AA"

    WG_DEFAULT_BG = "#F1F5F9"
    WG_DEFAULT_TEXT = "#475569"
    WG_DEFAULT_BORDER = "#CBD5E1"

    # --- Extended Semantic / Accent Palette ---
    INFO_SOFT_STRONG = "#DBEAFE"
    INFO_ACCENT = "#3B82F6"
    INFO_BORDER_MEDIUM = "#93C5FD"
    EMERALD_PRIMARY = "#059669"
    EMERALD_ACCENT = "#10B981"
    EMERALD_DARK = "#047857"
    GREEN_ACCENT = "#16A34A"
    SUCCESS_BORDER_MEDIUM = "#86EFAC"
    INDIGO_SOFT = "#E0E7FF"
    INDIGO_DARK = "#4338CA"
    VIOLET_SOFT = "#EDE9FE"
    VIOLET_SURFACE = "#F5F3FF"
    VIOLET_BORDER = "#DDD6FE"
    VIOLET_BORDER_STRONG = "#C4B5FD"
    VIOLET_PRIMARY = "#6D28D9"
    VIOLET_DARK = "#5B21B6"
    PURPLE_DARK = "#6B21A8"
    TEAL_SURFACE = "#F0FDFA"
    TEAL_SOFT = "#CCFBF1"
    TEAL_BORDER = "#99F6E4"
    TEAL_DARK = "#0F766E"
    WARNING_BORDER_STRONG = "#FCD34D"
    WARNING_AMBER = "#D97706"
    OFFICE_BLUE = "#005A9E"
    OFFICE_BLUE_DARK = "#004578"
    BG_ALT_LIGHT = "#FAFAFA"
    BORDER_GRAY = "#E0E0E0"
    ACCENT_BRAND_PINK = "#E20074"
    SKY_SURFACE = "#F0F9FF"
    SKY_SOFT = "#E0F2FE"
    SKY_BORDER = "#BAE6FD"
    HIGHLIGHT_YELLOW = "#FEF08A"
    ORANGE_SOFT = "#FFEDD5"
    ORANGE_BORDER = "#FDBA74"
    ORANGE_DARK = "#C2410C"
    ACCENT_SOFT_PRESSED = "#DCEBFA"
    SKY_BORDER_STRONG = "#7DD3FC"
    WARNING_YELLOW_SOFT = "#FEF9C3"
    WARNING_YELLOW_TEXT = "#854D0E"
    WARNING_BROWN_DEEP = "#78350F"
    DIFF_ADDED_BG = "#E8F5E9"
    DIFF_REMOVED_BG = "#FFEBEE"
    DIFF_MODIFIED_BG = "#FFF9C4"
    EMERALD_SOFT = "#D1FAE5"
    EMERALD_BORDER = "#6EE7B7"
    BLUE_TINT_SOFT = "#F0F7FF"


class ThemeTypography:
    """Fixed typography scale and font stacks."""

    FONT_FAMILY_UI = '"Inter", "Segoe UI", Arial, sans-serif'
    FONT_FAMILY_MONO = '"Cascadia Mono", "Cascadia Code", Consolas, "Courier New", monospace'

    SIZE_PAGE_TITLE = 18
    SIZE_SECTION_TITLE = 14
    SIZE_BODY = 13
    SIZE_TABLE = 12
    SIZE_CAPTION = 11


class ThemeRadii:
    """Standard border radii (px)."""

    SM = 4
    MD = 6
    LG = 8
    XL = 10
    PILL = 12


class ThemeSpacing:
    """Standard layout spacing and control sizing tokens (px)."""

    XS = 4
    SM = 8
    MD = 12
    LG = 16
    XL = 24

    RAIL_EXPANDED_WIDTH = 224
    RAIL_COLLAPSED_WIDTH = 56
    CONTROL_HEIGHT = 32
    CONTROL_HEIGHT_SM = 28


class ThemeBreakpoints:
    """
    Responsive layout thresholds (logical pixels after OS DPI scaling).
    Includes a 30 px hysteresis buffer around the Compact/Standard boundary
    to prevent resize jitter when dragging window borders.
    """

    COMPACT_MAX_PX = 960
    COMPACT_COLLAPSE_PX = 940
    COMPACT_EXPAND_PX = 970

    WIDE_MIN_PX = 1600
    FOUR_COLUMN_MIN_PX = 1920
    SHORT_HEIGHT_MAX_PX = 760

    SPLIT_MIN_WINDOW_EXPANDED_PX = 1080
    SPLIT_MIN_WINDOW_ICON_RAIL_PX = 920

    PANE_LIST_CONDENSED_MAX_PX = 900
    PANE_FILTERS_FOLD_MAX_PX = 980
    PANE_DETAIL_CONDENSED_MAX_PX = 820


# Short convenient aliases
C = ThemeColors
T = ThemeTypography
R = ThemeRadii
S = ThemeSpacing
B = ThemeBreakpoints


# Canonical mapping of high-frequency raw hex values to ThemeColors attribute names.
# Used by automated migration tooling (Phase 0.4) to replace hardcoded colors safely.
HEX_TO_TOKEN: Dict[str, str] = {
    "#FFFFFF": "BG_CARD",
    "#F8FAFC": "BG_SURFACE",
    "#F1F5F9": "BG_CANVAS",
    "#E2E8F0": "BORDER_SUBTLE",
    "#CBD5E1": "BORDER_STRONG",
    "#94A3B8": "TEXT_FAINT",
    "#64748B": "TEXT_MUTED",
    "#475569": "TEXT_SECONDARY",
    "#334155": "TEXT_BODY",
    "#1E293B": "TEXT_STRONG",
    "#0F172A": "TEXT_PRIMARY",
    "#1E5C99": "ACCENT_PRIMARY",
    "#15426E": "ACCENT_HOVER",
    "#EBF3FC": "ACCENT_SOFT",
    "#DFEDFA": "ACCENT_SOFT_HOVER",
    "#BFD7F0": "ACCENT_RING",
    "#2563EB": "ACCENT_ALT",
    "#0066CC": "ACCENT_BRIGHT",
    "#0055AA": "ACCENT_BRIGHT_BORDER",
    "#0284C7": "ACCENT_SKY",
    "#0369A1": "ACCENT_SKY_DARK",
    "#15803D": "SUCCESS",
    "#166534": "SUCCESS_DARK",
    "#137333": "SUCCESS_ALT",
    "#F0FDF4": "SUCCESS_SOFT",
    "#DCFCE7": "SUCCESS_SOFT_ALT",
    "#E6F4EA": "SUCCESS_SOFT_MINT",
    "#BBF7D0": "SUCCESS_BORDER",
    "#CEEAD6": "SUCCESS_BORDER_MINT",
    "#B45309": "WARNING",
    "#92400E": "WARNING_DARK",
    "#FFFBEB": "WARNING_SOFT",
    "#FEF3C7": "WARNING_SOFT_ALT",
    "#FEFCE8": "WARNING_SOFT_LIGHT",
    "#FDE68A": "WARNING_BORDER",
    "#F59E0B": "WARNING_ACCENT",
    "#B91C1C": "DANGER",
    "#DC2626": "DANGER_ALT",
    "#991B1B": "DANGER_DARK",
    "#FEF2F2": "DANGER_SOFT",
    "#FEE2E2": "DANGER_SOFT_HOVER",
    "#FECACA": "DANGER_BORDER",
    "#F87171": "DANGER_BORDER_HOVER",
    "#FCA5A5": "DANGER_BORDER_ALT",
    "#EF4444": "DANGER_ACCENT",
    "#1D4ED8": "INFO",
    "#1E40AF": "INFO_DARK",
    "#2B6CB0": "INFO_ALT",
    "#EFF6FF": "INFO_SOFT",
    "#EBF8FF": "INFO_SOFT_ALT",
    "#BFDBFE": "INFO_BORDER",
    "#BEE3F8": "INFO_BORDER_ALT",
    "#90CDF4": "INFO_BORDER_SKY",
    "#EEF2FF": "WG_SA_BG",
    "#3730A3": "WG_SA_TEXT",
    "#C7D2FE": "WG_SA_BORDER",
    "#ECFDF5": "WG_RAN_BG",
    "#065F46": "WG_RAN_TEXT",
    "#A7F3D0": "WG_RAN_BORDER",
    "#FFF7ED": "WG_CT_BG",
    "#9A3412": "WG_CT_TEXT",
    "#FED7AA": "WG_CT_BORDER",
    "#DBEAFE": "INFO_SOFT_STRONG",
    "#3B82F6": "INFO_ACCENT",
    "#93C5FD": "INFO_BORDER_MEDIUM",
    "#059669": "EMERALD_PRIMARY",
    "#10B981": "EMERALD_ACCENT",
    "#047857": "EMERALD_DARK",
    "#16A34A": "GREEN_ACCENT",
    "#86EFAC": "SUCCESS_BORDER_MEDIUM",
    "#E0E7FF": "INDIGO_SOFT",
    "#4338CA": "INDIGO_DARK",
    "#EDE9FE": "VIOLET_SOFT",
    "#F5F3FF": "VIOLET_SURFACE",
    "#DDD6FE": "VIOLET_BORDER",
    "#C4B5FD": "VIOLET_BORDER_STRONG",
    "#6D28D9": "VIOLET_PRIMARY",
    "#5B21B6": "VIOLET_DARK",
    "#6B21A8": "PURPLE_DARK",
    "#F0FDFA": "TEAL_SURFACE",
    "#CCFBF1": "TEAL_SOFT",
    "#99F6E4": "TEAL_BORDER",
    "#0F766E": "TEAL_DARK",
    "#FCD34D": "WARNING_BORDER_STRONG",
    "#D97706": "WARNING_AMBER",
    "#005A9E": "OFFICE_BLUE",
    "#004578": "OFFICE_BLUE_DARK",
    "#FAFAFA": "BG_ALT_LIGHT",
    "#E0E0E0": "BORDER_GRAY",
    "#E20074": "ACCENT_BRAND_PINK",
    "#F0F9FF": "SKY_SURFACE",
    "#E0F2FE": "SKY_SOFT",
    "#BAE6FD": "SKY_BORDER",
    "#FEF08A": "HIGHLIGHT_YELLOW",
    "#FFEDD5": "ORANGE_SOFT",
    "#FDBA74": "ORANGE_BORDER",
    "#C2410C": "ORANGE_DARK",
    "#DCEBFA": "ACCENT_SOFT_PRESSED",
    "#7DD3FC": "SKY_BORDER_STRONG",
    "#FEF9C3": "WARNING_YELLOW_SOFT",
    "#854D0E": "WARNING_YELLOW_TEXT",
    "#78350F": "WARNING_BROWN_DEEP",
    "#E8F5E9": "DIFF_ADDED_BG",
    "#FFEBEE": "DIFF_REMOVED_BG",
    "#FFF9C4": "DIFF_MODIFIED_BG",
    "#D1FAE5": "EMERALD_SOFT",
    "#6EE7B7": "EMERALD_BORDER",
    "#F0F7FF": "BLUE_TINT_SOFT",
}


def get_bundled_fonts_dir() -> Path:
    """Returns the directory containing bundled .ttf / .otf font assets."""
    try:
        from core.utils.paths import get_project_root
        return get_project_root() / "assets" / "fonts"
    except ImportError:
        return Path(__file__).resolve().parents[2] / "assets" / "fonts"


def load_app_fonts(app: Optional[object] = None) -> List[str]:
    """
    Registers any bundled .ttf or .otf fonts in src/assets/fonts/ with Qt's
    QFontDatabase and updates the QApplication default font if available.

    Must be called AFTER QApplication is instantiated. Safe to call in
    headless or missing-assets environments (falls back to Segoe UI).
    """
    loaded_families: List[str] = []
    try:
        from PyQt5.QtGui import QFont, QFontDatabase
    except Exception as exc:
        logger.debug("Skipping font registration (PyQt5 unavailable): %s", exc)
        return loaded_families

    fonts_dir = get_bundled_fonts_dir()
    if fonts_dir.is_dir():
        for ext in ("*.ttf", "*.otf"):
            for font_file in sorted(fonts_dir.glob(ext)):
                font_id = QFontDatabase.addApplicationFont(str(font_file))
                if font_id >= 0:
                    families = QFontDatabase.applicationFontFamilies(font_id)
                    loaded_families.extend(families)
                else:
                    logger.warning("Failed to load bundled font: %s", font_file)

    if app is not None and hasattr(app, "setFont"):
        available = set(QFontDatabase().families())
        preferred = "Inter" if "Inter" in available else "Segoe UI"
        default_font = QFont(preferred)
        default_font.setPixelSize(T.SIZE_BODY)
        app.setFont(default_font)

    if loaded_families:
        logger.info("Loaded bundled UI fonts: %s", ", ".join(sorted(set(loaded_families))))
    return loaded_families
