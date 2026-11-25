"""Working-group presets and shared button/card QSS for TopicDossierDialog."""

import html
from typing import Dict, List

from core.ui.theme import C, R

ALL_COMMON_WGS = ["SA2", "RAN2", "RAN3", "CT1", "RAN1", "SA3", "CT3", "CT4", "SA1", "SA4", "SA5", "RAN4"]

STYLE_CARD = f"""
    QFrame#CardFrame {{
        background-color: {C.BG_CARD};
        border: 1px solid {C.BORDER_SUBTLE};
        border-radius: 8px;
    }}
"""

STYLE_STATUS_BANNER = f"""
    QFrame#IncStatusBanner {{
        background-color: {C.BG_SURFACE};
        border: 1px solid {C.BORDER_SUBTLE};
        border-radius: {R.MD}px;
    }}
    QFrame#IncStatusBanner QLabel {{
        background: transparent;
        border: none;
        font-size: 11px;
    }}
"""

STYLE_PROGRESS_BAR = f"""
    QProgressBar {{
        background-color: {C.BG_MUTED};
        border: none;
        border-radius: 4px;
    }}
    QProgressBar::chunk {{
        background-color: {C.ACCENT_PRIMARY};
        border-radius: 4px;
    }}
"""

STYLE_CANCEL_BTN = f"""
    QPushButton {{
        background-color: {C.BG_SURFACE};
        color: {C.DANGER_DARK};
        border: 1px solid {C.DANGER_BORDER};
        border-radius: {R.MD}px;
        padding: 3px 10px;
        font-size: 11px;
        font-weight: 500;
    }}
    QPushButton:hover {{
        background-color: {C.DANGER_SOFT};
        border-color: {C.DANGER_BORDER_HOVER};
    }}
    QPushButton:disabled {{
        color: {C.TEXT_FAINT};
        border-color: {C.BORDER_SUBTLE};
    }}
"""


def format_wg_badges_html(wgs: List[str]) -> str:
    """Renders compact WG-family colored inline badges for QLabel rich text."""
    spans: List[str] = []
    for raw in wgs:
        wg = str(raw or "").strip().upper()
        if not wg:
            continue
        if wg.startswith("SA"):
            bg, fg = C.WG_SA_BG, C.WG_SA_TEXT
        elif wg.startswith("RAN"):
            bg, fg = C.WG_RAN_BG, C.WG_RAN_TEXT
        elif wg.startswith("CT"):
            bg, fg = C.WG_CT_BG, C.WG_CT_TEXT
        else:
            bg, fg = C.WG_DEFAULT_BG, C.WG_DEFAULT_TEXT
        safe_wg = html.escape(wg)
        spans.append(
            f"<span style='background-color: {bg}; color: {fg}; "
            f"font-weight: bold; font-size: 10px;'>&nbsp;{safe_wg}&nbsp;</span>"
        )
    return " ".join(spans)


def format_meeting_num(mtg: Dict[str, object]) -> str:
    """Formats a meeting number preserving '-bis' disambiguation from folder_name."""
    num = str(mtg.get("meeting_number") or "").strip()
    folder = str(mtg.get("folder_name") or "").strip().lower()
    if "bis" in folder and "bis" not in num.lower():
        return f"{num}-bis" if num else "bis"
    return num


def format_meeting_label(mtg: Dict[str, object]) -> str:
    """Formats a meeting display label like 'SA2 #178' or 'RAN3 #133-bis'."""
    wg = str(mtg.get("wg_name") or "").strip().upper()
    num = format_meeting_num(mtg)
    return f"{wg} #{num}" if num else wg


def build_incremental_banner_html(
    *,
    total_mtgs: int,
    dec_cnt: int,
    doc_cnt: int,
    stub_cnt: int,
    analyzed_cnt: int,
    stale_cnt: int,
    pending_cnt: int,
    no_act_cnt: int,
    syn_badge: str,
    syn_color: str,
) -> str:
    """Builds the semantic rich-text summary for the Existing Corpus Detected banner."""
    stage1_parts = [
        f"<span style='color: {C.SUCCESS_DARK}; font-weight: bold;'>✅ {analyzed_cnt} Analyzed</span>"
    ]
    if stale_cnt:
        stage1_parts.append(
            f"<span style='color: {C.WARNING_DARK}; font-weight: bold;'>🔄 {stale_cnt} Stale</span>"
        )
    stage1_parts.append(f"<span style='color: {C.TEXT_MUTED};'>⏳ {pending_cnt} Pending</span>")
    if no_act_cnt:
        stage1_parts.append(f"<span style='color: {C.TEXT_MUTED};'>⚪ {no_act_cnt} No Activity</span>")

    stub_note = (
        f" (<span style='color: {C.WARNING_DARK}; font-weight: bold;'>{stub_cnt} stubs</span>)"
        if stub_cnt
        else ""
    )
    return (
        f"<span style='color: {C.TEXT_STRONG}; font-weight: bold;'>🔄 Existing Corpus Detected:</span> "
        f"{total_mtgs} meetings · {dec_cnt} decisions · {doc_cnt} docs{stub_note}<br>"
        f"<b>LLM Stage-1:</b> {' · '.join(stage1_parts)} &nbsp;|&nbsp; "
        f"<span style='color: {syn_color}; font-weight: bold;'>{syn_badge}</span>"
    )

STYLE_PRIMARY_BTN = f"""
    QPushButton {{
        background-color: {C.EMERALD_ACCENT};
        color: {C.BG_CARD};
        font-weight: bold;
        border: 1px solid {C.EMERALD_PRIMARY};
        border-radius: 6px;
        padding: 6px 16px;
        font-size: 12px;
    }}
    QPushButton:hover {{
        background-color: {C.EMERALD_PRIMARY};
    }}
    QPushButton:disabled {{
        background-color: {C.TEXT_FAINT};
        border-color: {C.TEXT_FAINT};
    }}
"""

STYLE_SECONDARY_BTN = f"""
    QPushButton {{
        background-color: {C.BG_SURFACE};
        color: {C.TEXT_STRONG};
        border: 1px solid {C.BORDER_STRONG};
        border-radius: 6px;
        padding: 5px 12px;
        font-size: 11px;
        font-weight: 500;
    }}
    QPushButton:hover {{
        background-color: {C.BG_CANVAS};
        border-color: {C.TEXT_FAINT};
    }}
    QPushButton:disabled {{
        color: {C.TEXT_FAINT};
        background-color: {C.BG_SURFACE};
        border-color: {C.BORDER_SUBTLE};
    }}
"""
