"""Design-token HTML/CSS styles for TDocsHelpDialog rich-text help tabs."""

from __future__ import annotations

from core.ui.theme import C

HELP_CSS = f"""
<style>
    body {{
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
        color: {C.TEXT_STRONG};
        background-color: {C.BG_CARD};
        font-size: 13px;
        line-height: 1.5;
        margin: 12px;
    }}
    h2 {{
        color: {C.TEXT_PRIMARY};
        font-size: 16px;
        margin-top: 14px;
        margin-bottom: 8px;
        border-bottom: 1.5px solid {C.BORDER_SUBTLE};
        padding-bottom: 4px;
    }}
    .lead {{
        color: {C.TEXT_SECONDARY};
        font-size: 13px;
        margin-bottom: 14px;
    }}
    .card {{
        background: {C.BG_SURFACE};
        border: 1px solid {C.BORDER_SUBTLE};
        border-radius: 6px;
        padding: 10px 14px;
        margin-bottom: 12px;
    }}
    .card-title {{
        font-weight: bold;
        font-size: 14px;
        color: {C.TEXT_PRIMARY};
        margin-bottom: 4px;
    }}
    .badge {{
        display: inline-block;
        padding: 2px 7px;
        font-size: 11px;
        font-weight: 600;
        border-radius: 4px;
        margin-right: 6px;
    }}
    .badge-blue {{ background-color: {C.INFO_SOFT_STRONG}; color: {C.INFO_DARK}; border: 1px solid {C.INFO_BORDER}; }}
    .badge-green {{ background-color: {C.SUCCESS_SOFT_ALT}; color: {C.SUCCESS_DARK}; border: 1px solid {C.SUCCESS_BORDER}; }}
    .badge-amber {{ background-color: {C.WARNING_SOFT_ALT}; color: {C.WARNING_DARK}; border: 1px solid {C.WARNING_BORDER}; }}
    .badge-purple {{ background-color: {C.VIOLET_SOFT}; color: {C.VIOLET_DARK}; border: 1px solid {C.VIOLET_BORDER}; }}
    .badge-slate {{ background-color: {C.BG_CANVAS}; color: {C.TEXT_BODY}; border: 1px solid {C.BORDER_STRONG}; }}

    kbd {{
        display: inline-block;
        padding: 2px 6px;
        font-family: SFMono-Regular, Consolas, "Liberation Mono", Menlo, Courier, monospace;
        font-size: 11px;
        font-weight: 600;
        color: {C.TEXT_STRONG};
        background-color: {C.BG_CANVAS};
        border: 1px solid {C.BORDER_STRONG};
        border-radius: 4px;
        box-shadow: 0 1px 0 rgba(0,0,0,0.1);
        margin: 0 2px;
    }}

    table.help-table {{
        width: 100%;
        border-collapse: collapse;
        margin-top: 8px;
        margin-bottom: 14px;
    }}
    table.help-table th {{
        background-color: {C.BG_CANVAS};
        color: {C.TEXT_BODY};
        font-size: 12px;
        font-weight: 600;
        text-align: left;
        padding: 8px 10px;
        border-bottom: 2px solid {C.BORDER_STRONG};
    }}
    table.help-table td {{
        padding: 8px 10px;
        border-bottom: 1px solid {C.BORDER_SUBTLE};
        font-size: 12.5px;
        vertical-align: top;
    }}
    table.help-table tr:hover td {{
        background-color: {C.BG_SURFACE};
    }}

    .tip-box {{
        background-color: {C.SUCCESS_SOFT};
        border-left: 4px solid {C.GREEN_ACCENT};
        padding: 8px 12px;
        border-radius: 0 6px 6px 0;
        margin-top: 6px;
        font-size: 12px;
        color: {C.SUCCESS_DARK};
    }}
    .note-box {{
        background-color: {C.INFO_SOFT};
        border-left: 4px solid {C.INFO_ACCENT};
        padding: 8px 12px;
        border-radius: 0 6px 6px 0;
        margin-top: 6px;
        font-size: 12px;
        color: {C.INFO_DARK};
    }}
    .warn-box {{
        background-color: {C.WARNING_SOFT};
        border-left: 4px solid {C.WARNING_ACCENT};
        padding: 8px 12px;
        border-radius: 0 6px 6px 0;
        margin-top: 6px;
        font-size: 12px;
        color: {C.WARNING_DARK};
    }}
</style>
"""
