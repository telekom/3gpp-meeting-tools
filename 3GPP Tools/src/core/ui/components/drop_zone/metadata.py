"""Design tokens, file-type metadata archetypes, and size formatting helpers."""

from __future__ import annotations

from core.ui.theme import C

# Design Tokens (matching 3GPP Tools Slate/Blue design language)
COLOR_TEXT_PRIMARY = C.TEXT_PRIMARY       # Slate 900
COLOR_TEXT_SECONDARY = C.TEXT_SECONDARY   # Slate 600
COLOR_TEXT_MUTED = C.TEXT_MUTED           # Slate 500

COLOR_IDLE_BG = C.BG_SURFACE              # Slate 50
COLOR_IDLE_BORDER = C.BORDER_STRONG       # Slate 300
COLOR_IDLE_HOVER_BORDER = C.BORDER_HOVER  # Slate 400

COLOR_DRAG_BG = C.INFO_SOFT               # Blue 50
COLOR_DRAG_BORDER = C.ACCENT_ALT          # Blue 600
COLOR_DRAG_ACCENT = C.INFO                # Blue 700

COLOR_LOADED_BG = C.BG_CARD               # Pure White
COLOR_LOADED_BORDER = C.BORDER_SUBTLE     # Slate 200

COLOR_BUSY_BG = C.WARNING_SOFT            # Amber 50
COLOR_BUSY_BORDER = C.WARNING_ACCENT      # Amber 500

COLOR_ERROR_BG = C.DANGER_SOFT            # Red 50
COLOR_ERROR_BORDER = C.DANGER_ACCENT      # Red 500

# File Type Archetypes (Badge Color, Background, Icon)
FILE_TYPE_METADATA = {
    ".docx": {"icon": "📘", "name": "Word Document", "bg": C.INFO_SOFT, "fg": C.INFO, "border": C.INFO_BORDER},
    ".doc": {"icon": "📘", "name": "Legacy Word Doc", "bg": C.INFO_SOFT, "fg": C.INFO, "border": C.INFO_BORDER},
    ".docm": {"icon": "📘", "name": "Word Macro-Enabled", "bg": C.INFO_SOFT, "fg": C.INFO, "border": C.INFO_BORDER},
    ".md": {"icon": "📝", "name": "Markdown Document", "bg": C.VIOLET_SURFACE, "fg": C.VIOLET_PRIMARY, "border": C.VIOLET_BORDER},
    ".markdown": {"icon": "📝", "name": "Markdown Document", "bg": C.VIOLET_SURFACE, "fg": C.VIOLET_PRIMARY, "border": C.VIOLET_BORDER},
    ".puml": {"icon": "📐", "name": "PlantUML Diagram", "bg": C.WG_RAN_BG, "fg": C.EMERALD_DARK, "border": C.WG_RAN_BORDER},
    ".plantuml": {"icon": "📐", "name": "PlantUML Diagram", "bg": C.WG_RAN_BG, "fg": C.EMERALD_DARK, "border": C.WG_RAN_BORDER},
    ".pptx": {"icon": "📙", "name": "PowerPoint Deck", "bg": C.WG_CT_BG, "fg": C.ORANGE_DARK, "border": C.WG_CT_BORDER},
    ".ppt": {"icon": "📙", "name": "Legacy PowerPoint", "bg": C.WG_CT_BG, "fg": C.ORANGE_DARK, "border": C.WG_CT_BORDER},
    ".vsdx": {"icon": "🔷", "name": "Visio Drawing", "bg": C.SKY_SURFACE, "fg": C.ACCENT_SKY_DARK, "border": C.SKY_BORDER},
    ".vsd": {"icon": "🔷", "name": "Legacy Visio Drawing", "bg": C.SKY_SURFACE, "fg": C.ACCENT_SKY_DARK, "border": C.SKY_BORDER},
    ".pdf": {"icon": "📕", "name": "PDF Document", "bg": C.DANGER_SOFT, "fg": C.DANGER, "border": C.DANGER_BORDER},
    ".txt": {"icon": "📄", "name": "Text Document", "bg": C.BG_SURFACE, "fg": C.TEXT_BODY, "border": C.BORDER_STRONG},
    ".zip": {"icon": "🗜️", "name": "Zip Archive", "bg": C.BG_CANVAS, "fg": C.TEXT_BODY, "border": C.BORDER_STRONG},
}

DEFAULT_FILE_TYPE = {
    "icon": "📄",
    "name": "Document",
    "bg": C.BG_SURFACE,
    "fg": C.TEXT_SECONDARY,
    "border": C.BORDER_STRONG,
}


def format_file_size(size_bytes: int) -> str:
    """Format bytes into human-readable B, KB, MB, or GB."""
    if size_bytes < 1024:
        return f"{size_bytes} B"
    kb = size_bytes / 1024.0
    if kb < 1024:
        return f"{kb:.1f} KB"
    mb = kb / 1024.0
    if mb < 1024:
        return f"{mb:.2f} MB"
    gb = mb / 1024.0
    return f"{gb:.2f} GB"
