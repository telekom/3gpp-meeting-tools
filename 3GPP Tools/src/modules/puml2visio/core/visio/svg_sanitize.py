"""SVG attribute sanitization prior to Microsoft Visio import."""

import logging
import re
from pathlib import Path
from typing import Callable, Optional


def sanitize_svg_for_visio(
    svg_path: Path,
    log_fn: Optional[Callable[[str, int], None]] = None,
) -> None:
    """Sanitize SVG `<tspan>` tags to prevent Visio import crashes."""
    if not svg_path.exists():
        return
    try:
        with open(svg_path, "r", encoding="utf-8", errors="ignore") as f:
            content = f.read()

        content = re.sub(r'(<tspan[^>]*?)\s+font-family="[^"]*"', r"\1", content)
        content = re.sub(r'(<tspan[^>]*?)\s+baseline-shift="[^"]*"', r"\1", content)
        content = re.sub(r'(<tspan[^>]*?)\s+dy="[^"]*"', r"\1", content)
        content = re.sub(r'(<tspan[^>]*?)\s+dx="[^"]*"', r"\1", content)

        with open(svg_path, "w", encoding="utf-8") as f:
            f.write(content)
    except Exception as e:
        if log_fn is not None:
            log_fn(f"⚠️ Could not sanitize SVG: {e}", logging.WARNING)
