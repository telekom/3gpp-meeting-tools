"""Backward-compatible re-export shim for Markdown-to-Word conversion.

Canonical implementation lives in ``modules.word_tools.core.md_to_docx``.
"""

from __future__ import annotations

import logging

from modules.word_tools.core.md_to_docx import (
    CALLOUT_CONFIG,
    MarkdownToWordThread,
    PLANTUML_START_DIRECTIVES,
    _apply_run_style,
    _embed_image_tag,
    _extract_plantuml_title,
    _process_list,
    add_runs_from_element,
    convert_markdown_to_docx,
    render_metafile_to_png,
    render_plantuml_to_png,
    render_svg_to_png,
    resolve_image_path,
)

logger = logging.getLogger(__name__)

__all__ = [
    "CALLOUT_CONFIG",
    "MarkdownToWordThread",
    "PLANTUML_START_DIRECTIVES",
    "_apply_run_style",
    "_embed_image_tag",
    "_extract_plantuml_title",
    "_process_list",
    "add_runs_from_element",
    "convert_markdown_to_docx",
    "logger",
    "render_metafile_to_png",
    "render_plantuml_to_png",
    "render_svg_to_png",
    "resolve_image_path",
]
