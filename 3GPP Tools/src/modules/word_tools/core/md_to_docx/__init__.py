"""Modular Markdown-to-Word conversion package."""

from modules.word_tools.core.md_to_docx.converter import convert_markdown_to_docx
from modules.word_tools.core.md_to_docx.inline_and_lists import (
    _apply_run_style,
    _embed_image_tag,
    _process_list,
    add_runs_from_element,
)
from modules.word_tools.core.md_to_docx.media_renderer import (
    CALLOUT_CONFIG,
    PLANTUML_START_DIRECTIVES,
    _extract_plantuml_title,
    render_metafile_to_png,
    render_plantuml_to_png,
    render_svg_to_png,
    resolve_image_path,
)
from modules.word_tools.core.md_to_docx.worker import MarkdownToWordThread

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
    "render_metafile_to_png",
    "render_plantuml_to_png",
    "render_svg_to_png",
    "resolve_image_path",
]
