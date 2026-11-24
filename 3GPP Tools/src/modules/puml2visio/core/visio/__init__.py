"""Modular PlantUML <-> Visio conversion package."""

from .svg_sanitize import sanitize_svg_for_visio
from .threads import ConverterThread, SvgConverterThread, VisioReaderThread
from .visio_pipeline import (
    append_plantuml_source_page,
    convert_svg_to_vsdx,
    flatten_svg_groups,
    normalize_text_shapes,
    remove_background_shapes,
    resize_page_to_contents,
)

__all__ = [
    "ConverterThread",
    "SvgConverterThread",
    "VisioReaderThread",
    "append_plantuml_source_page",
    "convert_svg_to_vsdx",
    "flatten_svg_groups",
    "normalize_text_shapes",
    "remove_background_shapes",
    "resize_page_to_contents",
    "sanitize_svg_for_visio",
]
