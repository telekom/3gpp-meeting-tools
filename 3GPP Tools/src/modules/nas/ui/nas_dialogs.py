"""Backward-compatible shim: spec catalogue helpers and the Specification Import Studio dialog.

The catalogue now lives in `modules.nas.core.spec_catalog` (no Qt dependency) and the dialog in
`modules.nas.ui.spec_import`; existing imports from this module keep working.
"""

from modules.nas.core.spec_catalog import (
    PROTOCOL_SPEC_CATALOG,
    SPEC_BUNDLES,
    get_spec_display_title,
    get_spec_metadata,
    get_spec_title,
    get_spec_tooltip,
    is_importable_spec,
    spec_generation,
    spec_matches_filters,
)
from modules.nas.ui.spec_import import NASVersionSelectDialog

__all__ = [
    "PROTOCOL_SPEC_CATALOG",
    "SPEC_BUNDLES",
    "NASVersionSelectDialog",
    "get_spec_display_title",
    "get_spec_metadata",
    "get_spec_title",
    "get_spec_tooltip",
    "is_importable_spec",
    "spec_generation",
    "spec_matches_filters",
]
