"""Backward-compatible re-export shim for NAS UI delegates, cards, and inspector widgets.

The implementation is organized under ``modules.nas.ui.components``.
"""

from __future__ import annotations

from .components import (
    RE_EXPLICIT_SPEC_REF,
    RE_NAS_CROSS_REF,
    NASInspectorWidget,
    NASMatrixTreeDelegate,
    NASMessageItemDelegate,
    NASReleasePillWidget,
    NASSpecCardWidget,
    NASVersionCardsWidget,
    NASVersionTreeWidget,
    _extract_major_release,
    _NASCardCompatChild,
    _NASCardCompatSpec,
)

__all__ = [
    "NASInspectorWidget",
    "NASMatrixTreeDelegate",
    "NASMessageItemDelegate",
    "NASReleasePillWidget",
    "NASSpecCardWidget",
    "NASVersionCardsWidget",
    "NASVersionTreeWidget",
    "RE_EXPLICIT_SPEC_REF",
    "RE_NAS_CROSS_REF",
    "_NASCardCompatChild",
    "_NASCardCompatSpec",
    "_extract_major_release",
]
