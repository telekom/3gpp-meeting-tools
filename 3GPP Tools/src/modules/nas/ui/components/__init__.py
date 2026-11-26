"""Modular NAS / Protocol Inspector UI components package."""

from __future__ import annotations

from .delegates import NASMatrixTreeDelegate, NASMessageItemDelegate
from .inspector import NASInspectorWidget
from .inspector_render import RE_EXPLICIT_SPEC_REF, RE_NAS_CROSS_REF
from .packet_stack_widget import PacketStackWidget
from .spec_card import (
    NASReleasePillWidget,
    NASSpecCardWidget,
    _extract_major_release,
)
from .version_cards import (
    NASVersionCardsWidget,
    _NASCardCompatChild,
    _NASCardCompatSpec,
)
from .version_tree import NASVersionTreeWidget

__all__ = [
    "NASInspectorWidget",
    "NASMatrixTreeDelegate",
    "NASMessageItemDelegate",
    "NASReleasePillWidget",
    "NASSpecCardWidget",
    "NASVersionCardsWidget",
    "NASVersionTreeWidget",
    "PacketStackWidget",
    "RE_EXPLICIT_SPEC_REF",
    "RE_NAS_CROSS_REF",
    "_NASCardCompatChild",
    "_NASCardCompatSpec",
    "_extract_major_release",
]
