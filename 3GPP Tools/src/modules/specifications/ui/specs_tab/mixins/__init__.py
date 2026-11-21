"""Feature mixins composing SpecificationsTab."""

from __future__ import annotations

from .actions import SpecsTabActionsMixin
from .filters import SpecsTabFiltersMixin
from .row_documents_cell import SpecsTabRowDocumentsCellMixin
from .table_render import SpecsTabTableRenderMixin
from .ui_setup import SpecsTabUiSetupMixin

__all__ = [
    "SpecsTabActionsMixin",
    "SpecsTabFiltersMixin",
    "SpecsTabRowDocumentsCellMixin",
    "SpecsTabTableRenderMixin",
    "SpecsTabUiSetupMixin",
]
