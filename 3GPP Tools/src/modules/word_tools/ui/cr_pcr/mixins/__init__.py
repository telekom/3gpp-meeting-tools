"""Feature mixins composing CrPcrDialog."""

from modules.word_tools.ui.cr_pcr.mixins.generation_and_template import (
    CrPcrGenerationTemplateMixin,
)
from modules.word_tools.ui.cr_pcr.mixins.tree_and_naming import (
    CrPcrTreeNamingMixin,
)
from modules.word_tools.ui.cr_pcr.mixins.ui_setup import CrPcrUiSetupMixin

__all__ = [
    "CrPcrGenerationTemplateMixin",
    "CrPcrTreeNamingMixin",
    "CrPcrUiSetupMixin",
]
