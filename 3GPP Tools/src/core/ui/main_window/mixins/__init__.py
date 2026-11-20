# --- File: src/core/ui/main_window/mixins/__init__.py ---
"""Feature mixins composing DragDropUI (see ../README.md)."""
from core.ui.main_window.mixins.shell_setup import ShellSetupMixin
from core.ui.main_window.mixins.panels_navigation import PanelsNavigationMixin
from core.ui.main_window.mixins.puml_conversion import PumlConversionMixin
from core.ui.main_window.mixins.system_services import SystemServicesMixin

__all__ = ["ShellSetupMixin", "PanelsNavigationMixin", "PumlConversionMixin", "SystemServicesMixin"]
