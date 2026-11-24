"""Backward-compatible re-export shim for AssistantTab.

Canonical implementation lives in ``modules.assistant.ui.tab``.
"""

from __future__ import annotations

import logging

from modules.assistant.ui.tab.tab import AssistantTab

logger = logging.getLogger(__name__)

__all__ = ["AssistantTab", "logger"]
