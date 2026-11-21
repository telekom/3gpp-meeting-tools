"""Modular external agent domain tool handlers and registry package."""

from __future__ import annotations

from .common import ToolArgumentSpec, _resolve_db_paths
from .registry import PROFILE_TOOLS, build_external_agent_tool_registry

__all__ = [
    "PROFILE_TOOLS",
    "ToolArgumentSpec",
    "_resolve_db_paths",
    "build_external_agent_tool_registry",
]
