# --- File: src/modules/meetings/core/session_plan/db/__init__.py ---
"""Modular package for the persistent SQLite SessionPlanDatabase."""

from .database import SessionPlanDatabase, logger

__all__ = [
    "SessionPlanDatabase",
    "logger",
]
