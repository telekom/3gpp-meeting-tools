"""Backward-compatible re-export shim for SpecsDatabase.

Canonical implementation lives in ``modules.specifications.core.db``.
"""

from __future__ import annotations

from modules.specifications.core.db.database import SpecsDatabase

__all__ = ["SpecsDatabase"]
