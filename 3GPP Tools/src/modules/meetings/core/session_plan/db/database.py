# --- File: src/modules/meetings/core/session_plan/db/database.py ---
"""SessionPlanDatabase shell: initialization, test-isolated defaults path resolution, WAL connections."""

import contextlib
import logging
import sqlite3
import tempfile
from pathlib import Path
from typing import Optional

from .mixins import (
    SessionPlanDelegationMixin,
    SessionPlanLocationsMixin,
    SessionPlanSchemaDefaultsMixin,
    SessionPlanSnapshotsAnnotationsMixin,
)

logger = logging.getLogger("modules.meetings.core.session_plan.session_plan_db")


class SessionPlanDatabase(
    SessionPlanSchemaDefaultsMixin,
    SessionPlanSnapshotsAnnotationsMixin,
    SessionPlanDelegationMixin,
    SessionPlanLocationsMixin,
):
    """Persistent SQLite database interface for meeting session plans."""

    def __init__(self, db_path: Path, global_defaults_path: Optional[Path] = None):
        self.db_path = Path(db_path)
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        self.global_defaults_path = self._resolve_global_defaults_path(global_defaults_path)
        self._init_tables()
        self._sync_from_global_defaults()

    def _resolve_global_defaults_path(self, explicit_path: Optional[Path]) -> Path:
        """
        Resolves the cross-meeting persistent defaults JSON file.
        - In production, uses `<app_data_root>/delegation_defaults.json` so AI assignments
          and roster persist across all meetings (`<meeting_dir>/Agenda/session_plans.db`).
        - When `db_path` is inside the OS temporary directory (unit tests), scopes the
          defaults file to the top-level temporary test folder so multi-meeting tests
          share defaults across meeting subfolders while remaining 100% isolated from
          production data and other test cases.
        """
        if explicit_path is not None:
            return Path(explicit_path)
        try:
            temp_root = Path(tempfile.gettempdir()).resolve()
            db_resolved = self.db_path.resolve()
            if str(db_resolved).lower().startswith(str(temp_root).lower()):
                rel_parts = db_resolved.relative_to(temp_root).parts
                test_root = temp_root / rel_parts[0] if len(rel_parts) > 1 else self.db_path.parent
                return test_root / "delegation_defaults.json"
        except Exception:
            pass
        try:
            from core.utils.paths import get_app_data_root
            app_root = get_app_data_root()
            app_root.mkdir(parents=True, exist_ok=True)
            return app_root / "delegation_defaults.json"
        except Exception:
            return self.db_path.parent / "delegation_defaults.json"

    @contextlib.contextmanager
    def _get_connection(self):
        conn = sqlite3.connect(self.db_path, timeout=30.0)
        conn.row_factory = sqlite3.Row
        try:
            conn.execute("PRAGMA journal_mode=WAL;")
            conn.execute("PRAGMA busy_timeout=30000;")
            with conn:
                yield conn
        finally:
            conn.close()
