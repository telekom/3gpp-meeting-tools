"""SpecHistoryDatabase orchestrator combining schema/writes, queries, and TDoc cache mixins."""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Optional, Union

from modules.specifications.core.history_db.history_queries import (
    SpecHistoryQueriesMixin,
)
from modules.specifications.core.history_db.schema_and_writes import (
    SpecHistorySchemaAndWritesMixin,
)
from modules.specifications.core.history_db.tdoc_metadata import (
    SpecHistoryTdocMetadataMixin,
)

logger = logging.getLogger(__name__)


class SpecHistoryDatabase(
    SpecHistorySchemaAndWritesMixin,
    SpecHistoryQueriesMixin,
    SpecHistoryTdocMetadataMixin,
):
    """Thread-safe SQLite database manager for 3GPP Specification Change History and TDoc cross-referencing."""

    def __init__(self, db_path: Optional[Union[Path, str]] = None):
        if db_path is None:
            from core.utils.paths import get_project_root

            self.db_path = get_project_root() / "db" / "3gpp_spec_history.db"
        else:
            self.db_path = Path(db_path)
        self.logger = logger
        self._init_db()
