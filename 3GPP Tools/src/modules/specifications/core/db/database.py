"""SpecsDatabase facade assembling schema, ingestion, query, and annotation mixins."""

from __future__ import annotations

import logging
from pathlib import Path

from modules.specifications.core.db.annotations import SpecsAnnotationsMixin
from modules.specifications.core.db.ingestion import SpecsIngestionMixin
from modules.specifications.core.db.queries import SpecsQueryMixin
from modules.specifications.core.db.schema import SpecsSchemaMixin


class SpecsDatabase(
    SpecsSchemaMixin,
    SpecsIngestionMixin,
    SpecsQueryMixin,
    SpecsAnnotationsMixin,
):
    """SQLite storage layer for 3GPP specifications, archive files, WI links, and user annotations."""

    def __init__(self, db_path: Path):
        self.db_path = db_path
        self.logger = logging.getLogger(__name__)
        self._init_db()
        self._cleanup_orphans()
