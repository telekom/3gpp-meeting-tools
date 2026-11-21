"""Unified Chairman's Notes & Agreements Parser facade class."""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import docx

from .block_parser import ChairNotesBlockParserMixin
from .cache_and_scan import ChairNotesCacheAndScanMixin
from .paragraph_agreements import ChairNotesParagraphAgreementsMixin
from .table_parser import ChairNotesTableParserMixin

logger = logging.getLogger(__name__)


class ChairNotesAgreementsParser(
    ChairNotesBlockParserMixin,
    ChairNotesTableParserMixin,
    ChairNotesParagraphAgreementsMixin,
    ChairNotesCacheAndScanMixin,
):
    """
    Parses `.docx` and `.htm`/`.html` Chairman's Notes into:
      1. `agenda_data`: Dict[tdoc_id, dict] compatible with `TDocsTableModel.merge_agenda_data`
      2. `agreements`: List[dict] of structured `MeetingAgreement` records
    """

    @classmethod
    def parse_docx_with_agreements(
        cls,
        docx_path: str,
        mtg_info: Optional[dict] = None,
        agenda_map: Optional[dict] = None,
    ) -> Tuple[Dict[str, dict], List[dict]]:
        path = Path(docx_path)
        if not path.exists():
            logger.error(f"[ChairNotesAgreementsParser] File not found: {docx_path}")
            return {}, []

        try:
            doc = docx.Document(str(path))
        except Exception as e:
            logger.error(f"[ChairNotesAgreementsParser] Failed to open docx '{docx_path}': {e}")
            return {}, []

        mtg_info = mtg_info or {}
        agenda_map = agenda_map or {}
        source_file = path.name

        # Detect whether this document has a main multi-row TDoc table (RAN3 / SA2)
        # or is a document-order Heading + Paragraph + 1x1 Boxed Table document (RAN2 / RAN1)
        main_tables = cls._find_main_tdoc_tables(doc)

        if main_tables:
            agenda_data, agreements = cls._parse_table_based_docx(
                doc, main_tables, source_file, mtg_info, agenda_map
            )
        else:
            agenda_data, agreements = cls._parse_block_based_docx(
                doc, source_file, mtg_info, agenda_map
            )

        logger.info(
            f"[ChairNotesAgreementsParser] Parsed {len(agenda_data)} TDoc entries and "
            f"{len(agreements)} structured agreements from {source_file}"
        )
        return agenda_data, agreements
