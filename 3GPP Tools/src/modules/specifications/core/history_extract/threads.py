"""Background QThread workers for single-spec and batch Change History extraction."""

import logging
from pathlib import Path
from typing import Optional

from PyQt5.QtCore import QThread, pyqtSignal

from core.utils.paths import get_specs_root
from modules.specifications.core.spec_history_db import SpecHistoryDatabase

from .extractor import SpecHistoryExtractor

logger = logging.getLogger(__name__)


class SpecHistoryExtractThread(QThread):
    """Background worker for extracting change history from a single specification document."""

    finished_success = pyqtSignal(str, int)  # spec_number, records_count
    error = pyqtSignal(str)

    def __init__(
        self,
        spec_number: str,
        db: SpecHistoryDatabase,
        version: Optional[str] = None,
        docx_path: Optional[Path] = None,
        specs_root: Optional[Path] = None,
    ):
        super().__init__()
        self.spec_number = spec_number.strip()
        self.db = db
        self.version = version
        self.docx_path = docx_path
        self.specs_root = specs_root

    def run(self):
        try:
            target_file = self.docx_path
            if not target_file or not target_file.exists():
                target_file = SpecHistoryExtractor.resolve_spec_document(
                    self.spec_number, self.version, self.specs_root
                )

            if not target_file or not target_file.exists():
                self.error.emit(f"No local Word document found for {self.spec_number}")
                return

            records = SpecHistoryExtractor.extract_from_docx(
                target_file, self.spec_number, self.version
            )
            if records:
                count = self.db.upsert_history_records(records)
                self.finished_success.emit(self.spec_number, count)
            else:
                self.finished_success.emit(self.spec_number, 0)
        except Exception as e:
            logger.error(f"Error in SpecHistoryExtractThread for {self.spec_number}: {e}", exc_info=True)
            self.error.emit(str(e))


class SpecHistoryBatchScanThread(QThread):
    """Background worker that scans all locally cached specifications and extracts change history."""

    progress = pyqtSignal(int, int, str)     # current, total, message
    finished_success = pyqtSignal(int, int)  # specs_processed, total_records
    error = pyqtSignal(str)

    def __init__(self, db: SpecHistoryDatabase, specs_root: Optional[Path] = None):
        super().__init__()
        self.db = db
        self.specs_root = Path(specs_root or get_specs_root())

    def run(self):
        try:
            if not self.specs_root.exists():
                self.finished_success.emit(0, 0)
                return

            candidate_dirs = [
                d for d in self.specs_root.iterdir()
                if d.is_dir() and not d.name.startswith(".")
            ]
            total_dirs = len(candidate_dirs)
            if total_dirs == 0:
                self.finished_success.emit(0, 0)
                return

            total_records = 0
            specs_processed = 0

            for idx, spec_dir in enumerate(candidate_dirs):
                if self.isInterruptionRequested():
                    break

                spec_num = spec_dir.name
                self.progress.emit(idx + 1, total_dirs, f"Scanning {spec_num}...")

                doc_file = SpecHistoryExtractor.resolve_spec_document(spec_num, specs_root=self.specs_root)
                if doc_file and doc_file.exists():
                    try:
                        records = SpecHistoryExtractor.extract_from_docx(doc_file, spec_num)
                        if records:
                            count = self.db.upsert_history_records(records)
                            total_records += count
                            specs_processed += 1
                    except Exception as spec_err:
                        logger.warning(f"Skipping {spec_num} during batch history scan: {spec_err}")

            self.finished_success.emit(specs_processed, total_records)
        except Exception as e:
            logger.error(f"Error in SpecHistoryBatchScanThread: {e}", exc_info=True)
            self.error.emit(str(e))
