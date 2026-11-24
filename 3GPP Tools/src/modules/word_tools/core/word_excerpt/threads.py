"""Background QThreads for Word heading scanning and dual DOCX/Markdown excerpt export."""

from __future__ import annotations

import logging
import os
import shutil
import sys
from pathlib import Path
from typing import Optional

from PyQt5.QtCore import QThread, pyqtSignal

from core.utils.paths import to_extended_path
from .engine import WordExcerptEngine
from .scanner import WordHeadingScanner, order_sources

logger = logging.getLogger(__name__)


class WordExcerptThread(QThread):
    """Background worker extracting selected clauses to DOCX and optional Markdown+media."""

    finished_path = pyqtSignal(str)
    failed = pyqtSignal(str)
    progress = pyqtSignal(int, str)

    def __init__(
        self,
        source_paths,
        headings,
        selected_ids,
        output_path,
        export_markdown: bool = False,
        revisions: str = "accept",
    ):
        super().__init__()
        self.source_paths = source_paths
        self.headings = headings
        self.selected_ids = selected_ids
        self.output_path = output_path
        self.export_markdown = bool(export_markdown)
        self.revisions = revisions
        self.md_path: Optional[Path] = None
        self.images_count: int = 0

    def run(self):
        try:
            logger.info("Extracting %d selected heading(s)...", len(self.selected_ids))

            def _scale_progress(percent: int, message: str):
                scaled = int(percent * 0.85) if self.export_markdown else percent
                self.progress.emit(scaled, message)

            facade = sys.modules.get("modules.word_tools.core.word_excerpt_extractor")
            engine_cls = (
                getattr(facade, "WordExcerptEngine", WordExcerptEngine)
                if facade is not None
                else WordExcerptEngine
            )

            out = engine_cls(
                self.source_paths,
                self.headings,
                progress_callback=_scale_progress,
            ).extract(self.selected_ids, self.output_path)

            if self.export_markdown:
                self.progress.emit(
                    88, "Converting excerpt to Markdown and extracting figures..."
                )
                from modules.meetings.core.tdocs.tdoc_markdown_converter import (
                    convert_document_to_markdown,
                )

                docx_file = Path(out).resolve()
                out_dir = docx_file.parent
                stem = docx_file.stem
                media_dir = out_dir / f"{stem}_media"
                media_rel = f"{stem}_media"

                exported = convert_document_to_markdown(
                    docx_file,
                    out_dir=out_dir,
                    tdoc_id=stem,
                    revisions=self.revisions,
                    media_dir=media_dir,
                    media_rel=media_rel,
                )
                self.md_path = (
                    Path(exported.md_path)
                    if exported and exported.md_path
                    else out_dir / f"{stem}.md"
                )
                self.images_count = getattr(exported, "images", 0)

                if media_dir.exists():
                    try:
                        if self.images_count == 0 or not os.listdir(
                            to_extended_path(media_dir)
                        ):
                            shutil.rmtree(to_extended_path(media_dir), ignore_errors=True)
                    except Exception as clean_err:
                        logger.debug("Failed checking/pruning media dir: %s", clean_err)

                self.progress.emit(
                    100,
                    f"Extraction complete (Markdown & {self.images_count} figure(s) extracted).",
                )
            else:
                self.progress.emit(100, "Extraction complete.")

            logger.info(
                "Excerpt created: %s (markdown=%s, images=%d)",
                out,
                self.export_markdown,
                self.images_count,
            )
            self.finished_path.emit(str(out))
        except Exception as exc:
            logger.exception("Word excerpt extraction failed")
            self.failed.emit(str(exc))


class WordHeadingScanThread(QThread):
    """Background worker parsing Word headings across ordered source documents."""

    scanned = pyqtSignal(object, object)  # ordered paths, headings
    failed = pyqtSignal(str)
    progress = pyqtSignal(int, str)

    def __init__(self, source_paths, order_mode="filename"):
        super().__init__()
        self.source_paths = list(source_paths)
        self.order_mode = order_mode

    def run(self):
        try:
            ordered = order_sources(self.source_paths, self.order_mode)

            def report(file_index, total_files, block_index, total_blocks, path):
                file_fraction = block_index / max(1, total_blocks)
                percent = int(((file_index + file_fraction) / max(1, total_files)) * 100)
                percent = max(0, min(99, percent))
                self.progress.emit(
                    percent,
                    f"Parsing {Path(path).name} - document {file_index + 1} of {total_files}",
                )

            headings = WordHeadingScanner().scan(ordered, "drop", report)
            self.progress.emit(100, f"Parsed {len(ordered)} document(s)")
            self.scanned.emit(ordered, headings)
        except Exception as exc:
            logger.exception("Heading scan failed")
            self.failed.emit(str(exc))
