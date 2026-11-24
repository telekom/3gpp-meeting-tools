"""Background QThread worker for converting Markdown files/URLs to .docx."""

from __future__ import annotations

import logging
import os
import tempfile
from pathlib import Path
from typing import Optional, Union

from PyQt5.QtCore import QThread, pyqtSignal

from core.network.session import NetworkSession
from core.utils.paths import strip_extended_path
from modules.word_tools.core.document_opener import open_document
from modules.word_tools.core.md_to_docx.converter import convert_markdown_to_docx

logger = logging.getLogger(__name__)


class MarkdownToWordThread(QThread):
    """
    Background worker for compiling Markdown to Word with embedded media and diagrams.
    Compatible with the QueueManager task registry.
    """

    ui_log_msg = pyqtSignal(str, int)
    finished_path = pyqtSignal(str)
    conversion_failed = pyqtSignal(str)
    progress = pyqtSignal(int, str)

    def __init__(
        self,
        doc_source: str,
        output_dir: Optional[Union[str, Path]] = None,
        companion_media_dir: Optional[Union[str, Path]] = None,
        auto_detect_media: bool = True,
        render_diagrams: bool = True,
        apply_3gpp_styles: bool = True,
        auto_open: bool = False,
    ):
        super().__init__()
        self.doc_source = str(doc_source)
        self.output_dir = Path(output_dir).resolve() if output_dir else None
        self.companion_media_dir = Path(companion_media_dir).resolve() if companion_media_dir else None
        self.auto_detect_media = auto_detect_media
        self.render_diagrams = render_diagrams
        self.apply_3gpp_styles = apply_3gpp_styles
        self.auto_open = auto_open
        self._downloaded_temp: Optional[Path] = None

    def _log(self, message: str, level: int = logging.INFO) -> None:
        logger.log(level, message)
        self.ui_log_msg.emit(message, level)

    def _resolve_source_path(self, input_str: str) -> Path:
        clean_input = input_str.strip()
        if not clean_input:
            raise ValueError("Markdown document path is empty.")

        if clean_input.startswith("http://") or clean_input.startswith("https://"):
            self._log("⏳ Downloading Markdown document from URL...", logging.INFO)
            session = NetworkSession.get_instance()
            r = session.get(clean_input, allow_redirects=True, timeout=30)
            r.raise_for_status()

            fd, tmp_str = tempfile.mkstemp(prefix="puml2word_", suffix=".md")
            os.close(fd)
            tmp_path = Path(tmp_str)
            tmp_path.write_bytes(r.content)
            self._downloaded_temp = tmp_path
            return tmp_path

        p = Path(clean_input).resolve()
        if not p.is_file():
            raise FileNotFoundError(f"Markdown file does not exist: {p}")
        return p

    def run(self) -> None:
        self._downloaded_temp = None
        try:
            source_path = self._resolve_source_path(self.doc_source)

            if self.output_dir:
                self.output_dir.mkdir(parents=True, exist_ok=True)
                target_docx = self.output_dir / f"{source_path.stem}.docx"
            else:
                target_docx = source_path.with_suffix(".docx")

            res = convert_markdown_to_docx(
                md_path=source_path,
                output_docx_path=target_docx,
                companion_media_dir=self.companion_media_dir,
                auto_detect_media=self.auto_detect_media,
                render_diagrams=self.render_diagrams,
                apply_3gpp_styles=self.apply_3gpp_styles,
                log_callback=lambda msg: self._log(msg, logging.INFO),
            )

            out_path = res["output_path"]
            clean_out = strip_extended_path(out_path)

            if self.auto_open:
                try:
                    open_document(clean_out)
                    self._log(f"👁️ Auto-opened '{Path(clean_out).name}' in Microsoft Word.", logging.INFO)
                except Exception as e:
                    self._log(f"⚠️ Could not auto-open Word: {e}", logging.WARNING)

            self.finished_path.emit(clean_out)

        except Exception as e:
            err_msg = str(e)
            self._log(f"❌ Markdown to Word conversion failed: {err_msg}", logging.ERROR)
            self.conversion_failed.emit(err_msg)

        finally:
            if self._downloaded_temp and self._downloaded_temp.exists():
                try:
                    self._downloaded_temp.unlink(missing_ok=True)
                except Exception:
                    pass
