"""Filename preview and Word/Markdown section extraction mixin."""

from pathlib import Path

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QFileDialog, QMessageBox

from modules.word_tools.core.word_config import WordConfig
from modules.word_tools.core.word_excerpt_extractor import WordExcerptThread
from modules.word_tools.core.word_excerpt_naming import generate_differentiated_export_name

from .scanning_and_tree import _get_facade_attr


class ExcerptCardExtractionMixin:
    """Manages export naming previews and background Word/Markdown extraction."""

    def _on_naming_scheme_changed(self):
        scheme = self.excerpt_naming_combo.currentData()
        if scheme:
            WordConfig.set_excerpt_naming_scheme(scheme)
        self._custom_filename_active = False
        self._update_preview_filename(force=True)

    def _on_filename_text_edited(self, text: str):
        self._custom_filename_active = bool(text.strip())

    def _reset_filename_to_auto(self):
        self._custom_filename_active = False
        self._update_preview_filename(force=True)

    def _update_preview_filename(self, force: bool = False):
        if self._custom_filename_active and not force:
            return
        selected = self._selected_excerpt_ids()
        scheme = (
            self.excerpt_naming_combo.currentData()
            if hasattr(self, "excerpt_naming_combo")
            else "source_clauses_timestamp"
        )
        auto_name = generate_differentiated_export_name(
            source_paths=self._excerpt_paths or self._excerpt_input_paths,
            headings=self._excerpt_headings,
            selected_ids=selected,
            scheme=scheme or "source_clauses_timestamp",
            ext=".docx",
        )
        if hasattr(self, "excerpt_filename_edit"):
            self.excerpt_filename_edit.setText(auto_name)

    def _selected_excerpt_ids(self):
        selected = []

        def collect(item):
            heading_id = item.data(0, Qt.UserRole)
            if heading_id and item.checkState(0) == Qt.Checked:
                selected.append(heading_id)
            for i in range(item.childCount()):
                collect(item.child(i))

        for i in range(self.excerpt_tree.topLevelItemCount()):
            collect(self.excerpt_tree.topLevelItem(i))
        return selected

    def _extract_selected_sections(self):
        selected = self._selected_excerpt_ids()
        if not selected:
            QMessageBox.information(self, "No Selection", "Select at least one heading to extract.")
            return
        default_dir = str(Path(self._excerpt_paths[0]).parent) if self._excerpt_paths else ""

        proposed_name = (
            self.excerpt_filename_edit.text().strip()
            if hasattr(self, "excerpt_filename_edit")
            else ""
        )
        if not proposed_name:
            scheme = (
                self.excerpt_naming_combo.currentData()
                if hasattr(self, "excerpt_naming_combo")
                else "source_clauses_timestamp"
            )
            proposed_name = generate_differentiated_export_name(
                source_paths=self._excerpt_paths or self._excerpt_input_paths,
                headings=self._excerpt_headings,
                selected_ids=selected,
                scheme=scheme or "source_clauses_timestamp",
                ext=".docx",
            )
        if not proposed_name.lower().endswith(".docx"):
            proposed_name += ".docx"

        output, _ = QFileDialog.getSaveFileName(
            self,
            "Save Extracted Sections",
            str(Path(default_dir) / proposed_name),
            "Word Document (*.docx)",
        )
        if not output:
            return
        if not output.lower().endswith(".docx"):
            output += ".docx"
        self.excerpt_extract_btn.setEnabled(False)
        self.excerpt_extract_btn.setText("⏳ Extracting...")
        also_md = self.chk_also_markdown.isChecked()
        excerpt_thread_cls = _get_facade_attr("WordExcerptThread", WordExcerptThread)
        self._excerpt_extract_thread = excerpt_thread_cls(
            self._excerpt_paths,
            self._excerpt_headings,
            selected,
            output,
            export_markdown=also_md,
        )
        self.excerpt_progress.setRange(0, 100)
        self.excerpt_progress.setValue(0)
        self.excerpt_progress.setVisible(True)
        self.excerpt_progress_label.setText("Preparing extraction...")
        self.excerpt_progress_label.setVisible(True)
        self._excerpt_extract_thread.finished_path.connect(self._on_excerpt_finished)
        self._excerpt_extract_thread.failed.connect(self._on_excerpt_extract_failed)
        self._excerpt_extract_thread.progress.connect(self._on_excerpt_extract_progress)
        self._excerpt_extract_thread.finished.connect(self._reset_excerpt_button)
        self._excerpt_extract_thread.start()

    def _on_excerpt_extract_progress(self, percent, message):
        if message.startswith("Preparing excerpt from "):
            self.excerpt_progress.setRange(0, 0)
        else:
            if self.excerpt_progress.minimum() == 0 and self.excerpt_progress.maximum() == 0:
                self.excerpt_progress.setRange(0, 100)
            self.excerpt_progress.setValue(percent)
        self.excerpt_progress_label.setText(message)

    def _on_excerpt_finished(self, path):
        thread = self._excerpt_extract_thread
        if thread and getattr(thread, "export_markdown", False) and getattr(thread, "md_path", None):
            md_path = thread.md_path
            img_count = getattr(thread, "images_count", 0)
            img_msg = f"\n• Figures extracted: {img_count}" if img_count else "\n• Figures extracted: 0"
            msg = (
                f"Successfully extracted selected sections!\n\n"
                f"📄 Word Document:\n{path}\n\n"
                f"📝 Markdown Document:\n{md_path}"
                f"{img_msg}"
            )
            QMessageBox.information(self, "Extraction Complete", msg)
        else:
            QMessageBox.information(self, "Extraction Complete", f"Created:\n{path}")

    def _on_excerpt_extract_failed(self, message):
        QMessageBox.critical(self, "Extraction Failed", message)

    def _reset_excerpt_button(self):
        self.excerpt_extract_btn.setText("✂️ Extract Selected Sections")
        self.excerpt_extract_btn.setEnabled(bool(self._excerpt_headings))
