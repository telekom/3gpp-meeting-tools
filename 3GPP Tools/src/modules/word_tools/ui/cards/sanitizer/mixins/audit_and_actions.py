"""Preset handling, document inspection, author selection, and sanitizer trigger mixin."""

from __future__ import annotations

import logging
import os
import re
from pathlib import Path

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QFileDialog, QMessageBox, QTableWidgetItem

from modules.word_tools.core.document_sanitizer import (
    SanitizeConfig,
    close_word_document_for_update,
    inspect_document,
    is_file_open_in_word,
    sanitize_document,
)
from modules.word_tools.core.word_config import WordConfig

logger = logging.getLogger(__name__)


class SanitizerActionsMixin:
    """Implements document auditing, quick presets, author selection, and sanitization actions."""

    def _on_target_lang_combo_changed(self, index: int):
        data = self.combo_target_lang.currentData()
        self.edit_custom_lang.setVisible(data == "custom")

    def _apply_preset_3gpp(self):
        self.chk_sec_lang.setChecked(True)
        idx = self.combo_target_lang.findData("en-GB")
        if idx >= 0:
            self.combo_target_lang.setCurrentIndex(idx)
        self.chk_lang_all_stories.setChecked(True)
        self.chk_lang_styles.setChecked(True)
        self.chk_lang_purge_revisions.setChecked(True)
        self.chk_sec_authors.setChecked(True)
        self._apply_author_preset(
            WordConfig.load().get("default_source_company", "Deutsche Telekom")
        )
        self.chk_sec_cleanup.setChecked(False)

    def _apply_preset_lang_only(self):
        self.chk_sec_lang.setChecked(True)
        idx = self.combo_target_lang.findData("en-GB")
        if idx >= 0:
            self.combo_target_lang.setCurrentIndex(idx)
        self.chk_lang_all_stories.setChecked(True)
        self.chk_lang_styles.setChecked(True)
        self.chk_lang_purge_revisions.setChecked(True)
        self.chk_sec_authors.setChecked(False)
        self.chk_sec_cleanup.setChecked(False)

    def _apply_preset_anonymize(self):
        self.chk_sec_lang.setChecked(True)
        idx = self.combo_target_lang.findData("en-GB")
        if idx >= 0:
            self.combo_target_lang.setCurrentIndex(idx)
        self.chk_sec_authors.setChecked(True)
        self._apply_author_preset("Contributor", "")
        self.chk_sec_cleanup.setChecked(True)
        self.chk_remove_comments.setChecked(True)

    def _on_sanitizer_file_selected(self, files: list):
        if not files:
            return
        self._inspect_selected_author_document(files[0])

    def _inspect_selected_author_document(self, file_path: str):
        path = Path(file_path)
        if not path.is_file():
            return
        if path.suffix.lower() not in (".docx", ".docm"):
            self.lbl_audit_summary.setText(
                "⚠️ Please select a modern Word document (.docx or .docm). Legacy .doc is not supported."
            )
            return

        self._last_author_file = str(path.resolve())
        self.author_result_frame.setVisible(False)

        try:
            audit = inspect_document(path)
            self._current_author_stats = audit.authors
            self._populate_author_table(audit.authors)

            self.lbl_kpi_lang.setText(f"🌐 Languages: {audit.language_stats.summary_text}")
            self.lbl_kpi_lang_rev.setText(
                f"🏷️ Lang Changes: {audit.language_stats.language_change_marks}"
            )
            self.lbl_kpi_authors.setText(f"👥 Authors: {audit.total_authors}")
            self.lbl_kpi_text_rev.setText(
                f"📝 Ins: {audit.total_insertions} · Del: {audit.total_deletions}"
            )
            self.lbl_kpi_fmt.setText(f"🎨 Formatting: {audit.total_formatting}")
            self.lbl_kpi_comments.setText(f"💬 Comments: {audit.total_comments}")

            self.chk_sec_authors.setChecked(audit.total_authors > 0)

            self.lbl_audit_summary.setText(
                f"✅ Analyzed {path.name} ({audit.file_size_kb:.1f} KB): "
                f"{audit.language_stats.summary_text} · {audit.language_stats.language_change_marks} lang revision marks · "
                f"{audit.total_authors} author(s) · {audit.total_changes} total changes."
            )

        except Exception as exc:
            logger.error("Failed to inspect document %s: %s", path.name, exc)
            self.lbl_audit_summary.setText(f"❌ Failed to inspect document: {exc}")

    def _populate_author_table(self, stats: list):
        self.author_table.setRowCount(len(stats))
        if not stats:
            return

        for row, s in enumerate(stats):
            chk_item = QTableWidgetItem()
            chk_item.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled)
            chk_item.setCheckState(Qt.Checked)
            self.author_table.setItem(row, 0, chk_item)

            name_item = QTableWidgetItem(s.name)
            name_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            self.author_table.setItem(row, 1, name_item)

            for col, val in enumerate(
                [s.insertions, s.deletions, s.formatting, s.comments], start=2
            ):
                num_item = QTableWidgetItem(str(val))
                num_item.setTextAlignment(Qt.AlignCenter)
                num_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
                self.author_table.setItem(row, col, num_item)

            init_item = QTableWidgetItem(s.initials)
            init_item.setTextAlignment(Qt.AlignCenter)
            init_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            self.author_table.setItem(row, 6, init_item)

    def _select_all_authors(self, select: bool):
        for row in range(self.author_table.rowCount()):
            item = self.author_table.item(row, 0)
            if item:
                item.setCheckState(Qt.Checked if select else Qt.Unchecked)

    def _on_target_author_changed(self, text: str):
        suggested = self._derive_initials(text.strip())
        if suggested:
            self.edit_target_initials.setText(suggested)

    def _derive_initials(self, name: str) -> str:
        words = re.findall(r"[A-Za-z0-9]+", name or "")
        if not words:
            return ""
        if len(words) >= 2:
            return "".join(w[0].upper() for w in words[:3])
        return words[0][:3].title()

    def _apply_author_preset(self, author_name: str, initials: str = None):
        self.edit_target_author.setText(author_name)
        if initials is not None:
            self.edit_target_initials.setText(initials)
        else:
            self.edit_target_initials.setText(self._derive_initials(author_name))

    def _trigger_sanitizer(self, in_place: bool):
        source_inputs = self.pane_sanitizer.get_inputs()
        source_file = source_inputs[0] if source_inputs else self._last_author_file
        if not source_file or not Path(source_file).is_file():
            QMessageBox.warning(
                self, "No Document Selected", "Please select or drop a valid .docx file first."
            )
            return

        src_path = Path(source_file).resolve()

        sec_lang = self.chk_sec_lang.isChecked()
        sec_authors = self.chk_sec_authors.isChecked()
        sec_cleanup = self.chk_sec_cleanup.isChecked()

        if not (sec_lang or sec_authors or sec_cleanup):
            QMessageBox.warning(
                self,
                "No Operations Selected",
                "Please enable at least one section (Language, Authors, or Cleanup) to apply.",
            )
            return

        selected_authors = []
        if sec_authors:
            for row in range(self.author_table.rowCount()):
                chk = self.author_table.item(row, 0)
                name_item = self.author_table.item(row, 1)
                if chk and chk.checkState() == Qt.Checked and name_item:
                    selected_authors.append(name_item.text())

            if self.author_table.rowCount() > 0 and not selected_authors:
                QMessageBox.warning(
                    self,
                    "No Authors Selected",
                    "Author replacement is enabled, but no authors are checked in the table.",
                )
                return

        target_author = self.edit_target_author.text().strip()
        new_initials = (
            self.edit_target_initials.text().strip()
            if self.chk_author_initials.isChecked()
            else None
        )
        mapping = {old_name: target_author for old_name in selected_authors} if selected_authors else {}

        lang_data = self.combo_target_lang.currentData()
        if lang_data == "custom":
            target_lang = self.edit_custom_lang.text().strip() or "en-GB"
        else:
            target_lang = str(lang_data or "en-GB")

        reopen_after = False

        if in_place:
            dst_path = src_path
            if is_file_open_in_word(src_path):
                reply = QMessageBox.question(
                    self,
                    "Document Open in Word",
                    f"The document '{src_path.name}' is currently open in Microsoft Word.\n\n"
                    "To update it in-place, Word must save and close this document temporarily.\n\n"
                    "Would you like to save & close the document in Word now, apply sanitization, and re-open it?",
                    QMessageBox.Yes | QMessageBox.No,
                    QMessageBox.Yes,
                )
                if reply == QMessageBox.Yes:
                    closed = close_word_document_for_update(src_path, save_changes=True)
                    if not closed:
                        QMessageBox.warning(
                            self,
                            "File Lock",
                            "Could not automatically close the document in Word. Please save and close Word manually, or use 'Save As New Copy...'.",
                        )
                        return
                    reopen_after = True
                else:
                    return
        else:
            default_dst = src_path.parent / f"{src_path.stem}_sanitized{src_path.suffix}"
            chosen, _ = QFileDialog.getSaveFileName(
                self,
                "Save Sanitized Document As",
                str(default_dst),
                "Word Documents (*.docx *.docm)",
            )
            if not chosen:
                return
            dst_path = Path(chosen).resolve()

        config = SanitizeConfig(
            unify_language=sec_lang,
            target_language=target_lang,
            apply_to_all_stories=self.chk_lang_all_stories.isChecked(),
            update_styles_and_defaults=self.chk_lang_styles.isChecked(),
            purge_language_changes=self.chk_lang_purge_revisions.isChecked(),
            clear_no_proof=self.chk_clear_no_proof.isChecked(),
            update_authors=sec_authors and bool(mapping),
            author_mapping=mapping,
            new_initials=new_initials,
            update_metadata=self.chk_author_metadata.isChecked(),
            purge_all_formatting_changes=sec_cleanup and self.chk_purge_all_fmt.isChecked(),
            remove_all_comments=sec_cleanup and self.chk_remove_comments.isChecked(),
            backup=self.chk_sanitizer_backup.isChecked() if in_place else False,
        )

        success, _count, msg = sanitize_document(
            docx_path=src_path,
            output_path=dst_path,
            config=config,
        )

        if not success:
            QMessageBox.critical(self, "Sanitization Failed", msg)
            return

        self._last_modified_author_file = str(dst_path)
        self.author_result_lbl.setText(f"✅ {msg}")
        self.author_result_frame.setVisible(True)

        if reopen_after:
            from modules.word_tools.core.document_opener import open_document

            try:
                open_document(dst_path)
            except Exception as exc:
                logger.warning("Could not re-open document: %s", exc)

        self._inspect_selected_author_document(str(dst_path))

    def _open_author_result_file(self):
        if self._last_modified_author_file and Path(self._last_modified_author_file).is_file():
            from modules.word_tools.core.document_opener import open_document

            try:
                open_document(self._last_modified_author_file)
            except Exception as exc:
                QMessageBox.warning(self, "Open Failed", f"Could not open document: {exc}")

    def _open_author_result_folder(self):
        if self._last_modified_author_file:
            folder = Path(self._last_modified_author_file).parent
            if folder.is_dir() and hasattr(os, "startfile"):
                os.startfile(str(folder))
