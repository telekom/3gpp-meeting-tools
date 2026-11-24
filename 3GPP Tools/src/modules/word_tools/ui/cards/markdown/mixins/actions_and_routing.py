"""Studio mode routing, conversion triggers, result opening, and callbacks mixin."""

from __future__ import annotations

import os
from pathlib import Path

from PyQt5.QtWidgets import QApplication, QFileDialog, QMessageBox

from core.ui.theme import C
from modules.word_tools.core.document_opener import open_document


class MarkdownActionsMixin:
    """Implements direction routing, W2M/M2W triggers, result viewers, and worker callbacks."""

    def set_studio_mode(self, index: int):
        """Sets the active studio mode (0: Word -> Markdown, 1: Markdown -> Word)."""
        idx = max(0, min(index, 1))
        self.studio_stack.setCurrentIndex(idx)
        self.btn_mode_w2m.setChecked(idx == 0)
        self.btn_mode_m2w.setChecked(idx == 1)

    def _toggle_direction(self):
        new_idx = 1 if self.studio_stack.currentIndex() == 0 else 0
        self.set_studio_mode(new_idx)

    def _on_w2m_file_selected(self, files: list):
        if not files or self._routing_files:
            return
        ext = Path(files[0]).suffix.lower()
        if ext in (".md", ".markdown"):
            self._routing_files = True
            try:
                self.pane_markdown.clear_selection()
                self.set_studio_mode(1)
                self.pane_m2w._set_files(files)
            finally:
                self._routing_files = False
            return

        is_presentation = ext in (".pptx", ".ppt")
        self.md_revisions_combo.setEnabled(not is_presentation)
        if is_presentation:
            self.md_revisions_combo.setToolTip(
                "Track Changes is not applicable for PowerPoint presentations."
            )
        else:
            self.md_revisions_combo.setToolTip(
                "Select how Word Track Changes revisions should be rendered in Markdown."
            )

    def _on_m2w_file_selected(self, files: list):
        if not files or self._routing_files:
            return
        ext = Path(files[0]).suffix.lower()
        if ext in (".docx", ".doc", ".docm", ".pptx", ".ppt", ".pdf"):
            self._routing_files = True
            try:
                self.pane_m2w.clear_selection()
                self.set_studio_mode(0)
                self.pane_markdown._set_files(files)
            finally:
                self._routing_files = False

    # =========================================================================
    # Word -> Markdown Execution & Results
    # =========================================================================
    def _trigger_markdown_conversion(self):
        inputs = self.pane_markdown.get_inputs()
        if not inputs:
            QMessageBox.warning(
                self,
                "No Document",
                "Please select or drop a Word document or PowerPoint deck to convert to Markdown.",
            )
            return

        source_file = inputs[0]
        options = {
            "revisions_mode": self.md_revisions_combo.currentData() or "markup",
            "extract_media": self.md_media_cb.isChecked(),
            "include_header": self.md_header_cb.isChecked(),
            "output_dir": self.md_dest_edit.text().strip()
            if self.md_custom_dest_cb.isChecked()
            else None,
        }

        self._awaiting_w2m = True
        self.md_copy_btn.setText("📋 Copy Markdown")
        self.md_result_frame.setVisible(False)
        self.md_progress.setVisible(True)
        self.run_markdown_btn.setEnabled(False)
        self.convert_markdown_requested.emit(source_file, options)

    def _open_md_folder(self):
        if self._last_md_output_path:
            folder = Path(self._last_md_output_path).parent
            if folder.is_dir() and hasattr(os, "startfile"):
                os.startfile(str(folder))

    def _open_md_file(self):
        if self._last_md_output_path and Path(self._last_md_output_path).is_file():
            if hasattr(os, "startfile"):
                os.startfile(str(self._last_md_output_path))

    def _copy_md_to_clipboard(self):
        if self._last_md_output_path and Path(self._last_md_output_path).is_file():
            try:
                with open(self._last_md_output_path, "r", encoding="utf-8", errors="ignore") as f:
                    content = f.read()
                QApplication.clipboard().setText(content)
                self.md_copy_btn.setText("✓ Copied!")
            except Exception as e:
                QMessageBox.warning(self, "Copy Failed", f"Could not read file: {e}")

    def _toggle_md_preview(self):
        is_vis = self.md_preview_edit.isVisible()
        self.md_preview_edit.setVisible(not is_vis)
        self.md_preview_toggle_btn.setText(
            "📝 Hide Preview ▴" if not is_vis else "📝 Show Preview ▾"
        )

    def _on_md_custom_dest_toggled(self, checked: bool):
        self.md_dest_edit.setEnabled(checked)
        self.md_dest_browse_btn.setEnabled(checked)

    def _browse_md_dest(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Output Directory for Markdown")
        if folder:
            self.md_dest_edit.setText(folder)

    # =========================================================================
    # Markdown -> Word Execution & Results
    # =========================================================================
    def _trigger_m2w_conversion(self):
        inputs = self.pane_m2w.get_inputs()
        if not inputs:
            QMessageBox.warning(
                self,
                "No Document",
                "Please select or drop a Markdown (.md) document to convert to Word.",
            )
            return

        source_file = inputs[0]
        options = {
            "companion_media_dir": None,
            "auto_detect_media": self.m2w_media_cb.isChecked(),
            "render_diagrams": self.m2w_diagrams_cb.isChecked(),
            "apply_3gpp_styles": self.m2w_styles_cb.isChecked(),
            "auto_open": self.m2w_auto_open_cb.isChecked(),
            "output_dir": self.m2w_dest_edit.text().strip()
            if self.m2w_custom_dest_cb.isChecked()
            else None,
        }

        self._awaiting_m2w = True
        self.m2w_copy_path_btn.setText("🔗 Copy Path")
        self.m2w_result_frame.setVisible(False)
        self.m2w_progress.setVisible(True)
        self.run_m2w_btn.setEnabled(False)
        self.convert_markdown_to_word_requested.emit(source_file, options)

    def _open_m2w_folder(self):
        if self._last_docx_output_path:
            folder = Path(self._last_docx_output_path).parent
            if folder.is_dir() and hasattr(os, "startfile"):
                os.startfile(str(folder))

    def _open_m2w_file(self):
        if self._last_docx_output_path and Path(self._last_docx_output_path).is_file():
            try:
                open_document(self._last_docx_output_path)
            except Exception as e:
                QMessageBox.warning(self, "Open Failed", f"Could not open document in Word: {e}")

    def _copy_m2w_path_to_clipboard(self):
        if self._last_docx_output_path:
            QApplication.clipboard().setText(self._last_docx_output_path)
            self.m2w_copy_path_btn.setText("✓ Copied!")

    def _on_m2w_custom_dest_toggled(self, checked: bool):
        self.m2w_dest_edit.setEnabled(checked)
        self.m2w_dest_browse_btn.setEnabled(checked)

    def _browse_m2w_dest(self):
        folder = QFileDialog.getExistingDirectory(
            self, "Select Output Directory for Word Document"
        )
        if folder:
            self.m2w_dest_edit.setText(folder)

    # =========================================================================
    # Universal Task Callbacks
    # =========================================================================
    def on_conversion_success(self, out_path: str):
        p = Path(out_path).resolve()
        suffix = p.suffix.lower()

        # Word -> Markdown completion
        if suffix in (".md", ".markdown"):
            self._awaiting_w2m = False
            self.set_studio_mode(0)
            self._last_md_output_path = str(p)
            self._last_converted_md = str(p)
            self.md_progress.setVisible(False)
            self.run_markdown_btn.setEnabled(True)
            self.md_result_icon.setText("✅")
            self.md_result_lbl.setStyleSheet(f"font-weight: bold; color: {C.SUCCESS_DARK};")
            self.md_result_lbl.setText(f"Successfully converted to Markdown: {p.name}")
            self.md_open_folder_btn.setEnabled(True)
            self.md_open_file_btn.setEnabled(True)
            self.md_copy_btn.setEnabled(True)
            self.md_preview_toggle_btn.setEnabled(True)
            self.md_result_frame.setVisible(True)

            if p.is_file():
                try:
                    with open(p, "r", encoding="utf-8", errors="ignore") as f:
                        preview_lines = [f.readline() for _ in range(60)]
                    self.md_preview_edit.setPlainText("".join(preview_lines))
                except Exception:
                    self.md_preview_edit.setPlainText("")

        # Markdown -> Word completion (only handle if triggered from this card or currently in M2W mode)
        elif suffix in (".docx", ".doc") and (
            self._awaiting_m2w or self.studio_stack.currentIndex() == 1
        ):
            self._awaiting_m2w = False
            self.set_studio_mode(1)
            self._last_docx_output_path = str(p)
            self.m2w_progress.setVisible(False)
            self.run_m2w_btn.setEnabled(True)
            self.m2w_result_icon.setText("✅")
            self.m2w_result_lbl.setStyleSheet(f"font-weight: bold; color: {C.SUCCESS_DARK};")
            self.m2w_result_lbl.setText(f"Successfully compiled Word document: {p.name}")
            self.m2w_open_folder_btn.setEnabled(True)
            self.m2w_open_file_btn.setEnabled(True)
            self.m2w_copy_path_btn.setEnabled(True)
            self.m2w_result_frame.setVisible(True)

    def on_conversion_failed(self, error_msg: str):
        if self._awaiting_m2w or (not self._awaiting_w2m and self.studio_stack.currentIndex() == 1):
            self._awaiting_m2w = False
            self.m2w_progress.setVisible(False)
            self.run_m2w_btn.setEnabled(True)
            self.m2w_result_icon.setText("❌")
            self.m2w_result_lbl.setStyleSheet(f"font-weight: bold; color: {C.DANGER_ALT};")
            self.m2w_result_lbl.setText(f"Compilation failed: {error_msg}")
            self.m2w_open_folder_btn.setEnabled(False)
            self.m2w_open_file_btn.setEnabled(False)
            self.m2w_copy_path_btn.setEnabled(False)
            self.m2w_result_frame.setVisible(True)
        else:
            self._awaiting_w2m = False
            self.md_progress.setVisible(False)
            self.run_markdown_btn.setEnabled(True)
            self.md_result_icon.setText("❌")
            self.md_result_lbl.setStyleSheet(f"font-weight: bold; color: {C.DANGER_ALT};")
            self.md_result_lbl.setText(f"Conversion failed: {error_msg}")
            self.md_open_folder_btn.setEnabled(False)
            self.md_open_file_btn.setEnabled(False)
            self.md_copy_btn.setEnabled(False)
            self.md_preview_toggle_btn.setEnabled(False)
            self.md_preview_edit.setVisible(False)
            self.md_result_frame.setVisible(True)

    def on_processing_state_changed(self, is_processing: bool):
        self.run_markdown_btn.setEnabled(not is_processing)
        self.run_m2w_btn.setEnabled(not is_processing)
        if self.studio_stack.currentIndex() == 0:
            self.md_progress.setVisible(is_processing)
        else:
            self.m2w_progress.setVisible(is_processing)
