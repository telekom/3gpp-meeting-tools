"""Clause extraction worker invocation and CR/pCR draft launching mixin."""

from __future__ import annotations

from pathlib import Path

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QFileDialog, QMessageBox

from modules.word_tools.core.document_opener import open_document
from modules.word_tools.core.word_config import WordConfig
from modules.word_tools.core.word_excerpt_extractor import WordExcerptThread
from modules.word_tools.core.word_excerpt_naming import generate_differentiated_export_name


class SpecClauseExtractionMixin:
    """Runs asynchronous Word/Markdown extraction and launches modeless CrPcrDialog."""

    def _extract_selected_clauses(self):
        selected_ids = self._get_selected_ids()
        if not selected_ids:
            QMessageBox.information(
                self, "No Clauses Selected", "Please select at least one clause to extract."
            )
            return

        clean_spec = self.spec_num.replace(".", "_")
        scheme = WordConfig.get_excerpt_naming_scheme()
        prefix = f"{self.spec_type}_{clean_spec}_v{self.version}"
        default_out_name = generate_differentiated_export_name(
            source_paths=self._ordered_paths,
            headings=self._headings,
            selected_ids=selected_ids,
            scheme=scheme,
            custom_prefix=prefix,
            ext=".docx",
        )
        default_out_path = str(self.spec_dl_dir / default_out_name)

        output_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save Extracted Clauses",
            default_out_path,
            "Word Document (*.docx)",
        )

        if not output_path:
            return

        if not output_path.lower().endswith(".docx"):
            output_path += ".docx"

        self.btn_extract.setEnabled(False)
        self.btn_extract.setText("⏳ Extracting...")
        self.progress_bar.setValue(0)
        self.progress_bar.setVisible(True)
        self.status_label.setText("Preparing clause extraction...")

        self._extract_thread = WordExcerptThread(
            source_paths=self._ordered_paths,
            headings=self._headings,
            selected_ids=selected_ids,
            output_path=output_path,
            export_markdown=self.chk_also_markdown.isChecked(),
        )
        self._extract_thread.progress.connect(self._on_extract_progress)
        self._extract_thread.finished_path.connect(self._on_extract_finished)
        self._extract_thread.failed.connect(self._on_extract_failed)
        self._extract_thread.finished.connect(self._on_extract_done)
        self._extract_thread.start()

    def _on_extract_progress(self, percent: int, message: str):
        if message.startswith("Preparing excerpt from "):
            self.progress_bar.setRange(0, 0)
        else:
            if self.progress_bar.minimum() == 0 and self.progress_bar.maximum() == 0:
                self.progress_bar.setRange(0, 100)
            self.progress_bar.setValue(percent)
        self.status_label.setText(message)

    def _on_extract_finished(self, out_path_str: str):
        p = Path(out_path_str)
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(100)
        self.status_label.setText(f"✅ Excerpt saved: {p.name}")

        self.extraction_completed.emit(p)

        if self.open_after_cb.isChecked() and p.exists():
            open_document(p)

        thread = self._extract_thread
        if thread and getattr(thread, "export_markdown", False) and getattr(thread, "md_path", None):
            md_path = thread.md_path
            img_count = getattr(thread, "images_count", 0)
            img_msg = (
                f"\n• Figures extracted: {img_count}"
                if img_count
                else "\n• Figures extracted: 0"
            )
            msg = (
                f"Successfully extracted selected clauses!\n\n"
                f"📄 Word Document:\n{p}\n\n"
                f"📝 Markdown Document:\n{md_path}"
                f"{img_msg}"
            )
            QMessageBox.information(self, "Extraction Complete", msg)
        else:
            QMessageBox.information(
                self,
                "Extraction Complete",
                f"Successfully extracted selected clauses to:\n{p}",
            )

    def _on_extract_failed(self, err_msg: str):
        self.progress_bar.setRange(0, 100)
        self.status_label.setText("⚠️ Extraction failed.")
        QMessageBox.critical(self, "Extraction Error", f"Extraction failed:\n{err_msg}")

    def _on_extract_done(self):
        self._update_extract_button()

    def _open_cr_pcr_from_selection(self):
        from modules.word_tools.ui.cr_pcr_dialog import CrPcrDialog

        selected_ids = self._get_selected_ids()
        dlg = CrPcrDialog(
            spec_num=self.spec_num,
            spec_type=self.spec_type,
            version=self.version,
            fname=self.fname,
            url=self.url,
            spec_dl_dir=self.spec_dl_dir,
            ordered_paths=self._ordered_paths,
            headings=self._headings,
            parent=self._parent_ref or self.parent(),
        )
        dlg._populate_tree()
        dlg._is_updating_checks = True
        try:
            for i in range(dlg.clause_tree.topLevelItemCount()):
                item = dlg.clause_tree.topLevelItem(i)

                def restore_checks(it):
                    h_id = it.data(0, Qt.UserRole)
                    if h_id and h_id in selected_ids:
                        it.setCheckState(0, Qt.Checked)
                    for ci in range(it.childCount()):
                        restore_checks(it.child(ci))

                restore_checks(item)
        finally:
            dlg._is_updating_checks = False
        dlg._update_change_block_summary()
        if not hasattr(self, "_active_cr_dialogs"):
            self._active_cr_dialogs = []
        self._active_cr_dialogs.append(dlg)
        dlg.destroyed.connect(
            lambda: self._active_cr_dialogs.remove(dlg) if dlg in self._active_cr_dialogs else None
        )
        dlg.show()
        dlg.raise_()
        dlg.activateWindow()
