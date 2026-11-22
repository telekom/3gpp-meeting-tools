"""Change History, origin CR/TDoc attribution, clause diff, Word compare, and export mixin."""

from __future__ import annotations

import logging
import os
from pathlib import Path
from typing import Optional

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QMessageBox

from core.ui.toast import notify_info, notify_success
from modules.spec_search.core.spec_clause_exporter import (
    SpecClauseWordExtractorThread,
    get_default_clause_export_path,
)
from modules.spec_search.core.spec_doc_resolver import SpecDocResolver
from modules.spec_search.ui.spec_search_dialogs import SpecClauseDiffDialog

logger = logging.getLogger(__name__)


class SpecSearchTabClauseActionsMixin:
    """Implements clause diffing, origin TDoc lookup, Word/Markdown export, and doc opening."""

    def _open_version_history_dialog(
        self,
        spec_num: str,
        version: Optional[str] = None,
        search_term: Optional[str] = None,
        clause_num: Optional[str] = None,
        sort_by_origin_relevance: bool = False,
    ):
        """Opens the Change History & Integrated TDocs dialog modelessly directly from Spec Search."""
        from modules.specifications.ui.spec_history_dialog import SpecVersionHistoryDialog
        logger.info(f"📜 [Change History] Opening Change History & TDocs browser for TS {spec_num}{f' (v{version})' if version else ''}...")

        term = search_term if search_term is not None else self.search_input.text().strip()
        clause = clause_num if clause_num is not None else getattr(self.inspector, "_current_clause_num", "")

        if hasattr(self, "_history_dialog") and self._history_dialog:
            try:
                self._history_dialog.load_spec(
                    spec_num,
                    version,
                    initial_search=term,
                    initial_clause=clause,
                    sort_by_origin_relevance=sort_by_origin_relevance,
                )
                self._history_dialog.show()
                self._history_dialog.raise_()
                self._history_dialog.activateWindow()
                return
            except Exception:
                pass

        dlg = SpecVersionHistoryDialog(
            spec_number=spec_num,
            initial_version=version,
            parent=self.window() or self,
            specs_db=getattr(self, "specs_db", None),
            initial_search=term,
            initial_clause=clause,
            sort_by_origin_relevance=sort_by_origin_relevance,
        )
        self._history_dialog = dlg
        dlg.show()
        dlg.raise_()
        dlg.activateWindow()

    def _on_find_origin_cr_clicked(self):
        """Launches Origin CR ranking for the active specification and search query."""
        import re as _re
        active_spec = ""
        curr_tab = self.spec_results_tabs.currentWidget()
        if curr_tab and hasattr(curr_tab, "_spec_number") and curr_tab._spec_number:
            active_spec = curr_tab._spec_number
        elif hasattr(self, "_selected_spec_numbers") and self._selected_spec_numbers:
            active_spec = self._selected_spec_numbers[0]
        elif hasattr(self.inspector, "_current_spec") and self.inspector._current_spec:
            active_spec = self.inspector._current_spec

        if not active_spec:
            idx = self.spec_results_tabs.currentIndex()
            if idx >= 0:
                tab_txt = self.spec_results_tabs.tabText(idx)
                m = _re.search(r"\b(\d{2}\.\d{3}(?:-\d{2})?)\b", tab_txt)
                if m:
                    active_spec = m.group(1)

        if not active_spec:
            notify_info(
                self,
                "Select Specification",
                "Please select or search within a specification matrix tab first to find its origin CR.",
            )
            return

        query = self.search_input.text().strip()
        clause = getattr(self.inspector, "_current_clause_num", "")

        self._open_version_history_dialog(
            spec_num=active_spec,
            version=None,
            search_term=query,
            clause_num=clause,
            sort_by_origin_relevance=True,
        )

    def _open_clause_diff_dialog(
        self,
        spec_num: str,
        version: str,
        clause_num: str,
        clause_title: str,
        base_version: Optional[str] = None,
        focus_attribution: bool = False,
    ):
        """Opens the unified clause comparison & analysis dialog (Visual Diff, LLM prompt, Word Track Changes, TDoc Attribution)."""
        dialog = SpecClauseDiffDialog(
            db=self.db,
            spec_number=spec_num,
            current_version=version,
            clause_number=clause_num,
            clause_title=clause_title,
            cache_dir=self.cache_dir,
            specs_db=self.specs_db,
            base_version=base_version,
            focus_attribution=focus_attribution,
            parent=self,
        )
        self._diff_dialog = dialog
        dialog.show()
        dialog.raise_()
        dialog.activateWindow()

    def _handle_identify_origin_tdoc(
        self,
        spec_num: str,
        base_ver: str,
        target_ver: str,
        clause_num: str,
        clause_title: str,
    ):
        """Opens SpecClauseDiffDialog directly focused on the Modifying TDoc Attribution tab."""
        self._open_clause_diff_dialog(
            spec_num=spec_num,
            version=target_ver,
            clause_num=clause_num,
            clause_title=clause_title,
            base_version=base_ver,
            focus_attribution=True,
        )

    def _handle_open_spec_document(self, spec_number: str, version: str, clause_number: str, doc_format: str = "word"):
        """Dispatches specification document opening via SpecDocResolver."""
        SpecDocResolver.open_spec(
            spec_number=spec_number,
            version=version,
            clause_number=clause_number,
            doc_format=doc_format,
            cache_dir=self.cache_dir,
            specs_db=self.specs_db,
            parent_widget=self,
        )

    def _refresh_inspector_doc_menu(self):
        """Updates the Inspector Full Doc PDF action label depending on local PDF availability."""
        act_pdf = getattr(self.inspector, "act_pdf", None)
        if not act_pdf:
            return
        spec = getattr(self.inspector, "_current_spec", "")
        ver = getattr(self.inspector, "_current_version", "")
        clause = getattr(self.inspector, "_current_clause_num", "")
        if not spec or not ver:
            act_pdf.setText("📕 Open Converted PDF (.pdf)")
            return
        try:
            info = SpecDocResolver.resolve_spec_files(spec, ver, self.cache_dir, self.specs_db)
            pdf_target = SpecDocResolver.find_matching_pdf(info, clause)
            if pdf_target and pdf_target.exists():
                act_pdf.setText("📕 Open Converted PDF (.pdf)")
            else:
                act_pdf.setText("📕 Convert && Open PDF (.pdf)...")
        except Exception:
            act_pdf.setText("📕 Open Converted PDF (.pdf)")

    def _start_clause_word_compare(self, spec_num: str, base_ver: str, target_ver: str, clause_num: str, clause_title: str):
        """Extracts clause from both releases and executes Word comparison with Track Changes."""
        from modules.spec_search.core.spec_clause_word_comparator import SpecClauseWordCompareThread
        from PyQt5.QtWidgets import QProgressDialog

        prog = QProgressDialog(f"Comparing Clause {clause_num} in Microsoft Word (Track Changes)...", "Cancel", 0, 100, self)
        prog.setWindowTitle("Word Comparison (Track Changes)")
        prog.setWindowModality(1)
        prog.setValue(5)

        worker = SpecClauseWordCompareThread(
            spec_number=spec_num,
            base_version=base_ver,
            target_version=target_ver,
            clause_number=clause_num,
            clause_title=clause_title,
            cache_dir=self.cache_dir,
            specs_db=self.specs_db,
        )

        def on_progress(val, msg):
            prog.setValue(val)
            prog.setLabelText(msg)

        def on_success(msg):
            prog.close()
            logger.info(f"✅ {msg}")

        def on_error(err):
            prog.close()
            QMessageBox.critical(
                self,
                "Word Comparison Failed",
                f"Could not generate Microsoft Word comparison with Track Changes:\n{err}\n\nTip: You can inspect the clause differences directly in the Diff & AI Analysis dialog."
            )

        worker.progress.connect(on_progress)
        worker.finished_success.connect(on_success)
        worker.error.connect(on_error)
        prog.canceled.connect(worker.requestInterruption)

        self._active_word_workers = getattr(self, "_active_word_workers", [])
        self._active_word_workers.append(worker)
        worker.finished.connect(lambda: self._active_word_workers.remove(worker) if worker in self._active_word_workers else None)

        worker.start()
        prog.exec_()

    def _handle_extract_clause_word(
        self,
        spec_number: str,
        version: str,
        clause_number: str,
        clause_title: str,
        content: str,
        save_as: bool = False,
    ):
        """Extracts clause into a standalone Word (.docx) document with native formatting and tables."""
        from PyQt5.QtWidgets import QFileDialog, QProgressDialog

        dest_path = None
        if save_as:
            default_path = get_default_clause_export_path(spec_number, version, clause_number, ext=".docx")
            chosen, _ = QFileDialog.getSaveFileName(
                self,
                f"Save Clause {clause_number} as Word Document",
                str(default_path),
                "Word Document (*.docx)",
            )
            if not chosen:
                return
            dest_path = Path(chosen)

        prog = QProgressDialog(
            f"Extracting Clause {clause_number} from TS {spec_number} v{version}...",
            "Cancel",
            0,
            100,
            self,
        )
        prog.setWindowTitle("Extracting Clause to Word")
        prog.setWindowModality(Qt.WindowModal)
        prog.setValue(5)

        worker = SpecClauseWordExtractorThread(
            spec_number=spec_number,
            version=version,
            clause_number=clause_number,
            clause_title=clause_title,
            clause_content=content,
            output_path=dest_path,
            cache_dir=self.cache_dir,
            specs_db=self.specs_db,
        )

        def on_progress(val, msg):
            prog.setValue(val)
            prog.setLabelText(msg)

        def on_success(out_file):
            prog.close()
            logger.info(f"Clause excerpt created at: {out_file}")
            try:
                os.startfile(str(out_file))
            except Exception as e:
                logger.warning(f"Could not open file directly: {e}")
                notify_success(
                    self,
                    "Clause Extracted",
                    f"Clause {clause_number} saved to {Path(out_file).name}",
                )

        def on_error(err):
            prog.close()
            QMessageBox.critical(
                self,
                "Clause Extraction Failed",
                f"Could not extract Clause {clause_number} in Word:\n{err}",
            )

        worker.progress.connect(on_progress)
        worker.finished_success.connect(on_success)
        worker.error.connect(on_error)
        prog.canceled.connect(worker.requestInterruption)

        self._active_word_workers = getattr(self, "_active_word_workers", [])
        self._active_word_workers.append(worker)
        worker.finished.connect(
            lambda: self._active_word_workers.remove(worker)
            if worker in self._active_word_workers
            else None
        )

        worker.start()
        prog.exec_()

    def _handle_export_clause_markdown(
        self,
        spec_number: str,
        version: str,
        clause_number: str,
        clause_title: str,
        content: str,
    ):
        """Exports the clause to a standalone Markdown (.md) document file with pictures and diagrams."""
        from PyQt5.QtWidgets import QFileDialog, QProgressDialog

        default_path = get_default_clause_export_path(spec_number, version, clause_number, ext=".md")
        chosen, _ = QFileDialog.getSaveFileName(
            self,
            f"Export Clause {clause_number} to Markdown",
            str(default_path),
            "Markdown File (*.md);;All Files (*)",
        )
        if not chosen:
            return

        dest_path = Path(chosen)
        prog = QProgressDialog(
            f"Exporting Clause {clause_number} to Markdown with diagrams...",
            "Cancel",
            0,
            100,
            self,
        )
        prog.setWindowTitle("Exporting Clause to Markdown")
        prog.setWindowModality(Qt.WindowModal)
        prog.setValue(5)

        worker = SpecClauseWordExtractorThread(
            spec_number=spec_number,
            version=version,
            clause_number=clause_number,
            clause_title=clause_title,
            clause_content=content,
            output_path=dest_path,
            cache_dir=self.cache_dir,
            specs_db=self.specs_db,
            export_format="markdown",
        )

        def on_progress(val, msg):
            prog.setValue(val)
            prog.setLabelText(msg)

        def on_success(out_file):
            prog.close()
            img_count = getattr(worker, "images_count", 0)
            img_msg = f" with {img_count} diagram(s)/image(s) extracted into '{out_file.stem}_media/'" if img_count > 0 else ""
            logger.info(f"Clause exported to Markdown at: {out_file}{img_msg}")
            reply = QMessageBox.question(
                self,
                "Clause Exported to Markdown",
                f"Clause {clause_number} exported successfully to:\n{out_file}{img_msg}\n\nDo you want to open the file?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.Yes,
            )
            if reply == QMessageBox.Yes:
                try:
                    os.startfile(str(out_file))
                except Exception as e:
                    logger.warning(f"Could not open file: {e}")

        def on_error(err):
            prog.close()
            QMessageBox.critical(
                self,
                "Markdown Export Failed",
                f"Could not export Clause {clause_number} to Markdown:\n{err}",
            )

        worker.progress.connect(on_progress)
        worker.finished_success.connect(on_success)
        worker.error.connect(on_error)
        prog.canceled.connect(worker.requestInterruption)

        self._active_word_workers = getattr(self, "_active_word_workers", [])
        self._active_word_workers.append(worker)
        worker.finished.connect(
            lambda: self._active_word_workers.remove(worker)
            if worker in self._active_word_workers
            else None
        )

        worker.start()
        prog.exec_()

    def _handle_show_in_specs_tab(self, spec_number: str, version: str):
        """Emits signal to switch to Specifications tab with spec and version pre-selected."""
        self.navigate_to_specs_tab_requested.emit(spec_number, version)
