"""Version loading, visual diff rendering, AI prompt export, and Word Track Changes mixin."""

import html
import logging
from typing import Tuple

from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtWidgets import QApplication, QFileDialog, QMessageBox

from core.ui.theme import C
from modules.spec_search.core.spec_clause_diff import (
    build_llm_clause_prompt,
    generate_unified_diff,
)

logger = logging.getLogger(__name__)


class _DiffPreviewMixin:
    """Computes visual HTML diffs, builds AI prompts, and launches Word comparison."""

    def _load_available_versions(self):
        """Populates version dropdowns for the specification."""
        vers = self.db.get_versions_for_spec(self.spec_number)
        if not vers:
            return

        self.base_ver_combo.blockSignals(True)
        self.target_ver_combo.blockSignals(True)

        self.base_ver_combo.clear()
        self.target_ver_combo.clear()

        for v in vers:
            ver_str = v["version"]
            date_str = f" ({v['release_date']})" if v.get("release_date") else ""
            label = f"v{ver_str}{date_str}"
            self.base_ver_combo.addItem(label, v)
            self.target_ver_combo.addItem(label, v)

        # Default Target Version to current_version
        target_idx = 0
        for i in range(self.target_ver_combo.count()):
            v_data = self.target_ver_combo.itemData(i)
            if v_data and v_data.get("version") == self.current_version:
                target_idx = i
                break
        self.target_ver_combo.setCurrentIndex(target_idx)

        # Default Base Version: use initial_base_version if provided, else preceding version
        base_idx = -1
        if self.initial_base_version:
            for i in range(self.base_ver_combo.count()):
                v_data = self.base_ver_combo.itemData(i)
                if v_data and v_data.get("version") == self.initial_base_version:
                    base_idx = i
                    break

        if base_idx < 0:
            base_idx = min(target_idx + 1, self.base_ver_combo.count() - 1)

        self.base_ver_combo.setCurrentIndex(base_idx)

        self.base_ver_combo.blockSignals(False)
        self.target_ver_combo.blockSignals(False)

    def _format_diff_to_html(
        self, diff_text: str, base_ver: str, target_ver: str
    ) -> Tuple[str, int, int]:
        """
        Converts unified diff text into syntax-highlighted HTML.
        Returns: (html_str, additions_count, deletions_count)
        """
        if not diff_text or "No text differences detected" in diff_text:
            empty_html = (
                f"<div style='font-family: Segoe UI, sans-serif; padding: 18px; color: {C.SUCCESS_DARK}; "
                f"background-color: {C.SUCCESS_SOFT}; border: 1px solid {C.SUCCESS_BORDER}; border-radius: 6px; margin: 8px;'>"
                f"<h3 style='margin: 0 0 6px 0; font-size: 14px;'>✓ No Text Differences Detected</h3>"
                f"<p style='margin: 0; font-size: 12px; color: {C.SUCCESS};'>"
                f"Clause {html.escape(self.clause_number)} is identical between "
                f"<b>3GPP TS {html.escape(self.spec_number)} v{html.escape(base_ver)}</b> and "
                f"<b>v{html.escape(target_ver)}</b>."
                f"</p></div>"
            )
            return empty_html, 0, 0

        html_lines = []
        additions = 0
        deletions = 0

        for line in diff_text.splitlines():
            escaped = html.escape(line)
            if line.startswith("+++ ") or line.startswith("--- "):
                html_lines.append(
                    f"<div style='color: {C.TEXT_MUTED}; font-weight: bold; font-family: Consolas, monospace; line-height: 1.4;'>{escaped}</div>"
                )
            elif line.startswith("@@"):
                html_lines.append(
                    f"<div style='background-color: {C.INDIGO_SOFT}; color: {C.WG_SA_TEXT}; font-weight: bold; font-family: Consolas, monospace; "
                    f"padding: 3px 6px; margin: 4px 0; border-radius: 3px;'>{escaped}</div>"
                )
            elif line.startswith("+"):
                additions += 1
                html_lines.append(
                    f"<div style='background-color: {C.SUCCESS_SOFT_ALT}; color: {C.SUCCESS}; font-family: Consolas, monospace; padding: 1px 6px; white-space: pre-wrap;'>{escaped}</div>"
                )
            elif line.startswith("-"):
                deletions += 1
                html_lines.append(
                    f"<div style='background-color: {C.DANGER_SOFT_HOVER}; color: {C.DANGER}; font-family: Consolas, monospace; padding: 1px 6px; white-space: pre-wrap;'>{escaped}</div>"
                )
            else:
                html_lines.append(
                    f"<div style='color: {C.TEXT_STRONG}; font-family: Consolas, monospace; padding: 1px 6px; white-space: pre-wrap;'>{escaped}</div>"
                )

        full_html = (
            "<div style='font-family: Consolas, Courier New, monospace; font-size: 11px; line-height: 1.45;'>"
            + "".join(html_lines)
            + "</div>"
        )
        return full_html, additions, deletions

    def _generate_diff_preview(self):
        """Generates the unified diff and updates both the Visual Diff and the AI prompt preview."""
        base_data = self.base_ver_combo.currentData()
        target_data = self.target_ver_combo.currentData()

        if not base_data or not target_data:
            self.diff_stats_lbl.setText("<b>Diff Summary:</b> Select both Base and Target releases.")
            self.diff_browser.setText("Select both Base and Target specification releases.")
            self.prompt_browser.setText("Select both Base and Target specification releases.")
            return

        base_ver = base_data.get("version", "")
        base_date = base_data.get("release_date", "")
        target_ver = target_data.get("version", "")
        target_date = target_data.get("release_date", "")

        if hasattr(self, "btn_view_changes") and target_ver:
            self.btn_view_changes.setText(f"📜 v{target_ver} Changes / TDocs")
            self.btn_view_changes.setToolTip(f"View all integrated CRs, p-CRs, and TDocs that created v{target_ver}, and directly open the source contribution documents.")

        tier = self.tier_group.checkedId()
        focus_mode = self.focus_combo.currentData() or "standards"

        # Fetch Clause Text
        base_clause = self.db.get_clause_content_by_spec_ver(self.spec_number, base_ver, self.clause_number)
        target_clause = self.db.get_clause_content_by_spec_ver(self.spec_number, target_ver, self.clause_number)

        base_text = base_clause.get("content", "") if base_clause else ""
        target_text = target_clause.get("content", "") if target_clause else ""

        base_lbl = f"TS {self.spec_number} v{base_ver} (Clause {self.clause_number})"
        target_lbl = f"TS {self.spec_number} v{target_ver} (Clause {self.clause_number})"
        self._raw_diff_text = generate_unified_diff(base_text, target_text, base_lbl, target_lbl)

        # 1. Update Visual Diff Tab
        diff_html, additions, deletions = self._format_diff_to_html(self._raw_diff_text, base_ver, target_ver)
        self.diff_browser.setHtml(diff_html)

        if additions == 0 and deletions == 0:
            self.diff_stats_lbl.setText(
                f"<b>Diff Summary:</b> <span style='color: {C.SUCCESS_DARK}; font-weight: bold;'>✓ Identical</span> "
                f"<span style='color: {C.TEXT_MUTED};'>(TS {self.spec_number} v{base_ver} ➔ v{target_ver})</span>"
            )
        else:
            self.diff_stats_lbl.setText(
                f"<b>Diff Summary:</b> "
                f"<span style='color: {C.SUCCESS}; font-weight: bold;'>+{additions} addition(s)</span> &nbsp;|&nbsp; "
                f"<span style='color: {C.DANGER}; font-weight: bold;'>-{deletions} deletion(s)</span> &nbsp;&nbsp; "
                f"<span style='color: {C.TEXT_MUTED};'>(TS {self.spec_number} v{base_ver} ➔ v{target_ver})</span>"
            )

        # 2. Update AI Analysis Prompt Tab
        hierarchy = None
        if tier >= 2:
            hierarchy = self.db.get_clause_hierarchy(self.spec_number, target_ver, self.clause_number)

        branch_clauses = None
        if tier >= 3:
            branch_clauses = self.db.get_branch_clauses(self.spec_number, target_ver, self.clause_number)

        self._generated_prompt = build_llm_clause_prompt(
            spec_number=self.spec_number,
            clause_number=self.clause_number,
            clause_title=self.clause_title,
            base_version=base_ver,
            base_date=base_date,
            target_version=target_ver,
            target_date=target_date,
            diff_text=self._raw_diff_text,
            tier=tier,
            hierarchy=hierarchy,
            branch_clauses=branch_clauses,
            focus_mode=focus_mode,
        )

        self.prompt_browser.setPlainText(self._generated_prompt)

        # Invalidate cached attribution when version selection changes
        self._cached_attribution_results = None
        if hasattr(self, "tab_widget") and self.tab_widget.currentIndex() == 2:
            self._run_attribution()

    def _copy_diff(self):
        """Copies the raw unified diff text to clipboard."""
        if self._raw_diff_text:
            QApplication.clipboard().setText(self._raw_diff_text)
            self.btn_copy_diff.setText("✅ Copied Diff!")
            QTimer.singleShot(1500, lambda: self.btn_copy_diff.setText("📋 Copy Diff"))

    def _open_version_changes(self):
        """Opens Change History dialog modelessly pre-filtered to version B with directly openable TDocs."""
        from modules.specifications.ui.spec_history_dialog import SpecVersionHistoryDialog
        logger.info(f"📜 [Change History] Opening Change History & TDocs for TS {self.spec_number} v{self.version_b}...")
        dlg = SpecVersionHistoryDialog(
            spec_number=self.spec_number,
            initial_version=self.version_b,
            parent=self.window() or self,
            specs_db=self.specs_db,
            initial_clause=self.clause_number,
        )
        self._history_dialog = dlg
        dlg.show()
        dlg.raise_()
        dlg.activateWindow()

    def _copy_prompt(self):
        """Copies the formatted LLM prompt to clipboard."""
        if self._generated_prompt:
            QApplication.clipboard().setText(self._generated_prompt)
            self.btn_copy_prompt.setText("✅ Copied to Clipboard!")
            QTimer.singleShot(1500, lambda: self.btn_copy_prompt.setText("📋 Copy Prompt for LLM"))

    def _save_markdown(self):
        """Saves the LLM prompt to a Markdown file."""
        if not self._generated_prompt:
            return
        default_name = f"TS_{self.spec_number}_Clause_{self.clause_number}_Diff.md".replace(".", "_")
        path, _ = QFileDialog.getSaveFileName(self, "Save Clause Diff Markdown", default_name, "Markdown Files (*.md)")
        if path:
            try:
                with open(path, "w", encoding="utf-8") as f:
                    f.write(self._generated_prompt)
                QMessageBox.information(self, "Saved", f"Prompt saved successfully to:\n{path}")
            except Exception as e:
                QMessageBox.critical(self, "Save Error", f"Failed to write file:\n{e}")

    def _start_word_comparison(self):
        """Extracts this clause from Base and Target releases and runs native Word comparison with Track Changes."""
        base_data = self.base_ver_combo.currentData()
        target_data = self.target_ver_combo.currentData()
        if not base_data or not target_data:
            QMessageBox.warning(self, "Selection Required", "Please select both Base and Target specification releases.")
            return

        base_ver = base_data.get("version", "")
        target_ver = target_data.get("version", "")

        from modules.spec_search.core.spec_clause_word_comparator import SpecClauseWordCompareThread
        from PyQt5.QtWidgets import QProgressDialog

        prog = QProgressDialog(f"Comparing Clause {self.clause_number} in Microsoft Word (Track Changes)...", "Cancel", 0, 100, self)
        prog.setWindowTitle("Word Comparison (Track Changes)")
        prog.setWindowModality(Qt.WindowModal)
        prog.setValue(5)

        self._word_cmp_worker = SpecClauseWordCompareThread(
            spec_number=self.spec_number,
            base_version=base_ver,
            target_version=target_ver,
            clause_number=self.clause_number,
            clause_title=self.clause_title,
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
                f"Could not generate Microsoft Word comparison with Track Changes:\n{err}\n\n"
                f"Tip: You can still inspect the full text differences directly in the 'Visual Diff' tab."
            )

        self._word_cmp_worker.progress.connect(on_progress)
        self._word_cmp_worker.finished_success.connect(on_success)
        self._word_cmp_worker.error.connect(on_error)
        prog.canceled.connect(self._word_cmp_worker.requestInterruption)

        self._word_cmp_worker.start()
        prog.exec_()
