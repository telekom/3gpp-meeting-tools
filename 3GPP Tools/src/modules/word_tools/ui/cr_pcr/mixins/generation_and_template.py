"""Contribution generation worker lifecycle and 3gpp_70.dotx template controls mixin."""

from pathlib import Path

from PyQt5.QtWidgets import QFileDialog, QMessageBox

from core.ui.ui_components import BADGE_STYLE_MUTED
from modules.word_tools.core.cr_pcr_generator import (
    ContributionMetadata,
    CrPcrGeneratorThread,
)
from modules.word_tools.core.document_opener import open_document
from modules.word_tools.core.styles_3gpp import (
    install_3gpp_dotx_template,
    is_3gpp_dotx_installed,
    uninstall_3gpp_dotx_template,
)
from core.ui.theme import C


class CrPcrGenerationTemplateMixin:
    """Runs CrPcrGeneratorThread and manages 3gpp_70.dotx Word template installation."""

    def _generate_contribution(self):
        selected_ids = self._get_selected_ids()
        if not selected_ids:
            QMessageBox.information(self, "No Clauses Selected", "Please select at least one clause.")
            return

        # Prepare ContributionMetadata
        is_cr = self.current_mode == "cr"
        cat_code = self.combo_cat.currentText().split(" - ")[0].strip() if is_cr else "F"

        meta = ContributionMetadata(
            mode=self.current_mode,
            meeting_wg=self.input_wg.text().strip(),
            meeting_num=self.input_mtg_num.text().strip(),
            meeting_dates=self.input_dates.text().strip(),
            meeting_location=self.input_location.text().strip(),
            tdoc_id=self.input_tdoc_id.text().strip() or "Draft-TDoc",
            title=self.input_title.text().strip() or f"Change to {self.spec_type} {self.spec_num}",
            source_company=self.input_source.text().strip() or "Delegate",
            doc_for=self.combo_doc_for.currentText() if not is_cr else "Approval",
            agenda_item=self.input_ai.text().strip(),
            work_item=self.input_work_item.text().strip(),
            abstract=self.text_abstract.toPlainText().strip() if not is_cr else "",
            discussion_text=self.text_discussion.toPlainText().strip() if not is_cr else "",
            spec_type=self.spec_type,
            spec_num=self.spec_num,
            spec_version=self.version,
            spec_release=self.combo_release.currentText() if is_cr else self.release,
            cr_num=self.input_cr_num.text().strip() if is_cr else "<CR#>",
            cr_rev=self.input_cr_rev.text().strip() if is_cr else "-",
            category=cat_code,
            reason_for_change=self.text_reason.toPlainText().strip() if is_cr else "",
            summary_of_change=self.text_summary.toPlainText().strip() if is_cr else "",
            consequences=self.text_consequences.toPlainText().strip() if is_cr else "",
            affects_uicc=self.chk_uicc.isChecked() if is_cr else False,
            affects_me=self.chk_me.isChecked() if is_cr else False,
            affects_ran=self.chk_ran.isChecked() if is_cr else False,
            affects_cn=self.chk_cn.isChecked() if is_cr else False,
        )

        proposed_name = self.filename_edit.text().strip() if hasattr(self, "filename_edit") else ""
        if not proposed_name:
            proposed_name = self._compute_auto_filename()
        if not proposed_name.lower().endswith(".docx"):
            proposed_name += ".docx"

        default_out_path = str(self.spec_dl_dir / proposed_name)

        output_path, _ = QFileDialog.getSaveFileName(
            self,
            f"Save {self.current_mode.upper()} Document",
            default_out_path,
            "Word Document (*.docx)",
        )

        if not output_path:
            return

        if not output_path.lower().endswith(".docx"):
            output_path += ".docx"

        self.btn_generate.setEnabled(False)
        self.btn_generate.setText("⏳ Generating...")
        self.progress_bar.setValue(0)
        self.progress_bar.setVisible(True)
        self.status_label.setText("Starting contribution generation...")

        self._gen_thread = CrPcrGeneratorThread(
            metadata=meta,
            source_paths=self._ordered_paths,
            headings=self._headings,
            selected_ids=selected_ids,
            output_path=output_path,
            all_new_map=self._all_new_map,
        )
        self._gen_thread.progress.connect(self._on_gen_progress)
        self._gen_thread.finished_path.connect(self._on_gen_finished)
        self._gen_thread.failed.connect(self._on_gen_failed)
        self._gen_thread.finished.connect(lambda: self.btn_generate.setEnabled(True))
        self._gen_thread.start()

    def _on_gen_progress(self, pct: int, msg: str):
        self.progress_bar.setValue(pct)
        self.status_label.setText(msg)

    def _on_gen_finished(self, out_path_str: str):
        p = Path(out_path_str)
        self.progress_bar.setValue(100)
        self.status_label.setText(f"✅ Generated: {p.name}")

        self.contribution_generated.emit(p)

        if self.open_word_cb.isChecked() and p.exists():
            open_document(p)

        QMessageBox.information(
            self,
            "Generation Complete",
            f"Successfully generated {self.current_mode.upper()} document:\n{p}",
        )

    def _on_gen_failed(self, err_msg: str):
        self.progress_bar.setValue(100)
        self.status_label.setText("⚠️ Generation failed.")
        QMessageBox.critical(self, "Generation Error", f"Contribution generation failed:\n{err_msg}")

    def _update_dotx_indicator(self):
        installed = is_3gpp_dotx_installed()
        if installed:
            self.dotx_badge.setText("🟢 Installed (3gpp_70.dotx)")
            self.dotx_badge.setStyleSheet(f"""
                QLabel {{
                    font-size: 11px;
                    font-weight: bold;
                    color: {C.WG_RAN_TEXT};
                    background-color: {C.EMERALD_SOFT};
                    border: 1px solid {C.WG_RAN_BORDER};
                    border-radius: 10px;
                    padding: 3px 8px;
                }}
            """)
            self.dotx_badge.setToolTip("3gpp_70.dotx is installed in %APPDATA%\\Microsoft\\Templates")
            self.btn_dotx_action.setText("🗑️ Remove Template")
            self.btn_dotx_action.setToolTip("Remove 3gpp_70.dotx from your Microsoft Word templates folder")
        else:
            self.dotx_badge.setText("⚪ Not Installed")
            self.dotx_badge.setStyleSheet(BADGE_STYLE_MUTED)
            self.dotx_badge.setToolTip("3gpp_70.dotx is not installed in %APPDATA%\\Microsoft\\Templates")
            self.btn_dotx_action.setText("📥 Install Template")
            self.btn_dotx_action.setToolTip("Download and install 3gpp_70.dotx into %APPDATA%\\Microsoft\\Templates")

    def _on_toggle_dotx_clicked(self):
        if is_3gpp_dotx_installed():
            reply = QMessageBox.question(
                self,
                "Remove 3GPP Word Template",
                "Do you want to remove 3gpp_70.dotx from your Microsoft Word templates folder?\n\n"
                "(%APPDATA%\\Microsoft\\Templates\\3gpp_70.dotx)\n\n"
                "Word will no longer automatically attach the 3GPP macro ribbon to newly created documents.",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No,
            )
            if reply == QMessageBox.Yes:
                ok, msg = uninstall_3gpp_dotx_template()
                if ok:
                    QMessageBox.information(self, "Template Removed", msg)
                else:
                    QMessageBox.warning(self, "Removal Failed", msg)
                self._update_dotx_indicator()
        else:
            ok, msg = install_3gpp_dotx_template(force=True)
            if ok:
                QMessageBox.information(self, "Template Installed", msg)
            else:
                QMessageBox.warning(self, "Installation Failed", msg)
            self._update_dotx_indicator()
