# --- File: src/modules/meetings/ui/tdocs/window/mixins/exports.py ---
"""
Exports: reports/statistics, Markdown/LLM export, package & Excel export, export folders.

Mixin of TDocsWindow: relies on attributes set in TDocsWindow.__init__ / _setup_*.
"""
from pathlib import Path
from PyQt5.QtWidgets import (
    QVBoxLayout,
    QLabel,
    QPushButton,
    QMessageBox,
    QDialog,
    QTextEdit,
)
from core.ui.ui_components import BUTTON_STYLE_TOOLBAR_SECONDARY
from modules.meetings.core.exporters.markdown_exporter import MarkdownExporterThread
from modules.meetings.core.stats.exporter_thread import StatisticsExporterThread
from modules.meetings.ui.tdocs_dialogs import StatisticsSettingsDialog
from modules.meetings.ui.dialogs.tdocs_export_controller import (
    get_package_export_rows,
    export_tdocs_package,
    open_excel_export_dialog,
    export_markdown_visible,
    export_markdown_single,
    open_post_meeting_export_dialog,
    open_pre_meeting_export_dialog,
    launch_cli_for_tdocs,
)
from core.ai.cli_launcher import populate_cli_agent_menu
from core.utils.folder_opener import open_folder as _open_folder
from modules.meetings.ui.tdocs_agenda_controller import start_window_tdocs_caching


class ExportsMixin:
    """Exports: reports/statistics, Markdown/LLM export, package & Excel export, export folders."""

    def _export_reports(self):
        self.export_thread = MarkdownExporterThread(
            self.meeting_dir,
            self.model._data,
            self.docs_ftp_url,
            self.mtg_info,
            agenda_map=getattr(self, "agenda_map", None),
        )
        self.export_thread.finished.connect(lambda s, m: self._on_export_finished(s, m, False))
        self.export_thread.start()

    def _open_stats_config(self):
        StatisticsSettingsDialog(self).exec_()

    def _generate_statistics(self):
        self.stats_btn.setText("⏳ Generating...")
        self.stats_btn.setEnabled(False)

        config = StatisticsSettingsDialog().load_config()
        self.stats_thread = StatisticsExporterThread(self.meeting_dir, self.model._data, self.mtg_info, config)
        self.stats_thread.finished.connect(lambda s, m: self._on_export_finished(s, m, True))
        self.stats_thread.start()

    def _on_export_finished(self, success: bool, msg: str, is_stats: bool):
        if is_stats:
            self.stats_btn.setText("📊 Statistics")
            self.stats_btn.setEnabled(True)

        if not success:
            QMessageBox.warning(self, "Export Failed", msg)
            return

        info_msg = msg if "successfully" in msg.lower() else f"Successfully generated:\n{msg}"
        QMessageBox.information(self, "Export Complete", info_msg)

        target_path = None
        for line in reversed(str(msg).splitlines()):
            candidate = Path(line.strip().strip('"').strip("'"))
            if candidate.exists():
                target_path = candidate
                break

        if not target_path or not target_path.exists():
            fallback_dir = self.meeting_dir / "Export"
            if fallback_dir.exists():
                target_path = fallback_dir

        if target_path and target_path.exists():
            _open_folder(target_path)

    def _open_meeting_folder(self):
        _open_folder(self.meeting_dir)

    def _on_export_menu_about_to_show(self):
        export_dir = self.meeting_dir / "Export"
        has_export = export_dir.is_dir()
        self.export_menu_sep.setVisible(has_export)
        self.act_open_export_dir.setVisible(has_export)

    def _open_export_folder(self):
        export_dir = self.meeting_dir / "Export"
        if export_dir.is_dir():
            _open_folder(export_dir)
        else:
            QMessageBox.information(
                self,
                "Export Folder Not Found",
                f"No export folder exists yet for this meeting:\n{export_dir}\n\n"
                "Run an export first to create it.",
            )

    def _show_unmatched_companies(self):
        unmatched = self.model.get_unmatched_sources()
        if not unmatched:
            QMessageBox.information(self, "All Matched",
                                    "Great news! Every source in this meeting was successfully matched by the CompanySanitizer.")
            return

        dialog = QDialog(self)
        dialog.setWindowTitle(f"Unmatched Companies ({len(unmatched)})")
        dialog.resize(500, 400)

        layout = QVBoxLayout(dialog)
        layout.addWidget(QLabel("The following raw 'Source' strings were not recognized and grouped as 'Other':"))

        text_edit = QTextEdit()
        text_edit.setReadOnly(True)
        text_edit.setPlainText("\n".join(unmatched))
        layout.addWidget(text_edit)

        btn = QPushButton("Close")
        btn.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        btn.clicked.connect(dialog.accept)
        layout.addWidget(btn)

        dialog.exec_()

    def _open_agenda_file(self):
        _open_folder(self.meeting_dir / "Agenda" / "TdocsByAgenda.htm")

    def _cache_all_tdocs(self):
        start_window_tdocs_caching(self)

    def _open_excel(self):
        _open_folder(Path(self.filepath))

    def _export_markdown_visible(self):
        export_markdown_visible(self)

    # Backwards compatibility alias
    _export_llm_visible = _export_markdown_visible

    def _export_markdown_single(self, tdoc_id: str):
        export_markdown_single(self, tdoc_id)

    # Backwards compatibility alias
    _export_llm_single = _export_markdown_single

    def _open_pre_meeting_export_dialog(self):
        open_pre_meeting_export_dialog(self)

    def _open_post_meeting_export_dialog(self):
        open_post_meeting_export_dialog(self)

    def _export_dossier(self, target_rows=None):
        from modules.meetings.ui.dialogs.tdocs_export_controller import open_dossier_export_dialog
        open_dossier_export_dialog(self, target_rows)

    def _launch_cli_for_tdocs(self, rows_data, agent_id: str, task_mode: str = "interactive", custom_question: str = ""):
        launch_cli_for_tdocs(self, rows_data, agent_id, task_mode=task_mode, custom_question=custom_question)

    def _populate_export_cli_submenu(self):
        rows = get_package_export_rows(self)
        populate_cli_agent_menu(
            self.cli_export_submenu,
            on_launch_callback=lambda aid, mode, r=rows: self._launch_cli_for_tdocs(r, aid, mode),
            has_multiple_or_revs=(len(rows) > 1),
            multi_count=len(rows),
            parent=self,
        )

    def _on_md_single_finished(self, success: bool, msg: str):
        self.export_btn.setText("📤 Export ▾")
        self.export_btn.setEnabled(True)
        if success:
            QMessageBox.information(self, "Markdown Export", msg)
            target = getattr(self.md_export_thread, "export_dir", self.meeting_dir / "Export")
            _open_folder(target if target.exists() else self.meeting_dir / "Export")
        else:
            QMessageBox.warning(self, "Export Failed", msg)

    _on_llm_single_finished = _on_md_single_finished

    def _on_md_progress(self, msg: str):
        self.export_btn.setText(f"⏳ {msg}"[:35])

    _on_llm_progress = _on_md_progress

    def _on_md_export_finished(self, success: bool, msg: str):
        self.export_btn.setText("📤 Export ▾")
        self.export_btn.setEnabled(True)
        if success:
            QMessageBox.information(self, "Markdown Export Complete", msg)
            target = getattr(self.md_export_thread, "export_dir", self.meeting_dir / "Export")
            _open_folder(target if target.exists() else self.meeting_dir / "Export")
        else:
            QMessageBox.warning(self, "Export Failed", msg)

    _on_llm_export_finished = _on_md_export_finished

    def _get_package_export_rows(self) -> list:
        return get_package_export_rows(self)

    def _export_tdocs_package(self):
        export_tdocs_package(self)

    def _on_package_export_progress(self, msg: str):
        self.export_btn.setText(f"⏳ {msg}"[:34])

    def _open_excel_export_dialog(self):
        open_excel_export_dialog(self)

    def _on_excel_export_progress(self, msg: str):
        self.export_btn.setText(f"⏳ {msg}"[:25])
