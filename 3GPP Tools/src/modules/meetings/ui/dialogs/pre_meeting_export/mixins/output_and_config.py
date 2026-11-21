"""Destination directory helpers, validation, and export configuration getters."""

import datetime
from pathlib import Path
from typing import Dict, Optional, Set

from PyQt5.QtWidgets import QFileDialog, QMenu, QMessageBox

from core.ui.theme import C


class PreMeetingOutputConfigMixin:
    """Handles destination folder status, recent exports menu, validation, and config output."""

    def _browse_destination(self):
        d = QFileDialog.getExistingDirectory(self, "Select Export Directory", str(self.meeting_dir / "Export"))
        if d:
            self.dest_edit.setText(d)

    def _update_destination_status(self):
        text = self.dest_edit.text().strip()
        if not text:
            self.lbl_dest_status.setVisible(False)
            return

        p = Path(text)
        if p.is_dir() and ((p / "index.md").is_file() or (p / ".export_manifest.json").is_file()):
            from modules.meetings.core.exporters.export_manifest import load_manifest
            m = load_manifest(p)
            if m:
                n_docs = m.get_total_tdocs()
                n_updates = len(m.sessions) - 1
                up_str = f", {n_updates} previous update(s)" if n_updates > 0 else ""
                msg = (
                    f"🔄 <b>Existing Export Detected:</b> Incremental Append Mode active "
                    f"({n_docs} existing TDocs{up_str} will be preserved and merged into the index)."
                )
            else:
                md_count = len([f for f in p.glob("**/*.md") if f.name.lower() != "index.md"])
                msg = (
                    f"🔄 <b>Existing Export Detected:</b> Incremental Append Mode active "
                    f"({md_count} existing TDocs will be preserved and merged into the index)."
                )
            self.lbl_dest_status.setText(msg)
            self.lbl_dest_status.setStyleSheet(f"""
                QLabel {{
                    background-color: {C.SUCCESS_SOFT};
                    border: 1px solid {C.SUCCESS_BORDER};
                    border-radius: 4px;
                    padding: 4px 8px;
                    font-size: 11px;
                    color: {C.SUCCESS_DARK};
                }}
            """)
            self.lbl_dest_status.setVisible(True)
        else:
            self.lbl_dest_status.setVisible(False)

    def _show_recent_exports_menu(self):
        export_root = self.meeting_dir / "Export"
        if not export_root.is_dir():
            QMessageBox.information(self, "No Recent Exports", "No existing export folders found in this meeting directory.")
            return

        candidates = []
        for item in export_root.iterdir():
            if item.is_dir() and ((item / "index.md").is_file() or (item / ".export_manifest.json").is_file()):
                try:
                    candidates.append((item.stat().st_mtime, item))
                except Exception:
                    pass

        if not candidates:
            QMessageBox.information(self, "No Recent Exports", "No previous export folders with index.md found in Export/.")
            return

        candidates.sort(key=lambda x: x[0], reverse=True)
        menu = QMenu(self)
        for mtime, folder in candidates[:10]:
            dt_str = datetime.datetime.fromtimestamp(mtime).strftime("%Y-%m-%d %H:%M")
            action = menu.addAction(f"📁 {folder.name} ({dt_str})")
            action.triggered.connect(lambda checked=False, p=folder: self._select_recent_export(p))

        menu.exec_(self.btn_recent.mapToGlobal(self.btn_recent.rect().bottomLeft()))

    def _select_recent_export(self, folder: Path):
        self.dest_edit.setText(str(folder))

    def _on_start_clicked(self):
        if not any([
            self.chk_ls_in.isChecked(),
            self.chk_ls_out.isChecked(),
            self.chk_pcr.isChecked(),
            self.chk_cr.isChecked(),
            self.chk_disc.isChecked(),
        ]):
            QMessageBox.warning(self, "No Types Selected", "Please select at least one document category to export.")
            return

        if self.rb_custom.isChecked():
            if not self.selected_wi_ais:
                QMessageBox.warning(
                    self,
                    "No Items Selected",
                    "Please select at least one Work Item or Agenda Item from the custom selection tree."
                )
                return

        self.accept()

    def get_scope_mode(self) -> str:
        if self.rb_followed.isChecked():
            return "followed"
        if self.rb_whole.isChecked():
            return "whole"
        if self.rb_starred.isChecked():
            return "starred"
        if self.rb_tracked.isChecked():
            return "tracked"
        if self.rb_visible.isChecked():
            return "visible"
        if self.rb_custom.isChecked():
            return "custom"
        return "followed"

    def get_custom_wis(self) -> Set[str]:
        return {wi.lower() for wi, ai in self.selected_wi_ais}

    def get_custom_ais(self) -> Set[str]:
        return {ai for wi, ai in self.selected_wi_ais}

    def get_custom_wi_ais(self) -> Dict[str, Set[str]]:
        wi_map: Dict[str, Set[str]] = {}
        for wi, ai in self.selected_wi_ais:
            k = wi.lower()
            if k not in wi_map:
                wi_map[k] = set()
            wi_map[k].add(ai)
        return wi_map

    def get_company_filter(self) -> Optional[Set[str]]:
        c = self.combo_company.currentData()
        return {str(c).strip().lower()} if c else None

    def get_export_config(self) -> dict:
        dest_text = self.dest_edit.text().strip()
        output_dir = Path(dest_text) if dest_text else None

        return {
            "scope_mode": self.get_scope_mode(),
            "custom_wis": self.get_custom_wis(),
            "custom_ais": self.get_custom_ais(),
            "custom_wi_ais": self.get_custom_wi_ais(),
            "include_ls_in": self.chk_ls_in.isChecked(),
            "include_ls_out": self.chk_ls_out.isChecked(),
            "include_pcr": self.chk_pcr.isChecked(),
            "include_cr": self.chk_cr.isChecked(),
            "include_disc": self.chk_disc.isChecked(),
            "include_general_ls": self.chk_general_ls.isChecked(),
            "deduplicate_revisions": self.chk_dedup_revs.isChecked(),
            "include_reserved_in_index": self.chk_include_reserved.isChecked(),
            "company_filter": self.get_company_filter(),
            "include_user_notes": self.chk_user_notes.isChecked(),
            "generate_local_indices": self.chk_local_indices.isChecked(),
            "download_missing": self.chk_download.isChecked(),
            "force_reconvert": self.chk_force.isChecked(),
            "compile_dossier": self.chk_compile_dossier.isChecked(),
            "auto_open": self.chk_auto_open.isChecked(),
            "output_dir": output_dir,
        }
