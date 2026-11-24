"""Label preset management, COM/policy inspection, template toggles, and config persistence."""

from __future__ import annotations

import os
import webbrowser
from pathlib import Path
from typing import Optional

from PyQt5.QtWidgets import QDialog, QFileDialog, QMessageBox

from core.ui.theme import C
from modules.word_tools.core.libreoffice_converter import (
    find_libreoffice_executable,
    resolve_soffice_binary,
)
from modules.word_tools.core.styles_3gpp import (
    OFFICIAL_CR_TEMPLATE_URL,
    install_3gpp_dotx_template,
    is_3gpp_dotx_installed,
    uninstall_3gpp_dotx_template,
)
from modules.word_tools.core.word_config import WordConfig
from modules.word_tools.ui.word_settings.sensitivity_edit_dialog import (
    SensitivityLabelEditDialog,
)


class WordSettingsActionsMixin:
    """Mixin handling WordConfig loading/saving, Windows MIP policy sync, COM tests, and templates."""

    def _load_values(self):
        cfg = WordConfig.load()
        self.chk_enable_sensitivity.setChecked(cfg.get("enable_sensitivity_labels", True))
        self.chk_save_after_label.setChecked(cfg.get("save_document_after_setting_sensitivity_label", False))
        self.edit_company.setText(cfg.get("default_source_company", ""))
        self.edit_lo_path.setText(cfg.get("libreoffice_path", ""))

        scheme = cfg.get("excerpt_naming_scheme", "source_clauses_timestamp")
        idx = self.combo_excerpt_naming.findData(scheme)
        if idx >= 0:
            self.combo_excerpt_naming.setCurrentIndex(idx)

        cr_scheme = cfg.get("cr_naming_scheme", "cr_standard")
        cr_idx = self.combo_cr_naming.findData(cr_scheme)
        if cr_idx >= 0:
            self.combo_cr_naming.setCurrentIndex(cr_idx)

        self._tenant_id = cfg.get("sensitivity_tenant_id", "")
        self._policy_upn = cfg.get("sensitivity_policy_upn", "")
        self._labels_cache = [dict(lbl) for lbl in cfg.get("sensitivity_labels", [])]
        active_id = cfg.get("active_sensitivity_label_id") or cfg.get("sensitivity_level_label_id", "")

        self._refresh_policy_status()
        self._populate_presets(active_id=active_id)
        self._on_sensitivity_toggled(self.chk_enable_sensitivity.isChecked())
        self._refresh_dotx_status()
        self._refresh_lo_status()

    def _refresh_policy_status(self, policy: Optional[dict] = None):
        from modules.word_tools.core.sensitivity_label import discover_windows_sensitivity_policy

        if policy is None:
            policy = discover_windows_sensitivity_policy()

        if policy and policy.get("labels"):
            upn = policy.get("upn") or self._policy_upn or "Corporate Account"
            tid = policy.get("tenant_id") or self._tenant_id
            short_tid = f"...{tid[-8:]}" if len(tid) > 8 else tid
            tid_info = f", Tenant {short_tid}" if short_tid else ""
            self.lbl_policy_status.setText(
                f"🟢 Windows Policy: {upn} ({len(policy['labels'])} File labels{tid_info})"
            )
            self.lbl_policy_status.setStyleSheet(f"font-size: 11px; color: {C.WG_RAN_TEXT}; font-weight: 600;")
        else:
            self.lbl_policy_status.setText(
                "⚪ No Windows / Office sensitivity policy detected on this machine"
            )
            self.lbl_policy_status.setStyleSheet(f"font-size: 11px; color: {C.TEXT_MUTED};")

    def _update_label_description(self, lbl: Optional[dict]):
        if not lbl:
            self.lbl_label_desc.setText("")
            self.lbl_label_desc.setStyleSheet(f"font-size: 11px; color: {C.TEXT_MUTED}; font-style: italic;")
            return

        desc = (lbl.get("description") or "").strip()
        if lbl.get("is_encrypted"):
            warn = "🔒 Encrypted label (RMS Rights Protection) — may restrict external 3GPP sharing and headless ZIP inspection."
            text = f"{warn}\n{desc}" if desc else warn
            self.lbl_label_desc.setText(text)
            self.lbl_label_desc.setStyleSheet(f"font-size: 11px; color: {C.WARNING}; font-weight: 500;")
        else:
            self.lbl_label_desc.setText(desc)
            self.lbl_label_desc.setStyleSheet(f"font-size: 11px; color: {C.TEXT_MUTED}; font-style: italic;")

    def _populate_presets(self, active_id: Optional[str] = None):
        self.combo_label_presets.blockSignals(True)
        self.combo_label_presets.clear()

        if not self._labels_cache:
            self.combo_label_presets.addItem("No sensitivity labels detected or configured", "")
            self.combo_label_presets.setCurrentIndex(0)
            self.combo_label_presets.blockSignals(False)

            self.edit_label_id.blockSignals(True)
            self.edit_label_name.blockSignals(True)
            self.edit_label_id.setText("")
            self.edit_label_name.setText("")
            self._update_label_description(None)
            self.edit_label_id.blockSignals(False)
            self.edit_label_name.blockSignals(False)
            self._on_sensitivity_toggled(self.chk_enable_sensitivity.isChecked())
            return

        select_idx = 0
        for idx, lbl in enumerate(self._labels_cache):
            lid = lbl.get("id", "")
            display_text = WordConfig.format_label_display(lbl, include_guid=True)
            self.combo_label_presets.addItem(display_text, lid)
            if active_id and lid.lower() == active_id.lower():
                select_idx = idx

        self.combo_label_presets.setCurrentIndex(select_idx)
        self.combo_label_presets.blockSignals(False)

        cur = self._labels_cache[select_idx]
        self.edit_label_id.blockSignals(True)
        self.edit_label_name.blockSignals(True)
        self.edit_label_id.setText(cur.get("id", ""))
        self.edit_label_name.setText(cur.get("name", ""))
        self._update_label_description(cur)
        self.edit_label_id.blockSignals(False)
        self.edit_label_name.blockSignals(False)
        self._on_sensitivity_toggled(self.chk_enable_sensitivity.isChecked())

    def _on_sync_policy_clicked(self):
        temp_cfg = {
            "sensitivity_labels": [dict(lbl) for lbl in self._labels_cache],
            "active_sensitivity_label_id": self.edit_label_id.text().strip(),
        }
        policy = WordConfig.sync_from_windows_policy(temp_cfg, overwrite_active=False)
        self._refresh_policy_status(policy=policy)

        if policy and policy.get("labels"):
            self._labels_cache = temp_cfg.get("sensitivity_labels", [])
            self._tenant_id = temp_cfg.get("sensitivity_tenant_id", "")
            self._policy_upn = temp_cfg.get("sensitivity_policy_upn", "")
            self._populate_presets(active_id=temp_cfg.get("active_sensitivity_label_id", ""))
            QMessageBox.information(
                self,
                "Windows Sensitivity Policy Synced",
                f"Detected {len(policy['labels'])} File-scoped label(s) for {policy.get('upn') or 'current user'}.\n\n"
                f"Click 'Save Preferences' to persist your selection.",
            )
        else:
            QMessageBox.information(
                self,
                "No Windows Policy Detected",
                "No local Microsoft Office / Purview CLP sensitivity policy was found under "
                "%LOCALAPPDATA%\\Microsoft\\Office\\CLP.\n\n"
                "Sensitivity labeling will remain inactive unless you add custom label presets manually.",
            )

    def _on_preset_changed(self, index: int):
        if 0 <= index < len(self._labels_cache):
            lbl = self._labels_cache[index]
            self.edit_label_id.blockSignals(True)
            self.edit_label_name.blockSignals(True)
            self.edit_label_id.setText(lbl.get("id", ""))
            self.edit_label_name.setText(lbl.get("name", ""))
            self._update_label_description(lbl)
            self.edit_label_id.blockSignals(False)
            self.edit_label_name.blockSignals(False)

    def _on_label_field_edited(self):
        """Update current cache entry when user edits name or id in-place."""
        idx = self.combo_label_presets.currentIndex()
        if 0 <= idx < len(self._labels_cache):
            self._labels_cache[idx]["id"] = self.edit_label_id.text().strip()
            self._labels_cache[idx]["name"] = self.edit_label_name.text().strip()

    def _on_add_label_clicked(self):
        dlg = SensitivityLabelEditDialog(self, name="", guid="", description="")
        if dlg.exec_() == QDialog.Accepted:
            data = dlg.get_data()
            self._labels_cache.append(data)
            self._populate_presets(active_id=data["id"])

    def _on_edit_label_clicked(self):
        idx = self.combo_label_presets.currentIndex()
        if not (0 <= idx < len(self._labels_cache)):
            return
        cur = self._labels_cache[idx]
        dlg = SensitivityLabelEditDialog(
            self,
            name=cur.get("name", ""),
            guid=cur.get("id", ""),
            description=cur.get("description", ""),
            display_name=cur.get("display_name", ""),
            is_encrypted=bool(cur.get("is_encrypted", False)),
        )
        if dlg.exec_() == QDialog.Accepted:
            data = dlg.get_data()
            updated = dict(cur)
            updated.update(data)
            self._labels_cache[idx] = updated
            self._populate_presets(active_id=updated["id"])

    def _on_del_label_clicked(self):
        idx = self.combo_label_presets.currentIndex()
        if 0 <= idx < len(self._labels_cache):
            target = self._labels_cache[idx]
            reply = QMessageBox.question(
                self,
                "Remove Sensitivity Label",
                f"Remove label preset '{target.get('name')}'?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No,
            )
            if reply == QMessageBox.Yes:
                self._labels_cache.pop(idx)
                self._populate_presets()

    def _launch_batch_labeler(self):
        from modules.word_tools.ui.batch_label_dialog import BatchSensitivityLabelDialog
        dlg = BatchSensitivityLabelDialog(self)
        dlg.exec_()

    def _inspect_document_label(self):
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Word Document to Inspect Sensitivity Label",
            "",
            "Word Documents (*.docx *.doc *.docm);;All Files (*.*)",
        )
        if not path:
            return

        from modules.word_tools.core.sensitivity_label import get_document_sensitivity_label
        lbl = get_document_sensitivity_label(path)
        p = Path(path)
        if lbl:
            QMessageBox.information(
                self,
                "🛡️ Document Sensitivity Label",
                f"File: {p.name}\n\n"
                f"• Label Name: {lbl.get('name', 'N/A')}\n"
                f"• Label GUID: {lbl.get('id', 'N/A')}\n"
                f"• Assignment Method: {lbl.get('method', 'N/A')}\n"
                f"• Set Date: {lbl.get('set_date', 'N/A')}\n"
                f"• Site ID: {lbl.get('site_id', 'N/A')}\n\n"
                f"Status: Document is classified and compliant.",
            )
        else:
            QMessageBox.warning(
                self,
                "🛡️ Document Sensitivity Label",
                f"File: {p.name}\n\n"
                f"Status: No sensitivity label detected (Unclassified / NOT_SET).\n"
                f"In enterprise IT environments, this document may trigger DLP blocks when saving or emailing.",
            )

    def _test_active_word_label(self):
        from modules.word_tools.core.document_opener import get_active_word_document
        active_doc = get_active_word_document(allow_unsaved_snapshot=False)
        if not active_doc:
            QMessageBox.warning(
                self,
                "No Active Word Document",
                "No active Microsoft Word document could be detected.\n\n"
                "Please open a Word document first, then click Test.",
            )
            return

        import pythoncom
        import win32com.client
        from modules.word_tools.core.sensitivity_label import get_document_sensitivity_label, set_sensitivity_label

        label_id = self.edit_label_id.text().strip()
        label_name = self.edit_label_name.text().strip()
        if not label_id or not label_name:
            QMessageBox.warning(self, "Missing Label Info", "Please specify both a Label Name and Label ID.")
            return

        pythoncom.CoInitialize()
        try:
            word = win32com.client.GetActiveObject("Word.Application")
            doc = word.ActiveDocument
            doc_name = str(getattr(doc, "Name", "Active Document"))
            set_sensitivity_label(
                doc,
                label_id=label_id,
                label_name=label_name,
                return_changed=True,
                force_overwrite=True,
            )
            current = get_document_sensitivity_label(doc)
            QMessageBox.information(
                self,
                "Sensitivity Label Applied",
                f"Target Document: {doc_name}\n"
                f"Applied Label: {label_name} ({label_id})\n\n"
                f"Status: Successfully applied to active Word document!\n"
                f"Verified in Word COM: {current.get('name') if current else 'Classified'}",
            )
        except Exception as e:
            QMessageBox.critical(
                self,
                "Sensitivity Label Test Failed",
                f"Failed to apply sensitivity label via COM:\n\n{e}\n\n"
                "Please verify that Microsoft Word supports sensitivity labels for this tenant.",
            )
        finally:
            pythoncom.CoUninitialize()

    def _refresh_dotx_status(self):
        installed = is_3gpp_dotx_installed()
        if installed:
            self.lbl_dotx_status.setText("🟢 Installed (in %APPDATA%\\Microsoft\\Templates)")
            self.lbl_dotx_status.setStyleSheet(f"color: {C.WG_RAN_TEXT}; font-weight: bold;")
            self.btn_toggle_dotx.setText("🗑️ Remove 3gpp_70.dotx")
        else:
            self.lbl_dotx_status.setText("⚪ Not Installed")
            self.lbl_dotx_status.setStyleSheet(f"color: {C.TEXT_MUTED}; font-weight: bold;")
            self.btn_toggle_dotx.setText("📥 Install 3gpp_70.dotx")

    def _refresh_lo_status(self):
        exe = find_libreoffice_executable()
        if exe:
            self.lbl_lo_detected.setText(f"Active executable: {exe}")
        else:
            self.lbl_lo_detected.setText("No executable currently detected.")

    def _toggle_dotx(self):
        if is_3gpp_dotx_installed():
            ok, msg = uninstall_3gpp_dotx_template()
            if ok:
                QMessageBox.information(self, "Template Removed", msg)
            else:
                QMessageBox.warning(self, "Removal Failed", msg)
        else:
            ok, msg = install_3gpp_dotx_template(force=True)
            if ok:
                QMessageBox.information(self, "Template Installed", msg)
            else:
                QMessageBox.warning(self, "Installation Failed", msg)
        self._refresh_dotx_status()

    def _open_templates_folder(self):
        appdata = os.environ.get("APPDATA", "")
        if appdata:
            tpl_dir = Path(appdata) / "Microsoft" / "Templates"
            tpl_dir.mkdir(parents=True, exist_ok=True)
            os.startfile(str(tpl_dir))

    def _open_official_templates_url(self):
        webbrowser.open(OFFICIAL_CR_TEMPLATE_URL)

    def _on_sensitivity_toggled(self, checked: bool):
        has_labels = bool(self._labels_cache)
        self.btn_sync_policy.setEnabled(checked)
        self.combo_label_presets.setEnabled(checked and has_labels)
        self.btn_add_label.setEnabled(checked)
        self.btn_edit_label.setEnabled(checked and has_labels)
        self.btn_del_label.setEnabled(checked and has_labels)
        self.edit_label_id.setEnabled(checked and has_labels)
        self.edit_label_name.setEnabled(checked and has_labels)
        self.chk_save_after_label.setEnabled(checked)
        self.btn_inspect_doc.setEnabled(checked)
        self.btn_test_active.setEnabled(checked and has_labels)
        self.btn_launch_batch.setEnabled(checked)

    def _browse_lo(self):
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Select LibreOffice Executable",
            self.edit_lo_path.text(),
            "Executable (*soffice.exe *LibreOfficePortable.exe *.exe);;All Files (*.*)",
        )
        if path:
            resolved = resolve_soffice_binary(path)
            self.edit_lo_path.setText(str(resolved or path))
            self._refresh_lo_status()

    def _save_values(self):
        cfg = WordConfig.load()
        cfg["enable_sensitivity_labels"] = self.chk_enable_sensitivity.isChecked()
        active_lid = self.edit_label_id.text().strip() if self._labels_cache else ""
        active_name = self.edit_label_name.text().strip() if self._labels_cache else ""

        cfg["active_sensitivity_label_id"] = active_lid
        cfg["sensitivity_level_label_id"] = active_lid
        cfg["sensitivity_level_label_name"] = active_name
        cfg["sensitivity_tenant_id"] = self._tenant_id
        cfg["sensitivity_policy_upn"] = self._policy_upn
        cfg["sensitivity_policy_initialized"] = True
        cfg["save_document_after_setting_sensitivity_label"] = self.chk_save_after_label.isChecked()
        cfg["sensitivity_labels"] = self._labels_cache
        cfg["default_source_company"] = self.edit_company.text().strip()
        cfg["libreoffice_path"] = self.edit_lo_path.text().strip()
        cfg["excerpt_naming_scheme"] = self.combo_excerpt_naming.currentData() or "source_clauses_timestamp"
        cfg["cr_naming_scheme"] = self.combo_cr_naming.currentData() or "cr_standard"
        WordConfig.save(cfg)
        self.accept()
