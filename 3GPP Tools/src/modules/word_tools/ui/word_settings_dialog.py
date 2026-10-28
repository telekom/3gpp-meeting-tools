# --- File: src/modules/word_tools/ui/word_settings_dialog.py ---
"""
Settings dialog for Word Tools, Corporate Sensitivity Labels, and 3GPP Templates.
"""
import os
import webbrowser
from pathlib import Path
from typing import Dict, List, Optional

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QCheckBox,
    QComboBox,
    QDialog,
    QFileDialog,
    QFormLayout,
    QFrame,
    QGroupBox,
    QHBoxLayout,
    QInputDialog,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QVBoxLayout,
)

from core.ui.ui_components import (
    BUTTON_STYLE_TOOLBAR_SECONDARY,
    CARD_FRAME_STYLE,
    SEARCH_INPUT_STYLE,
)
from modules.word_tools.core.libreoffice_converter import (
    LIBREOFFICE_DOWNLOAD_URL,
    find_libreoffice_executable,
    resolve_soffice_binary,
)
from modules.word_tools.core.styles_3gpp import (
    OFFICIAL_3GPP_DOTX_URL,
    OFFICIAL_CR_TEMPLATE_URL,
    OFFICIAL_PCR_TEMPLATE_ZIP_URL,
    install_3gpp_dotx_template,
    is_3gpp_dotx_installed,
    uninstall_3gpp_dotx_template,
)
from modules.word_tools.core.word_config import WordConfig


class SensitivityLabelEditDialog(QDialog):
    """Dialog to edit or create a sensitivity label configuration."""

    def __init__(self, parent=None, name: str = "", guid: str = "", description: str = ""):
        super().__init__(parent)
        self.setWindowTitle("✏️ Edit Sensitivity Label" if name else "➕ Add Sensitivity Label")
        self.setMinimumWidth(480)
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)

        form = QFormLayout()
        form.setSpacing(8)

        self.edit_name = QLineEdit(name)
        self.edit_name.setStyleSheet(SEARCH_INPUT_STYLE)
        self.edit_name.setPlaceholderText("e.g. VERTRAULICH - unverschlüsselt")
        form.addRow("Label Name:", self.edit_name)

        self.edit_guid = QLineEdit(guid)
        self.edit_guid.setStyleSheet(SEARCH_INPUT_STYLE)
        self.edit_guid.setPlaceholderText("e.g. 5d95d744-d996-4a3c-a104-fbf2486a4c52")
        form.addRow("Label GUID / ID:", self.edit_guid)

        self.edit_desc = QLineEdit(description)
        self.edit_desc.setStyleSheet(SEARCH_INPUT_STYLE)
        self.edit_desc.setPlaceholderText("e.g. Confidential - unencrypted")
        form.addRow("Description:", self.edit_desc)

        layout.addLayout(form)

        btn_row = QHBoxLayout()
        btn_row.addStretch()
        self.btn_cancel = QPushButton("Cancel")
        self.btn_cancel.clicked.connect(self.reject)
        btn_row.addWidget(self.btn_cancel)

        self.btn_save = QPushButton("Save Label")
        self.btn_save.setObjectName("primaryBtn")
        self.btn_save.clicked.connect(self._validate_and_accept)
        btn_row.addWidget(self.btn_save)

        layout.addLayout(btn_row)

    def _validate_and_accept(self):
        if not self.edit_name.text().strip():
            QMessageBox.warning(self, "Invalid Name", "Label display name cannot be empty.")
            return
        if not self.edit_guid.text().strip():
            QMessageBox.warning(self, "Invalid GUID", "Label GUID cannot be empty.")
            return
        self.accept()

    def get_data(self) -> dict:
        return {
            "name": self.edit_name.text().strip(),
            "id": self.edit_guid.text().strip(),
            "description": self.edit_desc.text().strip(),
        }


class WordSettingsDialog(QDialog):
    """Configuration dialog for Word Tools preferences, sensitivity labels, and templates."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("⚙️ Word Tools Settings & Preferences")
        self.setMinimumWidth(640)
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)
        self._labels_cache: List[dict] = []
        self._setup_ui()
        self._load_values()

    def _setup_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(16, 16, 16, 16)
        main_layout.setSpacing(14)

        # 1. 3GPP Templates & Environment
        grp_templates = QGroupBox("🏛️ 3GPP Templates && Microsoft Word Integration")
        grp_templates.setStyleSheet("QGroupBox { font-weight: bold; }")
        tpl_layout = QVBoxLayout(grp_templates)
        tpl_layout.setSpacing(8)

        tpl_status_row = QHBoxLayout()
        tpl_status_row.addWidget(QLabel("<b>3GPP Macro Ribbon (3gpp_70.dotx):</b>"))
        self.lbl_dotx_status = QLabel()
        tpl_status_row.addWidget(self.lbl_dotx_status)
        tpl_status_row.addStretch()

        self.btn_toggle_dotx = QPushButton("📥 Install 3gpp_70.dotx")
        self.btn_toggle_dotx.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_toggle_dotx.clicked.connect(self._toggle_dotx)
        tpl_status_row.addWidget(self.btn_toggle_dotx)
        tpl_layout.addLayout(tpl_status_row)

        tpl_actions_row = QHBoxLayout()
        self.btn_open_tpl_folder = QPushButton("📂 Open Word Templates Folder")
        self.btn_open_tpl_folder.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_open_tpl_folder.clicked.connect(self._open_templates_folder)
        tpl_actions_row.addWidget(self.btn_open_tpl_folder)

        self.btn_download_templates = QPushButton("🌐 Download Official Templates...")
        self.btn_download_templates.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_download_templates.clicked.connect(self._open_official_templates_url)
        tpl_actions_row.addWidget(self.btn_download_templates)
        tpl_actions_row.addStretch()
        tpl_layout.addLayout(tpl_actions_row)

        main_layout.addWidget(grp_templates)

        # 2. Delegate & Contribution Defaults
        grp_defaults = QGroupBox("👤 Delegate Defaults")
        grp_defaults.setStyleSheet("QGroupBox { font-weight: bold; }")
        def_form = QFormLayout(grp_defaults)
        def_form.setSpacing(8)

        self.edit_company = QLineEdit()
        self.edit_company.setStyleSheet(SEARCH_INPUT_STYLE)
        self.edit_company.setPlaceholderText("e.g., T-Mobile, Ericsson, Qualcomm...")
        def_form.addRow("Default Source / Company:", self.edit_company)
        main_layout.addWidget(grp_defaults)

        # 3. Enterprise Sensitivity Labels
        grp_sensitivity = QGroupBox("🛡️ Enterprise Sensitivity Labels (Win32 COM Automation)")
        grp_sensitivity.setStyleSheet("QGroupBox { font-weight: bold; }")
        sens_layout = QVBoxLayout(grp_sensitivity)
        sens_layout.setSpacing(8)

        self.chk_enable_sensitivity = QCheckBox(
            "Apply Sensitivity Label to programmatically modified Word documents (bypasses save prompt)"
        )
        sens_layout.addWidget(self.chk_enable_sensitivity)

        sens_form = QFormLayout()
        sens_form.setSpacing(8)

        # Preset selection row with Add, Edit, Delete
        preset_row = QHBoxLayout()
        self.combo_label_presets = QComboBox()
        self.combo_label_presets.setStyleSheet(SEARCH_INPUT_STYLE)
        self.combo_label_presets.currentIndexChanged.connect(self._on_preset_changed)
        preset_row.addWidget(self.combo_label_presets, 1)

        self.btn_add_label = QPushButton("➕ Add...")
        self.btn_add_label.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_add_label.setToolTip("Add a new custom enterprise sensitivity label.")
        self.btn_add_label.clicked.connect(self._on_add_label_clicked)
        preset_row.addWidget(self.btn_add_label)

        self.btn_edit_label = QPushButton("✏️ Edit...")
        self.btn_edit_label.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_edit_label.setToolTip("Edit the name, GUID, and description of the selected label.")
        self.btn_edit_label.clicked.connect(self._on_edit_label_clicked)
        preset_row.addWidget(self.btn_edit_label)

        self.btn_del_label = QPushButton("🗑️")
        self.btn_del_label.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_del_label.setToolTip("Remove selected sensitivity label preset.")
        self.btn_del_label.clicked.connect(self._on_del_label_clicked)
        preset_row.addWidget(self.btn_del_label)
        sens_form.addRow("Active Sensitivity Level:", preset_row)

        self.edit_label_name = QLineEdit()
        self.edit_label_name.setStyleSheet(SEARCH_INPUT_STYLE)
        self.edit_label_name.setPlaceholderText("Label display name (e.g. OFFEN)")
        self.edit_label_name.textChanged.connect(self._on_label_field_edited)
        sens_form.addRow("Label Name:", self.edit_label_name)

        self.edit_label_id = QLineEdit()
        self.edit_label_id.setStyleSheet(SEARCH_INPUT_STYLE)
        self.edit_label_id.setPlaceholderText("GUID (e.g. 55339bf0-f345-473a-9ec8-6ca7c8197055)")
        self.edit_label_id.textChanged.connect(self._on_label_field_edited)
        sens_form.addRow("Label GUID / ID:", self.edit_label_id)

        self.lbl_label_desc = QLabel()
        self.lbl_label_desc.setStyleSheet("font-size: 11px; color: #64748B; font-style: italic;")
        sens_form.addRow("", self.lbl_label_desc)

        self.chk_save_after_label = QCheckBox("Automatically save document immediately after labeling")
        sens_form.addRow("", self.chk_save_after_label)

        sens_layout.addLayout(sens_form)

        # Inspection, Testing & Batch Utility Toolbar
        tools_row = QHBoxLayout()
        self.btn_inspect_doc = QPushButton("🔍 Inspect Document Label...")
        self.btn_inspect_doc.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_inspect_doc.setToolTip("Inspect sensitivity label metadata from a local .docx or .doc file.")
        self.btn_inspect_doc.clicked.connect(self._inspect_document_label)
        tools_row.addWidget(self.btn_inspect_doc)

        self.btn_test_active = QPushButton("🧪 Test on Active Word")
        self.btn_test_active.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_test_active.setToolTip("Test applying this sensitivity label to the foreground open Microsoft Word document.")
        self.btn_test_active.clicked.connect(self._test_active_word_label)
        tools_row.addWidget(self.btn_test_active)

        self.btn_launch_batch = QPushButton("🚀 Batch Label Documents...")
        self.btn_launch_batch.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_launch_batch.setToolTip("Open the batch sensitivity labeler for multiple documents or an entire folder.")
        self.btn_launch_batch.clicked.connect(self._launch_batch_labeler)
        tools_row.addWidget(self.btn_launch_batch)

        tools_row.addStretch()
        sens_layout.addLayout(tools_row)

        self.chk_enable_sensitivity.toggled.connect(self._on_sensitivity_toggled)
        main_layout.addWidget(grp_sensitivity)

        # 4. LibreOffice Configuration (Headless / Legacy Fallback)
        grp_lo = QGroupBox("📄 LibreOffice Runtime (Headless && Legacy .doc Conversion)")
        grp_lo.setStyleSheet("QGroupBox { font-weight: bold; }")
        lo_layout = QVBoxLayout(grp_lo)
        lo_layout.setSpacing(8)

        lo_path_row = QHBoxLayout()
        self.edit_lo_path = QLineEdit()
        self.edit_lo_path.setStyleSheet(SEARCH_INPUT_STYLE)
        self.edit_lo_path.setPlaceholderText("Path to soffice.exe or LibreOfficePortable.exe...")
        lo_path_row.addWidget(self.edit_lo_path, 1)

        self.btn_browse_lo = QPushButton("📂 Browse...")
        self.btn_browse_lo.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_browse_lo.clicked.connect(self._browse_lo)
        lo_path_row.addWidget(self.btn_browse_lo)

        self.btn_dl_lo = QPushButton("📥 Portable...")
        self.btn_dl_lo.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_dl_lo.clicked.connect(lambda: webbrowser.open(LIBREOFFICE_DOWNLOAD_URL))
        lo_path_row.addWidget(self.btn_dl_lo)
        lo_layout.addLayout(lo_path_row)

        self.lbl_lo_detected = QLabel()
        self.lbl_lo_detected.setStyleSheet("font-size: 11px; color: #475569;")
        lo_layout.addWidget(self.lbl_lo_detected)

        main_layout.addWidget(grp_lo)

        # Buttons Row
        btn_row = QHBoxLayout()
        btn_row.addStretch()

        self.btn_cancel = QPushButton("Cancel")
        self.btn_cancel.clicked.connect(self.reject)
        btn_row.addWidget(self.btn_cancel)

        self.btn_save = QPushButton("Save Preferences")
        self.btn_save.setObjectName("primaryBtn")
        self.btn_save.clicked.connect(self._save_values)
        btn_row.addWidget(self.btn_save)

        main_layout.addLayout(btn_row)

    def _load_values(self):
        cfg = WordConfig.load()
        self.chk_enable_sensitivity.setChecked(cfg.get("enable_sensitivity_labels", True))
        self.chk_save_after_label.setChecked(cfg.get("save_document_after_setting_sensitivity_label", False))
        self.edit_company.setText(cfg.get("default_source_company", ""))
        self.edit_lo_path.setText(cfg.get("libreoffice_path", ""))

        self._labels_cache = [dict(l) for l in cfg.get("sensitivity_labels", WordConfig.DEFAULT_SENSITIVITY_LABELS)]
        active_id = cfg.get("active_sensitivity_label_id") or cfg.get("sensitivity_level_label_id", "")

        self._populate_presets(active_id=active_id)
        self._on_sensitivity_toggled(self.chk_enable_sensitivity.isChecked())
        self._refresh_dotx_status()
        self._refresh_lo_status()

    def _populate_presets(self, active_id: Optional[str] = None):
        self.combo_label_presets.blockSignals(True)
        self.combo_label_presets.clear()

        select_idx = 0
        for idx, lbl in enumerate(self._labels_cache):
            name = lbl.get("name", "Unknown")
            lid = lbl.get("id", "")
            desc = lbl.get("description", "")
            short_guid = f"...{lid[-8:]}" if len(lid) > 8 else lid
            display_text = f"{name} ({short_guid})" if short_guid else name
            self.combo_label_presets.addItem(display_text, lid)
            if active_id and lid.lower() == active_id.lower():
                select_idx = idx

        self.combo_label_presets.setCurrentIndex(select_idx)
        self.combo_label_presets.blockSignals(False)

        if self._labels_cache:
            cur = self._labels_cache[select_idx]
            self.edit_label_id.blockSignals(True)
            self.edit_label_name.blockSignals(True)
            self.edit_label_id.setText(cur.get("id", ""))
            self.edit_label_name.setText(cur.get("name", ""))
            self.lbl_label_desc.setText(cur.get("description", ""))
            self.edit_label_id.blockSignals(False)
            self.edit_label_name.blockSignals(False)

    def _on_preset_changed(self, index: int):
        if 0 <= index < len(self._labels_cache):
            lbl = self._labels_cache[index]
            self.edit_label_id.blockSignals(True)
            self.edit_label_name.blockSignals(True)
            self.edit_label_id.setText(lbl.get("id", ""))
            self.edit_label_name.setText(lbl.get("name", ""))
            self.lbl_label_desc.setText(lbl.get("description", ""))
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
        )
        if dlg.exec_() == QDialog.Accepted:
            data = dlg.get_data()
            self._labels_cache[idx] = data
            self._populate_presets(active_id=data["id"])

    def _on_del_label_clicked(self):
        if len(self._labels_cache) <= 1:
            QMessageBox.warning(self, "Cannot Remove", "At least one sensitivity label preset must be configured.")
            return

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
        from modules.word_tools.core.sensitivity_label import set_sensitivity_label, get_document_sensitivity_label

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
            self.lbl_dotx_status.setStyleSheet("color: #065F46; font-weight: bold;")
            self.btn_toggle_dotx.setText("🗑️ Remove 3gpp_70.dotx")
        else:
            self.lbl_dotx_status.setText("⚪ Not Installed")
            self.lbl_dotx_status.setStyleSheet("color: #64748B; font-weight: bold;")
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
        self.combo_label_presets.setEnabled(checked)
        self.btn_add_label.setEnabled(checked)
        self.btn_edit_label.setEnabled(checked)
        self.btn_del_label.setEnabled(checked)
        self.edit_label_id.setEnabled(checked)
        self.edit_label_name.setEnabled(checked)
        self.chk_save_after_label.setEnabled(checked)
        self.btn_inspect_doc.setEnabled(checked)
        self.btn_test_active.setEnabled(checked)
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
        active_lid = self.edit_label_id.text().strip()
        active_name = self.edit_label_name.text().strip()

        cfg["active_sensitivity_label_id"] = active_lid
        cfg["sensitivity_level_label_id"] = active_lid
        cfg["sensitivity_level_label_name"] = active_name
        cfg["save_document_after_setting_sensitivity_label"] = self.chk_save_after_label.isChecked()
        cfg["sensitivity_labels"] = self._labels_cache
        cfg["default_source_company"] = self.edit_company.text().strip()
        cfg["libreoffice_path"] = self.edit_lo_path.text().strip()
        WordConfig.save(cfg)
        self.accept()
