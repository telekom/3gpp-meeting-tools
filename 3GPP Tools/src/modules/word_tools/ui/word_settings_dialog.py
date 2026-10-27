# --- File: src/modules/word_tools/ui/word_settings_dialog.py ---
"""
Settings dialog for Word Tools, Corporate Sensitivity Labels, and 3GPP Templates.
"""
import os
import webbrowser
from pathlib import Path
from typing import Optional

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QCheckBox,
    QDialog,
    QFileDialog,
    QFormLayout,
    QFrame,
    QGroupBox,
    QHBoxLayout,
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


class WordSettingsDialog(QDialog):
    """Configuration dialog for Word Tools preferences, sensitivity labels, and templates."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("⚙️ Word Tools Settings & Preferences")
        self.setMinimumWidth(560)
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)
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
        self.edit_label_id = QLineEdit()
        self.edit_label_id.setStyleSheet(SEARCH_INPUT_STYLE)
        self.edit_label_id.setPlaceholderText("GUID (e.g. 55339bf0-f345-473a-9ec8-6ca7c8197055)")
        sens_form.addRow("Label GUID / ID:", self.edit_label_id)

        self.edit_label_name = QLineEdit()
        self.edit_label_name.setStyleSheet(SEARCH_INPUT_STYLE)
        self.edit_label_name.setPlaceholderText("Label display name (e.g. OFFEN)")
        sens_form.addRow("Label Name:", self.edit_label_name)

        self.chk_save_after_label = QCheckBox("Automatically save document immediately after labeling")
        sens_form.addRow("", self.chk_save_after_label)

        sens_layout.addLayout(sens_form)
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
        self.edit_label_id.setText(cfg.get("sensitivity_level_label_id", "55339bf0-f345-473a-9ec8-6ca7c8197055"))
        self.edit_label_name.setText(cfg.get("sensitivity_level_label_name", "OFFEN"))
        self.chk_save_after_label.setChecked(cfg.get("save_document_after_setting_sensitivity_label", False))
        self.edit_company.setText(cfg.get("default_source_company", ""))
        self.edit_lo_path.setText(cfg.get("libreoffice_path", ""))

        self._on_sensitivity_toggled(self.chk_enable_sensitivity.isChecked())
        self._refresh_dotx_status()
        self._refresh_lo_status()

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
        self.edit_label_id.setEnabled(checked)
        self.edit_label_name.setEnabled(checked)
        self.chk_save_after_label.setEnabled(checked)

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
        cfg["sensitivity_level_label_id"] = self.edit_label_id.text().strip()
        cfg["sensitivity_level_label_name"] = self.edit_label_name.text().strip()
        cfg["save_document_after_setting_sensitivity_label"] = self.chk_save_after_label.isChecked()
        cfg["default_source_company"] = self.edit_company.text().strip()
        cfg["libreoffice_path"] = self.edit_lo_path.text().strip()
        WordConfig.save(cfg)
        self.accept()
