"""UI layout builder mixin for `WordSettingsDialog`."""

from __future__ import annotations

import webbrowser

from PyQt5.QtWidgets import (
    QCheckBox,
    QComboBox,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QVBoxLayout,
)

from core.ui.theme import C
from core.ui.ui_components import BUTTON_STYLE_TOOLBAR_SECONDARY, SEARCH_INPUT_STYLE
from modules.word_tools.core.libreoffice_converter import LIBREOFFICE_DOWNLOAD_URL
from modules.word_tools.core.word_excerpt_naming import CR_NAMING_SCHEMES, NAMING_SCHEMES


class WordSettingsUiSetupMixin:
    """Mixin constructing templates, defaults, sensitivity labels, and LibreOffice groups."""

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

        self.combo_excerpt_naming = QComboBox()
        self.combo_excerpt_naming.setStyleSheet(SEARCH_INPUT_STYLE)
        for scheme_id, scheme_label in NAMING_SCHEMES:
            self.combo_excerpt_naming.addItem(scheme_label, scheme_id)
        def_form.addRow("Default Excerpt Naming:", self.combo_excerpt_naming)

        self.combo_cr_naming = QComboBox()
        self.combo_cr_naming.setStyleSheet(SEARCH_INPUT_STYLE)
        for scheme_id, scheme_label in CR_NAMING_SCHEMES:
            self.combo_cr_naming.addItem(scheme_label, scheme_id)
        def_form.addRow("Default CR / pCR Naming:", self.combo_cr_naming)
        main_layout.addWidget(grp_defaults)

        # 3. Enterprise Sensitivity Labels
        grp_sensitivity = QGroupBox("🛡️ Enterprise Sensitivity Labels (Win32 COM Automation)")
        grp_sensitivity.setStyleSheet("QGroupBox { font-weight: bold; }")
        sens_layout = QVBoxLayout(grp_sensitivity)
        sens_layout.setSpacing(8)

        policy_row = QHBoxLayout()
        self.lbl_policy_status = QLabel()
        self.lbl_policy_status.setStyleSheet(f"font-size: 11px; color: {C.TEXT_SECONDARY};")
        policy_row.addWidget(self.lbl_policy_status, 1)

        self.btn_sync_policy = QPushButton("🔄 Sync from Windows Policy")
        self.btn_sync_policy.setStyleSheet(BUTTON_STYLE_TOOLBAR_SECONDARY)
        self.btn_sync_policy.setToolTip(
            "Discover sensitivity labels from your local Windows/Office Purview CLP policy cache."
        )
        self.btn_sync_policy.clicked.connect(self._on_sync_policy_clicked)
        policy_row.addWidget(self.btn_sync_policy)
        sens_layout.addLayout(policy_row)

        self.chk_enable_sensitivity = QCheckBox(
            "Apply Sensitivity Label to programmatically modified Word documents (bypasses save prompt)"
        )
        sens_layout.addWidget(self.chk_enable_sensitivity)

        sens_form = QFormLayout()
        sens_form.setSpacing(8)

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
        sens_form.addRow("Default Sensitivity Level:", preset_row)

        self.edit_label_name = QLineEdit()
        self.edit_label_name.setStyleSheet(SEARCH_INPUT_STYLE)
        self.edit_label_name.setPlaceholderText("Label policy name (e.g. PUBLIC)")
        self.edit_label_name.textChanged.connect(self._on_label_field_edited)
        sens_form.addRow("Label Name:", self.edit_label_name)

        self.edit_label_id = QLineEdit()
        self.edit_label_id.setStyleSheet(SEARCH_INPUT_STYLE)
        self.edit_label_id.setPlaceholderText("GUID (e.g. 00000000-0000-0000-0000-000000000000)")
        self.edit_label_id.textChanged.connect(self._on_label_field_edited)
        sens_form.addRow("Label GUID / ID:", self.edit_label_id)

        self.lbl_label_desc = QLabel()
        self.lbl_label_desc.setWordWrap(True)
        self.lbl_label_desc.setStyleSheet(f"font-size: 11px; color: {C.TEXT_MUTED}; font-style: italic;")
        sens_form.addRow("", self.lbl_label_desc)

        self.chk_save_after_label = QCheckBox("Automatically save document immediately after labeling")
        sens_form.addRow("", self.chk_save_after_label)

        sens_layout.addLayout(sens_form)

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
        self.lbl_lo_detected.setStyleSheet(f"font-size: 11px; color: {C.TEXT_SECONDARY};")
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
