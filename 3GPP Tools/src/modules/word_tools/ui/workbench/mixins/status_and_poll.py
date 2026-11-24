"""Ribbon switching, Active Word COM monitor, LibreOffice/DOTX status, and settings mixin."""

from __future__ import annotations

import logging
import os
import shutil
import sys
from pathlib import Path

from PyQt5.QtWidgets import QAction, QMenu, QMessageBox

from core.ui.theme import C, R, T
from core.ui.ui_components import TOGGLE_PILL_STYLE
from core.utils.folder_opener import open_folder
from modules.word_tools.core.document_opener import get_active_word_document
from modules.word_tools.core.styles_3gpp import (
    install_3gpp_dotx_template,
    is_3gpp_dotx_installed,
    uninstall_3gpp_dotx_template,
)
from modules.word_tools.core.word_config import WordConfig

logger = logging.getLogger(__name__)


class WordTabStatusAndPollMixin:
    """Manages workbench navigation, Active Word COM polling, and environment checks."""

    def set_active_operation(self, index: int):
        self.stack.setCurrentIndex(index)
        self._update_nav_pills(index)
        if self.op_combo.currentIndex() != index:
            self.op_combo.blockSignals(True)
            self.op_combo.setCurrentIndex(index)
            self.op_combo.blockSignals(False)

    def _on_op_combo_changed(self, index: int):
        self.set_active_operation(index)

    def open_workbench(self, key_or_index):
        """Switch Word Tools to the named workbench card or stack index."""
        key_map = {
            "doc_to_docx": 0,
            "visio": 1,
            "headings": 2,
            "excerpt": 3,
            "slice": 3,
            "compare": 4,
            "converter": 5,
            "markdown": 6,
            "cr_pcr": 7,
            "sanitizer": 8,
        }
        if isinstance(key_or_index, int):
            idx = key_or_index
        else:
            idx = key_map.get(str(key_or_index).strip().lower(), 4)
        self.set_active_operation(idx)

    def _update_nav_pills(self, active_index: int):
        for btn, idx in self._nav_buttons:
            is_active = idx == active_index
            btn.setChecked(is_active)
            if is_active:
                btn.setStyleSheet(
                    f"background-color: {C.ACCENT_SOFT}; border: 2px solid {C.ACCENT_INDIGO_BLUE}; "
                    f"color: {C.ACCENT_INDIGO_BLUE}; font-weight: 700; border-radius: {R.PILL}px; "
                    f"padding: 4px 12px; font-size: {T.SIZE_TABLE}px;"
                )
            else:
                btn.setStyleSheet(TOGGLE_PILL_STYLE)

    def showEvent(self, event):
        super().showEvent(event)
        self._poll_active_word(force=True)

    def _on_active_word_clicked(self):
        self._poll_active_word(force=True)
        if self._active_word_file and self._active_word_file.is_file():
            open_folder(self._active_word_file.parent)

    def _poll_active_word(self, force: bool = False):
        if not force and not self.isVisible():
            return
        try:
            facade = sys.modules.get("modules.word_tools.ui.word_tabs")
            get_active = (
                getattr(facade, "get_active_word_document", get_active_word_document)
                if facade is not None
                else get_active_word_document
            )
            # Background polling must pass allow_unsaved_snapshot=False to avoid disk writes
            doc_info = get_active(allow_unsaved_snapshot=False)
            if doc_info and doc_info.get("name"):
                doc_name = doc_info["name"]
                raw_path = doc_info.get("path")
                file_path = Path(raw_path) if raw_path and Path(raw_path).is_file() else None
                self._active_word_file = file_path
                mod_suffix = "*" if doc_info.get("modified") else ""
                self.btn_active_word.setText(f"🪟 Active Word: {doc_name}{mod_suffix}")
                self.btn_active_word.setStyleSheet(f"""
                    QPushButton {{
                        background-color: {C.WG_RAN_BG};
                        color: {C.WG_RAN_TEXT};
                        border: 1px solid {C.EMERALD_BORDER};
                        border-radius: 4px;
                        padding: 3px 10px;
                        font-weight: 600;
                        font-size: 11px;
                    }}
                """)
                if file_path:
                    self.btn_active_word.setToolTip(
                        f"Active in Microsoft Word:\n{file_path.resolve()}\n\nClick to reveal in folder."
                    )
                else:
                    self.btn_active_word.setToolTip(
                        f"Active in Microsoft Word:\n{doc_name} (Unsaved / Cloud-backed)"
                    )
            else:
                self._active_word_file = None
                self.btn_active_word.setText("🪟 Active Word: None Detected")
                self.btn_active_word.setStyleSheet(f"""
                    QPushButton {{
                        background-color: {C.BG_CANVAS};
                        color: {C.TEXT_MUTED};
                        border: 1px solid {C.BORDER_STRONG};
                        border-radius: 4px;
                        padding: 3px 10px;
                        font-weight: 500;
                        font-size: 11px;
                    }}
                """)
                self.btn_active_word.setToolTip(
                    "No active document currently open in Microsoft Word."
                )
        except Exception as exc:
            logger.debug("Active Word poll failed: %s", exc)

    def _setup_settings_menu(self):
        menu = QMenu(self)

        act_settings = QAction("⚙️ Preferences && Sensitivity Labels...", self)
        act_settings.triggered.connect(self._open_settings_dialog)
        menu.addAction(act_settings)

        act_batch_label = QAction("🏷️ Batch Apply Sensitivity Labels...", self)
        act_batch_label.triggered.connect(self._open_batch_label_dialog)
        menu.addAction(act_batch_label)

        menu.addSeparator()

        act_templates = QAction("📂 Open Word Templates Folder", self)
        act_templates.triggered.connect(self._open_word_templates_folder)
        menu.addAction(act_templates)

        self.settings_btn.setMenu(menu)

    def _open_settings_dialog(self):
        from modules.word_tools.ui.word_settings_dialog import WordSettingsDialog

        dlg = WordSettingsDialog(parent=self)
        dlg.exec_()
        self._check_dotx_status()

    def _open_batch_label_dialog(self):
        from modules.word_tools.ui.batch_label_dialog import BatchSensitivityLabelDialog

        dlg = BatchSensitivityLabelDialog(parent=self)
        dlg.exec_()

    def _open_word_templates_folder(self):
        roaming = os.environ.get("APPDATA", "")
        templates_dir = Path(roaming) / "Microsoft" / "Templates"
        if templates_dir.is_dir() and hasattr(os, "startfile"):
            os.startfile(str(templates_dir))
        else:
            QMessageBox.information(
                self, "Templates Folder", f"Microsoft Word Templates directory:\n{templates_dir}"
            )

    def _check_libreoffice_status(self):
        lo_path = WordConfig.load().get("libreoffice_path", "")
        if lo_path and Path(lo_path).is_file():
            self.lo_badge.setText("🟢 LibreOffice: Configured")
            self.lo_badge.setStyleSheet(
                f"color: {C.SUCCESS_DARK}; font-weight: 600; font-size: 11px;"
            )
        else:
            found = shutil.which("soffice") or shutil.which("libreoffice")
            if found:
                self.lo_badge.setText("🟢 LibreOffice: In PATH")
                self.lo_badge.setStyleSheet(
                    f"color: {C.SUCCESS_DARK}; font-weight: 600; font-size: 11px;"
                )
            else:
                self.lo_badge.setText("🟠 LibreOffice: Not Found")
                self.lo_badge.setStyleSheet(
                    f"color: {C.WARNING_DARK}; font-weight: 600; font-size: 11px;"
                )

    def _check_dotx_status(self):
        installed = is_3gpp_dotx_installed()
        if installed:
            self.dotx_badge.setText("🟢 3GPP Macro Template: Installed")
            self.dotx_badge.setStyleSheet(
                f"color: {C.SUCCESS_DARK}; font-weight: 600; font-size: 11px;"
            )
            self.dotx_btn.setText("Uninstall Template")
            self.dotx_btn.setStyleSheet(
                f"font-size: 10px; padding: 2px 6px; color: {C.DANGER_DARK};"
            )
            if hasattr(self, "cr_card_dotx_lbl"):
                self.cr_card_dotx_lbl.setText("🟢 Installed in Microsoft Templates")
                self.cr_card_dotx_lbl.setStyleSheet(
                    f"color: {C.WG_RAN_TEXT}; font-weight: bold; font-size: 11px;"
                )
                self.btn_cr_install_dotx.setText("🗑️ Remove")
        else:
            self.dotx_badge.setText("⚪ 3GPP Macro Template: Not Installed")
            self.dotx_badge.setStyleSheet(f"color: {C.TEXT_SECONDARY}; font-size: 11px;")
            self.dotx_btn.setText("Install Template")
            self.dotx_btn.setStyleSheet(
                f"font-size: 10px; padding: 2px 6px; color: {C.INFO_DARK};"
            )
            if hasattr(self, "cr_card_dotx_lbl"):
                self.cr_card_dotx_lbl.setText("⚪ Not installed in Word templates")
                self.cr_card_dotx_lbl.setStyleSheet(
                    f"color: {C.TEXT_MUTED}; font-size: 11px;"
                )
                self.btn_cr_install_dotx.setText("📥 Install")

    def _on_toggle_dotx_clicked(self):
        if is_3gpp_dotx_installed():
            reply = QMessageBox.question(
                self,
                "Uninstall 3GPP Template",
                "Remove 3gpp_70.dotx from your Microsoft Word Templates folder?",
                QMessageBox.Yes | QMessageBox.No,
            )
            if reply == QMessageBox.Yes:
                uninstall_3gpp_dotx_template()
                self._check_dotx_status()
        else:
            ok, msg = install_3gpp_dotx_template()
            if ok:
                QMessageBox.information(self, "Template Installed", msg)
            else:
                QMessageBox.warning(self, "Installation Failed", msg)
            self._check_dotx_status()
