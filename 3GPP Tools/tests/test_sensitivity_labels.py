# --- File: tests/test_sensitivity_labels.py ---
import os
import tempfile
import unittest
import zipfile
from pathlib import Path
from PyQt5.QtWidgets import QApplication

# Ensure QApplication exists for headless test runs
app = QApplication.instance() or QApplication([])

from modules.word_tools.core.word_config import WordConfig, DEFAULT_SENSITIVITY_LABELS
from modules.word_tools.core.sensitivity_label import (
    get_document_sensitivity_label,
    set_sensitivity_label,
    batch_set_sensitivity_label,
    MsoAssignmentMethod,
)
from modules.word_tools.ui.word_settings_dialog import WordSettingsDialog, SensitivityLabelEditDialog
from modules.word_tools.ui.batch_label_dialog import BatchSensitivityLabelDialog


class TestSensitivityLabels(unittest.TestCase):
    def test_word_config_multi_labels(self):
        """Verify WordConfig supports multiple sensitivity labels and active selection."""
        labels = WordConfig.get_sensitivity_labels()
        self.assertGreaterEqual(len(labels), 3)

        label_names = {l["name"] for l in labels}
        self.assertIn("OFFEN", label_names)
        self.assertIn("INTERN", label_names)
        self.assertTrue(any("VERTRAULICH" in n for n in label_names))

        # Test active label
        active = WordConfig.get_active_sensitivity_label()
        self.assertIsNotNone(active)
        self.assertIn("id", active)
        self.assertIn("name", active)

        # Test switching active label
        intern_label = next(l for l in labels if l["name"] == "INTERN")
        WordConfig.set_active_sensitivity_label_id(intern_label["id"])
        new_active = WordConfig.get_active_sensitivity_label()
        self.assertEqual(new_active["id"], intern_label["id"])
        self.assertEqual(new_active["name"], "INTERN")

        # Reset back to OFFEN
        offen_label = next(l for l in labels if l["name"] == "OFFEN")
        WordConfig.set_active_sensitivity_label_id(offen_label["id"])
        self.assertEqual(WordConfig.get_active_sensitivity_label()["name"], "OFFEN")

    def test_inspect_mock_docx_sensitivity_label(self):
        """Verify headless OpenXML inspection on synthetic .docx files."""
        custom_xml_content = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
    <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="MSIP_Label_5c6efe59-7e7c-4bd6-ba39-1850710f2efa_Enabled"><vt:lpwstr>true</vt:lpwstr></property>
    <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="3" name="MSIP_Label_5c6efe59-7e7c-4bd6-ba39-1850710f2efa_SetDate"><vt:lpwstr>2026-10-28T06:28:37Z</vt:lpwstr></property>
    <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="4" name="MSIP_Label_5c6efe59-7e7c-4bd6-ba39-1850710f2efa_Method"><vt:lpwstr>Privileged</vt:lpwstr></property>
    <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="5" name="MSIP_Label_5c6efe59-7e7c-4bd6-ba39-1850710f2efa_Name"><vt:lpwstr>INTERN</vt:lpwstr></property>
    <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="6" name="MSIP_Label_5c6efe59-7e7c-4bd6-ba39-1850710f2efa_SiteId"><vt:lpwstr>d313b56f-f400-44d3-8403-4b468b3d8ded</vt:lpwstr></property>
</Properties>"""

        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tf:
            temp_path = tf.name

        try:
            with zipfile.ZipFile(temp_path, "w") as zf:
                zf.writestr("docProps/custom.xml", custom_xml_content.encode("utf-8"))

            label_info = get_document_sensitivity_label(temp_path)
            self.assertIsNotNone(label_info)
            self.assertEqual(label_info["name"], "INTERN")
            self.assertEqual(label_info["id"], "5c6efe59-7e7c-4bd6-ba39-1850710f2efa")
            self.assertEqual(label_info["method"], "Privileged")
            self.assertEqual(label_info["site_id"], "d313b56f-f400-44d3-8403-4b468b3d8ded")
            self.assertTrue(label_info["enabled"])
        finally:
            if os.path.exists(temp_path):
                os.remove(temp_path)

    def test_inspect_unlabeled_docx(self):
        """Verify unclassified .docx returns None."""
        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tf:
            temp_path = tf.name

        try:
            with zipfile.ZipFile(temp_path, "w") as zf:
                zf.writestr("[Content_Types].xml", "<Types/>")

            label_info = get_document_sensitivity_label(temp_path)
            self.assertIsNone(label_info)
        finally:
            if os.path.exists(temp_path):
                os.remove(temp_path)

    def test_word_settings_dialog_presets_and_editing(self):
        """Verify WordSettingsDialog populates multiple presets and supports in-UI editing."""
        dlg = WordSettingsDialog()
        self.assertTrue(hasattr(dlg, "combo_label_presets"))
        self.assertGreaterEqual(dlg.combo_label_presets.count(), 3)
        self.assertTrue(hasattr(dlg, "btn_add_label"))
        self.assertTrue(hasattr(dlg, "btn_edit_label"))
        self.assertTrue(hasattr(dlg, "btn_del_label"))
        self.assertTrue(hasattr(dlg, "btn_inspect_doc"))
        self.assertTrue(hasattr(dlg, "btn_test_active"))
        self.assertTrue(hasattr(dlg, "btn_launch_batch"))

        # Verify SensitivityLabelEditDialog
        edit_dlg = SensitivityLabelEditDialog(dlg, name="CONFIDENTIAL", guid="11223344-5566-7788-9900-aabbccddeeff", description="Test Level")
        self.assertEqual(edit_dlg.edit_name.text(), "CONFIDENTIAL")
        self.assertEqual(edit_dlg.edit_guid.text(), "11223344-5566-7788-9900-aabbccddeeff")
        self.assertEqual(edit_dlg.edit_desc.text(), "Test Level")
        data = edit_dlg.get_data()
        self.assertEqual(data["name"], "CONFIDENTIAL")
        self.assertEqual(data["id"], "11223344-5566-7788-9900-aabbccddeeff")
        edit_dlg.close()

        dlg.close()

    def test_batch_sensitivity_label_dialog_dual_mode(self):
        """Verify BatchSensitivityLabelDialog supports both folder mode and specific document selection."""
        dlg = BatchSensitivityLabelDialog()
        self.assertTrue(hasattr(dlg, "tabs"))
        self.assertEqual(dlg.tabs.count(), 2)
        self.assertEqual(dlg.tabs.tabText(0), "📁 Entire Folder")
        self.assertEqual(dlg.tabs.tabText(1), "📄 Specific Documents")

        # Check folder tab controls
        self.assertTrue(hasattr(dlg, "edit_folder"))
        self.assertTrue(hasattr(dlg, "btn_browse_folder"))
        self.assertTrue(hasattr(dlg, "chk_recursive"))

        # Check file tab controls
        self.assertTrue(hasattr(dlg, "tbl_files"))
        self.assertTrue(hasattr(dlg, "btn_add_files"))
        self.assertTrue(hasattr(dlg, "btn_remove_file"))
        self.assertTrue(hasattr(dlg, "btn_clear_files"))

        # Test adding files programmatically
        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tf:
            f1 = Path(tf.name)
        try:
            dlg._add_file_paths([f1])
            self.assertEqual(len(dlg._selected_files), 1)
            self.assertEqual(dlg.tbl_files.rowCount(), 1)
            self.assertIn("1 document(s)", dlg.lbl_file_count.text())
        finally:
            if f1.exists():
                f1.unlink()

        dlg.close()

    def test_batch_set_sensitivity_label_file_paths_empty(self):
        """Verify batch processor handles empty file paths list cleanly without launching Word."""
        res = batch_set_sensitivity_label(file_paths=[])
        self.assertEqual(res["total"], 0)
        self.assertEqual(res["labeled"], 0)
        self.assertEqual(res["skipped"], 0)
        self.assertEqual(res["errors"], 0)


if __name__ == "__main__":
    unittest.main()
