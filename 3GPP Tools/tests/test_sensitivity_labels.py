# --- File: tests/test_sensitivity_labels.py ---
import gzip
import os
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest.mock import MagicMock, patch
from PyQt5.QtWidgets import QApplication

# Ensure QApplication exists for headless test runs
app = QApplication.instance() or QApplication([])

import modules.word_tools.core.word_config as word_config_mod
from modules.word_tools.core.word_config import WordConfig
from modules.word_tools.core.sensitivity_label import (
    discover_windows_sensitivity_policy,
    get_document_sensitivity_label,
    set_sensitivity_label,
    batch_set_sensitivity_label,
)
from modules.word_tools.ui.word_settings_dialog import WordSettingsDialog, SensitivityLabelEditDialog
from modules.word_tools.ui.batch_label_dialog import BatchSensitivityLabelDialog


SAMPLE_CLP_XML = """<?xml version="1.0" encoding="utf-8"?>
<SyncFile>
  <TenantId>11111111-2222-3333-4444-555555555555</TenantId>
  <Type>MipPolicies</Type>
  <Timestamp>2026-11-17T12:00:00Z</Timestamp>
  <Content>
    <Upn>delegate@example.org</Upn>
    <Uoid>aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee</Uoid>
    <labels>
      <settings>
        <setting key="mandatory" value="true" />
      </settings>
      <label id="aaaa0001-0000-0000-0000-000000000001" name="PUBLIC" enabled="true" isParentLabel="false">
        <description locale="en-US">Public documents</description>
        <settings>
          <setting key="displayName" type="localizedString">
            <value locale="default">PUBLIC - Unencrypted</value>
          </setting>
          <setting key="tooltip" type="localizedString">
            <value locale="default">Safe for external sharing</value>
          </setting>
          <setting key="order" value="0" />
          <setting key="isparent" value="False" />
          <setting key="contenttype" value="File, Email" />
        </settings>
      </label>
      <label id="aaaa0002-0000-0000-0000-000000000002" name="INTERNAL" enabled="true" isParentLabel="false">
        <settings>
          <setting key="displayName" type="localizedString">
            <value locale="default">INTERNAL</value>
          </setting>
          <setting key="order" value="1" />
          <setting key="isparent" value="False" />
          <setting key="contenttype" value="File, Email" />
        </settings>
      </label>
      <label id="aaaa0003-0000-0000-0000-000000000003" name="CONFIDENTIAL" enabled="true" isParentLabel="false">
        <settings>
          <setting key="displayName" type="localizedString">
            <value locale="default">CONFIDENTIAL - Encrypted</value>
          </setting>
          <setting key="tooltip" type="localizedString">
            <value locale="default">Restricted access with RMS encryption</value>
          </setting>
          <setting key="order" value="2" />
          <setting key="isparent" value="False" />
          <setting key="contenttype" value="File, Email" />
        </settings>
      </label>
      <label id="aaaa0004-0000-0000-0000-000000000004" name="MEETINGS_ONLY" enabled="true" isParentLabel="false">
        <settings>
          <setting key="order" value="3" />
          <setting key="isparent" value="False" />
          <setting key="contenttype" value="Teamwork" />
        </settings>
      </label>
      <label id="aaaa0005-0000-0000-0000-000000000005" name="DISABLED_LABEL" enabled="false" isParentLabel="false">
        <settings>
          <setting key="order" value="4" />
          <setting key="isparent" value="False" />
          <setting key="contenttype" value="File" />
        </settings>
      </label>
      <label id="aaaa0006-0000-0000-0000-000000000006" name="PARENT_GROUP" enabled="true" isParentLabel="true">
        <settings>
          <setting key="order" value="5" />
          <setting key="isparent" value="True" />
          <setting key="contenttype" value="File" />
        </settings>
      </label>
    </labels>
    <policies>
      <policy id="p1">
        <rules>
          <rule name="r1" id="r1">
            <condition>
              <containsClassification property="Item.ClassificationDiscovered">
                <keyValues>
                  <keyValue key="Label" value="aaaa0003-0000-0000-0000-000000000003" />
                </keyValues>
              </containsClassification>
            </condition>
            <action name="RightsProtectMessage">
              <argument key="Disabled" value="false" />
            </action>
          </rule>
        </rules>
      </policy>
    </policies>
  </Content>
</SyncFile>
"""


class TestSensitivityLabels(unittest.TestCase):
    def test_discover_windows_sensitivity_policy_from_gzip_xml(self):
        """Verify CLP gzip XML parsing filters File labels, orders them, detects RMS encryption, and picks default."""
        with tempfile.TemporaryDirectory() as tmp:
            gz_path = Path(tmp) / "policy.delegate@example.org_ADAL.gz"
            with gzip.open(gz_path, "wt", encoding="utf-8") as gz:
                gz.write(SAMPLE_CLP_XML)

            policy = discover_windows_sensitivity_policy(clp_dir=tmp)
            self.assertIsNotNone(policy)
            self.assertEqual(policy["tenant_id"], "11111111-2222-3333-4444-555555555555")
            self.assertEqual(policy["upn"], "delegate@example.org")
            self.assertTrue(policy["mandatory"])
            self.assertEqual(policy["recommended_default_id"], "aaaa0001-0000-0000-0000-000000000001")

            labels = policy["labels"]
            self.assertEqual([l["name"] for l in labels], ["PUBLIC", "INTERNAL", "CONFIDENTIAL"])
            self.assertFalse(labels[0]["is_encrypted"])
            self.assertFalse(labels[1]["is_encrypted"])
            self.assertTrue(labels[2]["is_encrypted"])
            self.assertEqual(labels[0]["display_name"], "PUBLIC - Unencrypted")

    def test_no_policy_returns_empty_labels_and_noops(self):
        """Verify machines without a Windows CLP policy get no hardcoded labels and set_sensitivity_label cleanly no-ops."""
        with tempfile.TemporaryDirectory() as tmp:
            cfg_file = Path(tmp) / "word_config.json"
            empty_clp = Path(tmp) / "empty_clp"
            empty_clp.mkdir()

            with patch.object(word_config_mod, "CONFIG_PATH", cfg_file), patch(
                "modules.word_tools.core.sensitivity_label.discover_windows_sensitivity_policy",
                return_value=None,
            ):
                labels = WordConfig.get_sensitivity_labels()
                self.assertEqual(labels, [])
                self.assertEqual(WordConfig.get_active_sensitivity_label(), {})

                mock_doc = MagicMock()
                changed = set_sensitivity_label(mock_doc, return_changed=True)
                self.assertFalse(changed)
                mock_doc.SensitivityLabel.SetLabel.assert_not_called()

    def test_word_config_auto_populates_and_switches_from_policy(self):
        """Verify WordConfig auto-populates from discovered Windows policy and uses real tenant_id as SiteId."""
        with tempfile.TemporaryDirectory() as tmp:
            cfg_file = Path(tmp) / "word_config.json"
            clp_dir = Path(tmp) / "clp"
            clp_dir.mkdir()
            with gzip.open(clp_dir / "policy.delegate@example.org.gz", "wt", encoding="utf-8") as gz:
                gz.write(SAMPLE_CLP_XML)

            parsed_policy = discover_windows_sensitivity_policy(clp_dir=clp_dir)

            with patch.object(word_config_mod, "CONFIG_PATH", cfg_file), patch(
                "modules.word_tools.core.sensitivity_label.discover_windows_sensitivity_policy",
                return_value=parsed_policy,
            ):
                labels = WordConfig.get_sensitivity_labels()
                self.assertEqual(len(labels), 3)
                self.assertEqual(WordConfig.get_active_sensitivity_label()["name"], "PUBLIC")

                # Switch active label to INTERNAL
                WordConfig.set_active_sensitivity_label_id("aaaa0002-0000-0000-0000-000000000002")
                self.assertEqual(WordConfig.get_active_sensitivity_label()["name"], "INTERNAL")

                # Verify set_sensitivity_label applies TenantId as SiteId
                mock_doc = MagicMock()
                mock_doc.SensitivityLabel.GetLabel.return_value.AssignmentMethod = -1
                new_sl = MagicMock()
                mock_doc.SensitivityLabel.CreateLabelInfo.return_value = new_sl

                with patch("modules.word_tools.core.sensitivity_label.platform.system", return_value="Windows"):
                    changed = set_sensitivity_label(mock_doc, return_changed=True)

                self.assertTrue(changed)
                self.assertEqual(new_sl.LabelId, "aaaa0002-0000-0000-0000-000000000002")
                self.assertEqual(new_sl.LabelName, "INTERNAL")
                self.assertEqual(new_sl.SiteId, "11111111-2222-3333-4444-555555555555")

    def test_encrypted_label_indicator_formatting(self):
        """Verify encrypted labels are clearly badged with lock icon and [Encrypted]."""
        unencrypted = {"id": "12345678-aaaa-bbbb-cccc-1234567890ab", "name": "PUBLIC", "display_name": "PUBLIC - Open", "is_encrypted": False}
        encrypted = {"id": "87654321-aaaa-bbbb-cccc-1234567890ab", "name": "CONFIDENTIAL", "display_name": "CONFIDENTIAL", "is_encrypted": True}

        self.assertEqual(WordConfig.format_label_display(unencrypted), "PUBLIC — PUBLIC - Open")
        self.assertEqual(WordConfig.format_label_display(encrypted), "🔒 CONFIDENTIAL [Encrypted]")
        self.assertIn("(...567890ab)", WordConfig.format_label_display(encrypted, include_guid=True))

    def test_inspect_mock_docx_sensitivity_label(self):
        """Verify headless OpenXML inspection on synthetic .docx files."""
        custom_xml_content = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
    <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="MSIP_Label_aaaa0002-0000-0000-0000-000000000002_Enabled"><vt:lpwstr>true</vt:lpwstr></property>
    <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="3" name="MSIP_Label_aaaa0002-0000-0000-0000-000000000002_SetDate"><vt:lpwstr>2026-10-28T06:28:37Z</vt:lpwstr></property>
    <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="4" name="MSIP_Label_aaaa0002-0000-0000-0000-000000000002_Method"><vt:lpwstr>Privileged</vt:lpwstr></property>
    <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="5" name="MSIP_Label_aaaa0002-0000-0000-0000-000000000002_Name"><vt:lpwstr>INTERNAL</vt:lpwstr></property>
    <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="6" name="MSIP_Label_aaaa0002-0000-0000-0000-000000000002_SiteId"><vt:lpwstr>11111111-2222-3333-4444-555555555555</vt:lpwstr></property>
</Properties>"""

        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tf:
            temp_path = tf.name

        try:
            with zipfile.ZipFile(temp_path, "w") as zf:
                zf.writestr("docProps/custom.xml", custom_xml_content.encode("utf-8"))

            label_info = get_document_sensitivity_label(temp_path)
            self.assertIsNotNone(label_info)
            self.assertEqual(label_info["name"], "INTERNAL")
            self.assertEqual(label_info["id"], "aaaa0002-0000-0000-0000-000000000002")
            self.assertEqual(label_info["method"], "Privileged")
            self.assertEqual(label_info["site_id"], "11111111-2222-3333-4444-555555555555")
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

    def test_word_settings_dialog_sync_and_presets(self):
        """Verify WordSettingsDialog displays policy status, sync button, and encrypted indicators."""
        with tempfile.TemporaryDirectory() as tmp:
            cfg_file = Path(tmp) / "word_config.json"
            clp_dir = Path(tmp) / "clp"
            clp_dir.mkdir()
            with gzip.open(clp_dir / "policy.delegate@example.org.gz", "wt", encoding="utf-8") as gz:
                gz.write(SAMPLE_CLP_XML)
            parsed_policy = discover_windows_sensitivity_policy(clp_dir=clp_dir)

            with patch.object(word_config_mod, "CONFIG_PATH", cfg_file), patch(
                "modules.word_tools.core.sensitivity_label.discover_windows_sensitivity_policy",
                return_value=parsed_policy,
            ):
                dlg = WordSettingsDialog()
                self.assertTrue(hasattr(dlg, "combo_label_presets"))
                self.assertTrue(hasattr(dlg, "btn_sync_policy"))
                self.assertTrue(hasattr(dlg, "lbl_policy_status"))
                self.assertIn("delegate@example.org", dlg.lbl_policy_status.text())
                self.assertEqual(dlg.combo_label_presets.count(), 3)
                items = [dlg.combo_label_presets.itemText(i) for i in range(dlg.combo_label_presets.count())]
                self.assertTrue(any("🔒" in t and "[Encrypted]" in t for t in items))

                # Verify SensitivityLabelEditDialog
                edit_dlg = SensitivityLabelEditDialog(
                    dlg,
                    name="CONFIDENTIAL",
                    guid="11223344-5566-7788-9900-aabbccddeeff",
                    description="Test Level",
                    is_encrypted=True,
                )
                self.assertEqual(edit_dlg.edit_name.text(), "CONFIDENTIAL")
                self.assertTrue(edit_dlg.chk_encrypted.isChecked())
                data = edit_dlg.get_data()
                self.assertTrue(data["is_encrypted"])
                edit_dlg.close()
                dlg.close()

    def test_batch_sensitivity_label_dialog_dual_mode(self):
        """Verify BatchSensitivityLabelDialog supports both folder mode and specific document selection."""
        with tempfile.TemporaryDirectory() as tmp:
            cfg_file = Path(tmp) / "word_config.json"
            with patch.object(word_config_mod, "CONFIG_PATH", cfg_file), patch(
                "modules.word_tools.core.sensitivity_label.discover_windows_sensitivity_policy",
                return_value=None,
            ):
                dlg = BatchSensitivityLabelDialog()
                self.assertFalse(dlg.btn_apply.isEnabled())
                self.assertEqual(dlg.tabs.count(), 2)

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
