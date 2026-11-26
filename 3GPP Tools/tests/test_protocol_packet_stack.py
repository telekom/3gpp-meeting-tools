"""Unit and UI tests for 5G UP Full Packet Stack, packetdiag export, and flexible NASTab layout."""

from __future__ import annotations

import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

import pandas as pd
from PyQt5.QtCore import QSettings, Qt
from PyQt5.QtWidgets import QApplication

from core.utils.utils import get_best_java
from modules.nas.core.nas_db import NASDatabase
from modules.nas.core.packet_stack import (
    UP_SCENARIO_PRESETS,
    PacketConfig,
    compose_up_packet,
    copy_packet_to_clipboard,
    render_packet_html,
    render_packet_markdown,
    render_packet_plain_text,
    render_packet_plantuml_packetdiag,
    resolve_up_packet_target,
)
from modules.nas.ui.nas_components import NASInspectorWidget
from modules.nas.ui.nas_tabs import NASTab

app = QApplication.instance() or QApplication(sys.argv)


class TestUPPacketStackEngine(unittest.TestCase):
    """Tests pure-Python 5G UP packet composition, 4-octet alignment, DB capabilities, and exporters."""

    def _find_ext_len(self, pkt, layer_prefix: str) -> str:
        for layer in pkt.layers:
            if layer.layer_id.startswith(layer_prefix) and layer.layer_id.endswith("_open"):
                return layer.fields[0].value_hint
        return ""

    def test_dl_pdu_session_padding_and_n_units(self):
        # Minimal: 2B inner -> 1 + 2 + 1 = 4B (0B pad, n=1) + 12B GTP-U = 16B
        pkt_min = compose_up_packet(PacketConfig(target_frame="dl_pdu_session", preset_id="n3_n9", transport_mode="none"))
        self.assertEqual(self._find_ext_len(pkt_min, "ext_pdu_session_dl"), "n=1 (4 B)")
        self.assertEqual(pkt_min.overhead_3gpp_octets, 16)

        # +PPP: 3B inner -> 1 + 3 + 1 = 5B -> 3B pad -> 8B (n=2) + 12B = 20B
        pkt_ppp = compose_up_packet(PacketConfig(target_frame="dl_pdu_session", preset_id="n3_n9", transport_mode="none", flags={"ppp": True}))
        self.assertEqual(self._find_ext_len(pkt_ppp, "ext_pdu_session_dl"), "n=2 (8 B)")
        self.assertEqual(pkt_ppp.overhead_3gpp_octets, 20)
        self.assertTrue(any(f.is_padding and f.octets == 3 for f in pkt_ppp.all_fields()))

        # +PPP+SNP: 6B inner -> 1 + 6 + 1 = 8B -> 0B pad -> 8B (n=2)
        pkt_ppp_snp = compose_up_packet(PacketConfig(target_frame="dl_pdu_session", preset_id="n3_n9", transport_mode="none", flags={"ppp": True, "snp": True}))
        self.assertEqual(self._find_ext_len(pkt_ppp_snp, "ext_pdu_session_dl"), "n=2 (8 B)")
        self.assertFalse(any(f.is_padding for f in pkt_ppp_snp.all_fields()))

        # +QMP: 10B inner -> 1 + 10 + 1 = 12B -> 0B pad (n=3)
        pkt_qmp = compose_up_packet(PacketConfig(target_frame="dl_pdu_session", preset_id="n3_n9", transport_mode="none", flags={"qmp": True}))
        self.assertEqual(self._find_ext_len(pkt_qmp, "ext_pdu_session_dl"), "n=3 (12 B)")

        # +BSSI (activates Octet 3 + 3B BSSize = 6B inner) -> 8B (n=2)
        pkt_bssi = compose_up_packet(PacketConfig(target_frame="dl_pdu_session", preset_id="n3_n9", transport_mode="none", flags={"bssi": True}))
        self.assertEqual(self._find_ext_len(pkt_bssi, "ext_pdu_session_dl"), "n=2 (8 B)")
        self.assertIsNotNone(pkt_bssi.find_field("BSSize"))

        # +TTNBI (activates Octet 3 + 2B TTNB = 5B inner) -> 1B pad -> 8B (n=2)
        pkt_ttnbi = compose_up_packet(PacketConfig(target_frame="dl_pdu_session", preset_id="n3_n9", transport_mode="none", flags={"ttnbi": True}))
        self.assertEqual(self._find_ext_len(pkt_ttnbi, "ext_pdu_session_dl"), "n=2 (8 B)")
        self.assertTrue(any(f.is_padding and f.octets == 1 for f in pkt_ttnbi.all_fields()))

    def test_ul_pdu_session_and_dl_pdu_set_padding(self):
        # UL with new_ie_1 auto-enables new_ie_flag
        pkt_ul = compose_up_packet(
            PacketConfig(target_frame="ul_pdu_session", preset_id="n3_n9", transport_mode="ipv4_udp", flags={"new_ie_1": True})
        )
        self.assertTrue(pkt_ul.config.flags["new_ie_flag"])
        self.assertIsNotNone(pkt_ul.find_field("UL Congestion Information"))
        self.assertEqual(pkt_ul.overhead_transport_octets, 28)

        # DL PDU Set minimal: 5B inner -> 1 + 5 + 1 = 7B -> 1B pad -> 8B (n=2)
        pkt_ps_min = compose_up_packet(
            PacketConfig(target_frame="dl_pdu_set", preset_id="f1u_cu_du", transport_mode="none", flags={"ext_nr_ran": False, "ext_pdu_set": True})
        )
        self.assertEqual(self._find_ext_len(pkt_ps_min, "ext_pdu_set"), "n=2 (8 B)")
        self.assertTrue(any(f.is_padding and f.octets == 1 for f in pkt_ps_min.all_fields()))

        # DL PDU Set +PSSI: 8B inner -> 1 + 8 + 1 = 10B -> 2B pad -> 12B (n=3)
        pkt_ps_sz = compose_up_packet(
            PacketConfig(target_frame="dl_pdu_set", preset_id="f1u_cu_du", transport_mode="none", flags={"ext_nr_ran": False, "ext_pdu_set": True, "pssi": True})
        )
        self.assertEqual(self._find_ext_len(pkt_ps_sz, "ext_pdu_set"), "n=3 (12 B)")
        self.assertTrue(any(f.is_padding and f.octets == 2 for f in pkt_ps_sz.all_fields()))

    def test_all_presets_and_extension_chaining(self):
        for preset in UP_SCENARIO_PRESETS:
            target = "end_marker" if preset.preset_id == "end_marker" else (
                "dl_pdu_set" if preset.preset_id == "n3_n9_pdu_set" else "dl_pdu_session"
            )
            pkt = compose_up_packet(
                PacketConfig(
                    target_frame=target,
                    preset_id=preset.preset_id,
                    flags=dict(preset.default_flags),
                    include_tpdu=preset.include_tpdu,
                )
            )
            self.assertGreater(len(pkt.layers), 0)
            if preset.preset_id == "xnu_forwarding":
                # 0x03 -> 0x84 -> 0x00
                self.assertIn("0x03", pkt.find_field("Next Extension Header Type").value_hint)
                self.assertTrue(any("0x84" in s for s in pkt.ext_header_summaries))
            elif preset.preset_id == "n3_n9_pdu_set":
                # Chained 0x85 -> 0x04 -> 0x00
                self.assertEqual(len(pkt.ext_header_summaries), 2)
            elif preset.preset_id == "end_marker":
                self.assertIsNone(pkt.find_field("T-PDU Payload"))
                self.assertIn("254", pkt.find_field("Message Type").value_hint)
            elif preset.preset_id == "dummy_qmp":
                self.assertIsNone(pkt.find_field("T-PDU Payload"))
                self.assertIsNotNone(pkt.find_field("DL Sending Time Stamp"))

    def test_target_resolver(self):
        # Non-UP specs return None
        self.assertIsNone(resolve_up_packet_target("24.501", "8.2.1", "REGISTRATION REQUEST"))
        self.assertIsNone(resolve_up_packet_target("38.331", "6.2.2", "RRCReconfiguration"))

        # TS 38.415 IE auto-enables trigger flag and highlights field
        res_ts = resolve_up_packet_target("38.415", "5.5.3.9", "DL Sending Time Stamp")
        self.assertEqual(res_ts["target_frame"], "dl_pdu_session")
        self.assertTrue(res_ts["auto_enable_flags"].get("qmp"))
        self.assertEqual(res_ts["highlight_field"], "DL Sending Time Stamp")

        res_ps = resolve_up_packet_target("38.415", "6.5.3.12", "PSSize")
        self.assertEqual(res_ps["target_frame"], "dl_pdu_set")
        self.assertTrue(res_ps["auto_enable_flags"].get("pssi"))

        res_em = resolve_up_packet_target("29.281", "7.3.2.1", "End Marker")
        self.assertEqual(res_em["target_frame"], "end_marker")

    def test_release_capability_db_awareness(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            db = NASDatabase(Path(tmpdir) / "rel_test.db")
            db.insert_parsed_spec(
                "38.415",
                "15.2.0",
                messages=[{
                    "clause": "5.5.2.1",
                    "message_name": "DL PDU SESSION INFORMATION (PDU Type 0)",
                    "ies": [
                        {"information_element": "PDU Type", "field_path": "PDU Type", "presence": "M", "format": "Bit field", "length": "4 bits"},
                        {"information_element": "QMP", "field_path": "QMP", "presence": "M", "format": "Flag", "length": "1 bit"},
                        {"information_element": "SNP", "field_path": "SNP", "presence": "M", "format": "Flag", "length": "1 bit"},
                        {"information_element": "PPP", "field_path": "PPP", "presence": "M", "format": "Flag", "length": "1 bit"},
                    ],
                }],
                ie_defs=[],
            )
            db.insert_parsed_spec(
                "38.415",
                "19.1.0",
                messages=[{
                    "clause": "5.5.2.1",
                    "message_name": "DL PDU SESSION INFORMATION (PDU Type 0)",
                    "ies": [
                        {"information_element": "PDU Type", "field_path": "PDU Type", "presence": "M", "format": "Bit field", "length": "4 bits"},
                        {"information_element": "QMP", "field_path": "QMP", "presence": "M", "format": "Flag", "length": "1 bit"},
                        {"information_element": "MSNP", "field_path": "MSNP", "presence": "M", "format": "Flag", "length": "1 bit"},
                        {"information_element": "BSSI", "field_path": "BSSI", "presence": "C", "format": "Flag", "length": "1 bit"},
                    ],
                }],
                ie_defs=[],
            )

            pkt_r15 = compose_up_packet(PacketConfig(target_frame="dl_pdu_session", spec_version_38415="15.2.0"), db=db)
            flags_r15 = {f.key: f.is_release_present for f in pkt_r15.available_flags}
            self.assertTrue(flags_r15["qmp"])
            self.assertFalse(flags_r15["bssi"])
            self.assertFalse(flags_r15["msnp"])

            pkt_r19 = compose_up_packet(PacketConfig(target_frame="dl_pdu_session", spec_version_38415="19.1.0"), db=db)
            flags_r19 = {f.key: f.is_release_present for f in pkt_r19.available_flags}
            self.assertTrue(flags_r19["bssi"])
            self.assertTrue(flags_r19["msnp"])

    def test_exporters_and_plantuml_packetdiag_compile(self):
        pkt = compose_up_packet(
            PacketConfig(
                target_frame="dl_pdu_session",
                preset_id="n3_n9_pdu_set",
                transport_mode="ipv6_udp",
                flags={"gtpu_seq": True, "ext_pdu_session": True, "ext_pdu_set": True, "qmp": True, "pssi": True},
            )
        )
        html_ui = render_packet_html(pkt, interactive_links=True)
        html_word = render_packet_html(pkt, interactive_links=False)
        self.assertIn("action:packet_field:", html_ui)
        self.assertNotIn("action:", html_word)

        md = render_packet_markdown(pkt)
        self.assertIn("| Octet | Layer / Field |", md)

        txt = render_packet_plain_text(pkt)
        self.assertIn("TEID", txt)

        puml = render_packet_plantuml_packetdiag(pkt)
        self.assertTrue(puml.startswith("@startpacketdiag"))
        self.assertTrue(puml.endswith("@endpacketdiag"))
        self.assertIn("TEID", puml)

        copy_packet_to_clipboard(pkt, mode="packetdiag")
        self.assertIn("@startpacketdiag", QApplication.clipboard().text())

        jar_path = Path("src/assets/plantuml.jar")
        java_exe, _ = get_best_java()
        if jar_path.exists() and java_exe:
            res = subprocess.run(
                [java_exe, "-jar", str(jar_path), "-pipe", "-tsvg", "-charset", "UTF-8"],
                input=puml,
                capture_output=True,
                text=True,
                timeout=20,
            )
            self.assertEqual(res.returncode, 0, f"PlantUML failed: {res.stderr}")
            self.assertIn("<svg", res.stdout)

    def test_ipv4_and_ipv6_html_table_strictly_11_columns_no_ts_rfc(self):
        from bs4 import BeautifulSoup

        for mode in ("ipv4_udp", "ipv6_udp"):
            pkt = compose_up_packet(
                PacketConfig(
                    target_frame="dl_pdu_session",
                    preset_id="n3_n9_pdu_set",
                    transport_mode=mode,
                    flags={"ext_pdu_session": True, "ext_pdu_set": True, "qmp": True},
                )
            )
            html_out = render_packet_html(pkt, interactive_links=True)
            self.assertNotIn("TS RFC", html_out)
            soup = BeautifulSoup(html_out, "html.parser")
            tbody = soup.find("tbody")
            self.assertIsNotNone(tbody)
            for tr in tbody.find_all("tr"):
                col_sum = sum(int(cell.get("colspan", "1")) for cell in tr.find_all(["td", "th"]))
                self.assertEqual(col_sum, 11, f"{mode} row had {col_sum} columns instead of 11: {tr}")


class TestProtocolsFlexibleUI(unittest.TestCase):
    """Tests NASInspectorWidget mode toggle and ProtocolWorkspaceLayoutController ergonomics."""

    def test_inspector_packet_mode_toggle_preserves_legacy_contracts(self):
        inspector = NASInspectorWidget()
        self.assertEqual(inspector.stack.currentIndex(), 0)
        self.assertEqual(inspector.copy_btn.text(), "📋 Copy Definition ▾")
        self.assertEqual(len([a for a in inspector.copy_btn.menu().actions() if not a.isSeparator()]), 4)

        # Non-UP item hides Full Packet segment
        inspector.display_definitions(
            clause="8.2.1",
            ie_name="REGISTRATION REQUEST",
            spec_number="24.501",
            defs=[{"ie_name": "REGISTRATION REQUEST", "spec_number": "24.501", "version": "19.0.0", "raw_description": "<p>NAS</p>"}],
            containing_msgs=[],
        )
        self.assertEqual(inspector.stack.currentIndex(), 1)
        self.assertTrue(inspector.mode_segment.isHidden())

        # UP item reveals Full Packet segment and switches sub-view cleanly
        inspector.display_definitions(
            clause="5.5.2.1",
            ie_name="DL PDU SESSION INFORMATION (PDU Type 0)",
            spec_number="38.415",
            defs=[{"ie_name": "DL PDU SESSION INFORMATION (PDU Type 0)", "spec_number": "38.415", "version": "19.1.0", "raw_description": "<p>UP</p>"}],
            containing_msgs=[],
        )
        self.assertFalse(inspector.mode_segment.isHidden())
        inspector.set_inspector_mode("packet")
        self.assertEqual(inspector.stack.currentIndex(), 1)
        self.assertEqual(inspector.content_stack.currentIndex(), 1)
        psw = inspector.packet_stack_widget
        pkt_default = psw.current_packet()
        self.assertIsNotNone(pkt_default)
        # Default N3/N9 has 0x85 ON and 0x04 OFF (16B 3GPP overhead)
        self.assertEqual(pkt_default.overhead_3gpp_octets, 16)
        self.assertTrue(psw._flag_buttons["ext_pdu_session"].isChecked())
        self.assertFalse(psw._flag_buttons["ext_pdu_set"].isChecked())

        # Switch to N3/N9 + PDU Set Chained (0x85 -> 0x04): both extensions + both flag groups
        idx_chain = psw.scenario_combo.findData("n3_n9_pdu_set")
        psw.scenario_combo.setCurrentIndex(idx_chain)
        pkt_chain = psw.current_packet()
        self.assertEqual(pkt_chain.overhead_3gpp_octets, 24)
        self.assertIn("qmp", psw._flag_buttons)
        self.assertIn("edb", psw._flag_buttons)

        # Switch back to N3/N9: cleanly restores 0x85 ON, 0x04 OFF, 16B overhead
        idx_n3 = psw.scenario_combo.findData("n3_n9")
        psw.scenario_combo.setCurrentIndex(idx_n3)
        pkt_back = psw.current_packet()
        self.assertEqual(pkt_back.overhead_3gpp_octets, 16)
        self.assertIn("DL PDU SESSION INFORMATION", pkt_back.title)

        # Toggle Scenarios guide
        self.assertTrue(psw.scenarios_guide_panel.isHidden())
        psw.btn_scenarios_help.setChecked(True)
        self.assertFalse(psw.scenarios_guide_panel.isHidden())

    def test_layout_controller_modes_popout_collapse_and_autocompact(self):
        settings = QSettings("3GPP_Tools_Test", "ProtocolsLayoutTest")
        settings.clear()

        with tempfile.TemporaryDirectory() as tmpdir:
            tab = NASTab(nas_db_path=Path(tmpdir) / "layout_ui.db", config_path=Path(tmpdir) / "nas_config.json")
            tab.resize(1320, 820)
            tab.show()
            app.processEvents()

            ctrl = tab.layout_controller
            ctrl._settings = settings

            ctrl.set_mode("split_h")
            self.assertEqual(ctrl.current_mode(), "split_h")
            self.assertEqual(tab.right_splitter.orientation(), Qt.Horizontal)

            ctrl.set_mode("inspector_only")
            self.assertTrue(tab.matrix_stack.isHidden())
            self.assertFalse(tab.inspector.stack.isHidden())

            ctrl.set_mode("matrix_only")
            self.assertFalse(tab.matrix_stack.isHidden())
            self.assertTrue(tab.inspector.stack.isHidden())

            # Pop out and dock back
            ctrl.set_mode("popout")
            app.processEvents()
            self.assertIsNotNone(ctrl._popout_win)
            self.assertIs(tab.inspector.parent(), ctrl._popout_win.body_container)

            ctrl.set_mode("split_v")
            app.processEvents()
            self.assertIsNone(ctrl._popout_win)
            self.assertIs(tab.inspector.parent(), tab.right_splitter)

            # Left sidebar collapse
            ctrl.set_left_collapsed(True)
            self.assertTrue(tab.left_pane_widget.isHidden())
            ctrl.set_left_collapsed(False)
            self.assertFalse(tab.left_pane_widget.isHidden())

            # Indexed Protocols summary bar collapse
            ctrl.update_specs_summary([{"spec_number": "38.415", "version": "19.1.0"}])
            ctrl.set_specs_card_collapsed(True)
            self.assertTrue(tab.ver_body_widget.isHidden())
            self.assertFalse(tab.ver_summary_bar.isHidden())
            self.assertIn("TS 38.415", tab.ver_summary_lbl.text())
            ctrl.set_specs_card_collapsed(False)
            self.assertFalse(tab.ver_body_widget.isHidden())

            # Auto-compact on single-row All IEs matrix
            df_one = pd.DataFrame([{
                "iei": "-", "ie_name": "DL PDU SESSION INFORMATION (PDU Type 0)",
                "type_reference": "5.5.2.1", "presence": "Defined", "format": "Clause 5.5.2.1",
                "length": "Rev #1", "spec_number": "38.415", "version": "19.1.0", "order_index": 0,
            }])
            tab._browse_mode = "ies"
            tab._current_matrix_df = df_one
            tab._rebuild_matrix_model(ignore_ie_filter=True)
            self.assertTrue(ctrl._auto_compacted)

            tab.shutdown()
            tab.close()


if __name__ == "__main__":
    unittest.main()
