"""Tests for Phase C: 5G NR RAN Uu CP/UP packet stacks, channel mapping, configurable depth, and RLC segmentation."""

from __future__ import annotations

import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

from bs4 import BeautifulSoup
from PyQt5.QtCore import QUrl
from PyQt5.QtWidgets import QApplication

from core.utils.utils import get_best_java
from modules.nas.core.nas_db import NASDatabase
from modules.nas.core.packet_stack import (
    PacketConfig,
    compose_packet,
    compose_ran_uu_packet,
    infer_rrc_channel_profile,
    render_packet_html,
    render_packet_plantuml_packetdiag,
    resolve_packet_target,
)
from modules.nas.ui.nas_components import NASInspectorWidget

app = QApplication.instance() or QApplication(sys.argv)


class TestRanUuPacketStack(unittest.TestCase):
    """Tests TS 38.300 / 38.331 / 37.324 / 38.323 / 38.322 / 38.321 CP & UP stack composition."""

    def test_target_resolution_and_channel_mapping(self):
        self.assertIsNone(resolve_packet_target("36.331", "6.2.2", "RRCConnectionReconfiguration"))

        p_dl = infer_rrc_channel_profile("DLInformationTransfer")
        self.assertEqual((p_dl["bearer"], p_dl["logical"], p_dl["rlc_mode"], p_dl["lcid"]), ("SRB2", "DCCH", "AM", 2))

        p_setup = infer_rrc_channel_profile("RRCSetup")
        self.assertEqual((p_setup["bearer"], p_setup["logical"], p_setup["rlc_mode"], p_setup["pdcp"]), ("SRB0", "CCCH", "TM", False))

        p_recfg = infer_rrc_channel_profile("RRCReconfiguration")
        self.assertEqual((p_recfg["bearer"], p_recfg["logical"], p_recfg["rlc_mode"], p_recfg["lcid"]), ("SRB1", "DCCH", "AM", 1))

        p_pg = infer_rrc_channel_profile("Paging")
        self.assertEqual((p_pg["logical"], p_pg["transport"], p_pg["rlc_mode"], p_pg["pdcp"]), ("PCCH", "PCH", "TM", False))

        p_sib = infer_rrc_channel_profile("SIB1")
        self.assertEqual((p_sib["logical"], p_sib["transport"], p_sib["rlc_mode"], p_sib["pdcp"]), ("BCCH", "DL-SCH", "TM", False))

        self.assertEqual(resolve_packet_target("37.324", "6.2.2.2", "DL Data PDU with SDAP header")["preset_id"], "uu_up_drb_dl")
        self.assertEqual(resolve_packet_target("38.323", "6.2.2.1", "Data PDU for SRBs")["preset_id"], "uu_cp_srb1_rrc")
        self.assertTrue(resolve_packet_target("38.323", "6.2.2.3", "Data PDU with 18 bits PDCP SN")["auto_enable_flags"]["pdcp_sn_18"])
        self.assertTrue(resolve_packet_target("38.322", "6.2.2.3", "UMD PDU")["auto_enable_flags"]["rlc_um"])

    def test_configurable_stack_depth(self):
        def layer_ids(depth: str) -> list[str]:
            pkt = compose_ran_uu_packet(
                PacketConfig(stack_family="ran_uu", target_frame="uu_up", preset_id="uu_up_drb_dl", transport_mode=depth)
            )
            return [l.layer_id for l in pkt.layers]

        ids_mac = layer_ids("uu_depth_mac")
        self.assertIn("uu_mac_sub", ids_mac)
        self.assertIn("uu_rlc_am", ids_mac)
        self.assertIn("uu_pdcp_drb12", ids_mac)
        self.assertIn("uu_sdap_dl", ids_mac)

        ids_rlc = layer_ids("uu_depth_rlc")
        self.assertNotIn("uu_mac_sub", ids_rlc)
        self.assertIn("uu_rlc_am", ids_rlc)
        self.assertIn("uu_pdcp_drb12", ids_rlc)

        ids_pdcp = layer_ids("uu_depth_pdcp")
        self.assertNotIn("uu_mac_sub", ids_pdcp)
        self.assertNotIn("uu_rlc_am", ids_pdcp)
        self.assertIn("uu_pdcp_drb12", ids_pdcp)

        ids_l3 = layer_ids("uu_depth_l3")
        self.assertNotIn("uu_mac_sub", ids_l3)
        self.assertNotIn("uu_rlc_am", ids_l3)
        self.assertNotIn("uu_pdcp_drb12", ids_l3)
        self.assertIn("uu_sdap_dl", ids_l3)

    def test_mac_ce_dl_front_vs_ul_tail_and_2tb_rlc_segmentation(self):
        # DL MAC CE precedes MAC SDU subheader; UL MAC CE follows payload
        dl_ce = compose_ran_uu_packet(
            PacketConfig(stack_family="ran_uu", target_frame="uu_up", preset_id="uu_up_drb_dl", flags={"mac_ce": True})
        )
        dl_ids = [l.layer_id for l in dl_ce.layers]
        self.assertLess(dl_ids.index("uu_mac_ce"), dl_ids.index("uu_rlc_am"))

        ul_ce = compose_ran_uu_packet(
            PacketConfig(stack_family="ran_uu", target_frame="uu_up", preset_id="uu_up_drb_ul", flags={"mac_ce": True})
        )
        ul_ids = [l.layer_id for l in ul_ce.layers]
        self.assertGreater(ul_ids.index("uu_mac_ce"), ul_ids.index("uu_up_tpdu"))

        # 2-TB RLC Segmentation on UP (18b SN + MAC-I) and CP (SRB2 + NAS)
        pkt_up_seg = compose_packet(
            PacketConfig(
                stack_family="ran_uu",
                target_frame="uu_up",
                preset_id="uu_up_drb_dl",
                flags={"sdap_header": True, "pdcp_sn_18": True, "pdcp_up_integrity": True, "rlc_segmented": True},
            )
        )
        up_ids = [l.layer_id for l in pkt_up_seg.layers]
        self.assertIn("uu_seg_divider", up_ids)
        div_idx = up_ids.index("uu_seg_divider")
        self.assertLess(up_ids.index("uu_pdcp_drb18"), div_idx)
        self.assertLess(up_ids.index("uu_sdap_dl"), div_idx)
        self.assertGreater(up_ids.index("uu_pdcp_maci"), div_idx)

        si_vals = [f.value_hint for f in pkt_up_seg.all_fields() if f.name == "SI"]
        self.assertEqual(si_vals, ["01", "10"])
        self.assertIsNotNone(pkt_up_seg.find_field("SO"))

        pkt_cp_seg = compose_packet(
            PacketConfig(
                stack_family="ran_uu",
                target_frame="uu_cp",
                preset_id="uu_cp_srb2_nas",
                message_name="DLInformationTransfer",
                flags={"expand_nas_pdu": True, "nas_security_hdr": True, "rlc_segmented": True},
            )
        )
        cp_ids = [l.layer_id for l in pkt_cp_seg.layers]
        self.assertIn("uu_seg_divider", cp_ids)
        self.assertIn("uu_nas_inner", cp_ids)
        self.assertGreater(cp_ids.index("uu_pdcp_maci"), cp_ids.index("uu_seg_divider"))

        for label, pkt in (("UP-SEG", pkt_up_seg), ("CP-SEG", pkt_cp_seg)):
            html_out = render_packet_html(pkt, interactive_links=True)
            self.assertIn("Channel Mapping:", html_out)
            self.assertIn("NR Uu =", html_out)
            soup = BeautifulSoup(html_out, "html.parser")
            table = soup.find("table")
            self.assertEqual(table.get("width"), "100%")
            for tr in table.find("tbody").find_all("tr"):
                col_sum = sum(int(c.get("colspan", "1")) for c in tr.find_all(["td", "th"]))
                self.assertEqual(col_sum, 11, f"{label} row sum {col_sum}: {tr}")
                tds = tr.find_all("td")
                if len(tds) >= 3:
                    self.assertIsNone(tds[-1].find("a"))

        jar_path = Path("src/assets/plantuml.jar")
        java_exe, _ = get_best_java()
        if jar_path.exists() and java_exe:
            for p in (pkt_up_seg, pkt_cp_seg):
                puml = render_packet_plantuml_packetdiag(p)
                res = subprocess.run(
                    [java_exe, "-jar", str(jar_path), "-pipe", "-tsvg", "-charset", "UTF-8"],
                    input=puml,
                    capture_output=True,
                    text=True,
                    timeout=20,
                )
                self.assertEqual(res.returncode, 0, f"PlantUML NR Uu failed: {res.stderr}")
                self.assertIn("<svg", res.stdout)

    def test_inspector_uu_depth_segmentation_and_side_panel(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            db = NASDatabase(Path(tmpdir) / "ran_uu_test.db")
            db.insert_parsed_spec(
                "38.331",
                "19.3.0",
                messages=[
                    {
                        "clause": "6.2.2",
                        "message_name": "RRCReconfiguration",
                        "direction": "Network to UE",
                        "description": "Signalling radio bearer: SRB1 RLC-SAP: AM Logical channel: DCCH",
                        "ies": [
                            {"information_element": "rrc-TransactionIdentifier", "depth": 0, "type_reference": "RRC-TransactionIdentifier", "presence": "M"},
                            {"information_element": "radioBearerConfig", "depth": 0, "type_reference": "RadioBearerConfig", "presence": "O"},
                        ],
                    }
                ],
                ie_defs=[{"clause": "6.2.2", "ie_name": "RRCReconfiguration", "raw_description": "<p>RRC def</p>"}],
            )
            db.insert_parsed_spec(
                "38.323",
                "19.1.0",
                messages=[],
                ie_defs=[{"clause": "6.3.2", "ie_name": "PDCP SN", "raw_description": f"<p>MARKER_PDCP_SN {'x' * 300}</p>"}],
            )

            inspector = NASInspectorWidget(db=db)
            inspector.display_definitions(
                clause="6.2.2",
                ie_name="RRCReconfiguration",
                spec_number="38.331",
                defs=[{"ie_name": "RRCReconfiguration", "spec_number": "38.331", "version": "19.3.0", "raw_description": "<p>RRC</p>"}],
                containing_msgs=[],
                message_name="RRCReconfiguration",
            )
            self.assertFalse(inspector.mode_segment.isHidden())
            inspector.set_inspector_mode("packet")
            psw = inspector.packet_stack_widget
            self.assertEqual(psw.current_packet().config.stack_family, "ran_uu")

            # Click PDCP SN in packet table -> Side Panel shows 38.323 Cl. 6.3.2 definition inside Packet mode
            psw._on_anchor_clicked(QUrl("action:packet_field:38.323:6.3.2:PDCP%20SN"))
            self.assertEqual(inspector.content_stack.currentIndex(), 1)
            self.assertFalse(psw.side_panel.isHidden())
            self.assertIn("MARKER_PDCP_SN", psw.side_panel.browser.toHtml())

            # Toggle segmentation pill & change depth to RLC
            psw._flag_buttons["rlc_segmented"].setChecked(True)
            self.assertIsNotNone(psw.current_packet().find_field("SO"))

            idx_rlc = psw.transport_combo.findData("uu_depth_rlc")
            psw.transport_combo.setCurrentIndex(idx_rlc)
            self.assertIsNone(psw.current_packet().find_field("LCID"))
