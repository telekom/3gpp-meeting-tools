"""Tests for Phase B: 5GS NAS & Encapsulated Payloads (URSP / UE Policies / 5GSM / SOR / SMS) with N1/N2/Uu carriers."""

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
    compose_nas_packet,
    compose_packet,
    render_packet_html,
    render_packet_plantuml_packetdiag,
    resolve_packet_target,
)
from modules.nas.ui.nas_components import NASInspectorWidget

app = QApplication.instance() or QApplication(sys.argv)


class TestNasPolicyPacketStack(unittest.TestCase):
    """Tests target resolution, 5-layer URSP/5GSM/SOR/SMS nesting, N1/N2/Uu carriers, HTML & PlantUML."""

    def test_resolve_nas_and_ue_policy_targets(self):
        self.assertIsNone(resolve_packet_target("24.301", "8.2.4", "ATTACH REQUEST"))

        t_ursp = resolve_packet_target("24.526", "5.2", "URSP rule")
        self.assertEqual(t_ursp["stack_family"], "nas")
        self.assertEqual(t_ursp["preset_id"], "nas_ue_policy_ursp")
        self.assertEqual(t_ursp["policy_part_type"], "ursp")

        t_andsp = resolve_packet_target("24.526", "5.3.2", "WLANSP (WLAN selection policy)")
        self.assertEqual(t_andsp["policy_part_type"], "andsp")

        for spec, expected in (("24.588", "v2xp"), ("24.555", "prosep"), ("24.514", "rslpp"), ("24.578", "a2xp")):
            t = resolve_packet_target(spec, "5.2", "Policy")
            self.assertEqual(t["stack_family"], "nas")
            self.assertEqual(t["policy_part_type"], expected)

        t_sm = resolve_packet_target("24.501", "8.3.2.1.1", "PDU SESSION ESTABLISHMENT ACCEPT", "PDU SESSION ESTABLISHMENT ACCEPT")
        self.assertEqual(t_sm["preset_id"], "nas_5gsm_in_5gmm")
        self.assertEqual(t_sm["payload_container_type"], "5gsm")

        t_pol = resolve_packet_target("24.501", "D.5.1.1.1", "MANAGE UE POLICY COMMAND", "MANAGE UE POLICY COMMAND")
        self.assertEqual(t_pol["preset_id"], "nas_ue_policy_ursp")
        self.assertEqual(t_pol["payload_container_type"], "ue_policy")
        self.assertEqual(t_pol["message_name"], "DL NAS TRANSPORT")

        t_prov = resolve_packet_target("24.501", "D.5.3.1.1", "UE POLICY PROVISIONING REQUEST", "UE POLICY PROVISIONING REQUEST")
        self.assertEqual(t_prov["message_name"], "UL NAS TRANSPORT")

        t_reg = resolve_packet_target("24.501", "8.2.6.1.1", "REGISTRATION REQUEST", "REGISTRATION REQUEST")
        self.assertEqual(t_reg["preset_id"], "nas_plain_5gmm")

    def test_ursp_five_layer_nesting_across_n1_n2_uu_carriers_and_plantuml(self):
        for carrier in ("n1_standalone", "over_n2_ngap", "over_uu_rrc"):
            pkt = compose_packet(
                PacketConfig(
                    stack_family="nas",
                    target_frame="nas_pdu",
                    preset_id="nas_ue_policy_ursp",
                    transport_mode=carrier,
                    message_name="DL NAS TRANSPORT",
                    payload_container_type="ue_policy",
                    policy_part_type="ursp",
                )
            )
            self.assertEqual(pkt.config.stack_family, "nas")
            self.assertIsNotNone(pkt.find_field("Message authentication code"))
            self.assertIsNotNone(pkt.find_field("DL NAS TRANSPORT message identity"))
            self.assertIsNotNone(pkt.find_field("UPSC"))
            self.assertIsNotNone(pkt.find_field("Length of URSP rule"))
            self.assertIsNotNone(pkt.find_field("Traffic descriptor"))
            self.assertIsNotNone(pkt.find_field("Route selection descriptor"))
            self.assertIsNotNone(pkt.find_field("S-NSSAI type (0x02, SST + SD)"))

            if carrier == "over_n2_ngap":
                self.assertIsNotNone(pkt.find_field("NAS-PDU"))
            elif carrier == "over_uu_rrc":
                self.assertIsNotNone(pkt.find_field("LCID"))
                self.assertIsNotNone(pkt.find_field("D/C"))
                self.assertIsNotNone(pkt.find_field("PDCP SN"))
                self.assertIsNotNone(pkt.find_field("criticalExtensions"))
                self.assertIsNotNone(pkt.find_field("MAC-I"))

            html_out = render_packet_html(pkt, interactive_links=True)
            self.assertIn("5GS NAS =", html_out)
            self.assertIn("Encapsulation Chain:", html_out)
            soup = BeautifulSoup(html_out, "html.parser")
            table = soup.find("table")
            self.assertEqual(table.get("width"), "100%")
            for tr in table.find("tbody").find_all("tr"):
                col_sum = sum(int(c.get("colspan", "1")) for c in tr.find_all(["td", "th"]))
                self.assertEqual(col_sum, 11, f"{carrier} row sum {col_sum}: {tr}")
                tds = tr.find_all("td")
                if len(tds) >= 3:
                    self.assertIsNone(tds[-1].find("a"))

        # PlantUML packetdiag compile check on full Over Uu (MAC -> RLC -> PDCP -> RRC -> 5GMM -> Policy -> URSP)
        puml = render_packet_plantuml_packetdiag(pkt)
        self.assertTrue(puml.startswith("@startpacketdiag"))
        self.assertTrue(puml.endswith("@endpacketdiag"))
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
            self.assertEqual(res.returncode, 0, f"PlantUML NAS+URSP failed: {res.stderr}")
            self.assertIn("<svg", res.stdout)

    def test_5gsm_sor_sms_and_ursp_toggles(self):
        pkt_sm = compose_nas_packet(
            PacketConfig(
                stack_family="nas",
                target_frame="nas_pdu",
                preset_id="nas_5gsm_in_5gmm",
                message_name="DL NAS TRANSPORT",
                nas_message_name="PDU SESSION ESTABLISHMENT ACCEPT",
                payload_container_type="5gsm",
            )
        )
        self.assertIn("PDU SESSION ESTABLISHMENT ACCEPT", pkt_sm.title)
        self.assertIsNotNone(pkt_sm.find_field("Authorized QoS rules"))
        sm_layer_ids = [l.layer_id for l in pkt_sm.layers]
        self.assertIn("nas_outer_5gmm_tail", sm_layer_ids)
        self.assertGreater(sm_layer_ids.index("nas_outer_5gmm_tail"), sm_layer_ids.index("nas_inner_5gsm"))

        pkt_sor = compose_nas_packet(
            PacketConfig(stack_family="nas", target_frame="nas_pdu", preset_id="nas_sor_container")
        )
        self.assertIsNotNone(pkt_sor.find_field("SOR-MAC-IAUSF"))

        pkt_sms = compose_nas_packet(
            PacketConfig(stack_family="nas", target_frame="nas_pdu", preset_id="nas_sms_over_nas")
        )
        self.assertIsNotNone(pkt_sms.find_field("RP-DATA / SMS TPDU"))

        pkt_ursp_opts = compose_nas_packet(
            PacketConfig(
                stack_family="nas",
                target_frame="nas_pdu",
                preset_id="nas_ue_policy_ursp",
                flags={
                    "nas_security_hdr": False,
                    "expand_payload_container": True,
                    "ursp_td_fqdn": True,
                    "ursp_rsd_location": True,
                },
            )
        )
        self.assertIsNone(pkt_ursp_opts.find_field("Message authentication code"))
        self.assertIsNotNone(pkt_ursp_opts.find_field("Destination FQDN type (0xA1)"))
        self.assertIsNotNone(pkt_ursp_opts.find_field("Location criteria type (0x40, TAI/NCGI)"))

        for pkey in ("andsp", "v2xp", "prosep", "rslpp", "a2xp"):
            p_other = compose_nas_packet(
                PacketConfig(
                    stack_family="nas",
                    target_frame="nas_pdu",
                    preset_id="nas_ue_policy_ursp",
                    payload_container_type="ue_policy",
                    policy_part_type=pkey,
                )
            )
            self.assertIn(pkey.upper(), p_other.title.upper())

    def test_inspector_side_panel_with_temp_nas_and_ursp_db(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            db = NASDatabase(Path(tmpdir) / "nas_ursp_test.db")
            db.insert_parsed_spec(
                "24.501",
                "19.7.0",
                messages=[
                    {
                        "clause": "8.2.11.1.1",
                        "message_name": "DL NAS TRANSPORT",
                        "direction": "network → UE",
                        "ies": [
                            {"iei": "", "information_element": "Extended protocol discriminator", "depth": 0, "type_reference": "Extended protocol discriminator 9.2", "presence": "M", "format": "V", "length": "1"},
                            {"iei": "", "information_element": "Security header type", "depth": 0, "type_reference": "Security header type 9.3", "presence": "M", "format": "V", "length": "1/2"},
                            {"iei": "", "information_element": "Spare half octet", "depth": 0, "type_reference": "Spare half octet 9.5", "presence": "M", "format": "V", "length": "1/2"},
                            {"iei": "", "information_element": "DL NAS TRANSPORT message identity", "depth": 0, "type_reference": "Message type 9.7", "presence": "M", "format": "V", "length": "1"},
                            {"iei": "", "information_element": "Payload container type", "depth": 0, "type_reference": "Payload container type 9.11.3.40", "presence": "M", "format": "V", "length": "1/2"},
                            {"iei": "", "information_element": "Spare half octet", "depth": 0, "type_reference": "Spare half octet 9.5", "presence": "M", "format": "V", "length": "1/2"},
                            {"iei": "", "information_element": "Payload container", "depth": 0, "type_reference": "Payload container 9.11.3.39", "presence": "M", "format": "LV-E", "length": "3-65537"},
                        ],
                    }
                ],
                ie_defs=[
                    {
                        "clause": "D.6.2",
                        "ie_name": "UE policy section management list",
                        "raw_description": f"<p>MARKER_D62_DEF {'x' * 300}</p>",
                    }
                ],
            )
            db.insert_parsed_spec(
                "24.526",
                "19.4.0",
                messages=[],
                ie_defs=[
                    {
                        "clause": "5.2",
                        "ie_name": "Encoding of UE policy part type URSP",
                        "raw_description": f"<p>MARKER_24526_URSP {'x' * 300}</p>",
                    }
                ],
            )

            inspector = NASInspectorWidget(db=db)
            inspector.display_definitions(
                clause="5.2",
                ie_name="URSP rule (UE route selection policy)",
                spec_number="24.526",
                defs=[
                    {
                        "ie_name": "Encoding of UE policy part type URSP",
                        "spec_number": "24.526",
                        "version": "19.4.0",
                        "raw_description": "<p>URSP</p>",
                    }
                ],
                containing_msgs=[],
            )
            self.assertFalse(inspector.mode_segment.isHidden())
            inspector.set_inspector_mode("packet")
            psw = inspector.packet_stack_widget
            self.assertEqual(psw.current_packet().config.stack_family, "nas")
            self.assertFalse(psw.policy_part_combo.isHidden())

            # Click inner TS 24.526 Traffic descriptor -> Side Panel renders 24.526 Cl. 5.2 in Packet mode
            psw._on_anchor_clicked(QUrl("action:packet_field:24.526:5.2:Traffic%20descriptor"))
            self.assertEqual(inspector.content_stack.currentIndex(), 1)
            self.assertFalse(psw.side_panel.isHidden())
            self.assertIn("MARKER_24526_URSP", psw.side_panel.browser.toHtml())

            # Click Annex D.6.2 UPSC -> Side Panel resolves TS 24.501 D.6.2 definition
            psw._on_anchor_clicked(QUrl("action:packet_field:24.501:D.6.2:UPSC"))
            self.assertIn("MARKER_D62_DEF", psw.side_panel.browser.toHtml())

            # Switch Carrier to Over Uu RRC and Policy to ANDSP cleanly
            idx_uu = psw.transport_combo.findData("over_uu_rrc")
            psw.transport_combo.setCurrentIndex(idx_uu)
            self.assertIsNotNone(psw.current_packet().find_field("MAC-I"))

            idx_andsp = psw.policy_part_combo.findData("andsp")
            psw.policy_part_combo.setCurrentIndex(idx_andsp)
            self.assertIn("ANDSP", psw.current_packet().title)
