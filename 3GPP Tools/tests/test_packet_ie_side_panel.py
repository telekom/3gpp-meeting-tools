"""Tests for responsive Packet Stack tables, non-blocking Word open, NGAP polish, and Inline IE Side Panel."""

from __future__ import annotations

import sys
import tempfile
import time
import unittest
from pathlib import Path
from unittest import mock

from bs4 import BeautifulSoup
from PyQt5.QtCore import QSettings, Qt, QUrl
from PyQt5.QtTest import QSignalSpy
from PyQt5.QtWidgets import QApplication

from modules.nas.core.nas_db import NASDatabase
from modules.nas.core.packet_stack import (
    PacketConfig,
    compose_ngap_packet,
    compose_up_packet,
    render_packet_html,
    render_packet_markdown,
    render_packet_plain_text,
)
from modules.nas.core.packet_stack.ngap_catalog import _normalize_direction
from modules.nas.ui.components import (
    NASInspectorWidget,
    PacketStackWidget,
    field_lookup_candidates,
    resolve_packet_field,
)
from modules.nas.ui.components.packet_ie_resolver import build_side_panel_html
from modules.spec_search.core.spec_doc_resolver import SpecDocResolver
from modules.word_tools.core.document_opener_async import open_document_async

app = QApplication.instance() or QApplication(sys.argv)


class TestResponsiveTableAndNgapPolish(unittest.TestCase):
    """Tests responsive HTML table width attrs, family-aware wording, and NGAP direction/clause fixes."""

    def test_html_table_has_width_100_and_th_widths_for_up_and_ngap(self):
        up_pkt = compose_up_packet(PacketConfig())
        ngap_pkt = compose_ngap_packet(
            PacketConfig(
                stack_family="ngap",
                target_frame="ngap_pdu",
                preset_id="n2_ue_signalling",
                message_name="DownlinkNASTransport",
            )
        )
        for label, pkt in (("UP", up_pkt), ("NGAP", ngap_pkt)):
            html_out = render_packet_html(pkt, interactive_links=True)
            soup = BeautifulSoup(html_out, "html.parser")
            table = soup.find("table")
            self.assertIsNotNone(table, f"{label} missing table")
            self.assertEqual(table.get("width"), "100%", f"{label} table missing width='100%'")
            ths = table.find("thead").find_all("th")
            widths = [th.get("width") for th in ths]
            self.assertEqual(widths[0], "7%")
            self.assertEqual(widths[2], "9%")
            self.assertEqual(widths[3], "16%")
            self.assertEqual(widths[4:], ["8.5%"] * 8)

            for tr in table.find("tbody").find_all("tr"):
                col_sum = sum(int(c.get("colspan", "1")) for c in tr.find_all(["td", "th"]))
                self.assertEqual(col_sum, 11)

            # Table links must never emit blocking action:open_spec hrefs
            self.assertNotIn("action:open_spec:", html_out)

    def test_family_aware_banner_markdown_and_plain_text(self):
        up_pkt = compose_up_packet(PacketConfig())
        ngap_pkt = compose_ngap_packet(
            PacketConfig(
                stack_family="ngap",
                target_frame="ngap_pdu",
                preset_id="n2_ue_signalling",
                message_name="DownlinkNASTransport",
            )
        )

        up_html = render_packet_html(up_pkt)
        self.assertIn("3GPP UP =", up_html)
        self.assertIn("Ext Header Chain:", up_html)
        self.assertIn("Fixed 3GPP UP Overhead:", render_packet_markdown(up_pkt))
        self.assertIn("3GPP UP =", render_packet_plain_text(up_pkt))

        ngap_html = render_packet_html(ngap_pkt)
        self.assertIn("N2 / NGAP =", ngap_html)
        self.assertIn("NGAP Summary:", ngap_html)
        self.assertNotIn("3GPP UP =", ngap_html)

        ngap_md = render_packet_markdown(ngap_pkt)
        self.assertIn("Fixed N2 / NGAP Overhead:", ngap_md)
        self.assertIn("NGAP Summary:", ngap_md)
        self.assertNotIn("3GPP UP", ngap_md)

        ngap_txt = render_packet_plain_text(ngap_pkt)
        self.assertIn("Overhead  : N2 / NGAP =", ngap_txt)
        self.assertNotIn("3GPP UP", ngap_txt)

    def test_normalize_direction_adds_arrow_for_aiot_and_ngran(self):
        self.assertEqual(_normalize_direction("A-IoT CN node NG-RAN node", 0), "A-IoT CN node → NG-RAN node")
        self.assertEqual(_normalize_direction("NG-RAN node A-IoT CN node", 1), "NG-RAN node → A-IoT CN node")
        self.assertEqual(_normalize_direction("AMF NG-RAN node", 0), "AMF → NG-RAN node")
        self.assertEqual(_normalize_direction("NG-RAN node AMF", 1), "NG-RAN node → AMF")
        self.assertEqual(_normalize_direction("NG-RAN AMF", 1), "NG-RAN node → AMF")
        self.assertEqual(_normalize_direction("AMF → NG-RAN node", 0), "AMF → NG-RAN node")


class TestPacketIESidePanelAndAsyncOpen(unittest.TestCase):
    """Tests per-IE clause lookup, candidate resolution, Side Panel UI, and non-blocking Word open."""

    def _build_temp_ngap_db(self, db_path: Path) -> NASDatabase:
        db = NASDatabase(db_path)
        db.insert_parsed_spec(
            "38.413",
            "19.3.0",
            messages=[
                {
                    "clause": "9.2.5.2",
                    "message_name": "DownlinkNASTransport",
                    "direction": "AMF NG-RAN node",
                    "ies": [
                        {
                            "iei": "4",
                            "information_element": "MessageType",
                            "field_path": "MessageType",
                            "depth": 0,
                            "type_reference": "MessageType (9.3.1.1)",
                            "presence": "M",
                            "format": "IE",
                            "length": "ignore",
                        },
                        {
                            "iei": "10",
                            "information_element": "AMF-UE-NGAP-ID",
                            "field_path": "AMF-UE-NGAP-ID",
                            "depth": 0,
                            "type_reference": "AMF-UE-NGAP-ID",
                            "presence": "M",
                            "format": "IE",
                            "length": "reject",
                        },
                        {
                            "iei": "38",
                            "information_element": "NAS-PDU",
                            "field_path": "NAS-PDU",
                            "depth": 0,
                            "type_reference": "NAS-PDU",
                            "presence": "M",
                            "format": "IE",
                            "length": "reject",
                        },
                    ],
                }
            ],
            ie_defs=[
                {
                    "clause": "9.3.3.1",
                    "ie_name": "AMF-UE-NGAP-ID",
                    "raw_description": f"<p>MARKER_AMF_DEF {'x' * 300}</p>",
                }
            ],
        )
        return db

    def test_field_lookup_candidates(self):
        cands = field_lookup_candidates("AIOTFIdentifier Value", None)
        self.assertIn("AIOTFIdentifier", cands)
        self.assertLess(cands.index("AIOTFIdentifier"), cands.index("AIOTFIdentifier Value"))

        cands_ie = field_lookup_candidates("IE ID = 467 (AIOTFIdentifier)", None)
        self.assertEqual(cands_ie[0], "AIOTFIdentifier")

    def test_ngap_db_per_ie_clause_and_side_panel_resolution(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            db = self._build_temp_ngap_db(Path(tmpdir) / "test_ngap.db")
            pkt = compose_ngap_packet(
                PacketConfig(
                    stack_family="ngap",
                    target_frame="ngap_pdu",
                    preset_id="n2_ue_signalling",
                    message_name="DownlinkNASTransport",
                ),
                db=db,
            )
            self.assertEqual(pkt.direction, "AMF → NG-RAN node")

            ie_fld = pkt.find_field("AMF-UE-NGAP-ID")
            self.assertIsNotNone(ie_fld)
            self.assertEqual(ie_fld.clause, "9.3.3.1")
            self.assertEqual(ie_fld.def_name, "AMF-UE-NGAP-ID")

            val_fld = pkt.find_field("AMF-UE-NGAP-ID Value")
            self.assertIsNotNone(val_fld)
            self.assertEqual(val_fld.clause, "9.3.3.1")

            res = resolve_packet_field(db, pkt, "38.413", "9.3.3.1", "AMF-UE-NGAP-ID Value")
            self.assertTrue(res.found)
            html_out = build_side_panel_html(res)
            self.assertIn("MARKER_AMF_DEF", html_out)

    def test_packet_stack_side_panel_widget_and_inspector_stay_in_packet_mode(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            db = self._build_temp_ngap_db(Path(tmpdir) / "widget_ngap.db")
            settings = QSettings(str(Path(tmpdir) / "layout.ini"), QSettings.IniFormat)
            psw = PacketStackWidget(db=db, settings=settings)
            ok = psw.set_context(
                db=db,
                spec_number="38.413",
                clause="9.2.5.2",
                ie_name="DownlinkNASTransport",
                message_name="DownlinkNASTransport",
            )
            self.assertTrue(ok)
            self.assertTrue(psw.side_panel.isHidden())

            spy = QSignalSpy(psw.open_spec_requested)

            # Clicking a field opens the inline side panel and does NOT open Word
            psw._on_anchor_clicked(QUrl("action:packet_field:38.413:9.3.3.1:AMF-UE-NGAP-ID%20Value"))
            self.assertFalse(psw.side_panel.isHidden())
            self.assertIn("AMF-UE-NGAP-ID", psw.side_panel.title_lbl.text())
            self.assertIn("MARKER_AMF_DEF", psw.side_panel.browser.toHtml())
            self.assertEqual(len(spy), 0)

            # Legacy open_spec link in packet view also routes to inline side panel cleanly
            psw._on_anchor_clicked(QUrl("action:open_spec:38.413::9.2.5.2"))
            self.assertEqual(len(spy), 0)
            self.assertFalse(psw.side_panel.isHidden())

            # Re-open AMF-UE-NGAP-ID, then explicitly click Word in the Side Panel header
            psw._on_anchor_clicked(QUrl("action:packet_field:38.413:9.3.3.1:AMF-UE-NGAP-ID%20Value"))
            psw.side_panel.word_btn.click()
            self.assertEqual(len(spy), 1)
            self.assertEqual(spy[0][0], "38.413")
            self.assertEqual(spy[0][3], "word")

            # Dock toggle switches splitter orientation Horizontal -> Vertical -> Horizontal
            self.assertEqual(psw.body_splitter.orientation(), Qt.Horizontal)
            psw.side_panel.dock_btn.click()
            self.assertEqual(psw.body_splitter.orientation(), Qt.Vertical)
            psw.side_panel.dock_btn.click()
            self.assertEqual(psw.body_splitter.orientation(), Qt.Horizontal)

            # Close hides the side panel
            psw.side_panel.close_btn.click()
            self.assertTrue(psw.side_panel.isHidden())

            # End-to-end inside NASInspectorWidget: clicking a field keeps Packet mode active (index 1)
            inspector = NASInspectorWidget(db=db)
            inspector.display_definitions(
                clause="9.2.5.2",
                ie_name="DownlinkNASTransport",
                spec_number="38.413",
                defs=[
                    {
                        "ie_name": "DownlinkNASTransport",
                        "spec_number": "38.413",
                        "version": "19.3.0",
                        "raw_description": "<p>NGAP Msg</p>",
                    }
                ],
                containing_msgs=[],
                message_name="DownlinkNASTransport",
            )
            inspector.set_inspector_mode("packet")
            self.assertEqual(inspector.content_stack.currentIndex(), 1)
            inspector.packet_stack_widget._on_anchor_clicked(
                QUrl("action:packet_field:38.413:9.3.3.1:AMF-UE-NGAP-ID%20Value")
            )
            self.assertEqual(inspector.content_stack.currentIndex(), 1)
            self.assertFalse(inspector.packet_stack_widget.side_panel.isHidden())

    def test_open_document_async_and_resolver_non_blocking(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            tmp_root = Path(tmpdir)
            spec_dir = tmp_root / "38.413"
            spec_dir.mkdir(parents=True, exist_ok=True)
            dummy_docx = spec_dir / "38413-j30.docx"
            dummy_docx.write_bytes(b"PK\x03\x04dummy")

            # 1. open_document_async returns immediately and invokes callback via Qt signal
            opened_paths = []
            with mock.patch(
                "modules.word_tools.core.document_opener_async.open_document",
                side_effect=lambda _p: time.sleep(0.25),
            ):
                t0 = time.monotonic()
                thread = open_document_async(dummy_docx, on_opened=opened_paths.append)
                elapsed = time.monotonic() - t0
                self.assertLess(elapsed, 0.15)
                self.assertTrue(thread.wait(3000))
                for _ in range(10):
                    app.processEvents()
                    if opened_paths:
                        break
                    time.sleep(0.01)
                self.assertEqual(opened_paths, [str(dummy_docx)])

            # 2. SpecDocResolver.open_spec routes Word open through open_document_async
            fake_info = {
                "spec_dir": spec_dir,
                "word_docs": [dummy_docx],
                "zip_file": None,
                "url": "",
                "filename": "",
                "stem": "38413-j30",
            }
            async_mock = mock.MagicMock()
            with (
                mock.patch.object(SpecDocResolver, "resolve_spec_files", staticmethod(lambda *a, **k: fake_info)),
                mock.patch("modules.spec_search.core.spec_doc_resolver.open_document_async", async_mock),
            ):
                ok = SpecDocResolver.open_spec(
                    spec_number="38.413",
                    version="19.3.0",
                    clause_number="9.3.3.1",
                    doc_format="word",
                    cache_dir=tmp_root,
                )
            self.assertTrue(ok)
            async_mock.assert_called_once()
            self.assertEqual(async_mock.call_args.args[0], dummy_docx)
