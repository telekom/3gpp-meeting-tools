"""Tests for Phase A: Spec/Clause plain column, compact Word paste, spec catalog, L2 parser, and Import Studio."""

from __future__ import annotations

import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from typing import Any, List, Tuple

from bs4 import BeautifulSoup
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QApplication, QPushButton

from modules.nas.core.nas_db import NASDatabase
from modules.nas.core.nas_threads import SpecCacheIndex
from modules.nas.core.nas_exporter import (
    copy_definition_to_clipboard,
    copy_matrix_to_clipboard,
)
from modules.nas.core.packet_stack import (
    PacketConfig,
    compose_ngap_packet,
    compose_up_packet,
    copy_packet_to_clipboard,
    render_packet_html,
)
from modules.nas.core.parsing.protocol_parser_common import ProtocolDocxDispatcher
from modules.nas.core.spec_catalog import (
    PROTOCOL_SPEC_CATALOG,
    spec_matches_filters,
)
from modules.nas.core.word_html import compact_html_for_word
from modules.nas.ui.nas_dialogs import NASVersionSelectDialog

app = QApplication.instance() or QApplication(sys.argv)


def _make_synthetic_l2_docx(docx_path: Path, paragraphs: List[Tuple[str, str, bool]], table_rows: List[List[str]] | None = None) -> None:
    """Create a minimal valid `.docx` containing styled Heading/Body paragraphs, optional drawing, and a table."""
    w_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    body_parts: List[str] = []
    for style, text, has_drawing in paragraphs:
        ppr = f'<w:pPr><w:pStyle w:val="{style}"/></w:pPr>' if style else ""
        run_inner = ""
        if has_drawing:
            run_inner += "<w:drawing/>"
        if text:
            # Escape XML chars cleanly
            esc = (
                text.replace("&", "&amp;")
                .replace("<", "&lt;")
                .replace(">", "&gt;")
            )
            run_inner += f"<w:t>{esc}</w:t>"
        body_parts.append(f"<w:p>{ppr}<w:r>{run_inner}</w:r></w:p>")
    if table_rows:
        trs = []
        for row in table_rows:
            tcs = "".join(
                f"<w:tc><w:p><w:r><w:t>{c.replace('&', '&amp;').replace('<', '&lt;')}</w:t></w:r></w:p></w:tc>"
                for c in row
            )
            trs.append(f"<w:tr>{tcs}</w:tr>")
        body_parts.append(f"<w:tbl>{''.join(trs)}</w:tbl>")

    doc_xml = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document xmlns:w="{w_ns}"><w:body>{"".join(body_parts)}</w:body></w:document>'
    )
    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        "</Types>"
    )
    with zipfile.ZipFile(docx_path, "w") as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("word/document.xml", doc_xml)


class _StubSpecsDB:
    """Minimal in-memory stand-in for `SpecsDatabase` implementing `search_files(spec_number=...)`."""

    def __init__(self, records: dict[str, List[Tuple[Any, ...]]]) -> None:
        self._records = records

    def search_files(self, spec_number: str = "", **_kwargs: Any) -> List[Tuple[Any, ...]]:
        return list(self._records.get(spec_number, []))


class TestPhaseAPacketWordAndL2(unittest.TestCase):
    """Verifies Spec/Clause column links removal, compact Word HTML, L2 catalog, and L2 parser."""

    def test_no_links_in_spec_clause_column_but_field_cells_remain_clickable(self):
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
            tbody = soup.find("tbody")
            self.assertIsNotNone(tbody, label)
            has_field_link = False
            for tr in tbody.find_all("tr"):
                tds = tr.find_all("td")
                if len(tds) < 3:
                    continue  # layer banner row
                last_td = tds[-1]
                self.assertIsNone(last_td.find("a"), f"{label} Spec/Clause cell must not contain hyperlinks: {last_td}")
                for td in tds[1:-2]:
                    for a in td.find_all("a"):
                        self.assertTrue(str(a.get("href", "")).startswith("action:packet_field:"))
                        has_field_link = True
            self.assertTrue(has_field_link, f"{label} should keep clickable field links")

    def test_compact_html_for_word_and_clipboard_exporters(self):
        raw = '<table><tr><th>H1</th><th>H2</th></tr><tr><td style="padding: 6px 10px;">Cell A</td><td><p style="margin: 8px 0;">Para B</p></td></tr></table>'
        compact = compact_html_for_word(raw, full_document=True)
        self.assertIn("p.MsoNormal", compact)
        self.assertIn("mso-padding-alt: 1.5pt 4pt 1.5pt 4pt", compact)
        self.assertIn("margin-bottom: .0001pt", compact)
        self.assertIn("line-height: 105%", compact)
        self.assertIn("Para B</p>", compact)
        self.assertIn("Cell A</p>", compact)

        cb = QApplication.clipboard()
        up_pkt = compose_up_packet(PacketConfig())
        copy_packet_to_clipboard(up_pkt, mode="word")
        pkt_clip = cb.mimeData().html()
        self.assertIn("MsoNormal", pkt_clip)
        self.assertIn("mso-padding-alt", pkt_clip)

        copy_matrix_to_clipboard("Matrix Title", ["Col1", "Col2"], [["Val1", "Val2"]], mode="word")
        mat_clip = cb.mimeData().html()
        self.assertIn("MsoNormal", mat_clip)
        self.assertIn("mso-padding-alt", mat_clip)

        copy_definition_to_clipboard(
            "<table><tr><td>DefCell</td></tr></table>",
            spec_number="38.413",
            version="19.3.0",
            ie_name="TestIE",
            clause="9.3.1",
            mode="word",
        )
        def_clip = cb.mimeData().html()
        self.assertIn("MsoNormal", def_clip)
        self.assertIn("mso-padding-alt", def_clip)

    def test_expanded_spec_catalog_and_faceted_filtering(self):
        for s_num in ("37.324", "38.323", "38.322", "38.321", "36.323", "36.322", "36.321"):
            self.assertIn(s_num, PROTOCOL_SPEC_CATALOG)
            meta = PROTOCOL_SPEC_CATALOG[s_num]
            self.assertTrue(meta["importable"], s_num)
            self.assertEqual(meta["domain"], "RAN", s_num)
            self.assertEqual(meta["layer"], "L2", s_num)

        for ref_spec in ("38.300", "38.425", "25.323"):
            self.assertIn(ref_spec, PROTOCOL_SPEC_CATALOG)
            meta = PROTOCOL_SPEC_CATALOG[ref_spec]
            self.assertFalse(meta["importable"], ref_spec)
            self.assertTrue(str(meta.get("note", "")).strip(), ref_spec)

        for s_num in ("37.324", "38.323", "38.322", "38.321"):
            self.assertTrue(spec_matches_filters(s_num, domain="RAN", layer="L2", generation="5G"))
        self.assertFalse(spec_matches_filters("36.323", domain="RAN", layer="L2", generation="5G"))
        self.assertFalse(spec_matches_filters("24.501", domain="RAN", layer="L2", generation="5G"))
        self.assertTrue(spec_matches_filters("38.323", plane="CP"))
        self.assertFalse(spec_matches_filters("37.324", plane="CP"))
        self.assertTrue(spec_matches_filters("38.323", text="5g pdcp"))

    def test_l2_pdu_parser_pdcp_sdap_and_rlc(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            tmp = Path(tmpdir)

            # 1. TS 38.323 NR PDCP
            pdcp_doc = tmp / "38323-j10.docx"
            _make_synthetic_l2_docx(
                pdcp_doc,
                paragraphs=[
                    ("Heading1", "6 Protocol data units, formats, and parameters", False),
                    ("Heading2", "6.2 Formats", False),
                    ("Heading4", "6.2.2.1 Data PDU for SRBs\u200e", False),
                    ("Normal", "Figure 6.2.2.1-1 shows the format of the PDCP Data PDU for SRBs with a 12-bit PDCP SN.", True),
                    ("Heading4", "6.2.2.2 Data PDU for DRBs and MRBs with 12 bits PDCP SN", False),
                    ("Normal", "Figure 6.2.2.2-1 shows the format for DRBs with 12-bit PDCP SN.", True),
                    ("Heading2", "6.3 Parameters", False),
                    ("Heading3", "6.3.2 PDCP SN", False),
                    ("Normal", "Length: 12 bits or 18 bits.", False),
                    ("Heading1", "7 Variables, constants, and timers", False),
                ],
                table_rows=[["Field", "Description"], ["PDCP SN", "Sequence number"]],
            )
            disp_pdcp = ProtocolDocxDispatcher(pdcp_doc)
            self.assertEqual(disp_pdcp.extract_spec_number(), "38.323")
            self.assertEqual(disp_pdcp.extract_version_from_filename(), "19.1.0")
            msgs, defs = disp_pdcp.parse()
            msg_names = [m["message_name"] for m in msgs]
            self.assertIn("Data PDU for SRBs", msg_names)
            self.assertIn("Data PDU for DRBs and MRBs with 12 bits PDCP SN", msg_names)
            self.assertFalse(any("\u200e" in m["message_name"] for m in msgs))

            srb_msg = next(m for m in msgs if m["message_name"] == "Data PDU for SRBs")
            self.assertEqual([f["name"] for f in srb_msg["ies"]], ["R", "PDCP SN", "Data", "MAC-I"])
            sn_ie = next(f for f in srb_msg["ies"] if f["name"] == "PDCP SN")
            self.assertIn("Clause 6.3.2", sn_ie["type_reference"])
            self.assertTrue(any("PDU format figure" in d["raw_description"] for d in defs))

            # 2. TS 37.324 SDAP
            sdap_doc = tmp / "37324-j00.docx"
            _make_synthetic_l2_docx(
                sdap_doc,
                paragraphs=[
                    ("Heading1", "6 Protocol data units, formats, and parameters", False),
                    ("Heading4", "6.2.2.2 DL Data PDU with SDAP header", False),
                    ("Normal", "Downlink Data PDU format with RDI, RQI and QFI fields.", True),
                    ("Heading3", "6.3.4 QFI", False),
                    ("Normal", "QoS Flow Identifier, 6 bits.", False),
                ],
            )
            sdap_msgs, _ = ProtocolDocxDispatcher(sdap_doc).parse()
            dl_sdap = next(m for m in sdap_msgs if m["message_name"] == "DL Data PDU with SDAP header")
            self.assertEqual([f["name"] for f in dl_sdap["ies"]], ["RDI", "RQI", "QFI", "Data"])
            qfi_ie = next(f for f in dl_sdap["ies"] if f["name"] == "QFI")
            self.assertIn("Clause 6.3.4", qfi_ie["type_reference"])

            # 3. TS 38.322 NR RLC (6.2.2.x formats + 6.2.3.x parameter definitions) & unknown fallback
            rlc_doc = tmp / "38322-j00.docx"
            _make_synthetic_l2_docx(
                rlc_doc,
                paragraphs=[
                    ("Heading 1", "6 Protocol data units, formats and parameters", False),
                    ("Heading 4", "6.2.2.4 AMD PDU", False),
                    ("Normal", "Acknowledged Mode Data PDU format carrying RLC SDU segments.", True),
                    ("Heading 4", "6.2.2.5 STATUS PDU", False),
                    ("Normal", "RLC STATUS PDU carrying ACK_SN and NACK_SN.", True),
                    ("Heading 4", "6.2.2.99 Vendor Experimental PDU", False),
                    ("Normal", "Custom experimental PDU format description.", False),
                    ("Heading 3", "6.2.3 Parameters", False),
                    ("Heading 4", "6.2.3.3 Sequence Number (SN) field", False),
                    ("Normal", "RLC Sequence Number.", False),
                    ("Heading 4", "6.2.3.6 Data/Control (D/C) field", False),
                    ("Normal", "Indicates whether the RLC PDU is an RLC data PDU or RLC control PDU.", False),
                ],
            )
            rlc_msgs, rlc_defs = ProtocolDocxDispatcher(rlc_doc).parse()
            rlc_msg_names = [m["message_name"] for m in rlc_msgs]
            self.assertNotIn("Sequence Number (SN) field", rlc_msg_names)
            self.assertNotIn("Data/Control (D/C) field", rlc_msg_names)
            self.assertTrue(any(d["clause"] == "6.2.3.6" and d["ie_name"] == "Data/Control (D/C) field" for d in rlc_defs))

            amd = next(m for m in rlc_msgs if "AMD PDU" in m["message_name"])
            self.assertEqual([f["name"] for f in amd["ies"]], ["D/C", "P", "SI", "R", "SN", "SO", "Data"])
            dc_ie = next(f for f in amd["ies"] if f["name"] == "D/C")
            sn_ie = next(f for f in amd["ies"] if f["name"] == "SN")
            self.assertIn("Clause 6.2.3.6", dc_ie["type_reference"])
            self.assertIn("Clause 6.2.3.3", sn_ie["type_reference"])
            status = next(m for m in rlc_msgs if "STATUS PDU" in m["message_name"])
            self.assertIn("ACK_SN", [f["name"] for f in status["ies"]])
            exp = next(m for m in rlc_msgs if "Vendor Experimental PDU" in m["message_name"])
            self.assertTrue(len(exp["ies"]) >= 1, "Unknown L2 format must still emit a non-empty fallback IE row")

            # 4. TS 38.321 NR MAC letter-suffixed clauses (e.g. 6.1.3.12a) preserve clean titles
            mac_doc = tmp / "38321-j30.docx"
            _make_synthetic_l2_docx(
                mac_doc,
                paragraphs=[
                    ("Heading1", "6 Protocol Data Units, formats and parameters", False),
                    ("Heading4", "6.1.3.12a SP CSI-RS Resource Set MAC CE", False),
                    ("Normal", "MAC Control Element for candidate cell activation.", True),
                ],
            )
            mac_msgs, mac_defs = ProtocolDocxDispatcher(mac_doc).parse()
            self.assertEqual(mac_msgs[0]["clause"], "6.1.3.12a")
            self.assertFalse(mac_defs[0]["ie_name"].startswith("a "))

    def test_modern_spec_import_dialog_facets_queue_bundle_and_preselect(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            tmp = Path(tmpdir)
            nas_db = NASDatabase(tmp / "proto.db")
            stub_db = _StubSpecsDB({
                "37.324": [(1, "37.324", "SDAP", "Rel-19", "37324-j00.zip", "19.0.0", "https://3gpp/37324-j00.zip", "2025-09-01")],
                "38.323": [(2, "38.323", "PDCP", "Rel-19", "38323-j10.zip", "19.1.0", "https://3gpp/38323-j10.zip", "2025-09-02")],
                "38.322": [(3, "38.322", "RLC", "Rel-19", "38322-j00.zip", "19.0.0", "https://3gpp/38322-j00.zip", "2025-09-03")],
                "38.321": [(4, "38.321", "MAC", "Rel-19", "38321-j40.zip", "19.4.0", "https://3gpp/38321-j40.zip", "2025-09-04")],
                "24.526": [(5, "24.526", "UE Policy", "Rel-19", "24526-j40.zip", "19.4.0", "https://3gpp/24526-j40.zip", "2025-09-05")],
            })
            dlg = NASVersionSelectDialog(stub_db, nas_db, tmp)

            # Filter RAN + L2: 24.526 hidden, 38.323 visible, reference-only 25.323/38.300 unselectable
            for btn in dlg.facets._groups["domain"].buttons():
                if btn.property("filter_value") == "RAN":
                    btn.click()
                    break
            for btn in dlg.facets._groups["layer"].buttons():
                if btn.property("filter_value") == "L2":
                    btn.click()
                    break
            visible_specs = [str(dlg.spec_list.item(i).data(Qt.UserRole)) for i in range(dlg.spec_list.count())]
            self.assertNotIn("24.526", visible_specs)
            self.assertIn("38.323", visible_specs)
            for i in range(dlg.spec_list.count()):
                it = dlg.spec_list.item(i)
                if str(it.data(Qt.UserRole)) == "25.323":
                    self.assertFalse(bool(it.flags() & Qt.ItemIsSelectable))

            # Select 37.324 and tick its row, then switch to 38.323 and tick its row -> multi-spec queue = 2
            dlg.preselect_spec("37.324", major_version=19)
            dlg.preselect_spec("38.323", major_version=19)
            self.assertEqual(len(dlg._queue), 2)
            self.assertIn("2 releases queued from 2 specifications", dlg.summary_lbl.text())
            dlg._on_accept()
            self.assertEqual({d["spec_number"] for d in dlg.selected_files_info}, {"37.324", "38.323"})

            # Clear & test quick bundle queueing across all 4 RAN L2 specs
            dlg._clear_queue()
            self.assertEqual(len(dlg._queue), 0)
            dlg._queue_bundle(["37.324", "38.323", "38.322", "38.321"])
            self.assertEqual({sp for sp, _v in dlg._queue}, {"37.324", "38.323", "38.322", "38.321"})

            # preselect_spec resets filters so UE Policy 24.526 is reachable directly
            dlg._clear_queue()
            self.assertTrue(dlg.preselect_spec("24.526", major_version=19))
            self.assertEqual(dlg.release_panel.current_spec(), "24.526")
            self.assertIn(("24.526", "19.4.0"), dlg._queue)

            # Quick bundle buttons escape literal '&' as '&&' for Qt mnemonics
            bundle_texts = [btn.text() for btn in dlg.findChildren(QPushButton) if "5G RAN L3" in btn.text()]
            self.assertEqual(len(bundle_texts), 1)
            self.assertIn("&&", bundle_texts[0])

            # Filter chip clicks must emit spec_selected at most once per change (not twice)
            emitted: List[str] = []
            dlg.spec_list.spec_selected.connect(emitted.append)
            for btn in dlg.facets._groups["domain"].buttons():
                if btn.property("filter_value") == "RAN":
                    btn.click()
                    break
            self.assertLessEqual(len(emitted), 1)

    def test_spec_cache_index_matches_archives_and_caches_scans(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            root = Path(tmpdir)
            spec_dir = root / "38.323"
            spec_dir.mkdir(parents=True)
            macd = root / "__MACOSX"
            macd.mkdir()

            zip_exact = spec_dir / "38323-j10.zip"
            zip_exact.write_bytes(b"zip1")
            docx_sub = root / "37324-j00_part1.docx"
            docx_sub.write_bytes(b"doc1")
            (spec_dir / "._38323-j10.zip").write_bytes(b"mac_resource")
            (macd / "38322-j00.zip").write_bytes(b"ignored")

            idx = SpecCacheIndex(candidate_dirs=[root])
            self.assertEqual(idx.find("38323-j10.zip", "38.323"), zip_exact)
            self.assertEqual(idx.find("38323-j10_foo.zip", "38.323"), None)
            self.assertEqual(idx.find("37324-j00.zip", "37.324"), docx_sub)
            self.assertIsNone(idx.find("38322-j00.zip", "38.322"))

            # Subsequent lookups hit the cached directory index rather than re-scanning
            zip_new = spec_dir / "38323-j20.zip"
            zip_new.write_bytes(b"zip2")
            self.assertIsNone(idx.find("38323-j20.zip", "38.323"))
