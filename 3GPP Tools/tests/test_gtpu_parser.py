"""Tests for TS 29.281 (GTPv1-U) parser package, TS 38.415 cross-spec unrolling, and DB wiring."""

from __future__ import annotations

import tempfile
import unittest
import zipfile
from pathlib import Path

try:
    from lxml import etree as ET
except ImportError:
    import xml.etree.ElementTree as ET

from modules.nas.core.nas_db import NASDatabase
from modules.nas.core.parsing.gtpu_parser import GTPUDocxParser
from modules.nas.core.parsing.protocol_parser_constants import W_NS

LIVE_TS_29281_DOCX = Path(r"C:\Users\colomikunoj\3GPP_Delegate_Helper\specs\29.281\29281-j20.docx")
W_URL = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _add_p(body: ET.Element, text: str, style: str = "") -> None:
    p = ET.SubElement(body, f"{{{W_URL}}}p")
    if style:
        p_pr = ET.SubElement(p, f"{{{W_URL}}}pPr")
        st = ET.SubElement(p_pr, f"{{{W_URL}}}pStyle")
        st.set(f"{{{W_URL}}}val", style)
    r = ET.SubElement(p, f"{{{W_URL}}}r")
    t = ET.SubElement(r, f"{{{W_URL}}}t")
    t.text = text


def _add_table(body: ET.Element, rows_data: list[list[str]]) -> None:
    tbl = ET.SubElement(body, f"{{{W_URL}}}tbl")
    for row in rows_data:
        tr = ET.SubElement(tbl, f"{{{W_URL}}}tr")
        for cell_txt in row:
            tc = ET.SubElement(tr, f"{{{W_URL}}}tc")
            p = ET.SubElement(tc, f"{{{W_URL}}}p")
            r = ET.SubElement(p, f"{{{W_URL}}}r")
            t = ET.SubElement(r, f"{{{W_URL}}}t")
            t.text = cell_txt


def _build_synthetic_gtpu_docx(path: Path) -> None:
    doc = ET.Element(f"{{{W_URL}}}document", nsmap={"w": W_URL} if hasattr(ET, "register_namespace") or "lxml" in ET.__name__ else {})
    body = ET.SubElement(doc, f"{{{W_URL}}}body")

    # Simulated TOC entries with trailing page numbers (must be skipped by 1 Scope guard)
    _add_p(body, "Contents", "TOCHeading")
    _add_p(body, "5.2.2 Extension Header types 19", "TOC3")
    _add_p(body, "8.5 Extension Header Type List 30", "TOC2")

    # Specification body starts at 1 Scope
    _add_p(body, "1 Scope", "Heading1")
    _add_p(body, "This document specifies the GTP-U protocol.", "Normal")

    _add_p(body, "5 GTP-U header", "Heading1")
    _add_p(body, "5.1 General format", "Heading2")
    _add_p(body, "The GTP-U header is a variable length header whose minimum length is 8 bytes.", "Normal")
    _add_table(
        body,
        [
            ["Octets", "Bits 8..1"],
            ["1", "Version | PT | (*) | E | S | PN"],
            ["2", "Message Type"],
            ["3-4", "Length"],
            ["5-8", "Tunnel Endpoint Identifier (TEID)"],
        ],
    )
    _add_p(body, "NOTE 0: (*) This bit is a spare bit.", "NO")
    _add_p(body, "Figure 5.1-1: Outline of the GTP-U Header", "TF")

    _add_p(body, "5.2 GTP-U Extension Header", "Heading2")
    _add_p(body, "5.2.1 General format of the GTP-U Extension Header", "Heading3")
    _add_p(body, "The Next Extension Header Type field defines the chained extension header.", "Normal")
    _add_table(
        body,
        [
            ["Next Extension Header Field Value", "Type of Extension Header"],
            ["0000 0000", "No more extension headers"],
            ["0000 0011", "Long PDCP PDU Number.See NOTE 2."],
            ["0000 0100", "PDU Set Information Container. See NOTE 5, NOTE 6, NOTE 7."],
            ["0100 0000", "UDP Port. Provides the UDP Source Port of the triggering message."],
            ["1000 0010", "Long PDCP PDU Number.See NOTE 3."],
            ["1000 0101", "PDU Session Container. See NOTE 4."],
            ["1000 0110", "PDU Set Information Container.See NOTE 5, NOTE 6, NOTE 8."],
            ["1100 0000", "PDCP PDU Number [4]-[5]. See NOTE 1."],
        ],
    )
    _add_p(body, "Figure 5.2.1-3: Definition of Extension Header Type", "TF")

    _add_p(body, "5.2.2 Extension Header types", "Heading3")
    _add_p(body, "Extension header types applicable to the GTP user plane.", "Normal")

    _add_p(body, "5.2.2.7 PDU Session Container", "Heading4")
    _add_p(body, "This extension header shall be transmitted in G-PDUs over N3 and N9.", "Normal")
    _add_table(body, [["Octets", "Bits 8..1"], ["1", "0xn"], ["2-(4n-1)", "PDU Session Container"]])

    _add_p(body, "7.3 Tunnel Management Messages", "Heading2")
    _add_p(body, "7.3.1 Error Indication", "Heading3")
    _add_p(body, "Table 7.3.1-1: Information Elements in an Error Indication", "TH")
    _add_table(
        body,
        [
            ["Information element", "Presence requirement", "Reference"],
            ["Tunnel Endpoint Identifier Data I", "Mandatory", "8.3"],
            ["GTP-U Peer Address", "Mandatory", "8.4"],
        ],
    )
    _add_p(body, "7.3.2 End Marker", "Heading3")
    _add_p(body, "7.3.2.1 General", "Heading4")
    _add_p(body, "Table 7.3.2.1-1: Information Elements in End Marker message", "TH")
    _add_table(
        body,
        [
            ["Information element", "Presence requirement", "Reference"],
            ["Private Extension", "Optional", "8.6"],
        ],
    )

    _add_p(body, "8 Information Elements", "Heading1")
    _add_p(body, "8.1 Information Element Types", "Heading2")
    _add_p(body, "Table 8.1-1: Information Elements", "TH")
    _add_table(
        body,
        [
            ["IE Type Value", "Format", "Information Element", "Reference"],
            ["16", "TV", "Tunnel Endpoint Identifier Data I", "8.3"],
            ["133", "TLV", "GSN Address. See NOTE 1.", "8.4"],
            ["141", "TLV", "Extension Header Type List", "8.5"],
        ],
    )
    # Clause 8.5 has no XML table (WMF image in real spec) -> fallback diagram must be injected
    _add_p(body, "8.5 Extension Header Type List", "Heading2")
    _add_p(body, "This information element contains a list of 'n' Extension Header Types.", "Normal")
    _add_p(body, "Figure 8.5-1: Extension Header Type List Information Element", "TF")

    _add_p(body, "8.8 Recovery Time Stamp", "Heading2")
    _add_p(body, "The Recovery Time Stamp IE indicates UTC time.", "Normal")

    # Clause 9 and Annex B must NOT bleed into Clause 8.8
    _add_p(body, "9 Error Handling", "Heading1")
    _add_p(body, "CLAUSE_9_SENTINEL_MUST_NOT_APPEAR_IN_8_8", "Normal")
    _add_p(body, "Annex B (informative): Change history", "Heading8")
    _add_p(body, "ANNEX_B_SENTINEL_MUST_NOT_APPEAR_IN_8_8", "Normal")

    xml_bytes = ET.tostring(doc, encoding="utf-8", xml_declaration=True)
    with zipfile.ZipFile(path, "w") as zf:
        zf.writestr("word/document.xml", xml_bytes)


class TestGTPUParser(unittest.TestCase):
    def test_synthetic_gtpu_parsing_structure_and_fallbacks(self):
        with tempfile.TemporaryDirectory() as tmp:
            docx_path = Path(tmp) / "29281-j20.docx"
            _build_synthetic_gtpu_docx(docx_path)

            parser = GTPUDocxParser([docx_path], spec_number="29.281")
            messages, ie_defs = parser.parse()

            ie_by_name = {d["ie_name"]: d for d in ie_defs}
            # 1. TOC ghost entries eliminated
            self.assertNotIn("Extension Header types 19", ie_by_name)
            self.assertNotIn("Extension Header Type List 30", ie_by_name)

            # 2. Clauses 5.1 & 5.2.1 captured cleanly
            self.assertIn("GTP-U Header", ie_by_name)
            self.assertEqual(ie_by_name["GTP-U Header"]["clause"], "5.1")
            self.assertIn("GTP-U Extension Header Format", ie_by_name)
            self.assertEqual(ie_by_name["GTP-U Extension Header Format"]["clause"], "5.2.1")

            # 3. 5.2.2 Extension Header types includes Figure 5.2.1-3 summary table
            self.assertIn("Extension Header types", ie_by_name)
            self.assertIn("Figure 5.2.1-3", ie_by_name["Extension Header types"]["raw_description"])
            self.assertIn("0000 0100", ie_by_name["Extension Header types"]["raw_description"])

            # 4. Clause 8.5 WMF fallback diagram injected
            self.assertIn("Extension Header Type List", ie_by_name)
            self.assertIn("Type = 141 (Decimal)", ie_by_name["Extension Header Type List"]["raw_description"])

            # 5. Clause 8.8 stopped cleanly before Clause 9 & Annex B
            rt_html = ie_by_name["Recovery Time Stamp"]["raw_description"]
            self.assertNotIn("CLAUSE_9_SENTINEL_MUST_NOT_APPEAR_IN_8_8", rt_html)
            self.assertNotIn("ANNEX_B_SENTINEL_MUST_NOT_APPEAR_IN_8_8", rt_html)

            # 6. Messages include GTP-U Header (Common), G-PDU, Error Indication, End Marker
            msg_by_name = {m["message_name"]: m for m in messages}
            self.assertIn("GTP-U Header (Common)", msg_by_name)
            self.assertEqual(len(msg_by_name["GTP-U Header (Common)"]["ies"]), 12)

            gpdu = msg_by_name["G-PDU"]
            gpdu_by_path = {ie["field_path"]: ie for ie in gpdu["ies"]}
            self.assertIn("GTP-U Header", gpdu_by_path)
            self.assertIn("GTP-U Header > Version", gpdu_by_path)
            self.assertIn("GTP-U Extension Headers", gpdu_by_path)
            self.assertIn("T-PDU Payload", gpdu_by_path)

            # Live Figure 5.2.1-3 dual codepoints & fixed PDCP PDU Number applicability
            self.assertEqual(
                gpdu_by_path["GTP-U Extension Headers > Long PDCP PDU Number"]["iei"],
                "0x03 / 0x82",
            )
            self.assertEqual(
                gpdu_by_path["GTP-U Extension Headers > PDU Set Information Container"]["iei"],
                "0x04 / 0x86",
            )
            self.assertEqual(
                gpdu_by_path["GTP-U Extension Headers > PDCP PDU Number"]["applicability"],
                "S1-U, X2, N3, Xn, Iu",
            )
            self.assertTrue(
                gpdu_by_path["GTP-U Extension Headers > PDU Session Container"]["field_description"]
            )

            # TS 38.415 cross-spec unrolling under G-PDU
            dl_sess = gpdu_by_path[
                "GTP-U Extension Headers > PDU Session Container > DL PDU SESSION INFORMATION (PDU Type 0)"
            ]
            self.assertEqual(dl_sess["depth"], 2)
            self.assertEqual(dl_sess["type_reference"], "TS 38.415 5.5.2.1")

            ul_sess = gpdu_by_path[
                "GTP-U Extension Headers > PDU Session Container > UL PDU SESSION INFORMATION (PDU Type 1)"
            ]
            self.assertEqual(ul_sess["type_reference"], "TS 38.415 5.5.2.2")

            dl_set = gpdu_by_path[
                "GTP-U Extension Headers > PDU Set Information Container > DL PDU SET INFORMATION (PDU Type 0)"
            ]
            self.assertEqual(dl_set["type_reference"], "TS 38.415 6.5.2.1")

            # Signalling messages carry applicable extension headers
            err_ies = {ie["information_element"]: ie for ie in msg_by_name["Error Indication"]["ies"]}
            self.assertIn("UDP Port", err_ies)
            self.assertEqual(err_ies["UDP Port"]["iei"], "0x40")

            end_ies = {ie["field_path"]: ie for ie in msg_by_name["End Marker"]["ies"]}
            self.assertIn("PDU Session Container", end_ies)
            self.assertIn(
                "PDU Session Container > DL PDU SESSION INFORMATION (PDU Type 0)",
                end_ies,
            )

    def test_db_roundtrip_field_description_and_ts38415_lookup(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_dir = Path(tmp)
            docx_path = tmp_dir / "29281-j20.docx"
            _build_synthetic_gtpu_docx(docx_path)
            db = NASDatabase(str(tmp_dir / "test_proto.db"))

            msgs, ies = GTPUDocxParser([docx_path], spec_number="29.281").parse()
            self.assertTrue(db.insert_parsed_spec("29.281", "19.2.0", msgs, ies))

            # Insert mock TS 38.415 target definition matching real DB naming
            self.assertTrue(
                db.insert_parsed_spec(
                    "38.415",
                    "19.1.0",
                    [],
                    [
                        {
                            "clause": "5.5.2.1 (Frame)",
                            "ie_name": "DL PDU SESSION INFORMATION (PDU Type 0)",
                            "raw_description": "<h2>DL PDU SESSION INFORMATION</h2>",
                            "structure_table": "[]",
                        }
                    ],
                )
            )

            ver_rows = db.get_imported_versions()
            v29281 = next(r["id"] for r in ver_rows if r["spec_number"] == "29.281")

            for include_desc in (True, False):
                df = db.get_message_evolution_df("G-PDU", [v29281], include_descriptions=include_desc)
                self.assertIn("field_description", df.columns)
                row = df[df["ie_name"] == "PDU Session Container"].iloc[0]
                self.assertIn("TS 38.415 Clause 5", row["field_description"])

            # Cross-spec resolution used by nas_matrix_controller RE_EXPLICIT_SPEC_REF
            resolved = db.get_ie_definitions_by_clause(
                clause="5.5.2.1",
                alt_name="DL PDU SESSION INFORMATION (PDU Type 0)",
                spec_number="38.415",
                version_ids=None,
            )
            self.assertEqual(len(resolved), 1)
            self.assertEqual(resolved[0]["clause"], "5.5.2.1 (Frame)")

    @unittest.skipUnless(LIVE_TS_29281_DOCX.exists(), "Local TS 29.281 v19.2.0 DOCX not present")
    def test_live_ts29281_j20_docx_smoke(self):
        msgs, ies = GTPUDocxParser([LIVE_TS_29281_DOCX], spec_number="29.281").parse()
        ie_by_name = {d["ie_name"]: d for d in ies}
        msg_by_name = {m["message_name"]: m for m in msgs}

        # No TOC ghost entries with trailing page numbers
        for name in ie_by_name:
            self.assertFalse(
                any(name.endswith(f" {n}") for n in ("19", "20", "21", "22", "23", "24", "29", "30", "31")),
                f"Unexpected TOC ghost IE: {name}",
            )

        # Clause 8.8 no longer swallows Clauses 9-13 & Annex B
        self.assertLess(len(ie_by_name["Recovery Time Stamp"]["raw_description"]), 10000)

        # Clause 8 WMF fallbacks active
        for ie_name, token in [
            ("Recovery", "Type = 14 (Decimal)"),
            ("GTP-U Peer Address", "Type = 133 (Decimal)"),
            ("Extension Header Type List", "Type = 141 (Decimal)"),
            ("Private Extension", "Type = 255 (Decimal)"),
        ]:
            self.assertIn(token, ie_by_name[ie_name]["raw_description"])

        self.assertIn("GTP-U Header (Common)", msg_by_name)
        self.assertIn("G-PDU", msg_by_name)
        gpdu_paths = {ie["field_path"]: ie for ie in msg_by_name["G-PDU"]["ies"]}
        self.assertEqual(
            gpdu_paths["GTP-U Extension Headers > PDU Set Information Container"]["iei"],
            "0x04 / 0x86",
        )
        self.assertEqual(
            gpdu_paths["GTP-U Extension Headers > Long PDCP PDU Number"]["iei"],
            "0x03 / 0x82",
        )


if __name__ == "__main__":
    unittest.main()
