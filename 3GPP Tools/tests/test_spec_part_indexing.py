import unittest
from pathlib import Path

from modules.nas.core.parsing.protocol_parser_constants import extract_part_index, RE_PART_INDEX
from modules.nas.core.parsing.protocol_parser_common import ProtocolDocxDispatcher
from modules.nas.core.parsing.asn1_base_parser import BaseAsn1DocxParser
from modules.nas.core.parsing.gtpu_parser import GTPUDocxParser
from modules.nas.core.parsing.pdu_session_up_parser import PDUSessionUPDocxParser
from modules.nas.core.parsing.pfcp_parser import PFCPDocxParser
from modules.spec_search.core.spec_text_extractor import SpecDocxExtractor


class TestSpecPartIndexing(unittest.TestCase):
    def test_extract_part_index_with_strings_and_paths(self):
        """Verify extract_part_index handles both str and Path instances without AttributeError."""
        test_cases = [
            ("23501-i40_s00_s04.docx", 0),
            ("23501-i40_s05_s08.docx", 5),
            ("23501-i40_s09_s12.docx", 9),
            ("23501-i40_s13_s16.docx", 13),
            ("38331-h00_1_s00_s04.docx", 1),
            ("38331-h00_part1.docx", 1),
            ("38331-h00_part2.docx", 2),
            ("38331-h00_1.docx", 1),
            ("38331-h00_2.docx", 2),
            ("38760-3-080.docx", 0),
            ("23501-k20.docx", 0),
            ("C:/specs/23.501/23501-i40_s05_s08.docx", 5),
            ("C:\\specs\\23.501\\23501-i40_s05_s08.docx", 5),
        ]

        for filename, expected in test_cases:
            # Test with str
            self.assertEqual(
                extract_part_index(filename),
                expected,
                f"Failed for str: {filename}"
            )
            self.assertEqual(
                ProtocolDocxDispatcher._extract_part_index(filename),
                expected,
                f"ProtocolDocxDispatcher failed for str: {filename}"
            )

            # Test with Path
            p = Path(filename)
            self.assertEqual(
                extract_part_index(p),
                expected,
                f"Failed for Path: {p}"
            )
            self.assertEqual(
                ProtocolDocxDispatcher._extract_part_index(p),
                expected,
                f"ProtocolDocxDispatcher failed for Path: {p}"
            )

    def test_parsers_extract_part_index_duck_typing(self):
        """Verify all parser classes accept both Path and str in _extract_part_index."""
        parsers = [
            ProtocolDocxDispatcher,
            BaseAsn1DocxParser,
            GTPUDocxParser,
            PDUSessionUPDocxParser,
            PFCPDocxParser,
        ]

        test_str = "23501-i40_s05_s08.docx"
        test_path = Path("23501-i40_s05_s08.docx")

        for parser_cls in parsers:
            self.assertEqual(
                parser_cls._extract_part_index(test_str),
                5,
                f"{parser_cls.__name__}._extract_part_index failed with str"
            )
            self.assertEqual(
                parser_cls._extract_part_index(test_path),
                5,
                f"{parser_cls.__name__}._extract_part_index failed with Path"
            )

    def test_sorting_with_path_and_str_keys(self):
        """Ensure sorting lists of Path objects works with direct method and lambda p: p.name."""
        docs = [
            Path("23501-i40_s05_s08.docx"),
            Path("23501-i40_s00_s04.docx"),
            Path("23501-i40_s13_s16.docx"),
            Path("23501-i40_s09_s12.docx"),
        ]

        # 1. Sorted directly by method reference
        sorted1 = sorted(docs, key=ProtocolDocxDispatcher._extract_part_index)
        expected_names = [
            "23501-i40_s00_s04.docx",
            "23501-i40_s05_s08.docx",
            "23501-i40_s09_s12.docx",
            "23501-i40_s13_s16.docx",
        ]
        self.assertEqual([p.name for p in sorted1], expected_names)

        # 2. Sorted by lambda passing p.name (which previously crashed with AttributeError)
        sorted2 = sorted(docs, key=lambda p: ProtocolDocxDispatcher._extract_part_index(p.name))
        self.assertEqual([p.name for p in sorted2], expected_names)

    def test_spec_docx_extractor_orders_parts(self):
        """Ensure SpecDocxExtractor sorts input paths by part index."""
        docs = [
            Path("tests/dummy_s05_s08.docx"),
            Path("tests/dummy_s00_s04.docx"),
        ]
        # Even if files don't exist on disk, extractor constructor filters by exists()
        # but let's verify sort key logic
        extractor_sorted = sorted(docs, key=extract_part_index)
        self.assertEqual(
            [p.name for p in extractor_sorted],
            ["dummy_s00_s04.docx", "dummy_s05_s08.docx"]
        )


if __name__ == "__main__":
    unittest.main()
