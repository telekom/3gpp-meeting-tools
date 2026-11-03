# --- File: tests/test_incremental_pcr_markdown.py ---
import sys
import unittest
from pathlib import Path

# Ensure src is on sys.path
_src_dir = Path(__file__).resolve().parent.parent / "src"
if str(_src_dir) not in sys.path:
    sys.path.insert(0, str(_src_dir))

from modules.meetings.core.tdoc2md import (
    DocxConverter,
    change_boundary_qualifier,
    BOUNDARY_RE,
    ALL_NEW_RE,
)


class TestIncrementalPcrMarkdown(unittest.TestCase):
    def test_change_boundary_qualifier_variations(self):
        """Verify boundary qualifier extraction handles parenthesized, unparenthesized, and hyphenated forms."""
        # Unparenthesized with ALL TEXT is NEW
        q1 = change_boundary_qualifier("* * * First Change ALL TEXT is NEW* * * *")
        self.assertEqual(q1, "ALL TEXT is NEW")

        # Parenthesized standard form
        q2 = change_boundary_qualifier("* * * First Change (ALL TEXT is NEW) * * * *")
        self.assertEqual(q2, "ALL TEXT is NEW")

        # Hyphenated form
        q3 = change_boundary_qualifier("* * * Next Change - All text is new * * *")
        self.assertEqual(q3, "All text is new")

        # Plain change without qualifier
        q4 = change_boundary_qualifier("* * * First Change * * *")
        self.assertEqual(q4, "")

        # End of changes
        q5 = change_boundary_qualifier("* * * End of Changes * * * *")
        self.assertEqual(q5, "")

        # Regular non-boundary text
        q6 = change_boundary_qualifier("This is normal proposal prose.")
        self.assertIsNone(q6)

    def test_incremental_all_new_conversion(self):
        """Verify that an 'ALL TEXT is NEW' block containing revisions suppresses [ADDED BLOCK]: and attributes baseline."""
        sample_docx = Path(
            r"C:\Users\colomikunoj\.gemini\antigravity\brain\26d0dbb0-a974-4b10-9d1b-c1e5d51b0b4e\.user_uploaded\media_1793604264619_d32fbf8b.docx"
        )
        if not sample_docx.exists():
            self.skipTest("Sample docx not found")

        # Test in markup mode
        conv_markup = DocxConverter(
            sample_docx,
            tdoc_id="S2-261",
            media_dir=Path("temp_test_media"),
            media_rel="media",
            revision_mode="markup",
        )
        res_markup = conv_markup.convert()

        # 1. Verify cover extracted revision of
        self.assertEqual(conv_markup.result.cover.get("revision of"), "S2-2611386")

        # 2. Verify marker recognized and correctly qualified
        self.assertIn("> **[* * * First Change ALL TEXT is NEW* * * *]**", res_markup.markdown)
        self.assertIn(
            "inline markers preserve Word tracked edits. Unmarked text is also within the all-new scope. Document revision of: S2-2611386.",
            res_markup.markdown,
        )

        # 3. Verify [ADDED BLOCK]: is suppressed on unmarked paragraphs in this incremental block
        self.assertNotIn("[ADDED BLOCK]:", res_markup.markdown)

        # 4. Verify inline revisions are retained
        self.assertIn("[INSERTED: Whether an event is treated as data handled by the DF", res_markup.markdown)
        self.assertIn("which event transfer use the data framework is FFS", res_markup.markdown)
        self.assertIn("[DELETED:", res_markup.markdown)

        # 5. Verify list numbering remains clean
        self.assertIn("c. File-based e.g. for bulk non-Real Time data transfer.", res_markup.markdown)
        self.assertIn("6. For Data Transfer within 6GC", res_markup.markdown)

        # Test in accept mode
        conv_accept = DocxConverter(
            sample_docx,
            tdoc_id="S2-261",
            media_dir=Path("temp_test_media"),
            media_rel="media",
            revision_mode="accept",
        )
        res_accept = conv_accept.convert()

        # In accept mode, insertions are merged, deletions dropped, no markup tags
        self.assertNotIn("[INSERTED:", res_accept.markdown)
        self.assertNotIn("[DELETED:", res_accept.markdown)
        self.assertNotIn("[ADDED BLOCK]:", res_accept.markdown)
        self.assertIn("6GC DF supports different modes of data transfer: GET/STREAMING/FILE.", res_accept.markdown)


if __name__ == "__main__":
    unittest.main()
