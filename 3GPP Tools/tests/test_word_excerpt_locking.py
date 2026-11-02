import io
import os
import shutil
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest.mock import MagicMock, patch

from modules.word_tools.core.styles_3gpp import ensure_3gpp_package_styles
from modules.word_tools.core.word_excerpt_extractor import WordExcerptEngine, HeadingInfo


class TestWordExcerptLocking(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp(prefix="test_excerpt_locking_")

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def _create_minimal_docx(self, path: Path) -> Path:
        styles_xml = b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:style w:type="paragraph" w:styleId="Normal">
                <w:name w:val="Normal"/>
            </w:style>
        </w:styles>"""
        with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as z:
            z.writestr("word/styles.xml", styles_xml)
            z.writestr("[Content_Types].xml", b"<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'/>")
        return path

    def test_ensure_3gpp_package_styles_retries_on_permission_error(self):
        """Simulate transient Windows file lock (PermissionError) on os.replace during style patching."""
        docx_path = Path(self.temp_dir) / "locked_doc.docx"
        self._create_minimal_docx(docx_path)

        real_replace = os.replace
        call_count = {"count": 0}

        def mock_replace(src, dst):
            call_count["count"] += 1
            if call_count["count"] <= 2:
                raise PermissionError(5, "Zugriff verweigert")
            return real_replace(src, dst)

        with patch("modules.word_tools.core.styles_3gpp.time.sleep", return_value=None), \
             patch("os.replace", side_effect=mock_replace):
            ensure_3gpp_package_styles(docx_path)

        self.assertGreaterEqual(call_count["count"], 3, "os.replace must have been retried on PermissionError")
        self.assertTrue(docx_path.exists())

    def test_ensure_3gpp_package_styles_copyfile_fallback(self):
        """Simulate persistent os.replace failure falling back to copyfile."""
        docx_path = Path(self.temp_dir) / "persistent_locked_doc.docx"
        self._create_minimal_docx(docx_path)

        def mock_replace_always_fail(src, dst):
            raise PermissionError(5, "Zugriff verweigert (Simulated Win32 MoveFileEx lock)")

        with patch("modules.word_tools.core.styles_3gpp.time.sleep", return_value=None), \
             patch("os.replace", side_effect=mock_replace_always_fail):
            ensure_3gpp_package_styles(docx_path)

        self.assertTrue(docx_path.exists())
        # Verify valid zip package after copyfile fallback
        with zipfile.ZipFile(docx_path, "r") as z:
            self.assertIn("word/styles.xml", z.namelist())

    def test_single_fragment_isolation_in_temp_dir(self):
        """Verify single-document excerpt applies sensitivity label and styles in temp work directory before destination."""
        source_doc = Path(self.temp_dir) / "source_spec.docx"
        self._create_minimal_docx(source_doc)

        dest_doc = Path(self.temp_dir) / "final_excerpt.docx"

        headings = [
            HeadingInfo(
                id="0:0",
                source_path=str(source_doc),
                source_index=0,
                block_index=0,
                end_index=5,
                level=1,
                text="1 Scope",
            )
        ]

        engine = WordExcerptEngine([str(source_doc)], headings)

        applied_label_targets = []
        ensured_styles_targets = []

        def fake_apply_label(p):
            applied_label_targets.append(Path(p).resolve())

        def fake_ensure_styles(p):
            ensured_styles_targets.append(Path(p).resolve())

        with patch.object(engine, "_make_fragment", side_effect=lambda src, rng, dst: shutil.copy2(src, dst)), \
             patch.object(engine, "_apply_label_in_fresh_word", side_effect=fake_apply_label), \
             patch("modules.word_tools.core.word_excerpt_extractor.ensure_3gpp_package_styles", side_effect=fake_ensure_styles), \
             patch("pythoncom.CoInitialize"), \
             patch("pythoncom.CoUninitialize"):
            result = engine.extract(["0:0"], str(dest_doc))

        self.assertEqual(result.resolve(), dest_doc.resolve())
        self.assertTrue(dest_doc.exists())

        # Crucial check: labeling and package style patching must have occurred on the
        # temporary working fragment inside temp dir, NOT directly on dest_doc!
        self.assertEqual(len(applied_label_targets), 1)
        self.assertEqual(len(ensured_styles_targets), 1)
        self.assertNotEqual(
            applied_label_targets[0],
            dest_doc.resolve(),
            "Sensitivity label MUST be applied in the isolated temp directory to avoid OneDrive/shell locks"
        )
        self.assertNotEqual(
            ensured_styles_targets[0],
            dest_doc.resolve(),
            "Package styles MUST be ensured in the isolated temp directory to avoid OneDrive/shell locks"
        )


if __name__ == "__main__":
    unittest.main()
