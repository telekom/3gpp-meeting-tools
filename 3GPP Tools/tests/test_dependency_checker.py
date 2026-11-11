# --- File: tests/test_dependency_checker.py ---
"""
Unit tests for the Startup Dependency Checker in 3GPP Tools.
"""

from __future__ import annotations

import io
import sys
import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch

from core.utils.dependency_checker import (
    DEFAULT_REQUIREMENTS,
    PACKAGE_IMPORT_MAP,
    Requirement,
    check_dependencies,
    check_dependencies_or_exit,
    check_missing_dependencies,
    find_requirements_file,
    format_dependency_report,
    is_module_available,
    load_requirements,
    parse_requirement_line,
    parse_version_tuple,
    should_show_gui_dialog,
)


class TestDependencyChecker(unittest.TestCase):
    """Test suite for core.utils.dependency_checker."""

    def test_parse_version_tuple(self):
        self.assertEqual(parse_version_tuple("5.15.11"), (5, 15, 11))
        self.assertEqual(parse_version_tuple("306"), (306,))
        self.assertEqual(parse_version_tuple("1.2.0-rc1"), (1, 2, 0, 1))
        self.assertEqual(parse_version_tuple(""), (0,))
        self.assertEqual(parse_version_tuple(None), (0,))

        # Verify tuple comparison behavior
        self.assertTrue(parse_version_tuple("5.15.11") >= parse_version_tuple("5.15.0"))
        self.assertTrue(parse_version_tuple("312") >= parse_version_tuple("306"))
        self.assertFalse(parse_version_tuple("1.0.0") >= parse_version_tuple("1.1.0"))

    def test_parse_requirement_line(self):
        # Normal versioned
        req = parse_requirement_line("PyQt5>=5.15.0")
        self.assertIsNotNone(req)
        self.assertEqual(req.dist_name, "PyQt5")
        self.assertEqual(req.module_name, "PyQt5")
        self.assertEqual(req.specifier, ">=5.15.0")
        self.assertEqual(req.min_version, "5.15.0")

        # Mapped module (python-docx -> docx)
        req2 = parse_requirement_line("python-docx>=1.1.0")
        self.assertIsNotNone(req2)
        self.assertEqual(req2.dist_name, "python-docx")
        self.assertEqual(req2.module_name, "docx")
        self.assertEqual(req2.min_version, "1.1.0")

        # Comment and empty lines
        self.assertIsNone(parse_requirement_line("# comment line"))
        self.assertIsNone(parse_requirement_line("   "))
        self.assertIsNone(parse_requirement_line(""))

        # Line with inline comment
        req3 = parse_requirement_line("requests>=2.25.1 # networking")
        self.assertIsNotNone(req3)
        self.assertEqual(req3.dist_name, "requests")
        self.assertEqual(req3.module_name, "requests")
        self.assertEqual(req3.min_version, "2.25.1")

    def test_package_import_map(self):
        # Ensure crucial distinct distribution-to-import mappings are present
        self.assertEqual(PACKAGE_IMPORT_MAP["pywin32"], "win32com")
        self.assertEqual(PACKAGE_IMPORT_MAP["python-docx"], "docx")
        self.assertEqual(PACKAGE_IMPORT_MAP["beautifulsoup4"], "bs4")
        self.assertEqual(PACKAGE_IMPORT_MAP["pillow"], "PIL")
        self.assertEqual(PACKAGE_IMPORT_MAP["pyqt5"], "PyQt5")

    def test_load_requirements_actual_file(self):
        req_file = find_requirements_file()
        self.assertIsNotNone(req_file)
        self.assertTrue(req_file.is_file())

        reqs = load_requirements(req_file)
        self.assertGreaterEqual(len(reqs), 14)

        # Verify key modules exist in parsed requirements
        dists = {r.dist_name.lower() for r in reqs}
        for expected in ["pyqt5", "pywin32", "python-docx", "requests", "beautifulsoup4", "openpyxl", "pandas", "plotly", "networkx", "lxml", "pillow", "pypdf", "pypdfium2", "markdown"]:
            self.assertIn(expected, dists, f"Missing {expected} in loaded requirements")

    def test_load_requirements_fallback(self):
        # Non-existent path falls back to DEFAULT_REQUIREMENTS
        non_existent = Path("/path/to/non_existent_requirements_12345.txt")
        reqs = load_requirements(non_existent)
        self.assertEqual(len(reqs), len(DEFAULT_REQUIREMENTS))
        dists = {r.dist_name for r in reqs}
        self.assertIn("PyQt5", dists)
        self.assertIn("python-docx", dists)

    def test_is_module_available(self):
        # Standard library is always available
        self.assertTrue(is_module_available("sys"))
        self.assertTrue(is_module_available("os"))
        self.assertTrue(is_module_available("pathlib"))

        # Non-existent dummy module
        self.assertFalse(is_module_available("definitely_non_existent_3gpp_dummy_module_999"))

    def test_format_dependency_report(self):
        missing = [
            Requirement("fake-pkg", "fake_mod", ">=1.0.0", "1.0.0"),
            Requirement("pywin32", "win32com", ">=306", "306"),
        ]
        outdated = [
            (Requirement("requests", "requests", ">=2.25.1", "2.25.1"), "2.20.0")
        ]

        report = format_dependency_report(missing, outdated)
        self.assertIn("[3GPP Tools] Missing Required Python Dependencies", report)
        self.assertIn(sys.executable, report)
        self.assertIn("fake-pkg", report)
        self.assertIn("win32com", report)
        self.assertIn("pip install", report)
        self.assertIn("requests", report)

    def test_should_show_gui_dialog(self):
        # In a test runner (unittest in sys.modules), GUI dialog should never show
        self.assertFalse(should_show_gui_dialog(allow_gui=True))
        # When allow_gui is False
        self.assertFalse(should_show_gui_dialog(allow_gui=False))

    @patch("core.utils.dependency_checker.is_module_available")
    def test_check_dependencies_all_present(self, mock_is_available):
        mock_is_available.return_value = True

        # When all packages are available and versions match (using high version 999.0.0 to satisfy pywin32>=306)
        with patch("core.utils.dependency_checker.get_installed_version", return_value="999.0.0"):
            res = check_dependencies(allow_gui=False, interactive=False)
            self.assertTrue(res)

    @patch("core.utils.dependency_checker.is_module_available")
    def test_check_dependencies_missing_package(self, mock_is_available):
        # Simulate missing PyQt5
        def side_effect(mod):
            return mod != "PyQt5"

        mock_is_available.side_effect = side_effect

        stream = io.StringIO()
        res = check_dependencies(allow_gui=False, interactive=False, stream=stream)
        self.assertFalse(res)
        output = stream.getvalue()
        self.assertIn("PyQt5", output)
        self.assertIn("pip install", output)

    @patch("core.utils.dependency_checker.check_dependencies", return_value=False)
    def test_check_dependencies_or_exit(self, mock_check):
        with self.assertRaises(SystemExit) as ctx:
            check_dependencies_or_exit(allow_gui=False, interactive=False, exit_code=42)
        self.assertEqual(ctx.exception.code, 42)


if __name__ == "__main__":
    unittest.main()
