import re
import sys
import unittest
from pathlib import Path

from PyQt5.QtWidgets import QApplication

from core.ui.theme import (
    HEX_TO_TOKEN,
    ThemeBreakpoints,
    ThemeColors,
    ThemeRadii,
    ThemeSpacing,
    ThemeTypography,
    B,
    C,
    R,
    S,
    T,
    get_bundled_fonts_dir,
    load_app_fonts,
)


HEX_PATTERN = re.compile(r"#[0-9A-Fa-f]{6}\b")


class TestThemeDesignTokens(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication(sys.argv)

    def test_aliases_match_classes(self):
        self.assertIs(C, ThemeColors)
        self.assertIs(T, ThemeTypography)
        self.assertIs(R, ThemeRadii)
        self.assertIs(S, ThemeSpacing)
        self.assertIs(B, ThemeBreakpoints)

    def test_all_color_tokens_are_valid_hex(self):
        color_attrs = [
            name
            for name in dir(ThemeColors)
            if name.isupper() and not name.startswith("_")
        ]
        self.assertGreaterEqual(len(color_attrs), 50)
        for attr in color_attrs:
            val = getattr(ThemeColors, attr)
            self.assertIsInstance(val, str, msg=f"{attr} is not a string")
            self.assertRegex(val, r"^#[0-9A-Fa-f]{6}$", msg=f"{attr}={val!r} invalid hex")

    def test_hex_to_token_references_valid_theme_attributes(self):
        self.assertGreaterEqual(len(HEX_TO_TOKEN), 50)
        for hex_val, token_name in HEX_TO_TOKEN.items():
            self.assertRegex(hex_val, r"^#[0-9A-Fa-f]{6}$")
            self.assertTrue(
                hasattr(ThemeColors, token_name),
                msg=f"HEX_TO_TOKEN[{hex_val!r}] -> unknown token {token_name!r}",
            )
            self.assertEqual(
                getattr(ThemeColors, token_name).upper(),
                hex_val.upper(),
                msg=f"HEX_TO_TOKEN mismatch for {token_name}",
            )

    def test_typography_scale_monotonic(self):
        self.assertEqual(
            [
                T.SIZE_PAGE_TITLE,
                T.SIZE_SECTION_TITLE,
                T.SIZE_BODY,
                T.SIZE_TABLE,
                T.SIZE_CAPTION,
            ],
            [18, 14, 13, 12, 11],
        )
        self.assertIn("Inter", T.FONT_FAMILY_UI)
        self.assertIn("Segoe UI", T.FONT_FAMILY_UI)
        self.assertIn("Cascadia Mono", T.FONT_FAMILY_MONO)

    def test_ui_components_uses_zero_raw_hex_literals(self):
        ui_comp_path = Path(__file__).resolve().parents[1] / "src" / "core" / "ui" / "ui_components.py"
        source = ui_comp_path.read_text(encoding="utf-8")
        # Strip comment-only lines just in case; ensure no raw #RRGGBB in code/QSS
        code_lines = [
            line for line in source.splitlines() if not line.lstrip().startswith("#")
        ]
        matches = HEX_PATTERN.findall("\n".join(code_lines))
        self.assertEqual(
            matches,
            [],
            msg=f"Found unexpected raw hex literals in ui_components.py: {matches}",
        )

    def test_load_app_fonts_safe_without_assets(self):
        fonts_dir = get_bundled_fonts_dir()
        self.assertEqual(fonts_dir.name, "fonts")
        loaded = load_app_fonts(self.app)
        self.assertIsInstance(loaded, list)
        self.assertEqual(self.app.font().pixelSize(), T.SIZE_BODY)


if __name__ == "__main__":
    unittest.main()
