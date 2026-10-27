import ast
import os
import re
import sys
import unittest
from pathlib import Path

# Ensure 'src' is in sys.path for test execution
project_root = Path(__file__).resolve().parent.parent
src_dir = project_root / "src"
if str(src_dir) not in sys.path:
    sys.path.insert(0, str(src_dir))


class TestUiAmpersandEscaping(unittest.TestCase):
    """
    Automated regression test verifying that all Qt UI text labels
    (buttons, checkboxes, radio buttons, groupboxes, menus, actions, tabs)
    escape literal '&' characters with '&&'.

    In Qt/PyQt, a single '&' is interpreted as an accelerator mnemonic prefix
    (which suppresses the character and renders an underscore or underbars the
    following character on Windows).
    """

    TARGET_CLASSES = {
        "QPushButton",
        "QToolButton",
        "QCommandLinkButton",
        "QCheckBox",
        "QRadioButton",
        "QGroupBox",
        "QAction",
        "QMenu",
        "QMenuBar",
    }

    TARGET_METHODS = {
        "setText",
        "setTitle",
        "addMenu",
        "addAction",
        "setTabText",
        "addTab",
        "insertTab",
    }

    def test_no_unescaped_ui_ampersands_in_src(self):
        self.assertTrue(src_dir.exists(), f"Source directory {src_dir} does not exist.")

        violations = []
        ampersand_pattern = re.compile(r"(?<!&)&(?!&)(?!amp;)(?!lt;)(?!gt;)(?!nbsp;)")

        for root, _, files in os.walk(src_dir):
            for file_name in files:
                if not file_name.endswith(".py"):
                    continue

                py_path = Path(root) / file_name
                try:
                    tree = ast.parse(py_path.read_text(encoding="utf-8"), filename=str(py_path))
                except Exception as e:
                    violations.append(f"{py_path}: Could not parse AST: {e}")
                    continue

                for node in ast.walk(tree):
                    if not isinstance(node, ast.Call):
                        continue

                    # Extract call name
                    func_name = ""
                    if isinstance(node.func, ast.Name):
                        func_name = node.func.id
                    elif isinstance(node.func, ast.Attribute):
                        func_name = node.func.attr

                    if func_name in self.TARGET_CLASSES or func_name in self.TARGET_METHODS:
                        for arg_idx, arg in enumerate(node.args):
                            strings_to_check = []
                            if isinstance(arg, ast.Constant) and isinstance(arg.value, str):
                                strings_to_check.append(arg.value)
                            elif isinstance(arg, ast.JoinedStr):
                                for part in arg.values:
                                    if isinstance(part, ast.Constant) and isinstance(part.value, str):
                                        strings_to_check.append(part.value)

                            for val in strings_to_check:
                                if ampersand_pattern.search(val):
                                    rel_path = py_path.relative_to(project_root)
                                    violations.append(
                                        f"{rel_path}:{node.lineno} {func_name}(arg {arg_idx}): {repr(val)} "
                                        f"contains unescaped '&'. Use '&&' for literal ampersands."
                                    )

        if violations:
            msg = (
                f"Found {len(violations)} unescaped UI ampersand(s) in PyQt widgets/actions.\n"
                "In Qt, a single '&' acts as a keyboard accelerator mnemonic prefix and renders as an underscore.\n"
                "To display a literal '&', use '&&'.\n\nViolations:\n" + "\n".join(violations)
            )
            self.fail(msg)

    def test_escape_qt_mnemonic_helper(self):
        """Verify that escape_qt_mnemonic converts '&' -> '&&' idempotently without double-escaping."""
        from core.ui.ui_components import escape_qt_mnemonic

        # Single ampersands must be doubled
        self.assertEqual(escape_qt_mnemonic("3GPP Work Plan & Release Timelines"), "3GPP Work Plan && Release Timelines")
        self.assertEqual(escape_qt_mnemonic("Documentation & Help"), "Documentation && Help")
        self.assertEqual(escape_qt_mnemonic("&"), "&&")

        # Already doubled ampersands must NOT become quadruple
        self.assertEqual(escape_qt_mnemonic("Already && Escaped"), "Already && Escaped")
        self.assertEqual(escape_qt_mnemonic("A & B && C & D"), "A && B && C && D")

        # Non-ampersand and empty strings
        self.assertEqual(escape_qt_mnemonic("Plain Text"), "Plain Text")
        self.assertEqual(escape_qt_mnemonic(""), "")
        self.assertEqual(escape_qt_mnemonic(None), "")

    def test_useful_links_menu_ampersand_escaping(self):
        """Verify that populate_help_menu escapes all dynamic category and link titles with '&&'."""
        from PyQt5.QtWidgets import QApplication, QMenu
        from core.config.useful_links import populate_help_menu

        app = QApplication.instance()
        if app is None:
            app = QApplication(sys.argv)

        menu = QMenu()
        populate_help_menu(menu, on_open_url=lambda u, t: None)

        ampersand_pattern = re.compile(r"(?<!&)&(?!&)")
        violations = []
        actions = menu.actions()
        self.assertGreater(len(actions), 0, "Useful links menu should have actions.")

        for act in actions:
            text = act.text()
            if not text:
                continue
            if ampersand_pattern.search(text):
                violations.append(f"QAction with text {repr(text)} contains an unescaped single '&'!")

        if violations:
            self.fail(
                "Dynamic actions in useful_links menu contain unescaped '&':\n"
                + "\n".join(violations)
                + "\nAll dynamic titles must be passed through escape_qt_mnemonic()."
            )


if __name__ == "__main__":
    unittest.main()
