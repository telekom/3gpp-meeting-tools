import ast
import os
import re
import unittest
from pathlib import Path


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
        project_root = Path(__file__).resolve().parent.parent
        src_dir = project_root / "src"
        self.assertTrue(src_dir.exists(), f"Source directory {src_dir} does not exist.")

        violations = []
        ampersand_pattern = re.compile(r"(?<!&)&(?!&)(?!amp;)(?!lt;)(?!gt;)")

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
                            if isinstance(arg, ast.Constant) and isinstance(arg.value, str):
                                val = arg.value
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


if __name__ == "__main__":
    unittest.main()
