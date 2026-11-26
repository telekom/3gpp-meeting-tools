"""Optional Ruff static-analysis gate for ``src/`` and ``tests/``.

Runs ``python -m ruff check src tests`` against ``pyproject.toml`` when Ruff is
installed in the active interpreter, and skips gracefully on end-user machines
that only have runtime dependencies from ``requirements.txt``.
"""

from __future__ import annotations

import importlib.util
import subprocess
import sys
import unittest
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent


class TestRuffLint(unittest.TestCase):
    def test_ruff_check_src_and_tests_clean(self):
        if importlib.util.find_spec("ruff") is None:
            self.skipTest("ruff is not installed in this Python environment")

        result = subprocess.run(
            [sys.executable, "-m", "ruff", "check", "src", "tests"],
            cwd=PROJECT_ROOT,
            capture_output=True,
            text=True,
        )
        self.assertEqual(
            result.returncode,
            0,
            f"ruff check failed:\n{result.stdout}\n{result.stderr}",
        )


if __name__ == "__main__":
    unittest.main()
