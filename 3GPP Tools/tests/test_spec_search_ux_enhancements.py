"""Unit tests for Spec Search UX enhancements:
- On-demand PDF conversion prompt & multi-part PDF matching
- Full <-> Minimal Scope Mode + 1-click Spec Solo + config persistence
- Clause-only search (no keyword) & normalized Modified badges
- Clickable in-text clause cross-references + Back/Forward inspector history
"""

from __future__ import annotations

import json
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import pandas as pd
from PyQt5.QtCore import QObject, Qt, QUrl, pyqtSignal
from PyQt5.QtWidgets import QApplication

from modules.spec_search.core.spec_doc_resolver import SpecDocResolver
from modules.spec_search.core.spec_search_db import SpecSearchDatabase
from modules.spec_search.ui.components.clause_inspector_render import (
    _highlight_html_text_nodes,
    linkify_clause_references,
)
from modules.spec_search.ui.spec_search_components import (
    ActiveSpecsScopeBar,
    SpecClauseInspector,
)
from modules.spec_search.ui.spec_search_models import SpecEvolutionMatrixModel
from modules.spec_search.ui.spec_search_tabs import SpecSearchTab
from modules.specifications.core.database import SpecsDatabase

app = QApplication.instance() or QApplication(sys.argv)


class _FakeWordConverterThread(QObject):
    """Synchronous stub replacing WordConverterThread during unit tests."""

    ui_log_msg = pyqtSignal(str, int)
    finished_path = pyqtSignal(str)
    conversion_failed = pyqtSignal(str)
    finished = pyqtSignal()

    def __init__(self, doc_source: str, target_format: str, **_kwargs):
        super().__init__()
        self.doc_source = Path(doc_source)
        self.target_format = target_format

    def requestInterruption(self) -> None:
        pass

    def start(self) -> None:
        out_pdf = self.doc_source.with_suffix(".pdf")
        out_pdf.write_bytes(b"%PDF-1.4\n%fake\n")
        self.finished_path.emit(str(out_pdf))
        self.finished.emit()


class TestSpecSearchUXEnhancements(unittest.TestCase):
    def test_pdf_matching_and_on_demand_conversion_prompt(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            root = Path(tmp_dir)
            spec_dir = root / "23.501"
            spec_dir.mkdir(parents=True)

            part1 = spec_dir / "23501-k30_s00_s04.docx"
            part2 = spec_dir / "23501-k30_s05_s08.docx"
            part1.write_bytes(b"docx1")
            part2.write_bytes(b"docx2")

            part1_pdf = spec_dir / "23501-k30_s00_s04.pdf"
            part1_pdf.write_bytes(b"%PDF-1.4")

            info = SpecDocResolver.resolve_spec_files("23.501", "20.3.0", cache_dir=root)
            self.assertEqual(SpecDocResolver.find_matching_pdf(info, "4.2"), part1_pdf)
            self.assertIsNone(SpecDocResolver.find_matching_pdf(info, "5.15.19"))

            opened: list[Path] = []
            with patch(
                "modules.spec_search.core.spec_pdf_converter.open_document",
                side_effect=lambda p: opened.append(Path(p)),
            ):
                ok = SpecDocResolver.open_spec(
                    "23.501", "20.3.0", clause_number="4.2", doc_format="pdf", cache_dir=root
                )
                self.assertTrue(ok)
                self.assertEqual(opened, [part1_pdf])
                self.assertEqual(QApplication.clipboard().text(), "4.2")

            # Missing PDF for part2 -> test Cancel, Open Word Instead, and Convert & Open PDF
            class _FakeMsgBox:
                Question = 1
                AcceptRole = 0
                ActionRole = 1
                Cancel = 2
                choice = "cancel"
                created_buttons: list[str] = []

                def __init__(self, *_a, **_kw):
                    self._buttons: dict[str, object] = {}
                    self._clicked = None

                def setIcon(self, *_a):
                    pass

                def setWindowTitle(self, *_a):
                    pass

                def setText(self, *_a):
                    pass

                def setInformativeText(self, *_a):
                    pass

                def addButton(self, label_or_std, _role=None):
                    btn = object()
                    key = str(label_or_std)
                    self._buttons[key] = btn
                    _FakeMsgBox.created_buttons.append(key)
                    return btn

                def setDefaultButton(self, *_a):
                    pass

                def exec_(self):
                    if _FakeMsgBox.choice == "convert":
                        self._clicked = self._buttons.get("Convert && Open PDF")
                    elif _FakeMsgBox.choice == "word":
                        self._clicked = self._buttons.get("Open Word Instead")
                    else:
                        self._clicked = self._buttons.get(str(_FakeMsgBox.Cancel))

                def clickedButton(self):
                    return self._clicked

            opened.clear()
            with (
                patch("modules.spec_search.core.spec_pdf_converter.QMessageBox", _FakeMsgBox),
                patch("modules.spec_search.core.spec_doc_resolver.open_document", side_effect=lambda p: opened.append(Path(p))),
                patch("modules.spec_search.core.spec_pdf_converter.open_document", side_effect=lambda p: opened.append(Path(p))),
                patch("modules.word_tools.core.word_converter.WordConverterThread", _FakeWordConverterThread),
            ):
                _FakeMsgBox.choice = "cancel"
                self.assertFalse(
                    SpecDocResolver.open_spec("23.501", "20.3.0", "5.15.19", "pdf", cache_dir=root)
                )
                self.assertEqual(opened, [])
                self.assertIn("Convert && Open PDF", _FakeMsgBox.created_buttons)

                _FakeMsgBox.choice = "word"
                self.assertTrue(
                    SpecDocResolver.open_spec("23.501", "20.3.0", "5.15.19", "pdf", cache_dir=root)
                )
                self.assertEqual(opened, [part2])

                opened.clear()
                _FakeMsgBox.choice = "convert"
                self.assertTrue(
                    SpecDocResolver.open_spec("23.501", "20.3.0", "5.15.19", "pdf", cache_dir=root)
                )
                expected_part2_pdf = spec_dir / "23501-k30_s05_s08.pdf"
                self.assertTrue(expected_part2_pdf.exists())
                self.assertEqual(opened, [expected_part2_pdf])
                self.assertEqual(QApplication.clipboard().text(), "5.15.19")

    def test_clause_only_search_and_modified_badge_detection(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            db = SpecSearchDatabase(Path(tmp_dir) / "search.db")
            db.insert_parsed_spec(
                "23.501",
                "18.2.0",
                [
                    {"clause_number": "5.15", "clause_title": "Network Slicing", "content": ""},
                    {"clause_number": "5.15.1", "clause_title": "General", "content": "Base   slicing\nrules."},
                    {"clause_number": "6.1", "clause_title": "AMF", "content": "Access and Mobility."},
                ],
                release_date="2024-03-01",
            )
            db.insert_parsed_spec(
                "23.501",
                "18.3.0",
                [
                    {"clause_number": "5.15", "clause_title": "Network Slicing", "content": ""},
                    # Only whitespace difference vs 18.2.0 -> should stay "✓ Present"
                    {"clause_number": "5.15.1", "clause_title": "General", "content": "Base slicing rules."},
                    {"clause_number": "6.1", "clause_title": "AMF", "content": "Access and Mobility."},
                ],
                release_date="2024-06-01",
            )
            db.insert_parsed_spec(
                "23.501",
                "18.4.0",
                [
                    {"clause_number": "5.15", "clause_title": "Network Slicing", "content": ""},
                    {"clause_number": "5.15.1", "clause_title": "General", "content": "Updated slicing rules with replacement."},
                    {"clause_number": "5.15.19", "clause_title": "Network Slice Replacement", "content": "Support of slice replacement."},
                    {"clause_number": "6.1", "clause_title": "AMF", "content": "Access and Mobility."},
                ],
                release_date="2024-09-01",
            )

            ver_ids = [v["id"] for v in db.get_imported_versions()]
            df = db.search_substring("", ver_ids, clause_filter="clause 5.15")
            self.assertFalse(df.empty)
            clauses = sorted(df["clause_number"].unique().tolist())
            self.assertEqual(clauses, ["5.15", "5.15.1", "5.15.19"])

            model = SpecEvolutionMatrixModel(raw_df=df)
            self.assertEqual(model.rowCount(), 3)
            # Row 0: 5.15 (empty body), Row 1: 5.15.1, Row 2: 5.15.19
            self.assertEqual(model.data(model.index(1, 2), Qt.DisplayRole), "🟢 Added")
            self.assertEqual(model.data(model.index(1, 3), Qt.DisplayRole), "✓ Present")
            self.assertEqual(model.data(model.index(1, 4), Qt.DisplayRole), "✏️ Modified")
            self.assertEqual(model.data(model.index(2, 2), Qt.DisplayRole), "—")
            self.assertEqual(model.data(model.index(2, 4), Qt.DisplayRole), "🟢 Added")
            self.assertIn("MODIFIED", model.data(model.index(1, 4), Qt.ToolTipRole))

        # Backward compatibility when synthetic DataFrame lacks content_hash
        compat_df = pd.DataFrame(
            [
                {
                    "spec_number": "23.501",
                    "version": "18.2.0",
                    "release_date": "2024-03-01",
                    "clause_number": "5.1",
                    "clause_title": "Overview",
                    "snippet_text": "Overview",
                    "clause_pk": 1,
                    "order_index": 1,
                },
                {
                    "spec_number": "23.501",
                    "version": "18.3.0",
                    "release_date": "2024-06-01",
                    "clause_number": "5.1",
                    "clause_title": "Overview",
                    "snippet_text": "Overview",
                    "clause_pk": 2,
                    "order_index": 1,
                },
            ]
        )
        compat_model = SpecEvolutionMatrixModel(raw_df=compat_df)
        self.assertEqual(compat_model.data(compat_model.index(0, 3), Qt.DisplayRole), "✓ Present")

    def test_scope_bar_minimal_mode_solo_and_persistence(self) -> None:
        # Overflow badge rendering test on ActiveSpecsScopeBar
        bar = ActiveSpecsScopeBar()
        many_specs = {f"23.{500 + i}": ["18.0.0"] for i in range(14)}
        bar.update_scope(many_specs, total_indexed_specs=14)
        self.assertIsNotNone(bar.btn_overflow)
        self.assertEqual(bar.btn_overflow.text(), "+2 more")
        self.assertEqual(bar.btn_expand_specs.text(), "📚 Specs (14/14) ▾")

        with tempfile.TemporaryDirectory() as tmp_dir:
            root = Path(tmp_dir)
            search_db_path = root / "search.db"
            specs_db_path = root / "specs.db"

            s_db = SpecSearchDatabase(search_db_path)
            s_db.insert_parsed_spec(
                "23.501",
                "18.4.0",
                [{"clause_number": "5.15", "clause_title": "Slicing", "content": "Slice text"}],
                release_date="2024-09-01",
            )
            s_db.insert_parsed_spec(
                "23.256",
                "19.5.0",
                [{"clause_number": "5.1", "clause_title": "UAS", "content": "UAS text"}],
                release_date="2025-12-01",
            )
            sp_db = SpecsDatabase(specs_db_path)
            with sp_db._get_connection() as conn:
                conn.execute("INSERT INTO specifications (id, number, title) VALUES (1, '23.501', '5GS Architecture')")
                conn.execute("INSERT INTO specifications (id, number, title) VALUES (2, '23.256', 'UAS Support')")
                conn.commit()

            with patch(
                "modules.spec_search.ui.spec_search_tab.tab.get_project_root",
                return_value=root,
            ):
                (root / "config").mkdir(exist_ok=True)
                tab = SpecSearchTab(search_db_path, specs_db_path)
                self.assertFalse(tab._specs_minimal_mode)
                self.assertTrue(tab.scope_bar.isHidden())

                tab.set_specs_minimal_mode(True)
                self.assertTrue(tab._specs_minimal_mode)
                self.assertFalse(tab.scope_bar.isHidden())
                self.assertTrue(tab.left_specs_panel.isHidden())
                self.assertEqual(tab.scope_bar.btn_expand_specs.text(), "📚 Specs (2/2) ▾")

                saved = json.loads((root / "config" / "spec_search_config.json").read_text(encoding="utf-8"))
                self.assertTrue(saved.get("specs_minimal_mode"))

                # Solo badge toggle
                badge_23501 = tab.scope_bar._spec_buttons["23.501"]
                badge_23501.click()
                self.assertEqual(tab._solo_spec, "23.501")
                self.assertIsNotNone(tab.scope_bar.btn_clear_solo)
                eff_ids = tab._get_effective_version_ids()
                expected_23501_ids = [v["id"] for v in s_db.get_versions_for_spec("23.501")]
                self.assertEqual(eff_ids, expected_23501_ids)

                tab.scope_bar.btn_clear_solo.click()
                self.assertIsNone(tab._solo_spec)
                self.assertEqual(len(tab._get_effective_version_ids()), 2)

                # Re-opening tab restores Minimal Mode
                tab.shutdown()
                tab2 = SpecSearchTab(search_db_path, specs_db_path)
                self.assertTrue(tab2._specs_minimal_mode)
                self.assertFalse(tab2.scope_bar.isHidden())
                tab2.shutdown()

    def test_linkify_inspector_history_and_cross_ref_navigation(self) -> None:
        raw = (
            "See clause 5.15.2.1, Annex A.1, TS 23.502 [3] clause 4.2.2, "
            "and clause 4.3.1 of TS 23.502 (not annex a)."
        )
        linked = linkify_clause_references(raw)
        self.assertIn('href="clause:5.15.2.1"', linked)
        self.assertIn('href="clause:Annex%20A.1"', linked)
        self.assertIn('href="spec-clause:23.502/4.2.2"', linked)
        self.assertIn('href="spec-clause:23.502/4.3.1"', linked)
        self.assertIn("(not annex a).", linked)

        highlighted = _highlight_html_text_nodes(linked, "clause")
        self.assertIn('href="clause:5.15.2.1"', highlighted)
        self.assertIn("<mark ", highlighted)

        inspector = SpecClauseInspector()
        emitted: list[tuple[str, str, str]] = []
        inspector.clause_reference_clicked.connect(lambda s, v, c: emitted.append((s, v, c)))

        inspector.display_clause("5.15.19", "Slice Replacement", "23.501", "18.4.0", "See clause 5.15.2.1.")
        self.assertFalse(inspector.btn_nav_back.isEnabled())
        self.assertFalse(inspector.btn_nav_forward.isEnabled())

        inspector._on_anchor_clicked(QUrl("clause:5.15.2.1"))
        inspector._on_anchor_clicked(QUrl("spec-clause:23.502/4.2.2"))
        self.assertEqual(emitted, [("23.501", "18.4.0", "5.15.2.1"), ("23.502", "", "4.2.2")])

        inspector.display_clause("5.15.2.1", "Slice Mod", "23.501", "18.4.0", "Body 2")
        self.assertTrue(inspector.btn_nav_back.isEnabled())
        self.assertFalse(inspector.btn_nav_forward.isEnabled())

        inspector.navigate_back()
        self.assertEqual(inspector._current_clause_num, "5.15.19")
        self.assertFalse(inspector.btn_nav_back.isEnabled())
        self.assertTrue(inspector.btn_nav_forward.isEnabled())

        inspector.navigate_forward()
        self.assertEqual(inspector._current_clause_num, "5.15.2.1")

        # Integration with SpecSearchTab._handle_clause_reference_clicked
        with tempfile.TemporaryDirectory() as tmp_dir:
            root = Path(tmp_dir)
            (root / "config").mkdir(exist_ok=True)
            search_db_path = root / "search.db"
            specs_db_path = root / "specs.db"
            s_db = SpecSearchDatabase(search_db_path)
            s_db.insert_parsed_spec(
                "23.501",
                "18.4.0",
                [
                    {"clause_number": "5.15.19", "clause_title": "Slice Replacement", "content": "See clause 5.15.2.1 and TS 23.502 clause 4.2.2."},
                    {"clause_number": "5.15.2.1", "clause_title": "Modification", "content": "Detailed modification procedure."},
                ],
                release_date="2024-09-01",
            )
            s_db.insert_parsed_spec(
                "23.502",
                "18.5.0",
                [
                    {"clause_number": "4.2.2", "clause_title": "Registration", "content": "Registration flow."},
                ],
                release_date="2024-09-02",
            )
            SpecsDatabase(specs_db_path)
            with patch("modules.spec_search.ui.spec_search_tab.tab.get_project_root", return_value=root):
                tab = SpecSearchTab(search_db_path, specs_db_path)
                tab.inspector.display_clause("5.15.19", "Slice Replacement", "23.501", "18.4.0", "See clause 5.15.2.1.")
                tab._handle_clause_reference_clicked("23.501", "18.4.0", "5.15.2.1")
                self.assertEqual(tab.inspector._current_spec, "23.501")
                self.assertEqual(tab.inspector._current_clause_num, "5.15.2.1")

                tab._handle_clause_reference_clicked("23.502", "", "4.2.2")
                self.assertEqual(tab.inspector._current_spec, "23.502")
                self.assertEqual(tab.inspector._current_version, "18.5.0")
                self.assertEqual(tab.inspector._current_clause_num, "4.2.2")

                tab.inspector.navigate_back()
                self.assertEqual(tab.inspector._current_clause_num, "5.15.2.1")
                tab.shutdown()


if __name__ == "__main__":
    unittest.main()
