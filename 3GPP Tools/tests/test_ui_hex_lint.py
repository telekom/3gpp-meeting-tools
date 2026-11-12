"""
Phase 0.4 & 0.5 tests:
1. Raw-hex lint ratchet guard for UI modules outside `core.ui.theme`.
2. Unit tests for `core.ui.hex_migrator` AST/tokenizer transformations.
"""

from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from typing import Dict

from core.ui.hex_migrator import (
    find_hex_literals_in_source,
    migrate_file,
    migrate_source,
    scan_ui_hex_counts,
)
from core.ui.theme import HEX_TO_TOKEN, C

# Maximum allowed legacy raw `#RRGGBB` literals per UI file while gradual
# tab-by-tab migration proceeds. New UI files and migrated files (`ui_components.py`)
# are NOT in this allowlist and must have 0 raw hex literals.
# Decrement or remove entries as each Phase 1/2 file is migrated.
LEGACY_UI_HEX_ALLOWLIST: Dict[str, int] = {
    "src/core/ai/ollama_dialog.py": 32,
    "src/core/ui/database_inspector.py": 6,
    "src/core/ui/database_maintenance_dialog.py": 66,
    "src/core/ui/drop_zone.py": 92,
    "src/core/ui/mcp_monitor_dialog.py": 50,
    "src/core/ui/process_manager_dialog.py": 8,
    "src/core/ui/system_status_bar.py": 38,
    "src/core/ui/ui_panels.py": 10,
    "src/modules/assistant/ui/assistant_components.py": 104,
    "src/modules/assistant/ui/assistant_tab.py": 15,
    "src/modules/emails/ui/config_dialog.py": 5,
    "src/modules/emails/ui/email_dialogs.py": 3,
    "src/modules/emails/ui/email_models.py": 1,
    "src/modules/emails/ui/email_summary_dialog.py": 18,
    "src/modules/emails/ui/email_toolbar.py": 17,
    "src/modules/emails/ui/email_window.py": 16,
    "src/modules/emails/ui/general_email_dialog.py": 10,
    "src/modules/emails/ui/tdoc_emails_dialog.py": 30,
    "src/modules/meetings/ui/agreement_tracker_dialog.py": 70,
    "src/modules/meetings/ui/chair_notes_picker_dialog.py": 20,
    "src/modules/meetings/ui/contributions/contribution_dialog.py": 12,
    "src/modules/meetings/ui/contributions/contribution_models.py": 16,
    "src/modules/meetings/ui/contributions/contribution_widgets.py": 12,
    "src/modules/meetings/ui/delegate_owner_picker.py": 18,
    "src/modules/meetings/ui/dialogs/ai_suggestion_dialog.py": 41,
    "src/modules/meetings/ui/dialogs/dossier_export_dialog.py": 13,
    "src/modules/meetings/ui/dialogs/excel_export_dialog.py": 10,
    "src/modules/meetings/ui/dialogs/post_meeting_export_dialog.py": 52,
    "src/modules/meetings/ui/dialogs/pre_meeting_export_dialog.py": 50,
    "src/modules/meetings/ui/dialogs/topic_dossier_dialog.py": 35,
    "src/modules/meetings/ui/dialogs/topic_inputs_dialog.py": 159,
    "src/modules/meetings/ui/meeting_agreements_dialog.py": 78,
    "src/modules/meetings/ui/meeting_details_dialogs.py": 16,
    "src/modules/meetings/ui/meetings_delegates.py": 17,
    "src/modules/meetings/ui/session_plan/session_plan_dialogs.py": 72,
    "src/modules/meetings/ui/session_plan/session_plan_window.py": 237,
    "src/modules/meetings/ui/tdoc_delegates.py": 17,
    "src/modules/meetings/ui/tdoc_info_dialog.py": 37,
    "src/modules/meetings/ui/tdoc_triage_dialog.py": 24,
    "src/modules/meetings/ui/tdocs_components.py": 46,
    "src/modules/meetings/ui/tdocs_dialogs.py": 11,
    "src/modules/meetings/ui/tdocs_drop_overlay.py": 6,
    "src/modules/meetings/ui/tdocs_help_dialog.py": 60,
    "src/modules/meetings/ui/tdocs_models.py": 53,
    "src/modules/meetings/ui/tdocs_window.py": 67,
    "src/modules/meetings/ui/topic_pills.py": 50,
    "src/modules/meetings/ui/ui_tabs.py": 21,
    "src/modules/nas/core/nas_exporter.py": 18,
    "src/modules/nas/ui/nas_components.py": 121,
    "src/modules/nas/ui/nas_dialogs.py": 8,
    "src/modules/nas/ui/nas_matrix_controller.py": 4,
    "src/modules/nas/ui/nas_models.py": 3,
    "src/modules/nas/ui/nas_tabs.py": 107,
    "src/modules/puml2visio/ui/batch_convert_tab.py": 32,
    "src/modules/puml2visio/ui/callflow_dialog.py": 35,
    "src/modules/puml2visio/ui/cards/office_bridge_card.py": 4,
    "src/modules/puml2visio/ui/cards/puml_batch_card.py": 2,
    "src/modules/puml2visio/ui/cards/recent_exports_strip.py": 8,
    "src/modules/puml2visio/ui/cards/stencils_card.py": 3,
    "src/modules/puml2visio/ui/cards/vsdx_extractor_card.py": 5,
    "src/modules/puml2visio/ui/elements_palette.py": 12,
    "src/modules/puml2visio/ui/java_dialog.py": 42,
    "src/modules/puml2visio/ui/preset_dialog.py": 17,
    "src/modules/puml2visio/ui/select_presentation_dialog.py": 55,
    "src/modules/puml2visio/ui/svg_preview_widget.py": 6,
    "src/modules/puml2visio/ui/syntax_highlighter.py": 8,
    "src/modules/puml2visio/ui/ui_components.py": 24,
    "src/modules/puml2visio/ui/ui_tabs.py": 3,
    "src/modules/spec_search/core/spec_clause_exporter.py": 5,
    "src/modules/spec_search/ui/spec_clause_diff_dialog.py": 110,
    "src/modules/spec_search/ui/spec_search_components.py": 136,
    "src/modules/spec_search/ui/spec_search_dialogs.py": 80,
    "src/modules/spec_search/ui/spec_search_tabs.py": 93,
    "src/modules/specifications/ui/dialogs.py": 21,
    "src/modules/specifications/ui/spec_annotations_dialog.py": 39,
    "src/modules/specifications/ui/spec_clause_dialog.py": 25,
    "src/modules/specifications/ui/spec_history_dialog.py": 37,
    "src/modules/specifications/ui/spec_info_dialog.py": 17,
    "src/modules/specifications/ui/spec_menu.py": 5,
    "src/modules/specifications/ui/spec_tdoc_report_dialog.py": 39,
    "src/modules/specifications/ui/ui_tabs.py": 68,
    "src/modules/word_tools/ui/batch_label_dialog.py": 17,
    "src/modules/word_tools/ui/cards/compare_card.py": 2,
    "src/modules/word_tools/ui/cards/converter_card.py": 14,
    "src/modules/word_tools/ui/cards/cr_pcr_card.py": 6,
    "src/modules/word_tools/ui/cards/excerpt_card.py": 6,
    "src/modules/word_tools/ui/cards/markdown_card.py": 14,
    "src/modules/word_tools/ui/cards/sanitizer_card.py": 13,
    "src/modules/word_tools/ui/cr_pcr_dialog.py": 31,
    "src/modules/word_tools/ui/document_selector_pane.py": 2,
    "src/modules/word_tools/ui/word_settings_dialog.py": 4,
    "src/modules/word_tools/ui/word_tabs.py": 29,
    "src/modules/work_items/ui/ui_tabs.py": 16,
    "src/modules/work_items/ui/work_items_delegates.py": 35,
    "src/modules/work_items/ui/work_items_dialogs.py": 78,
    "src/modules/work_items/ui/work_items_model.py": 5,
}


class TestUiHexLintRatchet(unittest.TestCase):
    def test_no_new_or_increased_raw_hex_colors_in_ui(self):
        repo_root = Path(__file__).resolve().parents[1]
        current = scan_ui_hex_counts(repo_root)

        violations = []
        for path, findings in sorted(current.items()):
            allowed = LEGACY_UI_HEX_ALLOWLIST.get(path, 0)
            if len(findings) > allowed:
                preview = ", ".join(
                    f"L{ln}:{hx}"
                    + (f" (-> C.{HEX_TO_TOKEN[hx]})" if hx in HEX_TO_TOKEN else "")
                    for ln, hx in findings[:8]
                )
                violations.append(
                    f"{path}: found {len(findings)} raw hexes (limit {allowed}). "
                    f"Examples: {preview}. Run `py -3.14 -m core.ui.hex_migrator {path} --dry-run --diff`"
                )

        self.assertEqual(
            violations,
            [],
            msg="Raw hex color lint check failed:\n" + "\n".join(violations),
        )


class TestHexMigrator(unittest.TestCase):
    def test_standalone_and_qss_braces_preserve_exact_runtime_value(self):
        src = (
            '"""Docstring with #FFFFFF should be ignored."""\n'
            "from PyQt5.QtWidgets import QWidget\n"
            "\n"
            "# Comment with #FFFFFF ignored\n"
            'BG = "#FFFFFF"\n'
            'QSS = "QPushButton { background: #FFFFFF; border: 1px solid #CBD5E1; color: #123456; }"\n'
        )
        res = migrate_source(src)
        self.assertTrue(res.changed)
        self.assertTrue(res.import_added)
        self.assertEqual(res.replacements, 3)
        self.assertIn("from core.ui.theme import C", res.updated)
        self.assertIn('"""Docstring with #FFFFFF should be ignored."""', res.updated)
        self.assertIn("BG = C.BG_CARD", res.updated)
        self.assertIn("{C.BORDER_STRONG}", res.updated)
        self.assertEqual(res.unmapped_hexes, {"#123456": 1})

        # Evaluate migrated expressions and confirm identical runtime strings
        ns = {"QWidget": object}
        # Replace PyQt5 import line with noop for standalone exec
        exec_code = res.updated.replace("from PyQt5.QtWidgets import QWidget", "")
        exec(exec_code, ns)
        self.assertEqual(ns["BG"], "#FFFFFF")
        self.assertEqual(
            ns["QSS"],
            "QPushButton { background: #FFFFFF; border: 1px solid #CBD5E1; color: #123456; }",
        )
        self.assertEqual(ns["C"], C)

    def test_existing_fstring_escaped_braces_preserved(self):
        src = (
            "from core.ui.theme import C\n"
            "w = 12\n"
            'S = f"QFrame {{ background: #F8FAFC; border: 1px solid #CBD5E1; width: {w}px; }}"\n'
        )
        res = migrate_source(src)
        self.assertFalse(res.import_added)
        self.assertEqual(res.replacements, 2)
        ns = {}
        exec(res.updated, ns)
        self.assertEqual(
            ns["S"],
            "QFrame { background: #F8FAFC; border: 1px solid #CBD5E1; width: 12px; }",
        )

    def test_find_hex_literals_ignores_comments_and_docstrings(self):
        src = (
            '"""Doc #FFFFFF"""\n'
            "# Comment #CBD5E1\n"
            'btn = "color: #0F172A;"\n'
        )
        self.assertEqual(find_hex_literals_in_source(src), [(3, "#0F172A")])

    def test_dry_run_vs_apply_on_disk(self):
        src = 'from core.ui.theme import C\nCOLOR = "#1E5C99"\n'
        with tempfile.TemporaryDirectory() as tmp:
            p = Path(tmp) / "sample_ui.py"
            p.write_text(src, encoding="utf-8")

            dry = migrate_file(p, apply=False)
            self.assertTrue(dry.changed)
            self.assertEqual(p.read_text(encoding="utf-8"), src)

            applied = migrate_file(p, apply=True)
            self.assertTrue(applied.changed)
            self.assertIn("COLOR = C.ACCENT_PRIMARY", p.read_text(encoding="utf-8"))


if __name__ == "__main__":
    unittest.main()
