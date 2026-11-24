# --- File: tests/test_code_modularity_lint.py ---
"""Modularity & anti-god-class ratchet tests for `src/` and root `AGENTS.md`."""

import unittest
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
SRC_ROOT = PROJECT_ROOT / "src"

MAX_NEW_FILE_LINES = 600
HARD_MAX_FILE_LINES = 1000
MAX_AGENTS_MD_BYTES = 16 * 1024  # 16 KB (2 KB safety margin below 18 KB context cutoff)

# Protected upstream-synced drop-in files exempt from line-count limits.
PROTECTED_FILES = {
    "src/modules/meetings/core/exporters/tdoc2md.py",
}

# Shrink-only allowlist of existing `src/**/*.py` files currently above 600 lines.
# Never increase a value here or add new entries; refactor into a package instead.
LEGACY_LARGE_FILE_ALLOWLIST = {
    "src/core/ai/cli_detector.py": 638,
    "src/core/ai/cli_launcher.py": 768,
    "src/core/ai/mcp_activity.py": 654,
    "src/core/ai/ollama_client.py": 681,
    "src/core/ai/ollama_dialog.py": 621,
    "src/core/ui/font_scale.py": 756,
    "src/core/ui/update_dialog.py": 787,
    "src/core/ui/window_session.py": 632,
    "src/modules/assistant/ui/assistant_components.py": 807,
    "src/modules/emails/core/tracked_threads/collection_and_pinning.py": 642,
    "src/modules/emails/ui/email_window.py": 799,
    "src/modules/emails/ui/meeting_inbox/mixins/data_and_filtering.py": 707,
    "src/modules/emails/ui/meeting_inbox/mixins/triage_actions.py": 683,
    "src/modules/emails/ui/meeting_inbox/mixins/ui_setup.py": 632,
    "src/modules/meetings/core/agreement_tracker_worker.py": 704,
    "src/modules/meetings/core/chair_notes_downloader.py": 617,
    "src/modules/meetings/core/exporters/topic_dossier/mixins/export_pipeline.py": 607,
    "src/modules/meetings/core/meetings_db.py": 749,
    "src/modules/meetings/core/scraper.py": 689,
    "src/modules/meetings/core/session_plan/session_plan_timezone.py": 707,
    "src/modules/meetings/ui/contributions/contribution_dialog.py": 710,
    "src/modules/meetings/ui/meeting_agreements_dialog.py": 741,
    "src/modules/meetings/ui/tdocs/window/mixins/filter_setup.py": 624,
    "src/modules/meetings/ui/tdocs/window/mixins/ui_header.py": 653,
    "src/modules/meetings/ui/tdocs_agenda_controller.py": 601,
    "src/modules/meetings/ui/tdocs_help_dialog.py": 642,
    "src/modules/nas/core/parsing/asn1_base_parser.py": 616,
    "src/modules/nas/core/parsing/nas_parser.py": 781,
    "src/modules/puml2visio/core/visio_converter.py": 615,
    "src/modules/puml2visio/templates/builtin_presets.py": 925,
    "src/modules/puml2visio/ui/callflow_dialog.py": 736,
    "src/modules/puml2visio/ui/ui_components.py": 765,
    "src/modules/spec_search/ui/quick_search_panel.py": 658,
    "src/modules/specifications/core/spec_history_db.py": 713,
    "src/modules/specifications/core/spec_history_extractor.py": 620,
    "src/modules/specifications/ui/dialogs.py": 619,
    "src/modules/specifications/ui/spec_clause_dialog.py": 679,
    "src/modules/specifications/ui/spec_tdoc_report_dialog.py": 677,
    "src/modules/word_tools/core/document_sanitizer.py": 674,
    "src/modules/word_tools/core/sensitivity_label.py": 686,
    "src/modules/word_tools/core/word_excerpt_extractor.py": 668,
    "src/modules/word_tools/ui/cards/excerpt_card.py": 609,
    "src/modules/word_tools/ui/cards/markdown_card.py": 653,
    "src/modules/word_tools/ui/cards/sanitizer_card.py": 650,
    "src/modules/word_tools/ui/word_settings_dialog.py": 726,
    "src/modules/word_tools/ui/word_tabs.py": 641,
}


def _count_lines(path: Path) -> int:
    return len(path.read_text(encoding="utf-8", errors="replace").splitlines())


class TestCodeModularityLint(unittest.TestCase):
    """Prevents new god classes/monoliths and keeps root AGENTS.md under 18 KB."""

    def test_no_new_or_growing_monolith_files_in_src(self) -> None:
        violations = []
        for py_path in sorted(SRC_ROOT.rglob("*.py")):
            rel = py_path.relative_to(PROJECT_ROOT).as_posix()
            if rel in PROTECTED_FILES:
                continue
            lines = _count_lines(py_path)
            if lines > HARD_MAX_FILE_LINES:
                violations.append(
                    f"{rel}: {lines} lines exceeds hard ceiling of {HARD_MAX_FILE_LINES} lines"
                )
                continue
            limit = LEGACY_LARGE_FILE_ALLOWLIST.get(rel, MAX_NEW_FILE_LINES)
            if lines > limit:
                violations.append(
                    f"{rel}: {lines} lines exceeds allowed ceiling of {limit} lines "
                    "(decompose into a focused package with a thin re-export shim)"
                )
        self.assertEqual(
            violations,
            [],
            "Modularity lint violations found:\n  - " + "\n  - ".join(violations),
        )

    def test_all_agents_md_files_under_16kb(self) -> None:
        oversized = []
        for agents_path in sorted(PROJECT_ROOT.rglob("AGENTS.md")):
            size_bytes = agents_path.stat().st_size
            if size_bytes >= MAX_AGENTS_MD_BYTES:
                rel = agents_path.relative_to(PROJECT_ROOT).as_posix()
                oversized.append(
                    f"{rel}: {size_bytes} bytes (>= {MAX_AGENTS_MD_BYTES} bytes). "
                    "Tighten prose or move detailed specs into docs/*.md / README.md."
                )
        self.assertEqual(
            oversized,
            [],
            "Oversized AGENTS.md files found:\n  - " + "\n  - ".join(oversized),
        )


if __name__ == "__main__":
    unittest.main()
