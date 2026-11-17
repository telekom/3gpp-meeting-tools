# --- File: tests/test_cli_agent_integration.py ---
from __future__ import annotations

import json
import os
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch

from PyQt5.QtWidgets import QApplication, QMenu

# Ensure 'src' is in sys.path for test execution
project_root = Path(__file__).resolve().parent.parent
src_dir = project_root / "src"
if str(src_dir) not in sys.path:
    sys.path.insert(0, str(src_dir))

from core.ai import cli_detector
from core.ai.cli_detector import (
    CliAgentStatus,
    build_interactive_cli_args,
    check_mcp_configured,
    detect_cli_agents,
    invalidate_cli_detection_cache,
)
from core.ai.cli_launcher import (
    CLI_CONTEXT_FILENAME,
    build_powershell_launch_command,
    launch_cli_in_terminal,
    populate_cli_agent_menu,
    write_cli_context_file,
)
from core.ai.mcp_activity import get_mcp_client_configurations
from modules.meetings.core.exporters.export_manifest import ExportManifest, save_manifest
from modules.meetings.core.meetings_db import MeetingsDatabase


class TestCliDetector(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.tmp_path = Path(self._tmp.name)
        invalidate_cli_detection_cache()

    def tearDown(self) -> None:
        invalidate_cli_detection_cache()
        self._tmp.cleanup()

    def test_detect_via_which_and_user_dir_fallback(self):
        fake_bin = self.tmp_path / "npm"
        fake_bin.mkdir(parents=True)
        codex_cmd = fake_bin / "codex.cmd"
        codex_cmd.write_text("@echo off\n", encoding="utf-8")

        claude_dir = self.tmp_path / "local_bin"
        claude_dir.mkdir(parents=True)
        claude_exe = claude_dir / "claude.exe"
        claude_exe.write_text("", encoding="utf-8")

        def fake_which(name: str):
            if name == "codex.cmd":
                return str(codex_cmd)
            return None

        with patch.object(cli_detector, "CONFIG_PATH", self.tmp_path / "missing_cfg.json"):
            statuses = detect_cli_agents(
                force_refresh=True,
                which_fn=fake_which,
                user_dirs=[claude_dir],
                home_dir=self.tmp_path,
            )

        self.assertTrue(statuses["codex"].installed)
        self.assertEqual(statuses["codex"].source, "path")
        self.assertEqual(statuses["codex"].executable_path, str(codex_cmd))

        self.assertTrue(statuses["claude"].installed)
        self.assertEqual(statuses["claude"].source, "user_dir")
        self.assertEqual(statuses["claude"].executable_path, str(claude_exe))

        self.assertFalse(statuses["gemini"].installed)
        self.assertEqual(statuses["gemini"].source, "not_found")

    def test_custom_config_override(self):
        custom_gemini = self.tmp_path / "my_gemini.cmd"
        custom_gemini.write_text("@echo off\n", encoding="utf-8")
        cfg_file = self.tmp_path / "cli_agents_config.json"
        cfg_file.write_text(
            json.dumps({"preferred_agent": "gemini", "custom_paths": {"gemini": str(custom_gemini)}}),
            encoding="utf-8",
        )

        with patch.object(cli_detector, "CONFIG_PATH", cfg_file):
            statuses = detect_cli_agents(
                force_refresh=True,
                which_fn=lambda _n: None,
                user_dirs=[],
                home_dir=self.tmp_path,
            )
            pref = cli_detector.get_preferred_cli_agent()

        self.assertTrue(statuses["gemini"].installed)
        self.assertEqual(statuses["gemini"].source, "config")
        self.assertEqual(statuses["gemini"].executable_path, str(custom_gemini))
        self.assertEqual(pref, "gemini")

    def test_interactive_args_builder(self):
        self.assertEqual(build_interactive_cli_args("codex", "codex.cmd", "hello"), ["codex.cmd", "hello"])
        self.assertEqual(build_interactive_cli_args("claude", "claude.exe", "hello"), ["claude.exe", "hello"])
        self.assertEqual(build_interactive_cli_args("gemini", "gemini.cmd", "hello"), ["gemini.cmd", "-i", "hello"])

    def test_check_mcp_configured(self):
        (self.tmp_path / ".codex").mkdir()
        (self.tmp_path / ".codex" / "config.toml").write_text('[mcp_servers."3gpp-tools"]\n', encoding="utf-8")
        (self.tmp_path / ".gemini").mkdir()
        (self.tmp_path / ".gemini" / "settings.json").write_text(
            json.dumps({"mcpServers": {"3gpp-tools": {"command": "python"}}}),
            encoding="utf-8",
        )
        (self.tmp_path / ".claude.json").write_text("{}", encoding="utf-8")

        self.assertTrue(check_mcp_configured("codex", home_dir=self.tmp_path))
        self.assertTrue(check_mcp_configured("gemini", home_dir=self.tmp_path))
        self.assertFalse(check_mcp_configured("claude", home_dir=self.tmp_path))


class TestCliLauncher(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.tmp_path = Path(self._tmp.name)

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def test_write_cli_context_file_for_tdocs_and_triage(self):
        md_file = self.tmp_path / "S2-2601720" / "S2-2601720.md"
        md_file.parent.mkdir(parents=True)
        md_file.write_text("# S2-2601720\nTest content", encoding="utf-8")

        ctx_path = write_cli_context_file(
            self.tmp_path,
            task_mode="compare",
            tdocs_meta=[
                {
                    "TDoc": "S2-2601720",
                    "Title": "URSP enhancement",
                    "Source": "Deutsche Telekom, Ericsson?",
                    "Type": "pCR",
                    "Agenda Item": "5.15.1",
                    "Related WIs": "FS_6G_ARC",
                    "TDoc Status": "Revised",
                }
            ],
            markdown_files=[md_file],
            local_triage_summary="Key issue: clarifies URSP rule precedence.",
        )

        self.assertTrue(ctx_path.is_file())
        self.assertEqual(ctx_path.name, CLI_CONTEXT_FILENAME)
        content = ctx_path.read_text(encoding="utf-8")
        self.assertIn("S2-2601720", content)
        self.assertIn("S2-2601720/S2-2601720.md", content)
        self.assertIn("clarifies URSP rule precedence", content)
        self.assertIn("Tentative / Carry-Over Co-Signers", content)
        self.assertIn("[INSERTED:", content)
        self.assertIn("cli_tools.py", content)

    def test_write_cli_context_file_for_topic_stage1(self):
        ctx_path = write_cli_context_file(
            self.tmp_path,
            task_mode="topic_stage1",
            topic_name="Ambient IoT (Rel-19)",
            manifest_summary={
                "analyzed_meetings": ["SA2_172"],
                "stale_meetings": ["RAN2_132"],
                "pending_meetings": ["SA2_173", "CT1_155"],
                "synthesis_status": "stale",
            },
        )
        content = ctx_path.read_text(encoding="utf-8")
        self.assertIn("Ambient IoT (Rel-19)", content)
        self.assertIn("Stage-1 meeting-by-meeting analysis", content)
        self.assertIn("SA2_173, CT1_155", content)
        self.assertIn("RAN2_132", content)

    def test_powershell_command_escaping_with_spaces_and_quotes(self):
        work_dir = Path("C:/Users/Test User/O'Brien Docs/S2-2601720")
        cmd = build_powershell_launch_command(
            "gemini",
            "C:/Program Files/nodejs/gemini.cmd",
            work_dir,
            "Read _CLI_CONTEXT.md and don't skip details",
        )
        self.assertIn("O''Brien Docs", cmd)
        self.assertIn("don''t skip details", cmd)
        self.assertIn("'-i'", cmd)

    def test_launch_cli_in_terminal_wt_and_powershell_fallback(self):
        statuses = {
            "codex": CliAgentStatus(
                agent_id="codex",
                label="Codex CLI",
                vendor="OpenAI",
                icon="🤖",
                installed=True,
                executable_path="C:/npm/codex.cmd",
                source="path",
            )
        }

        spawned = []

        def fake_popen(cmd, **kwargs):
            spawned.append((cmd, kwargs))
            return MagicMock()

        ok_wt = launch_cli_in_terminal(
            "codex",
            self.tmp_path,
            statuses=statuses,
            popen_fn=fake_popen,
            which_fn=lambda n: "wt.exe" if n == "wt.exe" else None,
        )
        self.assertTrue(ok_wt)
        self.assertEqual(spawned[0][0][0], "wt.exe")
        self.assertEqual(spawned[0][0][1], "-d")
        self.assertIn("-EncodedCommand", spawned[0][0])
        for arg in spawned[0][0]:
            self.assertNotIn(";", arg)

        spawned.clear()
        ok_ps = launch_cli_in_terminal(
            "codex",
            self.tmp_path,
            statuses=statuses,
            popen_fn=fake_popen,
            which_fn=lambda _n: None,
        )
        self.assertTrue(ok_ps)
        self.assertEqual(spawned[0][0][0], "powershell.exe")

        # Missing directory returns False
        self.assertFalse(
            launch_cli_in_terminal(
                "codex",
                self.tmp_path / "does_not_exist",
                statuses=statuses,
                popen_fn=fake_popen,
            )
        )


class TestCliUiIntegration(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.tmp_path = Path(self._tmp.name)

    def tearDown(self):
        self._tmp.cleanup()

    def test_populate_cli_agent_menu_and_tdocs_context_menu(self):
        from modules.meetings.ui.tdocs_menus import build_row_context_menu

        fake_statuses = {
            "codex": CliAgentStatus("codex", "Codex CLI", "OpenAI", "🤖", True, "C:/npm/codex.cmd", "path", True),
            "claude": CliAgentStatus("claude", "Claude CLI", "Anthropic", "🧠", False, None, "not_found", False, "npm i -g @anthropic-ai/claude-code"),
            "gemini": CliAgentStatus("gemini", "Gemini CLI", "Google", "✨", True, "C:/npm/gemini.cmd", "path", False),
        }

        launched = []
        menu = QMenu()
        populate_cli_agent_menu(
            menu,
            on_launch_callback=lambda aid, mode: launched.append((aid, mode)),
            has_multiple_or_revs=True,
            multi_count=3,
            statuses=fake_statuses,
        )
        texts = [a.text() for a in menu.actions() if not a.isSeparator()]
        self.assertTrue(any("Codex CLI" in t for t in texts))
        self.assertTrue(any("Gemini CLI" in t for t in texts))
        self.assertTrue(any("Claude CLI (Not Found)" in t for t in texts))

        # Trigger Summarize action on Codex submenu
        codex_act = next(a for a in menu.actions() if "Codex CLI" in a.text())
        codex_sub = codex_act.menu()
        sum_act = next(a for a in codex_sub.actions() if "Summarize" in a.text())
        sum_act.trigger()
        self.assertEqual(launched, [("codex", "summarize")])

        # Verify build_row_context_menu includes CLI submenu without executing popup
        with patch.object(QMenu, "exec_", return_value=None), patch(
            "core.ai.cli_launcher.detect_cli_agents", return_value=fake_statuses
        ):
            build_row_context_menu(
                parent=None,
                row_data={"TDoc": "S2-2601720", "Title": "Test"},
                revisions=["S2-2601720r01"],
                family=["S2-2601720", "S2-2601720r01"],
                docs_ftp_url="",
                revisions_url="",
                meeting_dir=self.tmp_path,
                download_callback=lambda *a, **kw: None,
                export_llm_callback=lambda *a, **kw: None,
                compose_email_callback=lambda *a, **kw: None,
                open_emails_callback=lambda *a, **kw: None,
                mark_read_callback=lambda: None,
                mark_unread_callback=lambda: None,
                open_details_callback=lambda: None,
                triage_callback=lambda *a: None,
                launch_cli_callback=lambda *a: None,
            )

    def test_triage_dialog_cli_escalation(self):
        from modules.meetings.ui.tdoc_triage_dialog import TDocTriageDialog

        tdoc_dir = self.tmp_path / "S2-2601720"
        tdoc_dir.mkdir(parents=True)
        (tdoc_dir / "S2-2601720.md").write_text("# S2-2601720\nProposal text.", encoding="utf-8")

        with patch("modules.meetings.ui.tdoc_triage_dialog.OllamaClient.ping_and_get_models", return_value=(False, [], "offline")):
            dlg = TDocTriageDialog(
                tdoc_data={"TDoc": "S2-2601720", "Title": "Test TDoc"},
                meeting_dir=self.tmp_path,
                auto_start=False,
            )
        dlg.output_text.setPlainText("Local triage summary text.")

        with patch("modules.meetings.ui.tdoc_triage_dialog.launch_cli_in_terminal", return_value=True) as mock_launch:
            dlg._launch_in_cli("codex", "interactive")

        mock_launch.assert_called_once()
        ctx_file = tdoc_dir / CLI_CONTEXT_FILENAME
        self.assertTrue(ctx_file.is_file())
        self.assertIn("Local triage summary text.", ctx_file.read_text(encoding="utf-8"))
        dlg.close()

    def test_topic_dossier_dialog_cli_button_and_manifest_summary(self):
        from modules.meetings.ui.dialogs.topic_dossier_dialog import TopicDossierDialog

        db_path = self.tmp_path / "meetings.db"
        mdb = MeetingsDatabase(db_path=db_path)

        export_dir = self.tmp_path / "Topic_Test"
        export_dir.mkdir(parents=True)
        (export_dir / "SA2_173").mkdir()
        (export_dir / "SA2_173" / "_company_positions_analysis.md").write_text("# Done", encoding="utf-8")
        (export_dir / "SA2_174").mkdir()

        manifest = ExportManifest(meeting="Test Topic", export_type="topic_dossier")
        manifest.exported_meetings = {
            "SA2_173": {"wg": "SA2", "decision_count": 2, "tdoc_count": 3, "analysis_status": "analyzed"},
            "SA2_174": {"wg": "SA2", "decision_count": 1, "tdoc_count": 2, "analysis_status": "pending"},
        }
        save_manifest(export_dir, manifest)

        dlg = TopicDossierDialog(meetings_db=mdb, parent=None)
        dlg.edit_dest_folder.setText(str(export_dir))
        dlg._check_incremental_status()
        self.assertTrue(dlg.btn_launch_cli.isEnabled())

        summary = dlg._collect_topic_manifest_summary(export_dir)
        self.assertEqual(summary["analyzed_meetings"], ["SA2_173"])
        self.assertEqual(summary["pending_meetings"], ["SA2_174"])

        with patch("modules.meetings.ui.dialogs.topic_dossier_dialog.launch_cli_in_terminal", return_value=True) as mock_launch:
            dlg._launch_cli_for_topic("gemini", "topic_stage1")
        mock_launch.assert_called_once()
        self.assertTrue((export_dir / CLI_CONTEXT_FILENAME).is_file())
        dlg.close()

    def test_mcp_configs_include_gemini_and_monitor_dialog_shows_strip(self):
        from core.ui.mcp_monitor_dialog import McpMonitorDialog

        cfgs = get_mcp_client_configurations()
        self.assertIn("gemini_cli", cfgs)
        self.assertIn("mcpServers", cfgs["gemini_cli"]["config_json"])

        dlg = McpMonitorDialog(parent=None)
        self.assertEqual(set(dlg._cli_status_widgets.keys()), {"codex", "claude", "gemini"})
        dlg.close()


if __name__ == "__main__":
    unittest.main()
