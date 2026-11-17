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
        with patch.object(cli_detector, "CONFIG_PATH", self.tmp_path / "missing_cfg.json"):
            self.assertEqual(build_interactive_cli_args("codex", "codex.cmd", "hello"), ["codex.cmd", "hello"])
            self.assertEqual(build_interactive_cli_args("claude", "claude.exe", "hello"), ["claude.exe", "hello"])
            self.assertEqual(build_interactive_cli_args("gemini", "gemini.cmd", "hello"), ["gemini.cmd", "-i", "hello"])

            self.assertEqual(
                build_interactive_cli_args("codex", "codex.exe", "hello", model="gpt-6.1-sol", reasoning_effort="low"),
                ["codex.exe", "-m", "gpt-6.1-sol", "-c", "model_reasoning_effort=low", "hello"],
            )
            self.assertEqual(
                build_interactive_cli_args("claude", "claude.exe", "hello", model="sonnet", reasoning_effort="high"),
                ["claude.exe", "--model", "sonnet", "--effort", "high", "hello"],
            )
            self.assertEqual(
                build_interactive_cli_args("gemini", "gemini.cmd", "hello", model="gemini-2.5-pro"),
                ["gemini.cmd", "-m", "gemini-2.5-pro", "-i", "hello"],
            )

    def test_discover_codex_models_from_cache_and_config_toml(self):
        codex_dir = self.tmp_path / ".codex"
        codex_dir.mkdir(parents=True)
        (codex_dir / "config.toml").write_text(
            'model = "gpt-6.1-sol"\nmodel_reasoning_effort = "high"\n[tui]\n',
            encoding="utf-8",
        )
        (codex_dir / "models_cache.json").write_text(
            json.dumps({
                "models": [
                    {
                        "slug": "gpt-6.1-sol",
                        "display_name": "GPT-6.1-Sol",
                        "visibility": "list",
                        "default_reasoning_level": "medium",
                        "supported_reasoning_levels": [
                            {"effort": "low"},
                            {"effort": "medium"},
                            {"effort": "high"},
                        ],
                    },
                    {
                        "slug": "hidden-model",
                        "display_name": "Hidden",
                        "visibility": "hide",
                    },
                ]
            }),
            encoding="utf-8",
        )

        with patch.object(cli_detector, "_CODEX_MODELS_CACHE", None), \
             patch.object(cli_detector, "CONFIG_PATH", self.tmp_path / "cli_agents_config.json"):
            models, def_model, def_effort = cli_detector.discover_cli_models("codex", home_dir=self.tmp_path)
            self.assertEqual([m.slug for m in models], ["gpt-6.1-sol"])
            self.assertEqual(models[0].reasoning_levels, ("low", "medium", "high"))
            self.assertEqual(def_model, "gpt-6.1-sol")
            self.assertEqual(def_effort, "high")
            self.assertEqual(
                cli_detector.format_active_model_badge("codex", home_dir=self.tmp_path),
                "default: GPT-6.1-Sol · high",
            )

            cli_detector.set_cli_agent_model_config("codex", "gpt-6.1-sol", "low")
            self.assertEqual(cli_detector.get_cli_agent_model_config("codex"), ("gpt-6.1-sol", "low"))
            self.assertEqual(
                cli_detector.format_active_model_badge("codex", home_dir=self.tmp_path),
                "GPT-6.1-Sol · low",
            )

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

    def test_single_tdoc_cli_launch_uses_canonical_folder_without_export_dir(self):
        from PyQt5.QtWidgets import QWidget
        from modules.meetings.core.tdocs.tdoc_markdown_converter import CONVERTER_VERSION
        from modules.meetings.ui.dialogs.tdocs_export_controller import launch_cli_for_tdocs

        meeting_dir = self.tmp_path / "SA2_173"
        base_dir = meeting_dir / "S2-2601720"
        base_dir.mkdir(parents=True)
        rev_md = base_dir / "S2-2601720r01.md"
        rev_md.write_text(
            f"<!--- tdoc_converter: version={CONVERTER_VERSION}; mode=markup --->\n"
            "# S2-2601720r01: URSP clarification\n\nBody text long enough for completeness check." + ("x" * 120),
            encoding="utf-8",
        )

        win = QWidget()
        win.meeting_dir = meeting_dir
        win.docs_ftp_url = ""
        win.revisions_url = ""
        win.mtg_info = {"wg_name": "SA2", "meeting_number": "173"}
        win.model = MagicMock()
        win.model.revisions = {}

        with patch("modules.meetings.ui.dialogs.tdocs_export_controller.TDocMarkdownExporterThread") as mock_thread_cls, \
             patch("modules.meetings.ui.dialogs.tdocs_export_controller.launch_cli_in_terminal", return_value=True) as mock_launch:
            launch_cli_for_tdocs(
                win,
                [{"TDoc": "S2-2601720r01", "Title": "URSP clarification", "Type": "pCR"}],
                "codex",
                task_mode="summarize",
            )

        # Instant canonical fast-path: no background exporter thread, no Export/ folder
        mock_thread_cls.assert_not_called()
        mock_launch.assert_called_once()
        launched_dir = mock_launch.call_args.args[1]
        self.assertEqual(launched_dir, base_dir)
        self.assertFalse((meeting_dir / "Export").exists())

        ctx_file = base_dir / CLI_CONTEXT_FILENAME
        self.assertTrue(ctx_file.is_file())
        self.assertIn("S2-2601720r01.md", ctx_file.read_text(encoding="utf-8"))
        win.close()

    def test_single_tdoc_cli_conversion_stays_in_canonical_folder_and_batch_uses_export(self):
        from PyQt5.QtWidgets import QWidget
        from modules.meetings.core.exporters.llm_exporter import TDocMarkdownExporterThread
        from modules.meetings.core.tdocs.tdoc_markdown_converter import CONVERTER_VERSION
        from modules.meetings.ui.dialogs.tdocs_export_controller import launch_cli_for_tdocs

        meeting_dir = self.tmp_path / "SA2_174"
        tdoc_id = "S2-2602222"
        base_dir = meeting_dir / tdoc_id
        base_dir.mkdir(parents=True)

        # Fake existing Word doc so Phase 2 converts locally without network
        (base_dir / f"{tdoc_id}.docx").write_bytes(b"dummy")

        def fake_convert(doc_path, out_dir=None, tdoc_id=None, **_kwargs):
            md_path = Path(out_dir) / f"{tdoc_id}.md"
            text = (
                f"<!--- tdoc_converter: version={CONVERTER_VERSION}; mode=markup --->\n"
                f"# {tdoc_id}: Generated\n\n" + ("content " * 30)
            )
            md_path.write_text(text, encoding="utf-8")
            return MagicMock(md_path=md_path, text=text)

        # 1) Single-TDoc conversion via stage_copies=False must NOT create Export/ or index.md
        exporter = TDocMarkdownExporterThread(
            meeting_dir,
            [{"TDoc": tdoc_id, "Title": "Generated", "Type": "pCR"}],
            "",
            "",
            is_bulk=False,
            stage_copies=False,
        )
        results = []
        exporter.finished.connect(lambda ok, msg: results.append((ok, msg)))
        with patch("modules.meetings.core.tdocs.tdoc_markdown_converter.convert_document_to_markdown", side_effect=fake_convert):
            exporter.run()

        self.assertEqual(results[0][0], True)
        self.assertTrue((base_dir / f"{tdoc_id}.md").is_file())
        self.assertFalse((base_dir / tdoc_id).exists())
        self.assertFalse((base_dir / "index.md").exists())
        self.assertFalse((meeting_dir / "Export").exists())

        # 2) Multi-TDoc / compare mode DOES stage under Export/CLI_Batch_*
        win = QWidget()
        win.meeting_dir = meeting_dir
        win.docs_ftp_url = ""
        win.revisions_url = ""
        win.mtg_info = {"wg_name": "SA2", "meeting_number": "174"}
        win.model = MagicMock()
        win.model.revisions = {tdoc_id: [f"{tdoc_id}r01"]}

        (base_dir / f"{tdoc_id}r01.md").write_text(
            f"<!--- tdoc_converter: version={CONVERTER_VERSION}; mode=markup --->\n# {tdoc_id}r01\n\n" + ("rev " * 40),
            encoding="utf-8",
        )

        class _SyncThread:
            def __init__(self, m_dir, rows, docs_url, rev_url, is_bulk=True, output_dir=None, stage_copies=True):
                self._real = TDocMarkdownExporterThread(
                    m_dir, rows, docs_url, rev_url,
                    is_bulk=is_bulk, output_dir=output_dir, stage_copies=stage_copies,
                )
                self.progress = self._real.progress
                self.finished = self._real.finished

            def start(self):
                self._real.run()

        with patch("modules.meetings.ui.dialogs.tdocs_export_controller.TDocMarkdownExporterThread", _SyncThread), \
             patch("modules.meetings.ui.dialogs.tdocs_export_controller.launch_cli_in_terminal", return_value=True) as mock_launch:
            launch_cli_for_tdocs(
                win,
                [{"TDoc": tdoc_id, "Title": "Compare test", "Type": "pCR"}],
                "codex",
                task_mode="compare",
            )

        mock_launch.assert_called_once()
        batch_dir = mock_launch.call_args.args[1]
        self.assertEqual(batch_dir.parent, meeting_dir / "Export")
        self.assertTrue(batch_dir.name.startswith("CLI_Batch_"))
        self.assertTrue((batch_dir / "index.md").is_file())
        self.assertTrue((batch_dir / CLI_CONTEXT_FILENAME).is_file())
        win.close()

    def test_standing_context_and_workspace_user_notes_in_cli_context(self):
        from core.ai import cli_launcher
        from core.ai.cli_launcher import (
            USER_NOTES_FILENAME,
            load_workspace_user_notes,
            save_standing_cli_context,
            save_workspace_user_notes,
        )

        meeting_dir = self.tmp_path / "SA2_175"
        t1_dir = meeting_dir / "S2-2603001"
        t1_dir.mkdir(parents=True)
        (t1_dir / "S2-2603001.md").write_text("# S2-2603001", encoding="utf-8")

        batch_dir = meeting_dir / "Export" / "CLI_Batch_test"
        batch_dir.mkdir(parents=True)

        standing_file = self.tmp_path / "cli_user_context.md"
        with patch.object(cli_launcher, "STANDING_CONTEXT_PATH", standing_file):
            save_standing_cli_context("Company rule: always inspect roaming & Rel-18 impacts.")
            save_workspace_user_notes(t1_dir, "Points 2 and 4 in S2-2603001 are critical for DT.")
            save_workspace_user_notes(batch_dir, "Compare overall direction across both contributions.")

            ctx_path = write_cli_context_file(
                batch_dir,
                task_mode="compare",
                tdocs_meta=[{"TDoc": "S2-2603001", "Title": "Roam"}, {"TDoc": "S2-2603002", "Title": "Alt"}],
                meeting_dir=meeting_dir,
            )
            text = ctx_path.read_text(encoding="utf-8")
            self.assertIn("Delegate & Company Priorities (Must Follow)", text)
            self.assertIn("Company rule: always inspect roaming & Rel-18 impacts.", text)
            self.assertIn("Compare overall direction across both contributions.", text)
            self.assertIn("Points 2 and 4 in S2-2603001 are critical for DT.", text)

            # Regenerating _CLI_CONTEXT.md must never clobber _USER_NOTES.md
            write_cli_context_file(t1_dir, task_mode="summarize", tdocs_meta=[{"TDoc": "S2-2603001"}])
            self.assertEqual(load_workspace_user_notes(t1_dir), "Points 2 and 4 in S2-2603001 are critical for DT.")
            self.assertTrue((t1_dir / USER_NOTES_FILENAME).is_file())

    def test_cli_context_editor_dialog_save_and_launch(self):
        from core.ai import cli_launcher
        from core.ui.cli_context_dialog import CliContextEditorDialog

        tdoc_dir = self.tmp_path / "S2-2604000"
        tdoc_dir.mkdir(parents=True)

        standing_file = self.tmp_path / "cli_user_context.md"
        cfg_file = self.tmp_path / "cli_agents_config.json"

        with patch.object(cli_detector, "CONFIG_PATH", cfg_file), \
             patch.object(cli_launcher, "STANDING_CONTEXT_PATH", standing_file):
            dlg = CliContextEditorDialog(
                agent_id="codex",
                scope_label="S2-2604000",
                target_dir=tdoc_dir,
                home_dir=self.tmp_path,
            )
            dlg.workspace_notes_edit.setPlainText("Focus on section 5.3 urgency.")
            dlg.standing_edit.setPlainText("We are Deutsche Telekom.")
            idx_m = dlg.model_combo.findData("gpt-6.1-sol")
            if idx_m >= 0:
                dlg.model_combo.setCurrentIndex(idx_m)
            idx_e = dlg.effort_combo.findData("low")
            if idx_e >= 0:
                dlg.effort_combo.setCurrentIndex(idx_e)

            dlg._on_save_and_launch()
            res = dlg.get_result()
            self.assertTrue(res.launch_requested)
            self.assertEqual(res.model, "gpt-6.1-sol")
            self.assertEqual(res.reasoning_effort, "low")
            self.assertEqual(cli_launcher.load_workspace_user_notes(tdoc_dir), "Focus on section 5.3 urgency.")
            self.assertEqual(cli_launcher.load_standing_cli_context(), "We are Deutsche Telekom.")
            self.assertEqual(cli_detector.get_cli_agent_model_config("codex"), ("gpt-6.1-sol", "low"))
            dlg.close()


if __name__ == "__main__":
    unittest.main()
