# --- File: tests/test_llm_runner.py ---
from __future__ import annotations

import io
import json
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch

project_root = Path(__file__).resolve().parent.parent
src_dir = project_root / "src"
if str(src_dir) not in sys.path:
    sys.path.insert(0, str(src_dir))

from core.ai.cli_detector import CliAgentStatus
from core.ai.llm_runner import (
    CancellationFlag,
    CliSubprocessRunner,
    ModelSpec,
    OllamaTextRunner,
    build_headless_cli_command,
    get_default_model_spec,
    get_m365_copilot_target,
    resolve_effective_model_spec,
    set_default_model_spec,
    set_m365_copilot_target,
)


class _FakeProc:
    def __init__(self, stdout_lines, stderr_lines=None, returncode=0, on_stdin_close=None):
        self.pid = 99999
        self._stdout_lines = list(stdout_lines)
        self._stderr_lines = list(stderr_lines or [])
        self.returncode = returncode
        self._on_stdin_close = on_stdin_close
        self.written = []
        self._polled = False

        class _Stdin:
            def __init__(self, parent):
                self._p = parent

            def write(self, data):
                self._p.written.append(data)

            def close(self):
                if self._p._on_stdin_close:
                    self._p._on_stdin_close()

        self.stdin = _Stdin(self)
        self.stdout = io.StringIO("".join(self._stdout_lines))
        self.stderr = io.StringIO("".join(self._stderr_lines))

    def poll(self):
        return self.returncode

    def terminate(self):
        pass

    def kill(self):
        pass

    def wait(self, timeout=None):
        return self.returncode


class TestLlmRunner(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.tmp_path = Path(self._tmp.name)

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def test_build_headless_cli_commands_and_flags(self):
        codex_cmd, codex_sys = build_headless_cli_command(
            "codex",
            "codex.exe",
            model="gpt-6.2-astra",
            reasoning_effort="low",
            system_prompt="SYS",
            work_dir=self.tmp_path,
            output_file=self.tmp_path / "out.txt",
        )
        self.assertFalse(codex_sys)
        self.assertEqual(codex_cmd[:2], ["codex.exe", "exec"])
        self.assertIn("-m", codex_cmd)
        self.assertIn("gpt-6.2-astra", codex_cmd)
        self.assertIn('model_reasoning_effort="low"', codex_cmd)
        self.assertIn("--ephemeral", codex_cmd)
        self.assertIn("mcp_servers={}", codex_cmd)
        self.assertEqual(codex_cmd[-1], "-")

        claude_cmd, claude_sys = build_headless_cli_command(
            "claude",
            "claude.exe",
            model="sonnet",
            reasoning_effort="high",
            system_prompt="Short system prompt",
        )
        self.assertTrue(claude_sys)
        self.assertIn("-p", claude_cmd)
        self.assertIn("--tools=", claude_cmd)
        self.assertNotIn("--tools", claude_cmd)  # must use --tools= so variadic arg never swallows next flags
        self.assertIn("--model", claude_cmd)
        self.assertIn("sonnet", claude_cmd)
        self.assertIn("--effort", claude_cmd)
        self.assertIn("high", claude_cmd)

        gemini_cmd, gemini_sys = build_headless_cli_command(
            "gemini",
            "gemini.cmd",
            model="gemini-2.5-pro",
        )
        self.assertFalse(gemini_sys)
        self.assertEqual(gemini_cmd, ["gemini.cmd", "--output-format", "text", "-m", "gemini-2.5-pro"])

    def test_config_roundtrip_leaf_and_synthesis_defaults(self):
        cfg_file = self.tmp_path / "email_summary_config.json"
        leaf = ModelSpec(provider="codex", model="gpt-6.2-astra", reasoning_effort="low")
        synth = ModelSpec(provider="claude", model="opus", reasoning_effort="high")

        set_default_model_spec("leaf", leaf, config_path=cfg_file)
        set_default_model_spec("synthesis", synth, config_path=cfg_file)

        self.assertEqual(get_default_model_spec("leaf", config_path=cfg_file), leaf)
        self.assertEqual(get_default_model_spec("synthesis", config_path=cfg_file), synth)
        self.assertEqual(resolve_effective_model_spec("leaf", config_path=cfg_file), leaf)
        override = ModelSpec(provider="ollama", model="qwen2.5:14b")
        self.assertEqual(resolve_effective_model_spec("leaf", override=override, config_path=cfg_file), override)

    def test_codex_subprocess_runner_parses_jsonl_and_output_file(self):
        jsonl = [
            json.dumps({"type": "thread.started", "thread_id": "t1"}) + "\n",
            json.dumps({"type": "item.completed", "item": {"type": "agent_message", "text": "### Summary\nOK"}}) + "\n",
            json.dumps({"type": "turn.completed", "usage": {"input_tokens": 120, "output_tokens": 15}}) + "\n",
        ]

        def fake_popen(cmd, **_kw):
            out_idx = cmd.index("-o") + 1
            Path(cmd[out_idx]).write_text("### Summary\nOK", encoding="utf-8")
            return _FakeProc(jsonl, returncode=0)

        statuses = {
            "codex": CliAgentStatus("codex", "Codex CLI", "OpenAI", "🤖", True, "codex.exe", "path", True, ""),
        }
        chunks = []
        res = CliSubprocessRunner().run(
            "SYS",
            "USR",
            ModelSpec("codex", "gpt-6.2-astra", "low"),
            on_chunk=chunks.append,
            popen_fn=fake_popen,
            statuses=statuses,
        )
        self.assertTrue(res.ok)
        self.assertEqual(res.text, "### Summary\nOK")
        self.assertEqual(res.usage.get("input_tokens"), 120)
        self.assertTrue(chunks)

    def test_cli_auth_error_is_classified_cleanly(self):
        def fake_popen(_cmd, **_kw):
            return _FakeProc([], stderr_lines=["Error: Your access token could not be refreshed because your refresh token was revoked.\n"], returncode=1)

        statuses = {
            "codex": CliAgentStatus("codex", "Codex CLI", "OpenAI", "🤖", True, "codex.exe", "path", True, ""),
        }
        res = CliSubprocessRunner().run(
            "SYS",
            "USR",
            ModelSpec("codex", "gpt-6.2-astra", "low"),
            popen_fn=fake_popen,
            statuses=statuses,
        )
        self.assertFalse(res.ok)
        self.assertIn("codex login", res.error.lower())

    def test_ollama_runner_streaming_and_cancellation(self):
        fake_client = MagicMock()
        fake_client.stream_chat.return_value = iter(["Hello ", "world!"])
        chunks = []
        res = OllamaTextRunner(client=fake_client).run(
            "SYS",
            "USR",
            ModelSpec("ollama", "qwen2.5:7b"),
            on_chunk=chunks.append,
        )
        self.assertTrue(res.ok)
        self.assertEqual(res.text, "Hello world!")
        self.assertEqual(chunks, ["Hello ", "world!"])

        # Cancellation mid-stream
        flag = CancellationFlag()

        def _gen(*_a, **_kw):
            yield "Part 1 "
            flag.cancel()
            yield "Part 2"

        fake_client.stream_chat.side_effect = _gen
        res_cancel = OllamaTextRunner(client=fake_client).run(
            "SYS",
            "USR",
            ModelSpec("ollama", "qwen2.5:7b"),
            cancel_flag=flag,
        )
        self.assertFalse(res_cancel.ok)
        self.assertTrue(res_cancel.cancelled)

    def test_m365_copilot_target_config_roundtrip(self):
        with tempfile.TemporaryDirectory() as tmp:
            cfg_p = Path(tmp) / "summary_cfg.json"
            self.assertEqual(get_m365_copilot_target(config_path=cfg_p), "web")
            set_default_model_spec("leaf", ModelSpec("ollama", "qwen2.5:7b"), config_path=cfg_p)
            set_m365_copilot_target("app", config_path=cfg_p)
            self.assertEqual(get_m365_copilot_target(config_path=cfg_p), "app")
            self.assertEqual(get_default_model_spec("leaf", config_path=cfg_p).model, "qwen2.5:7b")
            set_m365_copilot_target("invalid-target", config_path=cfg_p)
            self.assertEqual(get_m365_copilot_target(config_path=cfg_p), "web")


if __name__ == "__main__":
    unittest.main()
