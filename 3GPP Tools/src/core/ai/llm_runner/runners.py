"""Unified text-generation runners for local Ollama and headless external CLIs."""

from __future__ import annotations

import json
import logging
import os
import subprocess
import tempfile
import threading
import time
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Tuple

from core.ai.cli_detect import (
    CLI_AGENT_SPECS,
    detect_cli_agents,
    discover_cli_models,
)
from core.ai.ollama import OllamaClient, load_ollama_config

from .models import ModelSpec, PROVIDER_META, ProviderModelEntry, RunnerResult

logger = logging.getLogger(__name__)

ChunkCallback = Callable[[str], None]


class CancellationFlag:
    """Thread-safe cooperative cancellation token shared by UI workers and CLI runners."""

    def __init__(self) -> None:
        self._event = threading.Event()

    def cancel(self) -> None:
        self._event.set()

    def is_cancelled(self) -> bool:
        return self._event.is_set()


def _classify_cli_error(provider: str, rc: int, stderr_text: str, stdout_text: str = "") -> str:
    combined = f"{stderr_text}\n{stdout_text}".strip()
    lower = combined.lower()
    if "refresh token" in lower or "re-authenticate" in lower or "log out and sign in" in lower or "not logged in" in lower or "unauthorized" in lower:
        login_hint = {
            "codex": "run 'codex login'",
            "claude": "run 'claude auth login' or verify Bedrock credentials",
            "gemini": "run 'gemini' to sign in",
        }.get(provider, "re-authenticate the CLI")
        return f"{provider} authentication expired or failed ({login_hint}). Details: {combined[:300]}"
    if "model" in lower and ("not found" in lower or "unsupported" in lower or "invalid" in lower):
        return f"{provider} rejected the selected model/effort: {combined[:300]}"
    snippet = combined[:500] if combined else f"process exited with code {rc}"
    return f"{provider} CLI failed (exit {rc}): {snippet}"


def _terminate_process_tree(proc: subprocess.Popen) -> None:
    try:
        if os.name == "nt":
            subprocess.run(
                ["taskkill", "/F", "/T", "/PID", str(proc.pid)],
                capture_output=True,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
                timeout=5,
            )
            return
    except Exception:
        pass
    try:
        proc.terminate()
        proc.wait(timeout=2)
    except Exception:
        try:
            proc.kill()
        except Exception:
            pass


def _popen_kwargs_hidden() -> Dict[str, Any]:
    kwargs: Dict[str, Any] = {}
    if os.name == "nt":
        kwargs["creationflags"] = getattr(subprocess, "CREATE_NO_WINDOW", 0x08000000)
    return kwargs


def build_headless_cli_command(
    provider: str,
    executable: str,
    *,
    model: str = "",
    reasoning_effort: str = "",
    system_prompt: str = "",
    work_dir: Optional[Path] = None,
    output_file: Optional[Path] = None,
) -> Tuple[List[str], bool]:
    """
    Build non-interactive headless subprocess argv for `codex`, `claude`, or `gemini`.

    Returns `(argv, system_included_in_flag)`. When `system_included_in_flag` is False,
    the caller should prepend `system_prompt` into the stdin payload.
    """
    prov = (provider or "").strip().lower()
    clean_model = (model or "").strip()
    clean_effort = (reasoning_effort or "").strip().lower()
    clean_sys = (system_prompt or "").strip()

    if prov == "codex":
        cmd = [executable, "exec"]
        if work_dir is not None:
            cmd.extend(["-C", str(work_dir)])
        if clean_model:
            cmd.extend(["-m", clean_model])
        if clean_effort:
            cmd.extend(["-c", f'model_reasoning_effort="{clean_effort}"'])
        cmd.extend([
            "-c", "mcp_servers={}",
            "--ephemeral",
            "--skip-git-repo-check",
            "--sandbox", "read-only",
            "--color", "never",
            "--json",
        ])
        if output_file is not None:
            cmd.extend(["-o", str(output_file)])
        cmd.append("-")
        return cmd, False

    if prov == "claude":
        cmd = [
            executable,
            "-p",
            "--no-session-persistence",
            "--strict-mcp-config",
            "--disable-slash-commands",
            "--permission-mode", "dontAsk",
            "--tools=",
            "--output-format", "text",
        ]
        if clean_model:
            cmd.extend(["--model", clean_model])
        if clean_effort:
            cmd.extend(["--effort", clean_effort])
        sys_via_flag = bool(clean_sys) and len(clean_sys) <= 3500
        if sys_via_flag:
            cmd.extend(["--system-prompt", clean_sys])
        return cmd, sys_via_flag

    if prov == "gemini":
        cmd = [executable, "--output-format", "text"]
        if clean_model:
            cmd.extend(["-m", clean_model])
        return cmd, False

    raise ValueError(f"Unsupported CLI provider: {provider}")


class OllamaTextRunner:
    """Runs streaming text generation against local Ollama via `OllamaClient`."""

    def __init__(self, client: Optional[OllamaClient] = None) -> None:
        self.client = client or OllamaClient()

    def run(
        self,
        system_prompt: str,
        user_prompt: str,
        spec: ModelSpec,
        *,
        cancel_flag: Optional[CancellationFlag] = None,
        on_chunk: Optional[ChunkCallback] = None,
        num_ctx: int = 32768,
        temperature: float = 0.15,
    ) -> RunnerResult:
        t0 = time.perf_counter()
        model = spec.model.strip()
        if not model:
            cfg = load_ollama_config()
            model = str(cfg.get("selected_model") or "").strip()
        if not model:
            online, models, err = self.client.ping_and_get_models()
            if not online:
                return RunnerResult(
                    ok=False,
                    error=err or "Ollama server is offline.",
                    provider="ollama",
                    elapsed_s=time.perf_counter() - t0,
                )
            model = models[0] if models else ""
        if not model:
            return RunnerResult(
                ok=False,
                error="No Ollama model available.",
                provider="ollama",
                elapsed_s=time.perf_counter() - t0,
            )

        messages = []
        if system_prompt.strip():
            messages.append({"role": "system", "content": system_prompt.strip()})
        messages.append({"role": "user", "content": user_prompt})

        chunks: List[str] = []
        try:
            for delta in self.client.stream_chat(
                model=model,
                messages=messages,
                options={"temperature": temperature, "num_ctx": num_ctx},
            ):
                if cancel_flag and cancel_flag.is_cancelled():
                    return RunnerResult(
                        ok=False,
                        text="".join(chunks).strip(),
                        error="Cancelled by user.",
                        provider="ollama",
                        model=model,
                        elapsed_s=time.perf_counter() - t0,
                        cancelled=True,
                    )
                if delta:
                    chunks.append(delta)
                    if on_chunk:
                        on_chunk(delta)
            return RunnerResult(
                ok=True,
                text="".join(chunks).strip(),
                provider="ollama",
                model=model,
                elapsed_s=time.perf_counter() - t0,
            )
        except Exception as exc:
            return RunnerResult(
                ok=False,
                text="".join(chunks).strip(),
                error=f"Ollama request failed: {exc}",
                provider="ollama",
                model=model,
                elapsed_s=time.perf_counter() - t0,
            )


class CliSubprocessRunner:
    """Runs non-interactive text generation through `codex`, `claude`, or `gemini` CLIs."""

    def run(
        self,
        system_prompt: str,
        user_prompt: str,
        spec: ModelSpec,
        *,
        cancel_flag: Optional[CancellationFlag] = None,
        on_chunk: Optional[ChunkCallback] = None,
        timeout_s: float = 300.0,
        popen_fn: Any = None,
        statuses: Optional[Dict[str, Any]] = None,
    ) -> RunnerResult:
        t0 = time.perf_counter()
        prov = spec.provider.strip().lower()
        resolved_statuses = statuses if statuses is not None else detect_cli_agents()
        st = resolved_statuses.get(prov)
        if not st or not st.installed or not st.executable_path:
            cli_spec = CLI_AGENT_SPECS.get(prov)
            hint = f" Install: {cli_spec.install_hint}" if cli_spec else ""
            return RunnerResult(
                ok=False,
                error=f"{prov} CLI is not installed or not found on PATH.{hint}",
                provider=prov,
                model=spec.model,
                reasoning_effort=spec.reasoning_effort,
                elapsed_s=time.perf_counter() - t0,
            )

        spawn = popen_fn or subprocess.Popen
        with tempfile.TemporaryDirectory(prefix="3gpp_llm_") as tmp_dir:
            tmp_path = Path(tmp_dir)
            out_file = tmp_path / "last_message.txt" if prov == "codex" else None
            cmd, sys_in_flag = build_headless_cli_command(
                prov,
                st.executable_path,
                model=spec.model,
                reasoning_effort=spec.reasoning_effort,
                system_prompt=system_prompt,
                work_dir=tmp_path,
                output_file=out_file,
            )
            if sys_in_flag:
                stdin_payload = user_prompt
            elif system_prompt.strip():
                stdin_payload = (
                    f"<system_instructions>\n{system_prompt.strip()}\n</system_instructions>\n\n"
                    f"{user_prompt}"
                )
            else:
                stdin_payload = user_prompt

            try:
                proc = spawn(
                    cmd,
                    stdin=subprocess.PIPE,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    text=True,
                    encoding="utf-8",
                    errors="replace",
                    cwd=str(tmp_path),
                    **_popen_kwargs_hidden(),
                )
            except OSError as exc:
                return RunnerResult(
                    ok=False,
                    error=f"Failed to launch {prov} CLI: {exc}",
                    provider=prov,
                    model=spec.model,
                    reasoning_effort=spec.reasoning_effort,
                    elapsed_s=time.perf_counter() - t0,
                )

            stderr_lines: List[str] = []
            stderr_thread = threading.Thread(
                target=lambda: stderr_lines.extend(proc.stderr.readlines() if proc.stderr else []),
                daemon=True,
            )
            stderr_thread.start()

            try:
                if proc.stdin:
                    proc.stdin.write(stdin_payload)
                    proc.stdin.close()
            except Exception:
                pass

            stdout_chunks: List[str] = []
            agent_messages: List[str] = []
            usage: Dict[str, Any] = {}

            def _read_stdout() -> None:
                if not proc.stdout:
                    return
                for raw_line in proc.stdout:
                    stdout_chunks.append(raw_line)
                    if prov == "codex":
                        stripped = raw_line.strip()
                        if not stripped:
                            continue
                        try:
                            ev = json.loads(stripped)
                        except Exception:
                            continue
                        ev_type = ev.get("type")
                        if ev_type == "item.completed":
                            item = ev.get("item") or {}
                            if item.get("type") == "agent_message" and item.get("text"):
                                txt = str(item["text"])
                                agent_messages.append(txt)
                                if on_chunk:
                                    on_chunk(txt)
                        elif ev_type == "turn.completed" and isinstance(ev.get("usage"), dict):
                            usage.update(ev["usage"])
                        elif ev_type in ("error", "turn.failed"):
                            err_obj = ev.get("error") or ev
                            msg = str(err_obj.get("message") or stripped)
                            stderr_lines.append(msg + "\n")
                    else:
                        if on_chunk:
                            on_chunk(raw_line)

            reader_thread = threading.Thread(target=_read_stdout, daemon=True)
            reader_thread.start()

            deadline = time.monotonic() + float(timeout_s)
            timed_out = False
            while proc.poll() is None:
                if cancel_flag and cancel_flag.is_cancelled():
                    _terminate_process_tree(proc)
                    reader_thread.join(timeout=1.0)
                    stderr_thread.join(timeout=1.0)
                    return RunnerResult(
                        ok=False,
                        error="Cancelled by user.",
                        provider=prov,
                        model=spec.model,
                        reasoning_effort=spec.reasoning_effort,
                        elapsed_s=time.perf_counter() - t0,
                        cancelled=True,
                    )
                if time.monotonic() >= deadline:
                    timed_out = True
                    _terminate_process_tree(proc)
                    break
                time.sleep(0.05)

            reader_thread.join(timeout=2.0)
            stderr_thread.join(timeout=2.0)
            rc = proc.returncode if proc.returncode is not None else -1
            stderr_txt = "".join(stderr_lines).strip()

            if timed_out:
                return RunnerResult(
                    ok=False,
                    error=f"{prov} CLI timed out after {int(timeout_s)}s.",
                    provider=prov,
                    model=spec.model,
                    reasoning_effort=spec.reasoning_effort,
                    elapsed_s=time.perf_counter() - t0,
                )

            final_text = ""
            if prov == "codex":
                if out_file is not None and out_file.is_file():
                    final_text = out_file.read_text(encoding="utf-8", errors="replace").strip()
                if not final_text and agent_messages:
                    final_text = "\n\n".join(agent_messages).strip()
            else:
                final_text = "".join(stdout_chunks).strip()

            if rc != 0 or not final_text:
                return RunnerResult(
                    ok=False,
                    text=final_text,
                    error=_classify_cli_error(prov, rc, stderr_txt, "".join(stdout_chunks)),
                    provider=prov,
                    model=spec.model,
                    reasoning_effort=spec.reasoning_effort,
                    elapsed_s=time.perf_counter() - t0,
                    usage=usage,
                )

            return RunnerResult(
                ok=True,
                text=final_text,
                provider=prov,
                model=spec.model,
                reasoning_effort=spec.reasoning_effort,
                elapsed_s=time.perf_counter() - t0,
                usage=usage,
            )


def discover_all_runner_models(
    *,
    include_ollama: bool = True,
    force_refresh: bool = False,
) -> Dict[str, List[ProviderModelEntry]]:
    """Return available models across Ollama, Codex CLI, Claude CLI, and Gemini CLI."""
    grouped: Dict[str, List[ProviderModelEntry]] = {}

    if include_ollama:
        label, icon = PROVIDER_META["ollama"]
        try:
            online, models, _ = OllamaClient().ping_and_get_models()
        except Exception:
            online, models = False, []
        if online and models:
            grouped["ollama"] = [
                ProviderModelEntry(
                    provider="ollama",
                    provider_label=label,
                    provider_icon=icon,
                    model_slug=m,
                    display_name=m,
                    reasoning_levels=(),
                    default_effort="",
                    available=True,
                )
                for m in models
            ]
        else:
            grouped["ollama"] = []

    statuses = detect_cli_agents(force_refresh=force_refresh)
    for aid in ("codex", "claude", "gemini"):
        st = statuses.get(aid)
        label, icon = PROVIDER_META[aid]
        if not st or not st.installed:
            grouped[aid] = []
            continue
        models, _def_m, _def_e = discover_cli_models(aid)
        grouped[aid] = [
            ProviderModelEntry(
                provider=aid,
                provider_label=label,
                provider_icon=icon,
                model_slug=m.slug,
                display_name=m.display_name,
                reasoning_levels=m.reasoning_levels,
                default_effort=m.default_reasoning_level,
                available=True,
            )
            for m in models
        ]
    return grouped


def run_llm_text(
    system_prompt: str,
    user_prompt: str,
    spec: ModelSpec,
    *,
    cancel_flag: Optional[CancellationFlag] = None,
    on_chunk: Optional[ChunkCallback] = None,
    timeout_s: float = 300.0,
) -> RunnerResult:
    """Execute a unified text-generation request on Ollama or an installed external CLI."""
    norm = spec.normalized()
    if norm.provider == "ollama":
        return OllamaTextRunner().run(
            system_prompt,
            user_prompt,
            norm,
            cancel_flag=cancel_flag,
            on_chunk=on_chunk,
        )
    return CliSubprocessRunner().run(
        system_prompt,
        user_prompt,
        norm,
        cancel_flag=cancel_flag,
        on_chunk=on_chunk,
        timeout_s=timeout_s,
    )
