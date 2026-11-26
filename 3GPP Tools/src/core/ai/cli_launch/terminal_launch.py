"""PowerShell command quoting/encoding and terminal launch helpers for external LLM CLIs."""

from __future__ import annotations

import base64
import logging
import os
import shutil
import subprocess
import sys
from pathlib import Path
from typing import Any, Callable, Dict, Optional

from core.ai.cli_detector import (
    CLI_AGENT_SPECS,
    CliAgentStatus,
    build_interactive_cli_args,
    detect_cli_agents,
)
from core.ai.cli_launch.context_writer import DEFAULT_BOOTSTRAP_PROMPT
from core.utils.paths import strip_extended_path

logger = logging.getLogger(__name__)


def _resolve_detect_cli_agents():
    """Preserve unittest.mock.patch hooks targeting ``core.ai.cli_launcher.detect_cli_agents``."""
    facade = sys.modules.get("core.ai.cli_launcher")
    if facade is not None and hasattr(facade, "detect_cli_agents"):
        return facade.detect_cli_agents
    return detect_cli_agents


def _ps_quote(value: str) -> str:
    """Single-quote a string safely for PowerShell `-Command`."""
    return "'" + str(value).replace("'", "''") + "'"


def build_powershell_launch_command(
    agent_id: str,
    executable: str,
    working_dir: Path,
    bootstrap_prompt: str = DEFAULT_BOOTSTRAP_PROMPT,
    *,
    model: Optional[str] = None,
    reasoning_effort: Optional[str] = None,
    stage_clipboard_prompt: Optional[str] = None,
) -> str:
    """
    Construct a quote-safe PowerShell command string that moves into `working_dir`,
    prepends `executable`'s parent directory to `$env:PATH`, optionally stages a prompt
    on the Windows clipboard, and invokes the CLI in interactive mode.
    """
    clean_dir = strip_extended_path(working_dir)
    clean_exe = strip_extended_path(executable)
    exe_dir = str(Path(clean_exe).parent)
    cli_args = build_interactive_cli_args(
        agent_id,
        clean_exe,
        bootstrap_prompt,
        model=model,
        reasoning_effort=reasoning_effort,
    )

    quoted_exe = _ps_quote(cli_args[0])
    quoted_rest = " ".join(_ps_quote(a) for a in cli_args[1:])
    invoke_expr = f"& {quoted_exe}" + (f" {quoted_rest}" if quoted_rest else "")

    parts = [f"Set-Location -LiteralPath {_ps_quote(clean_dir)}"]
    if exe_dir:
        parts.append(f"$env:PATH = {_ps_quote(exe_dir)} + ';' + $env:PATH")
    staged = (stage_clipboard_prompt or "").strip()
    if staged:
        preview = staged.replace("\r", " ").replace("\n", " ")
        if len(preview) > 140:
            preview = preview[:137] + "..."
        parts.append(f"try {{ Set-Clipboard -Value {_ps_quote(staged)} }} catch {{}}")
        parts.append(
            "Write-Host '[3GPP Tools] Prompt copied to clipboard (press Ctrl+V to paste when ready):' -ForegroundColor Cyan"
        )
        parts.append(f"Write-Host {_ps_quote('  ' + preview)} -ForegroundColor DarkGray")
    parts.append(invoke_expr)
    return "; ".join(parts)


def encode_powershell_command(ps_command: str) -> str:
    """
    Encode a PowerShell command string as UTF-16LE Base64 for `-EncodedCommand`.

    Using `-EncodedCommand` eliminates all semicolons (`;`), quotes, spaces, and
    `&` operators from the outer command line so `wt.exe` (which treats unescaped
    `;` as a Windows Terminal subcommand separator) never splits the command.
    """
    return base64.b64encode(ps_command.encode("utf-16-le")).decode("ascii")


def _show_warning(parent: Any, title: str, msg: str) -> None:
    try:
        from PyQt5.QtWidgets import QApplication, QMessageBox

        if QApplication.instance() is not None:
            QMessageBox.warning(parent, title, msg)
            return
    except Exception:
        pass
    logger.warning("%s: %s", title, msg)


def _show_critical(parent: Any, title: str, msg: str) -> None:
    try:
        from PyQt5.QtWidgets import QApplication, QMessageBox

        if QApplication.instance() is not None:
            QMessageBox.critical(parent, title, msg)
            return
    except Exception:
        pass
    logger.error("%s: %s", title, msg)


def launch_cli_in_terminal(
    agent_id: str,
    working_dir: Path,
    bootstrap_prompt: Optional[str] = None,
    parent: Optional[Any] = None,
    *,
    model: Optional[str] = None,
    reasoning_effort: Optional[str] = None,
    auto_execute: bool = False,
    statuses: Optional[Dict[str, CliAgentStatus]] = None,
    popen_fn: Optional[Callable[..., Any]] = None,
    which_fn: Optional[Callable[[str], Optional[str]]] = None,
) -> bool:
    """
    Open Windows Terminal (`wt.exe`) or PowerShell (`powershell.exe`) in `working_dir`
    and launch the requested interactive LLM CLI (`codex`, `claude`, or `gemini`).

    When ``auto_execute`` is ``False`` (default), the CLI starts idle with zero
    automatic prompt submission (preventing unintended token burn), and any
    prepared prompt is staged onto the Windows clipboard for ``Ctrl+V``.
    """
    clean_dir = Path(strip_extended_path(working_dir))
    if not clean_dir.is_dir():
        _show_warning(parent, "Directory Not Found", f"The target directory does not exist:\n{clean_dir}")
        return False

    resolved_statuses = statuses if statuses is not None else _resolve_detect_cli_agents()()
    agent_key = (agent_id or "").strip().lower()
    st = resolved_statuses.get(agent_key)

    if not st or not st.installed or not st.executable_path:
        spec = CLI_AGENT_SPECS.get(agent_key)
        hint = f"\n\nInstall hint:\n{spec.install_hint}" if spec else ""
        _show_warning(
            parent,
            "CLI Agent Not Installed",
            f"{spec.label if spec else agent_id} was not found on this machine.{hint}",
        )
        return False

    raw_prompt = bootstrap_prompt if bootstrap_prompt is not None else DEFAULT_BOOTSTRAP_PROMPT
    if auto_execute:
        exec_prompt = raw_prompt
        staged_prompt = None
    else:
        exec_prompt = ""
        staged_prompt = raw_prompt.strip() or None

    ps_command = build_powershell_launch_command(
        agent_id=agent_key,
        executable=st.executable_path,
        working_dir=clean_dir,
        bootstrap_prompt=exec_prompt,
        model=model,
        reasoning_effort=reasoning_effort,
        stage_clipboard_prompt=staged_prompt,
    )
    encoded_cmd = encode_powershell_command(ps_command)

    spawn = popen_fn or subprocess.Popen
    which_lookup = which_fn or shutil.which

    wt_exe = which_lookup("wt.exe") or which_lookup("wt")
    if wt_exe:
        try:
            spawn(
                [wt_exe, "-d", str(clean_dir), "powershell.exe", "-NoExit", "-EncodedCommand", encoded_cmd],
                cwd=str(clean_dir),
            )
            logger.info("Launched %s via Windows Terminal in %s", st.label, clean_dir)
            return True
        except FileNotFoundError:
            pass
        except OSError as exc:
            logger.warning("wt.exe launch failed (%s); falling back to powershell.exe", exc)

    try:
        kwargs: Dict[str, Any] = {"cwd": str(clean_dir)}
        if os.name == "nt" and hasattr(subprocess, "CREATE_NEW_CONSOLE"):
            kwargs["creationflags"] = subprocess.CREATE_NEW_CONSOLE
        spawn(["powershell.exe", "-NoExit", "-EncodedCommand", encoded_cmd], **kwargs)
        logger.info("Launched %s via PowerShell in %s", st.label, clean_dir)
        return True
    except OSError as exc:
        _show_critical(parent, "Terminal Launch Error", f"Could not launch {st.label}:\n{exc}")
        return False
