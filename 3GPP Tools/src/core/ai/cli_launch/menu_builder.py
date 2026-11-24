"""Qt QMenu builders for external LLM CLI agents, model selectors, and MCP setup."""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Any, Callable, Dict, Optional

from core.ai.cli_detector import (
    CliAgentStatus,
    discover_cli_models,
    format_active_model_badge,
    get_cli_agent_model_config,
    set_cli_agent_model_config,
)
from core.ai.cli_launch.terminal_launch import _resolve_detect_cli_agents

logger = logging.getLogger(__name__)


def open_mcp_agent_setup_dialog(parent: Any = None) -> None:
    """Open `McpMonitorDialog` directly on the AI Agent Client Setup tab."""
    try:
        from core.ui.mcp_monitor_dialog import McpMonitorDialog

        dlg = McpMonitorDialog(parent=parent)
        if hasattr(dlg, "tabs") and hasattr(dlg, "setup_tab"):
            dlg.tabs.setCurrentWidget(dlg.setup_tab)
        dlg.show()
        dlg.raise_()
        dlg.activateWindow()
    except Exception as exc:
        logger.warning("Could not open MCP & AI Agent Cockpit: %s", exc)


def add_model_selector_submenu(
    agent_submenu: Any,
    agent_id: str,
    *,
    home_dir: Optional[Path] = None,
) -> Any:
    """
    Append a nested ``🧠 Model / Reasoning Effort ▸`` selector menu inside an agent's QMenu.
    Selecting an item updates ``cli_agents_config.json`` immediately via ``set_cli_agent_model_config``.
    """
    active_model, active_effort = get_cli_agent_model_config(agent_id)
    models, _cli_def_model, _cli_def_effort = discover_cli_models(agent_id, home_dir=home_dir)
    badge = format_active_model_badge(agent_id, home_dir=home_dir)

    model_menu = agent_submenu.addMenu(f"🧠 Model / Effort ({badge})")
    model_menu.setToolTipsVisible(True)

    act_default = model_menu.addAction("🔄 Use CLI Default")
    act_default.setCheckable(True)
    act_default.setChecked(not active_model and not active_effort)
    act_default.setToolTip("Clear any 3GPP Tools override and inherit the CLI's own configured default.")
    act_default.triggered.connect(lambda _=False, aid=agent_id: set_cli_agent_model_config(aid, "", ""))
    model_menu.addSeparator()

    for opt in models:
        is_this_model = active_model == opt.slug
        if opt.reasoning_levels:
            label = f"{opt.display_name} ({opt.slug})" if opt.display_name != opt.slug else opt.slug
            if is_this_model and active_effort:
                label = f"✓ {label} [{active_effort}]"
            m_sub = model_menu.addMenu(label)
            if opt.description:
                m_sub.setToolTip(opt.description)

            act_m_def = m_sub.addAction(
                "Default Effort" + (f" ({opt.default_reasoning_level})" if opt.default_reasoning_level else "")
            )
            act_m_def.setCheckable(True)
            act_m_def.setChecked(is_this_model and not active_effort)
            act_m_def.triggered.connect(
                lambda _=False, aid=agent_id, slug=opt.slug: set_cli_agent_model_config(aid, slug, "")
            )
            m_sub.addSeparator()
            for lvl in opt.reasoning_levels:
                act_lvl = m_sub.addAction(lvl)
                act_lvl.setCheckable(True)
                act_lvl.setChecked(is_this_model and active_effort == lvl)
                act_lvl.triggered.connect(
                    lambda _=False, aid=agent_id, slug=opt.slug, eff=lvl: set_cli_agent_model_config(aid, slug, eff)
                )
        else:
            label = f"{opt.display_name} ({opt.slug})" if opt.display_name != opt.slug else opt.slug
            act_m = model_menu.addAction(label)
            act_m.setCheckable(True)
            act_m.setChecked(is_this_model)
            act_m.triggered.connect(
                lambda _=False, aid=agent_id, slug=opt.slug: set_cli_agent_model_config(aid, slug, "")
            )

    return model_menu


def populate_cli_agent_menu(
    menu: Any,
    on_launch_callback: Callable[[str, str], None],
    *,
    has_multiple_or_revs: bool = False,
    multi_count: int = 0,
    parent: Any = None,
    statuses: Optional[Dict[str, CliAgentStatus]] = None,
) -> Any:
    """
    Populate a QMenu with detected CLI agents (`Codex CLI`, `Claude CLI`, `Gemini CLI`)
    and task presets (`interactive`, `summarize`, optional `compare`, `custom`, `edit_context`),
    quick Model / Reasoning Effort switchers, disabled entries for missing CLIs, and Cockpit link.
    """
    menu.clear()
    menu.setToolTipsVisible(True)

    resolved = statuses if statuses is not None else _resolve_detect_cli_agents()()
    installed_agents = [st for st in resolved.values() if st.installed]
    missing_agents = [st for st in resolved.values() if not st.installed]

    if installed_agents:
        for st in installed_agents:
            badge = format_active_model_badge(st.agent_id)
            sub = menu.addMenu(f"{st.icon} {st.label}  [{badge}]")
            sub.setToolTipsVisible(True)
            mcp_note = (
                "MCP configured"
                if st.mcp_configured
                else "MCP not yet registered (local Markdown + cli_tools.py ready)"
            )
            sub.setToolTip(f"{st.executable_path}\nModel: {badge}\n{mcp_note}")

            act_inter = sub.addAction("💬 Interactive Session (Read && Stand By)")
            act_inter.setToolTip(f"Prepare Markdown and start an interactive {st.label} session.")
            act_inter.triggered.connect(lambda _=False, aid=st.agent_id: on_launch_callback(aid, "interactive"))

            act_sum = sub.addAction("⚡ Summarize && Check Spec Impact")
            act_sum.setToolTip(f"Launch {st.label} with an immediate prompt to summarize mechanisms and spec impact.")
            act_sum.triggered.connect(lambda _=False, aid=st.agent_id: on_launch_callback(aid, "summarize"))

            if has_multiple_or_revs:
                cmp_label = (
                    f"⚖️ Compare {multi_count} Selected TDocs"
                    if multi_count > 1
                    else "⚖️ Compare Base vs. Revisions"
                )
                act_cmp = sub.addAction(cmp_label)
                act_cmp.setToolTip(f"Stage all relevant Markdown files and ask {st.label} to compare them.")
                act_cmp.triggered.connect(lambda _=False, aid=st.agent_id: on_launch_callback(aid, "compare"))

            sub.addSeparator()
            act_custom = sub.addAction("✏️ Ask Custom Question...")
            act_custom.setToolTip(f"Enter a custom question before launching {st.label}.")
            act_custom.triggered.connect(lambda _=False, aid=st.agent_id: on_launch_callback(aid, "custom"))

            act_ctx = sub.addAction("📝 Edit Context / Company Notes && Launch...")
            act_ctx.setToolTip(
                "Edit persistent _USER_NOTES.md / standing company guidance and select Model / Reasoning Effort before launching."
            )
            act_ctx.triggered.connect(lambda _=False, aid=st.agent_id: on_launch_callback(aid, "edit_context"))

            sub.addSeparator()
            add_model_selector_submenu(sub, st.agent_id)
    else:
        no_cli = menu.addAction("⚪ No LLM CLI Detected (codex / claude / gemini)")
        no_cli.setEnabled(False)
        no_cli.setToolTip("Install Codex CLI, Claude CLI, or Gemini CLI and click Re-scan in MCP Agent Setup.")

    if missing_agents:
        menu.addSeparator()
        for st in missing_agents:
            act_miss = menu.addAction(f"⚪ {st.label} (Not Found)")
            act_miss.setEnabled(False)
            act_miss.setToolTip(f"Not detected in PATH or user directories.\nInstall: {st.install_hint}")

    menu.addSeparator()
    act_cfg = menu.addAction("⚙️ Configure CLI Agents && MCP...")
    act_cfg.setToolTip(
        "Open the MCP Server & External AI Agent Cockpit to inspect CLI detection, default models, and MCP setup."
    )
    act_cfg.triggered.connect(lambda _=False: open_mcp_agent_setup_dialog(parent))

    return menu
