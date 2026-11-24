"""Static Tools, Word Workbenches, and Shell/System commands for the Ctrl+K palette."""

from __future__ import annotations

from typing import List

from core.ui.palette.models import CommandItem
from core.ui.right_side_panel import MODE_ASSISTANT, MODE_SPEC_SEARCH


def build_static_command_items(window: object) -> List[CommandItem]:
    """Assemble Tools, Word Workbenches, and Shell/Action commands for ``window``."""
    items: List[CommandItem] = []

    def _switch_tab(idx: int) -> None:
        tabs = getattr(window, "tabs", None)
        if tabs is not None and 0 <= idx < tabs.count():
            tabs.setCurrentIndex(idx)

    # 1. Core Tools (Ctrl+1..Ctrl+9)
    tools = [
        ("tool:plantuml", "PlantUML Studio", "Sequence diagrams, SVG live preview & AI generator", "Ctrl+1", "code", 0),
        ("tool:visio", "Visio & Office Studio", "Batch .puml conversion, EMF bridge & 3GPP stencils", "Ctrl+2", "layers", 1),
        ("tool:word", "Word Tools", "Compare docs, Markdown studio, CR/pCR & Sanitizer", "Ctrl+3", "file-text", 2),
        ("tool:specs", "Specifications Archive", "3GPP TS/TR browser, release downloads & change history", "Ctrl+4", "book-open", 3),
        ("tool:spec_search", "Spec Search", "Full-text FTS5 clause search & version comparison", "Ctrl+5", "search", 4),
        ("tool:protocols", "Protocol Explorer", "NAS, RRC, NGAP & PFCP evolution matrix", "Ctrl+6", "network", 5),
        ("tool:meetings", "Meetings & TDocs", "Agenda browser, TDocs, agreements & Session Plan", "Ctrl+7", "calendar", 6),
        ("tool:work_items", "Work Items Tracker", "Feature & WI search, tracking & linked specs", "Ctrl+8", "clipboard-list", 7),
        ("tool:assistant", "3GPP Research Assistant", "Local evidence-grounded Ollama AI assistant", "Ctrl+9", "sparkles", 8),
    ]
    for cid, title, sub, sc, icon, idx in tools:
        items.append(
            CommandItem(
                id=cid,
                category="Tools",
                title=title,
                subtitle=sub,
                shortcut=sc,
                icon_name=icon,
                callback=lambda i=idx: _switch_tab(i),
                keywords=f"tab switch goto {title} {sub}",
            )
        )

    # 2. Side Panels & Shell Actions
    def _dock(mode: str) -> None:
        if hasattr(window, "open_right_panel"):
            window.open_right_panel(mode)

    def _open_word_workbench(wb_key: str) -> None:
        _switch_tab(2)
        wt = getattr(window, "word_tab", None)
        if wt is not None and hasattr(wt, "open_workbench"):
            wt.open_workbench(wb_key)

    items.extend([
        CommandItem(
            id="action:dock_assistant",
            category="Side Panels & Actions",
            title="Dock Assistant Side Panel",
            subtitle="Open compact Assistant beside current workspace",
            shortcut="Ctrl+Shift+A",
            icon_name="sparkles",
            callback=lambda: _dock(MODE_ASSISTANT),
            keywords="side panel right dock ai assistant",
        ),
        CommandItem(
            id="action:dock_search",
            category="Side Panels & Actions",
            title="Dock Quick Spec Search",
            subtitle="Open compact clause search beside current workspace",
            shortcut="Ctrl+Shift+F",
            icon_name="search",
            callback=lambda: _dock(MODE_SPEC_SEARCH),
            keywords="side panel right dock quick spec search",
        ),
        CommandItem(
            id="action:toggle_drawer",
            category="Side Panels & Actions",
            title="Toggle Console & Queue Drawer",
            subtitle="Show or collapse bottom system output and task queue",
            shortcut="Ctrl+J",
            icon_name="terminal",
            callback=lambda: getattr(window, "toggle_bottom_drawer", lambda: None)(),
            keywords="logs output queue drawer console",
        ),
        CommandItem(
            id="action:toggle_rail",
            category="Side Panels & Actions",
            title="Toggle Navigation Sidebar",
            subtitle="Expand or collapse left icon navigation rail",
            shortcut="Ctrl+\\",
            icon_name="panel-left-close",
            callback=lambda: getattr(getattr(window, "nav_rail", None), "toggle_collapsed", lambda: None)(),
            keywords="sidebar nav rail collapse expand",
        ),
    ])

    # 3. Word Workbenches (7)
    workbenches = [
        ("wb:compare", "Word: Compare Documents", "Native Track Changes Word COM diff", "git-compare", "compare"),
        ("wb:markdown", "Word: Markdown <-> Word", "Bidirectional docx/pptx/pdf to Markdown", "file-code", "markdown"),
        ("wb:cr_pcr", "Word: 3GPP CR & pCR Workbench", "Draft Change Request / Pseudo-CR contributions", "file-plus", "cr_pcr"),
        ("wb:sanitizer", "Word: Sanitizer & Revisions", "Proofing language & Track Changes author cleanup", "sliders", "sanitizer"),
        ("wb:excerpt", "Word: Slice & Excerpt", "Clause slicing & heading tree extraction", "scissors", "excerpt"),
        ("wb:converter", "Word: Format Converter", "DOC / DOCX / PDF / HTML conversion", "refresh-cw", "converter"),
        ("wb:visio", "Word: Visio & Media Extractor", "Extract embedded Visio & image assets", "image", "visio"),
    ]
    for cid, title, sub, icon, wb_key in workbenches:
        items.append(
            CommandItem(
                id=cid,
                category="Word Workbenches",
                title=title,
                subtitle=sub,
                icon_name=icon,
                callback=lambda k=wb_key: _open_word_workbench(k),
                keywords=f"word workbench {title} {sub}",
            )
        )

    # 4. Sync & System Dialogs
    def _sync_meetings() -> None:
        _switch_tab(6)
        mt = getattr(window, "meetings_tab", None)
        if mt is not None and hasattr(mt, "update_db_requested"):
            mt.update_db_requested.emit(True, True, False)

    def _sync_work_items() -> None:
        _switch_tab(7)
        wt = getattr(window, "work_items_tab", None)
        if wt is not None and hasattr(wt, "sync_btn"):
            wt.sync_btn.click()

    sys_actions = [
        ("sys:sync_meetings", "Sync 3GPP Meetings Database", "Fetch latest WG meetings and TDocs metadata", "refresh-cw", _sync_meetings),
        ("sys:sync_wi", "Sync 3GPP Work Items", "Refresh Work Items catalog from portal", "refresh-cw", _sync_work_items),
        ("sys:ollama", "Ollama AI Settings", "Configure local model host, timeout & model selection", "sparkles", getattr(window, "open_ollama_settings", None)),
        ("sys:proxy", "Proxy Settings", "Configure HTTP/HTTPS proxy & corporate bypass", "shield", getattr(window, "open_proxy_settings", None)),
        ("sys:db_maint", "Database Maintenance", "Vacuum, inspect & repair SQLite databases", "database", getattr(window, "open_db_maintenance", None)),
        ("sys:task_mgr", "COM Task Manager", "Inspect Word/Visio/PPT background processes", "activity", getattr(window, "open_task_manager", None)),
        ("sys:java", "Java & PlantUML Engine", "Verify Java runtime & PlantUML JAR status", "cpu", getattr(window, "open_java_maintenance", None)),
        ("sys:mcp", "MCP Server & AI Cockpit", "Inspect stdio MCP tools and agent activity", "terminal", getattr(window, "open_mcp_monitor", None)),
        ("sys:display_settings", "Display & Text Size", "Configure global font size presets (90%–140%) & readability scale", "settings", getattr(window, "open_display_settings", None)),
        ("sys:world_clock", "World Clock & Meeting Venues", "Laptop time, UTC, active 3GPP venue clocks, WG filter & overrides", "clock", getattr(window, "open_world_clock", None)),
        ("sys:check_updates", "Check for App Updates", "Check GitHub master for new changes and update", "download", getattr(window, "open_app_updater", None)),
    ]
    for cid, title, sub, icon, fn in sys_actions:
        if fn is None:
            continue
        items.append(
            CommandItem(
                id=cid,
                category="Side Panels & Actions",
                title=title,
                subtitle=sub,
                icon_name=icon,
                callback=fn,
                keywords=f"system settings dialog {title} {sub}",
            )
        )

    restart_fn = getattr(window, "restart_application", None)
    if callable(restart_fn):
        items.append(
            CommandItem(
                id="sys:restart_app",
                category="Side Panels & Actions",
                title="Restart Application",
                subtitle="Save workspace state and relaunch 3GPP Delegate Tools",
                shortcut="Ctrl+Shift+R",
                icon_name="rotate-ccw",
                callback=restart_fn,
                keywords="system settings restart reload reboot relaunch app",
            )
        )

    return items
