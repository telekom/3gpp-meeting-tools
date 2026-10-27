# --- File: src/core/config/useful_links.py ---
import json
import logging
import os
import webbrowser
from pathlib import Path
from typing import Callable, Dict, List, Optional

from PyQt5.QtWidgets import QAction, QMenu

from core.utils.paths import get_project_root

logger = logging.getLogger(__name__)

CONFIG_FILENAME = "useful_links.json"

DEFAULT_LINKS_CONFIG = {
    "version": 1,
    "description": "Configurable external web resources, tools, and specification viewers.",
    "categories": [
        {
            "name": "Documentation & Help",
            "links": [
                {
                    "title": "Application Documentation",
                    "url": "https://github.com/telekom/3gpp-meeting-tools/tree/master/3GPP%20Tools",
                    "icon": "📖",
                    "shortcut": "F1",
                    "tooltip": "Open official 3GPP Tools documentation, user guides, and architecture manuals.",
                }
            ],
        },
        {
            "name": "Specification Viewers & Search",
            "links": [
                {
                    "title": "3GPP Guru Viewer",
                    "url": "https://3gpp.guru/",
                    "icon": "🌐",
                    "tooltip": "Fast web-based 3GPP specification viewer with syntax highlighting and cross-clause hyperlinks.",
                },
                {
                    "title": "SpecViewer (specviewer.org)",
                    "url": "https://specviewer.org/",
                    "icon": "🔍",
                    "tooltip": "3GPP specifications by version with clause-level search and diffing.",
                },
                {
                    "title": "TDocHamster",
                    "url": "https://tdochamster.com/",
                    "icon": "🐹",
                    "tooltip": "3GPP TDoc search, comparison, and meeting contribution visualizer.",
                },
            ],
        },
        {
            "name": "Official 3GPP Portals & APIs",
            "links": [
                {
                    "title": "3GPP Forge OpenAPI Specs",
                    "url": "https://forge.3gpp.org/rep/all/apis",
                    "icon": "🛠️",
                    "tooltip": "Official 3GPP OpenAPI 3.0 / YAML specifications repository for 5G Core, NWDAF, and SBA.",
                },
                {
                    "title": "3GPP Portal (3GU)",
                    "url": "https://portal.3gpp.org/",
                    "icon": "🏛️",
                    "tooltip": "Official 3GPP Portal: meetings calendar, work items, and Change Request (CR) database.",
                },
                {
                    "title": "3GPP Work Plan & Release Timelines",
                    "url": "https://www.3gpp.org/ftp/Information/WORK_PLAN/",
                    "icon": "📅",
                    "tooltip": "Official 3GPP Work Plan, Release timelines, study item schedules, and GANTT charts.",
                },
            ],
        },
    ],
}


def get_useful_links_config_path() -> Path:
    """Return the absolute path to the useful_links.json configuration file."""
    return get_project_root() / "config" / CONFIG_FILENAME


def load_useful_links() -> dict:
    """
    Load external links from src/config/useful_links.json.
    Falls back to DEFAULT_LINKS_CONFIG if the file is missing or contains invalid JSON.
    """
    config_path = get_useful_links_config_path()
    if not config_path.exists():
        logger.info(f"Useful links config not found at {config_path}; creating default.")
        try:
            config_path.parent.mkdir(parents=True, exist_ok=True)
            with open(config_path, "w", encoding="utf-8") as f:
                json.dump(DEFAULT_LINKS_CONFIG, f, indent=2, ensure_ascii=False)
        except Exception as exc:
            logger.warning(f"Failed to persist default useful links config: {exc}")
        return DEFAULT_LINKS_CONFIG

    try:
        with open(config_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, dict) and "categories" in data:
            return data
        logger.warning(f"Invalid schema in {config_path}; using fallback defaults.")
        return DEFAULT_LINKS_CONFIG
    except Exception as exc:
        logger.error(f"Failed to parse {config_path}: {exc}; using fallback defaults.")
        return DEFAULT_LINKS_CONFIG


def populate_help_menu(
    menu: QMenu,
    on_open_url: Callable[[str, str], None],
    on_edit_config: Optional[Callable[[], None]] = None,
) -> None:
    """
    Populate a QMenu with categorized links loaded from useful_links.json.

    :param menu: The QMenu instance to populate.
    :param on_open_url: Callback with signature (url: str, title: str) -> None.
    :param on_edit_config: Optional callback to open useful_links.json in an external editor.
    """
    menu.clear()
    menu.setToolTipsVisible(True)

    config = load_useful_links()
    categories = config.get("categories", [])

    for cat_idx, category in enumerate(categories):
        cat_name = category.get("name", "External Links")
        links = category.get("links", [])
        if not links:
            continue

        # Add section header
        menu.addSection(cat_name)

        for link in links:
            title = link.get("title", "Untitled Link")
            url = link.get("url", "").strip()
            icon = link.get("icon", "").strip()
            tooltip = link.get("tooltip", "").strip()
            shortcut = link.get("shortcut", "").strip()

            if not url:
                continue

            # Build label with icon/emoji and optional shortcut display
            label = f"{icon}  {title}" if icon else title
            if shortcut:
                label = f"{label}\t{shortcut}"

            act = menu.addAction(label)
            act_tooltip = f"{tooltip}\n{url}" if tooltip else url
            act.setToolTip(act_tooltip)
            act.setStatusTip(url)

            # Bind URL and title to triggered callback
            act.triggered.connect(
                lambda checked=False, u=url, t=title: on_open_url(u, t)
            )

    # Bottom configuration and reload actions
    menu.addSeparator()

    if on_edit_config:
        act_edit = menu.addAction("⚙️ Configure Useful Links (JSON)...")
        act_edit.setToolTip("Open useful_links.json in your default text editor to add, remove, or customize links.")
        act_edit.triggered.connect(lambda: on_edit_config())

    act_reload = menu.addAction("🔄 Reload Links")
    act_reload.setToolTip("Reload external links immediately from useful_links.json.")
    act_reload.triggered.connect(
        lambda: populate_help_menu(menu, on_open_url, on_edit_config)
    )
