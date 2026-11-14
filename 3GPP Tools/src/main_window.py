# --- File: src/main_window.py ---
import datetime
import logging
import os
import subprocess
import urllib.request
import webbrowser
from pathlib import Path

from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtGui import QKeySequence, QTextCursor
from PyQt5.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QSplitter, QStatusBar, QApplication,
    QDialog, QTabWidget, QPushButton, QShortcut, QLabel, QMessageBox, QMenu, QSizePolicy
)

from core.config.config import HELP_URL
from core.config.useful_links import populate_help_menu, get_useful_links_config_path
from core.network.session import NetworkConfigDialog, NetworkSession, ProxyProfileManager
from core.network.wifi_monitor import WifiMonitorThread
from core.queue_manager import QueueManager
from core.ui.bottom_drawer import BottomDrawer
from core.ui.command_palette import (
    CommandPaletteDialog,
    DynamicPaletteSearcher,
    build_static_command_items,
)
from core.ui.icons import lucide_icon
from core.ui.nav_rail import NavRail
from core.ui.right_side_panel import MODE_ASSISTANT, MODE_SPEC_SEARCH, RightSidePanel
from core.ui.shell_state import restore_shell_state, save_shell_state
from core.ui.theme import C
from core.ui.toast import ToastAction, ToastManager
from core.ui.ui_components import ProxyDialog, create_app_icon
from core.ui.ui_panels import (
    ConsolePanel, QueuePanel, ProcessManagerDialog, DatabaseMaintenanceDialog, GuiLogHandler
)
from core.ui.system_status_bar import SystemStatusBar
from core.utils.paths import get_project_root
from modules.meetings.ui.ui_tabs import MeetingsTab
from modules.puml2visio.config.paths import PLANTUML_JAR_NAME
from modules.puml2visio.core.live_preview import LivePreviewManager
from modules.puml2visio.core.visio_converter import VisioReaderThread
from modules.puml2visio.templates.plantuml_templates import PLANTUML_TYPES
from modules.puml2visio.ui.ui_tabs import CodeEditorTab, BatchConvertTab
from modules.puml2visio.utils.paths import get_puml2visio_asset_path
from modules.puml2visio.utils.utils import encode_plantuml, InitializationThread
from modules.specifications.ui.ui_tabs import SpecificationsTab
from modules.spec_search.ui.spec_search_tabs import SpecSearchTab
from modules.word_tools.ui.word_tabs import WordExtractorTab
from modules.work_items.ui.ui_tabs import WorkItemsTab
from modules.nas.ui.nas_tabs import NASTab
from modules.assistant.ui.assistant_tab import AssistantTab
from modules.spec_search.ui.quick_search_panel import QuickSpecSearchPanel
from core.ai.ollama_client import OllamaClient, OllamaMonitorThread
from core.ai.ollama_dialog import OllamaConfigDialog
from modules.puml2visio.ui.java_dialog import JavaMaintenanceDialog


class DragDropUI(QMainWindow):
    def __init__(self, gui_log_handler=None):
        super().__init__()
        self.setWindowTitle("3GPP Delegate Tools")
        self.setWindowIcon(create_app_icon())
        self.setMinimumSize(780, 560)
        self.resize(1160, 780)

        self.jar_path = get_puml2visio_asset_path(PLANTUML_JAR_NAME)
        self.last_out_path = ""
        self.init_thread = None
        self._last_update_check_result = None
        # Background monitors are managed by self.status_bar
        self._services_started = False  # Guard to prevent duplicate background starts
        self._visio_ready = True
        self._core_ready = True
        self._assistant_docked = False

        logging.info("🏁 [STARTUP:UI] Starting _setup_ui()...")
        self._setup_ui()
        restore_shell_state(self)
        if (
            hasattr(self, "right_panel")
            and self.right_panel.is_open
            and self.right_panel.mode == MODE_ASSISTANT
            and hasattr(self, "tabs")
            and hasattr(self, "assistant_tab")
            and self.tabs.currentWidget() is self.assistant_tab
        ):
            self.right_panel.close_panel()
        else:
            self._sync_assistant_dock_slot()
        logging.info("🏁 [STARTUP:UI] _setup_ui() finished successfully.")

        # ==========================================
        # --- GLOBAL GUI LOGGER ---
        # ==========================================
        #
        # main_tools normally creates GuiLogHandler before application startup
        # so that logs generated before ConsolePanel exists are not lost.
        #
        # Keep a fallback here so DragDropUI can still be instantiated directly
        # by development tools or other callers without having to provide a
        # handler explicitly.
        if gui_log_handler is None:
            self.gui_logger = GuiLogHandler()
            self.gui_logger.setLevel(logging.NOTSET)
            self.gui_logger.setFormatter(
                logging.Formatter("%(message)s")
            )
            logging.getLogger().addHandler(self.gui_logger)
        else:
            self.gui_logger = gui_log_handler

        # ConsolePanel now exists because _setup_ui() has completed.
        #
        # attach() connects the Qt signal and immediately replays all startup
        # records that were buffered before the panel was available.
        self.gui_logger.attach(self.console_panel.log_message)
        self.gui_logger.log_emitted.connect(self.bottom_drawer.record_log_event)

        # Wire Queue Manager
        logging.info("🏁 [STARTUP:QUEUE] Initializing QueueManager...")
        app_context = {"jar_path": self.jar_path}
        self.queue_manager = QueueManager(app_context=app_context)
        self.queue_manager.queue_updated.connect(self._on_queue_updated)
        self.queue_manager.processing_state_changed.connect(self._update_system_status)
        self.queue_manager.conversion_success.connect(self.on_conversion_success)
        self.queue_manager.conversion_failed.connect(self.on_conversion_failed)

        self.queue_panel.remove_requested.connect(self.queue_manager.remove_items)
        self.queue_panel.clear_requested.connect(self.queue_manager.clear_queue)
        self.queue_panel.abort_requested.connect(self.queue_manager.abort_current_task)

        self.queue_manager.processing_state_changed.connect(
            lambda is_proc, status: self.queue_panel.abort_btn.setEnabled(is_proc)
        )

        logging.info("🏁 [STARTUP:CACHE] Loading previous editor session cache...")
        self.cache_file = get_project_root() / ".editor_cache.puml"
        self._load_cache()
        logging.info("🏁 [STARTUP:CACHE] Previous editor cache step complete.")

        self.save_timer = QTimer()
        self.save_timer.setSingleShot(True)
        self.save_timer.setInterval(2000)
        self.save_timer.timeout.connect(self.save_cache)
        self.code_tab.text_input.textChanged.connect(self.save_timer.start)

        logging.info("🏁 [STARTUP:PREVIEW] Initializing LivePreviewManager...")
        self.live_preview = LivePreviewManager(self.code_tab.text_input, self.jar_path)
        self.live_preview.log_msg.connect(self.log_message)
        self.live_preview.svg_updated.connect(self.code_tab.svg_preview.load_svg)
        self.live_preview.error_detected.connect(self.code_tab.text_input.set_error_line)

    def _setup_ui(self):
        central_widget = QWidget()
        root_layout = QHBoxLayout()
        root_layout.setContentsMargins(0, 0, 0, 0)
        root_layout.setSpacing(0)

        self.nav_rail = NavRail(self)
        root_layout.addWidget(self.nav_rail)

        main_container = QWidget()
        main_layout = QVBoxLayout(main_container)
        main_layout.setContentsMargins(12, 10, 12, 8)
        main_layout.setSpacing(8)

        self.splitter = QSplitter(Qt.Vertical)

        # --- TOP HALF: TABS (tab bar hidden; driven by NavRail) ---
        self.tabs = QTabWidget()
        self.tabs.setObjectName("shellWorkspaceTabs")
        self.tabs.setDocumentMode(True)
        self.tabs.tabBar().hide()
        # Allow window to shrink cleanly to laptop screens without largest child tab clamping min width
        self.tabs.setMinimumSize(0, 0)
        self.tabs.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Ignored)
        self.tabs.setStyleSheet(
            f"QTabWidget#shellWorkspaceTabs::pane {{ border: none; background: {C.BG_CANVAS}; top: 0px; }}"
        )

        logging.info("🏁 [STARTUP:UI:TAB] 1/8 Initializing CodeEditorTab...")
        self.code_tab = CodeEditorTab()
        self.code_tab.template_requested.connect(self.insert_template)
        self.code_tab.docs_requested.connect(self.open_template_docs)
        self.code_tab.clear_requested.connect(self.clear_editor)
        self.code_tab.undo_requested.connect(self.code_tab.text_input.undo)
        self.code_tab.redo_requested.connect(self.code_tab.text_input.redo)
        self.code_tab.copy_code_requested.connect(self.copy_editor_code)
        self.code_tab.live_view_toggled.connect(self.toggle_live_view)
        self.code_tab.browser_preview_requested.connect(lambda: self.live_preview.open_in_browser())
        self.code_tab.planttext_requested.connect(self.show_in_planttext)
        self.code_tab.copy_path_requested.connect(self.copy_out_path)
        self.code_tab.open_folder_requested.connect(self.open_export_folder)
        self.code_tab.export_requested.connect(self._save_and_queue_pasted_text)
        self.code_tab.file_dropped.connect(self._on_editor_file_dropped)

        logging.info("🏁 [STARTUP:UI:TAB] 2/8 Initializing BatchConvertTab...")
        self.batch_tab = BatchConvertTab()
        self.batch_tab.files_dropped.connect(
            lambda paths, fmt: self.queue_manager.add_batch(paths, target_format=fmt)
        )
        self.batch_tab.send_to_editor_requested.connect(self._on_send_code_to_editor)
        self.batch_tab.insert_snippet_requested.connect(self._on_insert_snippet_from_visio)
        self.batch_tab.java_maintenance_requested.connect(self.open_java_maintenance)
        self.batch_tab.open_folder_requested.connect(self.open_export_folder)

        logging.info("🏁 [STARTUP:UI:TAB] 3/8 Initializing WordExtractorTab...")
        self.word_tab = WordExtractorTab()
        self.word_tab.extract_visio_requested.connect(
            lambda fp: self.queue_manager.add_item(Path(fp), "extract_visio")
        )
        self.word_tab.split_doc_requested.connect(
            lambda fp, pref, depth: self.queue_manager.add_item(
                Path(fp), "split_docx", {"prefix": pref, "depth": depth}
            )
        )
        self.word_tab.compare_doc_requested.connect(
            lambda a, b, keep_open: self.queue_manager.add_item(
                Path("Word_Comparison_Task"),
                "compare_docx",
                {"doc_a": a, "doc_b": b, "keep_open": keep_open}
            )
        )
        self.word_tab.convert_doc_requested.connect(
            lambda source_doc, target_fmt: self.queue_manager.add_item(
                Path(source_doc),
                "word_convert",
                {"fmt": target_fmt}
            )
        )
        self.word_tab.convert_markdown_requested.connect(
            self._on_word_convert_markdown_requested
        )
        self.word_tab.convert_markdown_to_word_requested.connect(
            self._on_word_convert_markdown_to_word_requested
        )

        db_path = get_project_root() / "db" / "3gpp_data.db"
        logging.info("🏁 [STARTUP:UI:TAB] 4/8 Initializing SpecificationsTab...")
        self.specs_tab = SpecificationsTab(db_path)
        self.specs_tab.update_db_requested.connect(
            lambda force_meta: self.queue_manager.add_item(
                Path("3GPP_Archive"),
                "update_specs_db",
                {"db_path": db_path, "force_metadata": force_meta, "target_specs": []}
            )
        )
        self.specs_tab.update_specific_requested.connect(
            lambda target_specs, force_meta: self.queue_manager.add_item(
                Path("3GPP_Archive_Targeted"),
                "update_specs_db",
                {"db_path": db_path, "force_metadata": force_meta, "target_specs": target_specs}
            )
        )

        logging.info("🏁 [STARTUP:UI:TAB] 5/8 Initializing MeetingsTab...")
        self.meetings_tab = MeetingsTab(db_path)

        logging.info("🏁 [STARTUP:UI:TAB] 6/8 Initializing WorkItemsTab...")
        self.work_items_tab = WorkItemsTab(db_path)
        self.work_items_tab.global_action_requested.connect(self.meetings_tab._handle_global_action_from_window)

        self.meetings_tab.update_db_requested.connect(
            lambda wg, docs, dyna: self.queue_manager.add_item(
                Path("3GPP_Meetings"),
                "update_meetings_db",
                {"db_path": db_path, "sync_wg": wg, "sync_docs": docs, "sync_dyna": dyna}
            )
        )
        self.meetings_tab.update_specific_requested.connect(
            lambda target_list, wg, docs, dyna: self.queue_manager.add_item(
                Path("3GPP_Meetings_Targeted"),
                "update_meetings_db",
                {
                    "db_path": db_path,
                    "target_meetings": target_list,
                    "sync_wg": wg,
                    "sync_docs": docs,
                    "sync_dyna": dyna
                }
            )
        )

        nas_db_path = get_project_root() / "db" / "3gpp_protocol_data.db"
        logging.info("🏁 [STARTUP:UI:TAB] 7/9 Initializing NASTab...")
        self.nas_tab = NASTab(nas_db_path, db_path)

        spec_search_db_path = get_project_root() / "db" / "3gpp_spec_search.db"
        logging.info("🏁 [STARTUP:UI:TAB] 8/9 Initializing SpecSearchTab...")
        self.spec_search_tab = SpecSearchTab(spec_search_db_path, db_path)

        # Wire bidirectional deep-linking between Specifications and Spec Search tabs
        self.spec_search_tab.navigate_to_specs_tab_requested.connect(self._navigate_from_search_to_specs)
        self.specs_tab.search_in_spec_search_requested.connect(self._navigate_from_specs_to_search)
        self.specs_tab.index_in_spec_search_requested.connect(self._index_spec_from_specs_tab)
        self.specs_tab.batch_index_in_spec_search_requested.connect(self._batch_index_specs_from_specs_tab)
        self.specs_tab.smart_index_in_spec_search_requested.connect(self._smart_index_spec_from_specs_tab)

        # Wire deep-linking from Protocols tab to Specifications and Spec Search tabs
        self.nas_tab.navigate_to_specs_tab_requested.connect(self._navigate_from_search_to_specs)
        self.nas_tab.search_in_spec_search_requested.connect(self._navigate_from_specs_to_search)

        logging.info("🏁 [STARTUP:UI:TAB] 9/9 Initializing AssistantTab...")
        self.assistant_tab = AssistantTab(
            specs_db_path=db_path,
            spec_search_db_path=spec_search_db_path,
            protocol_db_path=nas_db_path,
        )
        self.assistant_tab.open_ollama_settings_requested.connect(self.open_ollama_settings)
        self.assistant_tab.open_spec_search_requested.connect(self._navigate_from_assistant_to_spec_search)
        self.assistant_tab.open_specs_requested.connect(self._navigate_from_assistant_to_specs)
        self.assistant_tab.open_protocol_requested.connect(self._navigate_from_assistant_to_protocols)
        self.assistant_tab.undock_requested.connect(lambda: self._on_right_panel_promote(MODE_ASSISTANT))

        self.tabs.addTab(self.code_tab, "📝 PlantUML")
        self.tabs.addTab(self.batch_tab, "🔄 Visio && Office")
        self.tabs.addTab(self.word_tab, "📘 Word")
        self.tabs.addTab(self.specs_tab, "📚 Specifications")
        self.tabs.addTab(self.spec_search_tab, "🔎 Spec Search")
        self.tabs.addTab(self.nas_tab, "🔬 Protocols")
        self.tabs.addTab(self.meetings_tab, "🗓️ Meetings")
        self.tabs.addTab(self.work_items_tab, "📋 Work Items")
        self.tabs.addTab(self.assistant_tab, "🤖 Assistant")

        # Sync hidden QTabWidget <-> NavRail and trigger lazy loading
        self.nav_rail.tab_selected.connect(self.tabs.setCurrentIndex)
        if hasattr(self.nav_rail, "command_palette_requested") and hasattr(self, "open_command_palette"):
            self.nav_rail.command_palette_requested.connect(self.open_command_palette)
        self.tabs.currentChanged.connect(self._on_tab_changed)
        self.nav_rail.set_active_index(self.tabs.currentIndex())

        # Navigation rail footer: Help & Resources + Settings menus
        self.help_btn = QPushButton("Help && Docs")
        self.help_btn.setObjectName("navRailFooterBtn")
        self.help_btn.setIcon(
            lucide_icon("help-circle", color=C.TEXT_SECONDARY, active_color=C.ACCENT_PRIMARY, size=16)
        )
        self.help_btn.setToolTip("Useful 3GPP tools, external portals, and documentation (F1)")

        # Dynamic Help & Resources dropdown menu
        self.help_menu = QMenu(self)
        self.help_menu.setToolTipsVisible(True)
        self.help_menu.aboutToShow.connect(self._rebuild_help_menu)
        self.help_btn.setMenu(self.help_menu)
        self._rebuild_help_menu()

        self.settings_btn = QPushButton("Settings")
        self.settings_btn.setObjectName("navRailFooterBtn")
        self.settings_btn.setIcon(
            lucide_icon("settings", color=C.TEXT_SECONDARY, active_color=C.ACCENT_PRIMARY, size=16)
        )
        self.settings_btn.setToolTip("Configure AI, Proxy, Network, Java, Databases, and System Tools")
        self.settings_menu = QMenu(self)
        self.settings_menu.setToolTipsVisible(True)
        self._populate_settings_menu()
        self.settings_btn.setMenu(self.settings_menu)

        self.nav_rail.register_footer_button(self.help_btn, "Help && Docs")
        self.nav_rail.register_footer_button(self.settings_btn, "Settings")

        self.help_shortcut = QShortcut(QKeySequence("F1"), self)
        self.help_shortcut.activated.connect(self.open_documentation)

        self.splitter.addWidget(self.tabs)

        # --- BOTTOM HALF: COLLAPSIBLE CONSOLE & QUEUE DRAWER ---
        logging.info("🏁 [STARTUP:UI:PANELS] Initializing ConsolePanel and QueuePanel...")

        # Wire Console Panel buttons in _setup_ui():
        self.console_panel = ConsolePanel()
        self.console_panel.proxy_requested.connect(self.open_proxy_settings)
        self.console_panel.network_config_requested.connect(lambda: NetworkConfigDialog(self).exec_())
        self.console_panel.java_info_requested.connect(self.open_java_maintenance)
        self.console_panel.update_requested.connect(self.open_java_maintenance)
        self.console_panel.task_manager_requested.connect(self.open_task_manager)
        self.console_panel.db_maintenance_requested.connect(self.open_db_maintenance)
        self.console_panel.mcp_monitor_requested.connect(self.open_mcp_monitor)

        self.queue_panel = QueuePanel()

        self.bottom_drawer = BottomDrawer(self.console_panel, self.queue_panel, self)
        self.bottom_splitter = self.bottom_drawer.bottom_splitter
        self.bottom_drawer.toggled.connect(self._on_bottom_drawer_toggled)
        self.bottom_drawer.badges_changed.connect(self._on_drawer_badges_changed)

        self.splitter.addWidget(self.bottom_drawer)
        self.splitter.setSizes([700, 0])

        # Horizontal workspace splitter: main vertical workspace on left, contextual RightSidePanel on right
        self.workspace_splitter = QSplitter(Qt.Horizontal)
        self.workspace_splitter.setObjectName("workspaceSplitter")
        self.workspace_splitter.setChildrenCollapsible(False)
        self.workspace_splitter.setHandleWidth(4)
        self.right_panel = RightSidePanel(self)
        self.quick_search_panel = QuickSpecSearchPanel(
            get_project_root() / "db" / "3gpp_spec_search.db",
            self,
        )
        self.quick_search_panel.open_in_spec_search_requested.connect(
            self._navigate_from_quick_search_to_spec_search
        )
        self.quick_search_panel.open_in_specs_requested.connect(
            self._navigate_from_search_to_specs
        )
        self.right_panel.set_slot_widget(MODE_SPEC_SEARCH, self.quick_search_panel)

        self.workspace_splitter.addWidget(self.splitter)
        self.workspace_splitter.addWidget(self.right_panel)
        self.workspace_splitter.setStretchFactor(0, 1)
        self.workspace_splitter.setStretchFactor(1, 0)
        self.workspace_splitter.setSizes([900, 0])
        self.workspace_splitter.splitterMoved.connect(self._on_workspace_splitter_moved)

        self.right_panel.toggled.connect(self._on_right_panel_state_changed)
        self.right_panel.mode_changed.connect(
            lambda _mode: self._on_right_panel_state_changed(self.right_panel.is_open)
        )
        self.right_panel.promote_requested.connect(self._on_right_panel_promote)

        main_layout.addWidget(self.workspace_splitter)
        root_layout.addWidget(main_container, 1)

        central_widget.setLayout(root_layout)
        self.setCentralWidget(central_widget)

        # Bottom System Status Bar (decoupled component)
        self.status_bar = SystemStatusBar(self)
        self.setStatusBar(self.status_bar)
        self.status_bar.reinit_requested.connect(self._on_status_reinit_requested)
        self.status_bar.drawer_toggle_requested.connect(self.toggle_bottom_drawer)
        self.status_bar.assistant_dock_requested.connect(
            lambda: self.toggle_right_panel(MODE_ASSISTANT)
        )
        self.status_bar.search_dock_requested.connect(
            lambda: self.toggle_right_panel(MODE_SPEC_SEARCH)
        )
        self.status_bar.task_manager_requested.connect(self.open_task_manager)
        self.status_bar.db_maintenance_requested.connect(self.open_db_maintenance)
        self.status_bar.network_config_requested.connect(lambda: NetworkConfigDialog(self).exec_())
        self.status_bar.java_info_requested.connect(self.open_java_maintenance)
        self.status_bar.update_check_requested.connect(self.open_app_updater)
        self.status_bar.update_proxy_indicator(self.console_panel)
        self.status_bar.update_drawer_badge(
            self.bottom_drawer.is_open,
            self.bottom_drawer.queue_count,
            self.bottom_drawer.unread_errors,
        )
        self.status_bar.update_right_panel_badges(
            self.right_panel.is_open,
            self.right_panel.mode,
        )

        # Floating non-blocking Toast notification overlay manager
        self.toast_manager = ToastManager(self)

        # Keyboard shortcuts (Ctrl+1..9, Ctrl+J, Ctrl+\) & predictable tab focus order
        self._setup_shortcuts_and_focus()

        logging.info("🏁 [STARTUP:UI] _setup_ui() layout complete.")

    def _setup_shortcuts_and_focus(self):
        """Register shell keyboard shortcuts (`Ctrl+1..9`, `Ctrl+J`, `Ctrl+\\`) and tab focus order."""
        self.nav_shortcuts = []
        for idx in range(min(9, self.tabs.count())):
            sc = QShortcut(QKeySequence(f"Ctrl+{idx + 1}"), self)
            sc.setContext(Qt.WindowShortcut)
            sc.activated.connect(lambda tab_idx=idx: self._activate_tab_by_shortcut(tab_idx))
            self.nav_shortcuts.append(sc)

        self.drawer_shortcut = QShortcut(QKeySequence("Ctrl+J"), self)
        self.drawer_shortcut.setContext(Qt.WindowShortcut)
        self.drawer_shortcut.activated.connect(self.toggle_bottom_drawer)

        self.rail_shortcut = QShortcut(QKeySequence("Ctrl+\\"), self)
        self.rail_shortcut.setContext(Qt.WindowShortcut)
        self.rail_shortcut.activated.connect(self.nav_rail.toggle_collapsed)

        if hasattr(self, "open_command_palette"):
            self.command_palette_shortcut = QShortcut(QKeySequence("Ctrl+K"), self)
            self.command_palette_shortcut.setContext(Qt.WindowShortcut)
            self.command_palette_shortcut.activated.connect(self.open_command_palette)

        if hasattr(self, "toggle_right_panel"):
            self.assistant_dock_shortcut = QShortcut(QKeySequence("Ctrl+Shift+A"), self)
            self.assistant_dock_shortcut.setContext(Qt.WindowShortcut)
            self.assistant_dock_shortcut.activated.connect(
                lambda: self.toggle_right_panel(MODE_ASSISTANT)
            )

            self.search_dock_shortcut = QShortcut(QKeySequence("Ctrl+Shift+F"), self)
            self.search_dock_shortcut.setContext(Qt.WindowShortcut)
            self.search_dock_shortcut.activated.connect(
                lambda: self.toggle_right_panel(MODE_SPEC_SEARCH)
            )

        extra_status_focus = []
        if hasattr(self, "right_panel") and hasattr(self.status_bar, "assistant_dock_btn"):
            extra_status_focus = [
                self.status_bar.assistant_dock_btn,
                self.status_bar.search_dock_btn,
            ]

        focus_chain = [
            *self.nav_rail.ordered_focus_widgets(),
            self.tabs,
            self.bottom_drawer.close_btn,
            self.status_bar.ollama_indicator,
            self.status_bar.mcp_indicator,
            self.status_bar.proxy_indicator,
            self.status_bar.system_btn,
            *extra_status_focus,
            self.status_bar.drawer_toggle_btn,
        ]
        for prev_w, next_w in zip(focus_chain, focus_chain[1:]):
            QWidget.setTabOrder(prev_w, next_w)

    def _activate_tab_by_shortcut(self, index: int):
        """Switch active workspace tab via Ctrl+1..9 when shell navigation is enabled."""
        if self.tabs.isEnabled() and 0 <= index < self.tabs.count():
            self.tabs.setCurrentIndex(index)
            if hasattr(self, "nav_rail"):
                self.nav_rail.set_active_index(index)

    def toggle_bottom_drawer(self):
        """Toggle the bottom Console & Queue drawer cleanly."""
        if hasattr(self, "bottom_drawer"):
            self.bottom_drawer.toggle()

    def _on_bottom_drawer_toggled(self, opened: bool):
        if hasattr(self, "bottom_drawer") and hasattr(self, "splitter"):
            self.bottom_drawer.apply_splitter_sizes(self.splitter)
        if hasattr(self, "status_bar") and hasattr(self, "bottom_drawer"):
            self.status_bar.update_drawer_badge(
                opened,
                self.bottom_drawer.queue_count,
                self.bottom_drawer.unread_errors,
            )

    def _on_drawer_badges_changed(self, queue_count: int, unread_errors: int):
        if hasattr(self, "status_bar") and hasattr(self, "bottom_drawer"):
            self.status_bar.update_drawer_badge(
                self.bottom_drawer.is_open,
                queue_count,
                unread_errors,
            )

    def _on_queue_updated(self, display_items: list):
        self.queue_panel.update_list(display_items)
        if hasattr(self, "bottom_drawer"):
            self.bottom_drawer.set_queue_count(len(display_items))

    def _on_tab_changed(self, index: int):
        """Triggers lazy loading when an inactive tab is selected and syncs NavRail."""
        if hasattr(self, "nav_rail"):
            self.nav_rail.set_active_index(index)
        active_widget = self.tabs.widget(index)
        if active_widget is getattr(self, "assistant_tab", None) and getattr(self, "_assistant_docked", False):
            self._set_assistant_docked(False)
            if (
                hasattr(self, "right_panel")
                and self.right_panel.is_open
                and self.right_panel.mode == MODE_ASSISTANT
            ):
                self.right_panel.close_panel()
        if hasattr(active_widget, 'ensure_loaded'):
            active_widget.ensure_loaded()

    def _set_assistant_docked(self, docked: bool):
        """Move single-instance AssistantTab body between full tab and RightSidePanel."""
        if (
            not hasattr(self, "assistant_tab")
            or not hasattr(self.assistant_tab, "set_docked_mode")
            or not hasattr(self, "right_panel")
        ):
            return
        want = bool(docked)
        if want == getattr(self, "_assistant_docked", False):
            return
        if want:
            self.assistant_tab.set_docked_mode(True)
            self.right_panel.set_slot_widget(MODE_ASSISTANT, self.assistant_tab.body_container)
            self._assistant_docked = True
        else:
            self._assistant_docked = False
            self.assistant_tab.restore_body_into_tab()

    def _sync_assistant_dock_slot(self):
        if not hasattr(self, "right_panel") or not hasattr(self, "_set_assistant_docked"):
            return
        want = bool(self.right_panel.is_open and self.right_panel.mode == MODE_ASSISTANT)
        self._set_assistant_docked(want)

    def toggle_right_panel(self, mode: str):
        """Toggle or switch the contextual RightSidePanel (Assistant / Quick Spec Search)."""
        if hasattr(self, "right_panel"):
            self.right_panel.toggle_mode(mode)

    def open_right_panel(self, mode: str):
        """Open the contextual RightSidePanel on the requested mode."""
        if hasattr(self, "right_panel"):
            self.right_panel.open_mode(mode)

    def close_right_panel(self):
        """Close the contextual RightSidePanel."""
        if hasattr(self, "right_panel"):
            self.right_panel.close_panel()

    def _on_workspace_splitter_moved(self, _pos: int, _index: int):
        if hasattr(self, "right_panel") and hasattr(self, "workspace_splitter"):
            self.right_panel.capture_width_from_splitter(self.workspace_splitter)
        if hasattr(self, "toast_manager"):
            self.toast_manager.reposition()

    def _on_right_panel_state_changed(self, opened: bool):
        self._sync_assistant_dock_slot()
        if (
            opened
            and hasattr(self, "right_panel")
            and self.right_panel.mode == MODE_SPEC_SEARCH
            and hasattr(self, "quick_search_panel")
        ):
            self.quick_search_panel.refresh_versions()
            QTimer.singleShot(0, self.quick_search_panel.focus_input)
        if hasattr(self, "right_panel") and hasattr(self, "workspace_splitter"):
            self.right_panel.apply_splitter_sizes(self.workspace_splitter)
        if hasattr(self, "status_bar") and hasattr(self, "right_panel"):
            self.status_bar.update_right_panel_badges(opened, self.right_panel.mode)
        if hasattr(self, "nav_rail"):
            self.nav_rail.apply_responsive_width(self.width(), dock_open=opened)
        if hasattr(self, "toast_manager"):
            self.toast_manager.reposition()

    def _on_right_panel_promote(self, mode: str):
        """Promote docked panel mode into its full workspace tab and close the dock."""
        if mode == MODE_ASSISTANT:
            if hasattr(self, "_set_assistant_docked"):
                self._set_assistant_docked(False)
            if hasattr(self, "assistant_tab"):
                self.tabs.setCurrentWidget(self.assistant_tab)
        elif mode == MODE_SPEC_SEARCH and hasattr(self, "spec_search_tab"):
            self.tabs.setCurrentWidget(self.spec_search_tab)
            if hasattr(self, "quick_search_panel"):
                q = self.quick_search_panel.current_query()
                if q and hasattr(self.spec_search_tab, "search_input"):
                    self.spec_search_tab.search_input.setText(q)
        if hasattr(self, "right_panel"):
            self.right_panel.close_panel()

    def _navigate_from_quick_search_to_spec_search(self, spec_number: str, version: str, clause: str):
        """Switch to full Spec Search tab and focus spec, version, clause, and query."""
        self.tabs.setCurrentWidget(self.spec_search_tab)
        if spec_number:
            self.spec_search_tab.focus_spec(spec_number, version)
        if clause and hasattr(self.spec_search_tab, "clause_filter_input"):
            self.spec_search_tab.clause_filter_input.setText(clause)
        if hasattr(self, "quick_search_panel"):
            q = self.quick_search_panel.current_query()
            if q and hasattr(self.spec_search_tab, "search_input"):
                self.spec_search_tab.search_input.setText(q)
        if hasattr(self.spec_search_tab, "_execute_search"):
            self.spec_search_tab._execute_search()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        dock_open = hasattr(self, "right_panel") and self.right_panel.is_open
        if hasattr(self, "nav_rail"):
            self.nav_rail.apply_responsive_width(self.width(), dock_open=dock_open)
        if hasattr(self, "toast_manager"):
            self.toast_manager.reposition()

    def _navigate_from_search_to_specs(self, spec_number: str, version: str):
        """Switches to Specifications tab and selects the requested spec and version."""
        self.tabs.setCurrentWidget(self.specs_tab)
        self.specs_tab.select_and_show_spec(spec_number, version)

    def _navigate_from_specs_to_search(self, spec_number: str, version: str):
        """Switches to Spec Search tab and focuses on the requested spec and version."""
        self.tabs.setCurrentWidget(self.spec_search_tab)
        self.spec_search_tab.focus_spec(spec_number, version)

    def _navigate_from_assistant_to_spec_search(self, spec_number: str, version: str, clause: str):
        """Switches to Spec Search tab and loads the requested spec, version, and clause."""
        self.tabs.setCurrentWidget(self.spec_search_tab)
        if spec_number:
            self.spec_search_tab.focus_spec(spec_number, version)
        if clause and hasattr(self.spec_search_tab, "clause_filter_input"):
            self.spec_search_tab.clause_filter_input.setText(clause)

    def _navigate_from_assistant_to_specs(self, spec_number: str, version: str):
        """Switches to Specifications tab and selects the requested spec and version."""
        self.tabs.setCurrentWidget(self.specs_tab)
        self.specs_tab.select_and_show_spec(spec_number, version)

    def open_command_palette(self):
        """Launch the modal Ctrl+K quick switcher & command palette."""
        if not hasattr(self, "_palette_searcher") or self._palette_searcher is None:
            root_db = get_project_root() / "db"
            self._palette_searcher = DynamicPaletteSearcher(
                db_path=root_db / "3gpp_data.db",
                protocol_db_path=root_db / "3gpp_protocol_data.db",
                on_open_spec=lambda s: self._navigate_from_search_to_specs(s, ""),
                on_open_meeting=self._open_meeting_from_palette,
                on_open_work_item=self._open_work_item_from_palette,
                on_open_protocol=self._navigate_from_assistant_to_protocols,
            )
        dlg = CommandPaletteDialog(
            build_static_command_items(self),
            dynamic_searcher=self._palette_searcher.search,
            parent=self,
        )
        dlg.exec_()

    def _open_meeting_from_palette(self, query_or_tdoc: str):
        """Switch to Meetings tab and filter-safely focus the requested meeting/TDoc."""
        if hasattr(self, "meetings_tab"):
            self.tabs.setCurrentWidget(self.meetings_tab)
            if hasattr(self.meetings_tab, "focus_meeting"):
                self.meetings_tab.focus_meeting(query_or_tdoc)

    def _open_work_item_from_palette(self, query: str):
        """Switch to Work Items tab and filter-safely focus the requested WI."""
        if hasattr(self, "work_items_tab"):
            self.tabs.setCurrentWidget(self.work_items_tab)
            if hasattr(self.work_items_tab, "focus_work_item"):
                self.work_items_tab.focus_work_item(query)

    def _navigate_from_assistant_to_protocols(self, protocol: str, message: str):
        """Switches to Protocols tab and selects the requested message."""
        self.tabs.setCurrentWidget(self.nas_tab)
        if hasattr(self.nas_tab, "ensure_loaded"):
            self.nas_tab.ensure_loaded()
        if message and hasattr(self.nas_tab, "_jump_to_message"):
            self.nas_tab._jump_to_message(message)

    def _smart_index_spec_from_specs_tab(self, spec_number: str):
        """Switches to Spec Search tab and opens the targeted Smart Import dialog for this spec."""
        self.tabs.setCurrentWidget(self.spec_search_tab)
        self.spec_search_tab._open_smart_import_dialog(initial_spec=spec_number)

    def _index_spec_from_specs_tab(self, spec_number: str, version: str):
        """Directly indexes a single specification version from Specifications tab into Spec Search."""
        self._batch_index_specs_from_specs_tab([(spec_number, version)])

    def _batch_index_specs_from_specs_tab(self, spec_items: list):
        """Directly indexes one or more specification versions from Specifications tab into Spec Search."""
        from modules.nas.core.parsing.protocol_parser_common import ProtocolDocxDispatcher
        from modules.spec_search.core.spec_search_threads import is_change_mark_file

        if not spec_items:
            return

        tasks = []
        for spec_number, version in spec_items:
            clean_spec = str(spec_number).strip()
            clean_ver = str(version).lstrip("vV").strip()

            # 1. Exact lookup via get_spec_files
            files = self.specs_tab.db.get_spec_files(clean_spec)
            matching_file = None
            if clean_ver:
                matching_file = next(
                    (f for f in files if str(f.get("version", "")).lstrip("vV").strip() == clean_ver),
                    None
                )
            elif files:
                matching_file = files[0]
                clean_ver = str(matching_file.get("version", "")).lstrip("vV").strip()

            filename = ""
            file_url = ""
            release_date = ""

            if matching_file:
                filename = matching_file.get("filename", "")
                file_url = matching_file.get("url", "")
                release_date = matching_file.get("upload_date", "")
            else:
                rows = self.specs_tab.db.search_files(spec_number=clean_spec, release_version=clean_ver)
                if rows:
                    filename = rows[0][4]
                    file_url = rows[0][6]
                    release_date = rows[0][7]
                    if not clean_ver and len(rows[0]) > 5:
                        clean_ver = str(rows[0][5]).lstrip("vV").strip()

            if not filename and not file_url:
                continue

            task = {
                "spec_number": clean_spec,
                "version": clean_ver,
                "filename": filename,
                "file_url": file_url,
                "release_date": release_date,
            }

            # Check if Word document is already extracted locally in download_dir
            spec_dir = Path(self.specs_tab.download_dir) / clean_spec
            stem = filename.replace(".zip", "")
            if spec_dir.exists() and stem:
                local_docs = [
                    f for f in spec_dir.iterdir()
                    if f.is_file()
                    and f.stem.lower().startswith(stem.lower())
                    and f.suffix.lower() in [".docx", ".doc"]
                    and not f.name.startswith("~$")
                    and not is_change_mark_file(f.name)
                ]
                if local_docs:
                    local_docs.sort(key=ProtocolDocxDispatcher._extract_part_index)
                    task["local_docx_paths"] = local_docs

            tasks.append(task)

        if not tasks:
            QMessageBox.warning(
                self,
                "Spec Not Found",
                "Could not find specification metadata or archive records for the requested version(s)."
            )
            return

        self.tabs.setCurrentWidget(self.spec_search_tab)
        self.spec_search_tab._start_batch_ingestion(tasks)

    def _launch_init_thread(self, check_updates=False):
        """Thread-safe launcher that prevents duplicate initialization checks."""
        if self.init_thread is not None and self.init_thread.isRunning():
            logging.warning("⚠️ [STARTUP] InitializationThread already running. Skipping duplicate launch.")
            return

        self.init_thread = InitializationThread(self.jar_path, check_updates=check_updates)
        self.init_thread.ui_log_msg.connect(self.log_message)
        self.init_thread.init_complete.connect(self.on_init_complete)
        self.init_thread.network_error.connect(self.open_proxy_settings)
        self.init_thread.start()

        # Failsafe: Ensure UI is fully interactive within 3 seconds regardless of thread status
        QTimer.singleShot(3000, lambda: self.tabs.setEnabled(True))

    # --- DIALOG & THREAD MANAGEMENT ---
    def open_db_maintenance(self):
        """Spawns the Database Maintenance & Compaction dialog as a non-blocking modeless window."""
        if not hasattr(self, 'db_maintenance_dialog') or not self.db_maintenance_dialog.isVisible():
            self.db_maintenance_dialog = DatabaseMaintenanceDialog(self)
            self.db_maintenance_dialog.show()
            self.db_maintenance_dialog.raise_()
            self.db_maintenance_dialog.activateWindow()
        else:
            self.db_maintenance_dialog.raise_()
            self.db_maintenance_dialog.activateWindow()

    def open_task_manager(self):
        """Spawns the COM Process Manager dialog as a non-blocking floating window."""
        if not hasattr(self, 'task_manager_dialog') or not self.task_manager_dialog.isVisible():
            self.task_manager_dialog = ProcessManagerDialog(self)
            self.task_manager_dialog.show()
            self.task_manager_dialog.raise_()
            self.task_manager_dialog.activateWindow()
        else:
            self.task_manager_dialog.raise_()
            self.task_manager_dialog.activateWindow()

    def open_mcp_monitor(self):
        """Spawns the MCP Server & External AI Agent Cockpit dialog."""
        self.status_bar.open_mcp_monitor()

    # --- AUTO-SAVE LOGIC ---
    def _load_cache(self):
        if self.cache_file.exists():
            try:
                text = self.cache_file.read_text(encoding="utf-8")
                if text.strip():
                    self.code_tab.text_input.setPlainText(text)
                    self.log_message("♻️ Restored main window previous session.", logging.INFO)
            except Exception as e:
                self.log_message(f"⚠️ Could not load previous session: {e}", logging.WARNING)

    def save_cache(self):
        try:
            self.cache_file.write_text(self.code_tab.get_text(), encoding="utf-8")
        except Exception:
            pass

    # --- UI INTERACTION LOGIC ---
    def _populate_settings_menu(self):
        """Populates the rail footer Settings & System Tools dropdown menu."""
        self.settings_menu.clear()
        self.settings_menu.addAction("🦙 Ollama AI Settings...", self.open_ollama_settings)
        self.settings_menu.addAction("🛡️ Proxy Settings...", self.open_proxy_settings)
        self.settings_menu.addAction("🌐 Network Profiles && Connectivity...", lambda: NetworkConfigDialog(self).exec_())
        self.settings_menu.addAction("☕ Java && PlantUML Engine...", self.open_java_maintenance)
        self.settings_menu.addSeparator()
        self.settings_menu.addAction("🗄️ Database Maintenance...", self.open_db_maintenance)
        self.settings_menu.addAction("🖥️ COM Task Manager...", self.open_task_manager)
        self.settings_menu.addAction("🤖 MCP Server && AI Cockpit...", self.open_mcp_monitor)
        self.settings_menu.addSeparator()
        self.settings_menu.addAction("🔄 Check for Updates...", self.open_app_updater)
        self.settings_menu.addAction("📝 Edit Useful Links Config...", self.edit_useful_links_config)

    def _rebuild_help_menu(self):
        populate_help_menu(
            self.help_menu,
            on_open_url=self.open_external_url,
            on_edit_config=self.edit_useful_links_config,
        )
        self.help_menu.addSeparator()
        mcp_action = self.help_menu.addAction("🤖 MCP Server && AI Agent Cockpit...")
        mcp_action.setToolTip("Inspect real-time MCP server tool calls, external AI agent activity, and client configurations.")
        mcp_action.triggered.connect(self.open_mcp_monitor)
        upd_action = self.help_menu.addAction("🔄 Check for Updates...")
        upd_action.setToolTip("Check GitHub master for recent changes and update the application.")
        upd_action.triggered.connect(self.open_app_updater)

    def open_external_url(self, url: str, title: str = "Web Resource"):
        try:
            webbrowser.open(url)
            self.log_message(f"🌐 Opened {title} in your web browser: {url}", logging.INFO)
        except Exception as e:
            self.log_message(f"❌ Failed to open URL '{url}': {e}", logging.ERROR)

    def edit_useful_links_config(self):
        config_path = get_useful_links_config_path()
        try:
            if hasattr(os, "startfile"):
                os.startfile(str(config_path))
            else:
                webbrowser.open(f"file:///{config_path}")
            self.log_message(f"📝 Opened useful links configuration: {config_path.name}", logging.INFO)
        except Exception as e:
            self.log_message(f"❌ Failed to open config file '{config_path}': {e}", logging.ERROR)

    def open_documentation(self):
        try:
            webbrowser.open(HELP_URL)
            self.log_message("🌐 Opened Documentation (README) in your web browser.", logging.INFO)
        except Exception as e:
            self.log_message(f"❌ Failed to open documentation URL: {e}", logging.ERROR)

    def _update_system_status(self, is_processing: bool, status_text: str):
        self.status_bar.showMessage(status_text)
        if hasattr(self, 'word_tab') and hasattr(self.word_tab, 'on_processing_state_changed'):
            self.word_tab.on_processing_state_changed(is_processing)
        if is_processing:
            self.batch_tab.set_state(
                "busy", "⚙️ Processing Queue...\n\nPlease wait until finished or drop more files to queue them."
            )
        else:
            if not getattr(self, "_core_ready", True):
                self.batch_tab.set_state("error", "❌ Core engine error. Java or PlantUML engine is unavailable.")
            elif not getattr(self, "_visio_ready", True):
                self.batch_tab.set_state("error", "⚠️ Visio unavailable. Visio batch conversions will be disabled.")
            else:
                self.batch_tab.set_state(
                    "ready", "📥 Drag & Drop your .puml, .txt, .pptx, or .vsdx file(s) here\n\n(Batch converts between Visio & PowerPoint)"
                )

    def _on_send_code_to_editor(self, code: str):
        self.tabs.setCurrentWidget(self.code_tab)
        self._set_editor_text(code)
        self.log_message("📝 Loaded extracted PlantUML code into Editor tab.", logging.INFO)

    def _on_insert_snippet_from_visio(self, snippet):
        self.tabs.setCurrentWidget(self.code_tab)
        if hasattr(self.code_tab, "text_input"):
            self.code_tab.text_input.insert_snippet(snippet)
        label = getattr(snippet, 'label', 'snippet')
        self.log_message(f"🧩 Inserted '{label}' into PlantUML editor.", logging.INFO)

    def _set_editor_text(self, text):
        if hasattr(self.code_tab, "set_text"):
            self.code_tab.set_text(text)
        else:
            cursor = self.code_tab.text_input.textCursor()
            cursor.beginEditBlock()
            cursor.select(QTextCursor.Document)
            cursor.insertText(text)
            cursor.endEditBlock()
            cursor.setPosition(0)
            self.code_tab.text_input.setTextCursor(cursor)
        self.live_preview.update_now()

    def insert_template(self, selected_template):
        tpl = PLANTUML_TYPES[selected_template]["template"]
        curr_text = self.code_tab.get_text()
        if curr_text and curr_text.strip() != tpl.strip():
            reply = QMessageBox.question(
                self,
                "Replace Diagram?",
                f"The editor already contains PlantUML code.\n\n"
                f"Do you want to replace it completely with the '{selected_template}' boilerplate?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No
            )
            if reply != QMessageBox.Yes:
                return

        self._set_editor_text(tpl)
        self.log_message(f"📝 Inserted boilerplate for '{selected_template}' diagram.", logging.INFO)

    def clear_editor(self):
        self._set_editor_text("")
        if hasattr(self.code_tab, "_current_file_path"):
            self.code_tab._current_file_path = None
        self.log_message("🗑️ Editor cleared.", logging.INFO)

    def copy_editor_code(self):
        code_text = self.code_tab.get_text()
        if code_text:
            QApplication.clipboard().setText(code_text)
            self.log_message("📄 Source code copied to clipboard.", logging.INFO)
        else:
            self.log_message("⚠️ Editor is empty. Nothing to copy.", logging.WARNING)

    def open_export_folder(self):
        export_dir = get_project_root().parent / "output"
        export_dir.mkdir(parents=True, exist_ok=True)
        try:
            os.startfile(export_dir)
            self.log_message(f"📂 Opened export directory: {export_dir}", logging.INFO)
        except Exception as e:
            self.log_message(f"❌ Failed to open directory: {e}", logging.ERROR)

    def open_template_docs(self, selected_template):
        url = PLANTUML_TYPES[selected_template]["url"]
        try:
            webbrowser.open(url)
            self.log_message(f"🌐 Opened documentation for '{selected_template}'.", logging.INFO)
        except Exception as e:
            self.log_message(f"❌ Failed to open URL: {e}", logging.ERROR)

    def toggle_live_view(self, checked):
        open_browser = self.code_tab.live_view_btn.isChecked()
        self.live_preview.toggle(checked, open_browser=open_browser)

    def _on_status_reinit_requested(self):
        self.log_message("✅ Proxy configuration updated. Retrying system checks...", logging.INFO)
        self.batch_tab.set_state("ready", "⏳ Re-initializing system checks...")
        self.status_bar.showMessage("⏳ Re-initializing...")
        self.tabs.setEnabled(False)
        self._launch_init_thread(check_updates=False)

    def open_proxy_settings(self):
        self.status_bar.open_proxy_settings()


    def check_for_jar_updates(self):
        self.log_message("\n🔄 Initiating manual update check...", logging.INFO)
        self.batch_tab.set_state("ready", "⏳ Checking online for updates...")
        self.status_bar.showMessage("⏳ Checking for updates...")
        self.tabs.setEnabled(False)
        self._launch_init_thread(check_updates=True)

    def _on_editor_file_dropped(self, file_path: str):
        """Dispatches files dropped onto the PlantUML editor to the appropriate loader."""
        path = Path(file_path)
        ext = path.suffix.lower()
        if ext in (".vsdx", ".vsdm"):
            self.extract_code_from_visio(file_path)
        else:
            self.load_puml_file(file_path)

    def load_puml_file(self, file_path):
        """Loads a PlantUML file into the editor, updates live preview, and logs details."""
        path = Path(file_path)
        if not path.is_file():
            self.log_message(f"❌ File not found: {file_path}", logging.ERROR)
            return

        try:
            success = self.code_tab.load_file(path)
            if success:
                self.live_preview.update_now()
                from core.ui.drop_zone import format_file_size
                size_str = format_file_size(path.stat().st_size)
                self.log_message(f"📂 Loaded PlantUML diagram from: {path.name} ({size_str})", logging.INFO)
                self.status_bar.showMessage(f"Loaded {path.name}", 4000)
            else:
                self.log_message(f"❌ Could not load PlantUML file: {path.name}", logging.ERROR)
        except Exception as e:
            self.log_message(f"❌ Failed to load PlantUML file '{path.name}': {e}", logging.ERROR)
            QMessageBox.warning(self, "Load Error", f"Failed to load PlantUML file:\n{e}")

    def extract_code_from_visio(self, file_path):
        self.code_tab.text_input.clear()
        self.code_tab.text_input.setPlaceholderText(
            f"⏳ Extracting source from {Path(file_path).name}...\nPlease wait..."
        )
        self.code_tab.text_input.setEnabled(False)
        self.log_message(f"📂 Reading embedded source from: {Path(file_path).name}")
        self.reader_thread = VisioReaderThread(file_path)
        self.reader_thread.text_extracted.connect(self.on_visio_code_read)
        self.reader_thread.error_occurred.connect(self.on_visio_code_error)
        self.reader_thread.start()

    def on_visio_code_read(self, source_code):
        self.code_tab.text_input.setEnabled(True)
        self._set_editor_text(source_code)
        self.log_message("✅ Successfully extracted PlantUML source from Visio file.")

    def on_visio_code_error(self, error_msg):
        self.code_tab.text_input.setEnabled(True)
        self.code_tab.text_input.setPlaceholderText(
            "Paste PlantUML code OR drag & drop a .puml or .vsdx file here to load..."
        )
        self.log_message(f"❌ {error_msg}")

    def show_in_planttext(self):
        raw_text = self.code_tab.get_text()
        if not raw_text: return
        try:
            url = f"https://www.planttext.com/?text={encode_plantuml(raw_text)}"
            webbrowser.open(url)
            self.log_message("🌐 Opened code in planttext.com")
        except Exception as e:
            self.log_message(f"❌ Failed to open: {e}")

    def copy_out_path(self):
        if self.last_out_path:
            QApplication.clipboard().setText(self.last_out_path)
            self.log_message(f"📋 Copied to clipboard: {self.last_out_path}")

    def _save_and_queue_pasted_text(self, target_format):
        raw_text = self.code_tab.get_text()
        if not raw_text: return

        params = {}
        if target_format == "pptx_append":
            from modules.puml2visio.core.powerpoint_converter import get_open_powerpoint_presentations
            from modules.puml2visio.ui.select_presentation_dialog import SelectPresentationDialog

            open_presentations = get_open_powerpoint_presentations()
            if not open_presentations:
                reply = QMessageBox.question(
                    self,
                    "No Open PowerPoint Presentation",
                    "No open Microsoft PowerPoint presentation could be detected.\n\n"
                    "Would you like to export to a new PowerPoint presentation instead?",
                    QMessageBox.Yes | QMessageBox.No,
                    QMessageBox.Yes
                )
                if reply == QMessageBox.Yes:
                    target_format = "pptx"
                else:
                    return
            elif len(open_presentations) == 1:
                target_pres = open_presentations[0]
                params = {
                    "target_presentation": target_pres["name"],
                    "insert_at": "end"
                }
            else:
                dialog = SelectPresentationDialog(open_presentations, parent=self)
                if dialog.exec_() != QDialog.Accepted:
                    return
                if getattr(dialog, "is_new_presentation_requested", lambda: False)():
                    target_format = "pptx"
                else:
                    chosen = dialog.get_selected_presentation()
                    if not chosen:
                        return
                    params = {
                        "target_presentation": chosen["name"],
                        "insert_at": chosen.get("insert_at", "end")
                    }

        out_dir = get_project_root().parent / "output"
        out_dir.mkdir(parents=True, exist_ok=True)
        timestamp = datetime.datetime.now().strftime("%Y.%m.%d %H-%M-%S")
        base_name = f"{timestamp} diagram"
        puml_path = out_dir / f"{base_name}.puml"

        counter = 1
        while (
            puml_path.exists()
            or puml_path.with_suffix(".vsdx").exists()
            or puml_path.with_suffix(".pptx").exists()
            or puml_path.with_suffix(".svg").exists()
        ):
            puml_path = out_dir / f"{base_name}_{counter}.puml"
            counter += 1

        with open(puml_path, "w", encoding="utf-8") as f:
            f.write(raw_text)

        self.queue_manager.add_item(puml_path, target_format, params=params)

    def _on_word_convert_markdown_requested(self, doc: str, *args, **kwargs) -> None:
        """Enqueues Word to Markdown conversion supporting dict or positional arguments."""
        if args and isinstance(args[0], dict):
            opts = args[0]
        elif kwargs:
            opts = kwargs
        elif len(args) >= 1 and isinstance(args[0], str):
            rev = args[0]
            media = args[1] if len(args) > 1 else True
            out_dir = args[2] if len(args) > 2 else None
            hdr = args[3] if len(args) > 3 else True
            opts = {"revisions": rev, "extract_media": media, "output_dir": out_dir, "include_header": hdr}
        else:
            opts = {}

        params = {
            "fmt": "md",
            "revisions": opts.get("revisions_mode") or opts.get("revisions", "markup"),
            "extract_media": opts.get("extract_media", True),
            "output_dir": opts.get("output_dir"),
            "include_header": opts.get("include_header", True),
        }
        self.queue_manager.add_item(Path(doc), "word_to_markdown", params=params)

    def _on_word_convert_markdown_to_word_requested(self, doc: str, *args, **kwargs) -> None:
        """Enqueues Markdown to Word conversion supporting dict or positional arguments."""
        if args and isinstance(args[0], dict):
            opts = args[0]
        elif kwargs:
            opts = kwargs
        else:
            opts = {}

        params = {
            "output_dir": opts.get("output_dir"),
            "companion_media_dir": opts.get("companion_media_dir"),
            "render_diagrams": opts.get("render_diagrams", True),
            "apply_3gpp_styles": opts.get("apply_3gpp_styles", True),
            "auto_detect_media": opts.get("auto_detect_media", True),
            "auto_open": opts.get("auto_open", False),
        }
        self.queue_manager.add_item(Path(doc), "markdown_to_word", params=params)

    def on_conversion_failed(self, error_msg: str):
        if hasattr(self, 'word_tab') and hasattr(self.word_tab, 'on_conversion_failed'):
            self.word_tab.on_conversion_failed(error_msg)

    def on_conversion_success(self, out_path: str):
        if out_path == "SPECS_DB_PASS_ONE":
            self.log_message("🔄 Basic files loaded! You can now search while metadata downloads.", logging.INFO)
            if hasattr(self, 'specs_tab'):
                self.specs_tab.refresh_table()
                if hasattr(self.specs_tab, 'set_bg_sync_active'):
                    self.specs_tab.set_bg_sync_active(True)

        elif out_path == "SPECS_DB_PASS_TWO":
            self.log_message("✅ Deep metadata fully updated!", logging.INFO)
            if hasattr(self, 'specs_tab'):
                if hasattr(self.specs_tab, 'set_bg_sync_active'):
                    self.specs_tab.set_bg_sync_active(False)
                else:
                    self.specs_tab.refresh_table()
            self.show_toast("Specifications Updated", "Deep metadata synchronization completed.", variant="success")

        elif out_path == "OPENED_IN_PPT":
            self.log_message("👁️ PowerPoint is open with your new slide. You can copy it directly.")
            self.show_toast("PowerPoint Updated", "Added diagram slide to open presentation.", variant="success")

        elif out_path == "SPECS_DB_UPDATED":
            self.log_message("🔄 Reloading specifications table...")
            if hasattr(self, 'specs_tab'):
                self.specs_tab.refresh_table()
            self.show_toast("Specifications Reloaded", "Specification catalogue refreshed.", variant="success")

        elif out_path and out_path.startswith("MEETINGS_DB"):
            self.log_message("🔄 Reloading meetings table...", logging.INFO)
            if hasattr(self, 'meetings_tab'):
                # Only Phase 1 can discover/add meeting folders and therefore
                # potentially change the available WG choices. Phases 2 and 3
                # update Docs/DynaReport metadata only.
                if out_path == "MEETINGS_DB_PHASE_1":
                    self.meetings_tab._populate_filters()
                self.meetings_tab.refresh_table()

                # Phase 2 updates the Docs-derived first/last TDoc ranges.
                # Re-evaluate an unchanged Quick TDoc Jump immediately so a
                # newly resolved meeting appears without requiring retyping.
                if out_path == "MEETINGS_DB_PHASE_2":
                    self.meetings_tab.refresh_quick_tdoc_lookup()

        elif out_path:
            self.last_out_path = out_path
            self.code_tab.set_copy_path_enabled(True, out_path)

            if hasattr(self, 'word_tab') and hasattr(self.word_tab, 'on_conversion_success'):
                self.word_tab.on_conversion_success(out_path)

            if hasattr(self, 'batch_tab') and hasattr(self.batch_tab, 'on_conversion_success'):
                self.batch_tab.on_conversion_success(out_path)

            file_name = Path(out_path).name
            if out_path.lower().endswith(('.svg', '.txt')):
                try:
                    os.startfile(out_path)
                    ext = Path(out_path).suffix[1:].upper()
                    self.log_message(f"👁️ Opened {ext} in default system viewer.")
                except Exception as e:
                    self.log_message(f"⚠️ Could not automatically open file: {e}")
            else:
                self.log_message(f"✅ Successfully saved/downloaded: {file_name}", logging.INFO)

            self.show_toast(
                f"Saved {file_name}",
                str(Path(out_path).parent),
                variant="success",
                actions=[
                    ToastAction("Open", lambda p=out_path: self._open_output_file(p)),
                    ToastAction("Show in Folder", lambda p=out_path: self._reveal_output_folder(p)),
                ],
            )

    def show_toast(self, title: str, message: str = "", variant: str = "info", actions=None, duration_ms: int = 6000):
        """Show a non-blocking toast notification if ToastManager is initialized."""
        if hasattr(self, "toast_manager") and self.toast_manager is not None:
            return self.toast_manager.show_toast(
                title=title,
                message=message,
                variant=variant,
                actions=actions,
                duration_ms=duration_ms,
            )
        return None

    def _open_output_file(self, path_str: str):
        try:
            if hasattr(os, "startfile"):
                os.startfile(path_str)
            else:
                webbrowser.open(Path(path_str).as_uri())
        except Exception as exc:
            self.log_message(f"⚠️ Could not open file '{path_str}': {exc}", logging.WARNING)

    def _reveal_output_folder(self, path_str: str):
        p = Path(path_str)
        try:
            if os.name == "nt" and p.exists():
                subprocess.Popen(["explorer", "/select,", str(p)])
            elif hasattr(os, "startfile"):
                os.startfile(str(p.parent if p.parent.exists() else p))
        except Exception as exc:
            self.log_message(f"⚠️ Could not reveal folder for '{path_str}': {exc}", logging.WARNING)

    def log_message(self, message: str, level=logging.INFO):
        logging.log(level, message)

    def showEvent(self, event):
        super().showEvent(event)

        if not self._services_started:
            self._services_started = True

            logging.info(
                "🏁 [STARTUP:WINDOW] DragDropUI displayed. "
                "Scheduling background services..."
            )

            QTimer.singleShot(0, self._start_background_services)

    def _start_background_services(self):
        """Launches all background monitors independently so one never blocks another."""
        logging.info("🏁 [STARTUP:SERVICES] Launching background services...")

        logging.info("🏁 [STARTUP:SERVICES] 1/3 Starting system init...")
        self._start_system_init_check()
        logging.info("🏁 [STARTUP:SERVICES] 1/3 System init launched.")

        logging.info("🏁 [STARTUP:SERVICES] 2/3 Starting Ollama monitor...")
        self._start_ollama_monitor()
        logging.info("🏁 [STARTUP:SERVICES] 2/3 Ollama monitor launched.")

        logging.info("🏁 [STARTUP:SERVICES] 3/3 Starting Wi-Fi monitor...")
        self._start_wifi_monitor()
        logging.info("🏁 [STARTUP:SERVICES] 3/3 Wi-Fi monitor launched.")

        logging.info("🏁 [STARTUP:SERVICES] All background service threads active.")

    def open_app_updater(self):
        """Opens the Application Updates dialog and checks GitHub master manually."""
        from core.ui.update_dialog import AppUpdateDialog
        from core.updater.app_updater import load_update_state

        dlg = AppUpdateDialog(initial_result=self._last_update_check_result, parent=self)
        dlg.update_status_cleared.connect(self.status_bar.clear_update_available)
        dlg.exec_()
        if dlg._result is not None:
            self._last_update_check_result = dlg._result
            state = load_update_state()
            if (
                dlg._result.status == "update_available"
                and dlg._result.latest_commit is not None
                and state.get("skipped_sha") != dlg._result.latest_commit.sha
            ):
                self.status_bar.set_update_available(
                    len(dlg._result.pending_commits),
                    dlg._result.latest_commit.short_sha,
                    dlg._result.latest_commit.title,
                )
            else:
                self.status_bar.clear_update_available()

    def _start_system_init_check(self):
        """Starts Visio / PlantUML engine verification independently."""
        try:
            self._launch_init_thread(check_updates=False)
        except Exception as e:
            logging.error(f"❌ Failed to start InitializationThread: {e}")

    def _start_wifi_monitor(self):
        self.status_bar.start_wifi_monitor()

    def _start_ollama_monitor(self):
        self.status_bar.start_ollama_monitor()

    def open_ollama_settings(self):
        self.status_bar.open_ollama_settings()


    def closeEvent(self, event):
        logging.info("🏁 [SHUTDOWN] Window closeEvent triggered. Saving cache...")
        try:
            self.save_cache()
        except Exception as e:
            logging.warning(f"⚠️ [SHUTDOWN] Could not save cache: {e}")

        try:
            save_shell_state(self)
        except Exception as e:
            logging.warning(f"⚠️ [SHUTDOWN] Could not save shell layout state: {e}")

        self.status_bar.stop_monitors()
        if hasattr(self, "toast_manager") and self.toast_manager is not None:
            self.toast_manager.dismiss_all()

        # 3. Stop initialization thread independently
        if hasattr(self, 'init_thread') and self.init_thread is not None and self.init_thread.isRunning():
            try:
                self.init_thread.quit()
                self.init_thread.wait(300)
            except Exception:
                pass

        # 4. Cooperatively finish short-lived NAS query workers.
        if hasattr(self, "nas_tab") and self.nas_tab is not None:
            try:
                self.nas_tab.shutdown(wait_ms=1500)
            except Exception as e:
                logging.warning(f"⚠️ [SHUTDOWN] Could not stop NAS workers cleanly: {e}")

        # 5. Cooperatively stop an active Assistant research run.
        if hasattr(self, "assistant_tab") and self.assistant_tab is not None:
            try:
                self.assistant_tab.shutdown(wait_ms=500)
            except Exception as e:
                logging.warning(f"⚠️ [SHUTDOWN] Could not stop Assistant worker cleanly: {e}")

        # 6. Close CodeEditorTab modeless AI dialog if open.
        if hasattr(self, "code_tab") and self.code_tab is not None:
            try:
                ai_dlg = getattr(self.code_tab, "_ai_dialog", None)
                if ai_dlg is not None:
                    ai_dlg.close()
                self.code_tab.close()
            except Exception:
                pass

        # 7. Cooperatively stop any active WordExtractorTab threads.
        if hasattr(self, "word_tab") and self.word_tab is not None:
            for attr in ("_excerpt_scan_thread", "_excerpt_extract_thread"):
                th = getattr(self.word_tab, attr, None)
                if th is not None and th.isRunning():
                    try:
                        th.requestInterruption()
                        th.wait(500)
                    except Exception:
                        pass

        # 8. Shut down remaining feature tabs and docked panels if they define shutdown or close.
        for tab_name in ("quick_search_panel", "spec_search_tab", "work_items_tab", "meetings_tab", "assistant_tab"):
            tab = getattr(self, tab_name, None)
            if tab is not None:
                try:
                    if hasattr(tab, "shutdown") and callable(tab.shutdown):
                        tab.shutdown()
                    elif hasattr(tab, "close") and callable(tab.close):
                        tab.close()
                except Exception as e:
                    logging.warning(f"⚠️ [SHUTDOWN] Could not stop {tab_name} cleanly: {e}")

        # 9. Clean up any remaining top-level windows (e.g. SessionPlanWindow, standalone dialogs)
        try:
            for widget in QApplication.topLevelWidgets():
                if widget is not self and widget.isVisible():
                    try:
                        widget.close()
                    except Exception:
                        pass
        except Exception as e:
            logging.warning(f"⚠️ [SHUTDOWN] Could not close top-level windows cleanly: {e}")

        logging.info("🏁 [SHUTDOWN] Window closing. Releasing Qt window handle...")
        super().closeEvent(event)
        event.accept()

    def display_java_info(self):
        """Inspects and logs full Java runtime details to the terminal output."""
        from core.utils.utils import get_detailed_java_info
        from modules.puml2visio.utils.utils import get_plantuml_version

        info = get_detailed_java_info()
        self.log_message("\n" + "=" * 55)
        if not info["installed"]:
            self.log_message("❌ Java Check: No working Java runtime found!", logging.ERROR)
            self.log_message(
                "   Please install Java (JRE/JDK 11+ recommended) from https://adoptium.net or oracle.com.",
                logging.ERROR
            )
        else:
            self.log_message("☕ Java Runtime Diagnostics:", logging.INFO)
            self.log_message(f"   • Version:      Java {info['major']} ({info['version_str']})", logging.INFO)
            self.log_message(f"   • Architecture: {info['arch']}", logging.INFO)
            self.log_message(f"   • VM / Vendor:  {info['vendor']}", logging.INFO)
            self.log_message(f"   • Executable:   {info['path']}", logging.INFO)

            ver_file = self.jar_path.with_suffix('.version')
            req_type = "modern" if info['major'] >= 11 else "legacy"
            cur_type = ver_file.read_text(encoding="utf-8").strip() if ver_file.exists() else "unspecified"
            self.log_message(f"   • Profile Mode: {req_type.capitalize()} (Active profile: {cur_type})", logging.INFO)
            if self.jar_path.exists():
                jar_v = get_plantuml_version(info['path'], self.jar_path)
                self.log_message(f"   • PlantUML JAR: Version {jar_v or 'Detected'} ({self.jar_path.name})",
                                 logging.INFO)
            else:
                self.log_message(f"   • PlantUML JAR: ⚠️ Missing ({self.jar_path.name})", logging.WARNING)
        self.log_message("=" * 55 + "\n")

    def on_init_complete(self, core_ready: bool, visio_ready: bool = True):
        # Always re-enable tabs so non-Visio modules remain fully functional
        self._core_ready = core_ready
        self._visio_ready = visio_ready
        self.tabs.setEnabled(True)

        if core_ready and visio_ready:
            self.batch_tab.set_state(
                "ready",
                "📥 Drag & Drop your .puml, .txt, .pptx, or .vsdx file(s) here\n\n(Batch converts between Visio & PowerPoint)"
            )
            self._update_system_status(False, "🟢 System Idle.")
            self.log_message("🚀 System Ready. Paste code or drop files to begin.\n" + "-" * 45)
        elif core_ready and not visio_ready:
            self.batch_tab.set_state("error", "⚠️ Visio unavailable. Visio batch conversions will be disabled.")
            self._update_system_status(False, "🟡 System Ready (Visio Unavailable).")
            self.log_message(
                "⚠️ System initialized: Visio is not registered (Visio batch conversions disabled), "
                "but PlantUML, Specs, Protocols, and Word tools are ready to use.\n" + "-" * 45
            )
        else:
            self.batch_tab.set_state("error", "❌ Core engine error. Java or PlantUML engine is unavailable.")
            self._update_system_status(False, "🔴 Engine Error (Java Unavailable).")
            self.log_message(
                "❌ System initialized with errors. Java runtime was not detected. Please verify your Java installation.\n" + "-" * 45
            )

    def open_java_maintenance(self):
        """Spawns the consolidated Java Runtime & PlantUML Maintenance dialog."""
        dialog = JavaMaintenanceDialog(self.jar_path, parent=self)
        dialog.configuration_changed.connect(lambda: self._launch_init_thread(check_updates=False))
        dialog.exec_()