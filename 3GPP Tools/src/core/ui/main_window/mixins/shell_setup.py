# --- File: src/core/ui/main_window/mixins/shell_setup.py ---
"""
Main window UI construction, navigation rail, tabs, splitters, shortcuts and focus chain.

Mixin of DragDropUI: relies on attributes set in DragDropUI.__init__ / _setup_ui.
"""
import logging
from pathlib import Path
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QKeySequence
from PyQt5.QtWidgets import (
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QSplitter,
    QTabWidget,
    QPushButton,
    QShortcut,
    QMenu,
    QSizePolicy,
)
from core.network.session import NetworkConfigDialog
from core.ui.bottom_drawer import BottomDrawer
from core.ui.icons import lucide_icon
from core.ui.font_scale import get_font_scale_notifier
from core.ui.nav_rail import NavRail
from core.ui.right_side_panel import (
    MODE_ASSISTANT,
    MODE_SPEC_SEARCH,
    RightSidePanel,
)
from core.ui.theme import C
from core.ui.toast import ToastManager
from core.ui.ui_panels import ConsolePanel, QueuePanel
from core.ui.system_status_bar import SystemStatusBar
from core.utils.paths import get_project_root
from modules.meetings.ui.ui_tabs import MeetingsTab
from modules.puml2visio.ui.ui_tabs import CodeEditorTab, BatchConvertTab
from modules.specifications.ui.ui_tabs import SpecificationsTab
from modules.spec_search.ui.spec_search_tabs import SpecSearchTab
from modules.word_tools.ui.word_tabs import WordExtractorTab
from modules.work_items.ui.ui_tabs import WorkItemsTab
from modules.nas.ui.nas_tabs import NASTab
from modules.assistant.ui.assistant_tab import AssistantTab
from modules.spec_search.ui.quick_search_panel import QuickSpecSearchPanel


class ShellSetupMixin:
    """Main window UI construction, navigation rail, tabs, splitters, shortcuts and focus chain."""

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
        self.status_bar.restart_requested.connect(self.restart_application)
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
        self.status_bar.apply_responsive_width(self.width())

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

        self.restart_shortcut = QShortcut(QKeySequence("Ctrl+Shift+R"), self)
        self.restart_shortcut.setContext(Qt.WindowShortcut)
        self.restart_shortcut.activated.connect(self.restart_application)

        self.zoom_in_shortcut = QShortcut(QKeySequence.ZoomIn, self)
        self.zoom_in_shortcut.setContext(Qt.WindowShortcut)
        self.zoom_in_shortcut.activated.connect(lambda: self._zoom_active_reader(1))

        self.zoom_eq_shortcut = QShortcut(QKeySequence("Ctrl+="), self)
        self.zoom_eq_shortcut.setContext(Qt.WindowShortcut)
        self.zoom_eq_shortcut.activated.connect(lambda: self._zoom_active_reader(1))

        self.zoom_out_shortcut = QShortcut(QKeySequence.ZoomOut, self)
        self.zoom_out_shortcut.setContext(Qt.WindowShortcut)
        self.zoom_out_shortcut.activated.connect(lambda: self._zoom_active_reader(-1))

        self.zoom_reset_shortcut = QShortcut(QKeySequence("Ctrl+0"), self)
        self.zoom_reset_shortcut.setContext(Qt.WindowShortcut)
        self.zoom_reset_shortcut.activated.connect(lambda: self._zoom_active_reader(0, reset=True))

        get_font_scale_notifier().scale_changed.connect(self._on_font_scale_changed)

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
