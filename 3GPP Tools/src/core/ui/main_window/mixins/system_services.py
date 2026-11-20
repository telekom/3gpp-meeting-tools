# --- File: src/core/ui/main_window/mixins/system_services.py ---
"""
Settings/Help menus, useful links, DB/Task/MCP/Proxy/Ollama/Display/WorldClock/Java
dialogs, background monitors & init thread, font scaling, reader zoom, updater and restart.

Mixin of DragDropUI: relies on attributes set in DragDropUI.__init__ / _setup_ui.
"""
import logging
import os
import webbrowser
from PyQt5.QtCore import QTimer
from PyQt5.QtGui import QKeySequence
from PyQt5.QtWidgets import (
    QApplication,
    QMessageBox,
    QActionGroup,
)
from core.config.config import HELP_URL
from core.config.useful_links import populate_help_menu, get_useful_links_config_path
from core.network.session import NetworkConfigDialog
from core.ui.font_scale import find_active_zoom_target, zoom_text_widget
from core.ui.window_session import (
    RESTORE_MODE_ALWAYS,
    RESTORE_MODE_OFF,
    RESTORE_MODE_RESTART_ONLY,
    get_window_restore_mode,
    mark_pending_restart_restore,
    set_window_restore_mode,
)
from core.ui.toast import notify_info
from core.ui.ui_panels import ProcessManagerDialog, DatabaseMaintenanceDialog
from core.utils.paths import get_project_root
from modules.puml2visio.utils.utils import InitializationThread
from modules.puml2visio.ui.java_dialog import JavaMaintenanceDialog


class SystemServicesMixin:
    """Settings/Help menus, useful links, DB/Task/MCP/Proxy/Ollama/Display/WorldClock/Java"""

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

    # --- UI INTERACTION LOGIC ---
    def _populate_settings_menu(self):
        """Populates the rail footer Settings & System Tools dropdown menu."""
        self.settings_menu.clear()
        self.settings_menu.addAction("🔤 Display && Text Size...", self.open_display_settings)
        self.settings_menu.addAction("🕒 World Clock && Meeting Venues...", self.open_world_clock)
        self.settings_menu.addSeparator()
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
        self.settings_menu.addSeparator()

        restore_menu = self.settings_menu.addMenu("🪟 Restore Open Windows")
        restore_menu.setToolTipsVisible(True)
        restore_group = QActionGroup(restore_menu)
        restore_group.setExclusive(True)
        active_mode = get_window_restore_mode()
        for mode_val, label, tip in (
            (
                RESTORE_MODE_ALWAYS,
                "Always on Startup",
                "Reopen workspace windows automatically every time the application starts.",
            ),
            (
                RESTORE_MODE_RESTART_ONLY,
                "Only on App Restart / Update",
                "Reopen workspace windows only after Restart Application or an in-app update.",
            ),
            (
                RESTORE_MODE_OFF,
                "Off",
                "Do not reopen secondary workspace windows automatically on startup.",
            ),
        ):
            act = restore_menu.addAction(label)
            act.setCheckable(True)
            act.setChecked(active_mode == mode_val)
            act.setToolTip(tip)
            act.triggered.connect(lambda _=False, m=mode_val: self._set_window_restore_mode(m))
            restore_group.addAction(act)

        self.settings_menu.addSeparator()
        restart_action = self.settings_menu.addAction("🔁 Restart Application", self.restart_application)
        restart_action.setShortcut(QKeySequence("Ctrl+Shift+R"))
        restart_action.setToolTip("Save current workspace state and restart the application (Ctrl+Shift+R)")

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

    def _on_status_reinit_requested(self):
        self.log_message("✅ Proxy configuration updated. Retrying system checks...", logging.INFO)
        self.batch_tab.set_state("ready", "⏳ Re-initializing system checks...")
        self.status_bar.showMessage("⏳ Re-initializing...")
        self.tabs.setEnabled(False)
        self._launch_init_thread(check_updates=False)

    def open_proxy_settings(self):
        self.status_bar.open_proxy_settings()

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
        from core.updater.app_updater import (
            evaluate_update_status,
            inspect_local_install,
            load_update_state,
        )

        initial_result = self._last_update_check_result
        if initial_result is not None:
            if initial_result.status == "error" or not initial_result.all_commits:
                initial_result = None
            else:
                initial_result = evaluate_update_status(
                    inspect_local_install(),
                    initial_result.all_commits,
                )

        dlg = AppUpdateDialog(initial_result=initial_result, parent=self)
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

    def restart_application(self):
        """Save workspace state, confirm if background tasks are active, and relaunch the application."""
        if getattr(self, "_restarting", False):
            return

        qm = getattr(self, "queue_manager", None)
        is_busy = bool(qm is not None and (qm.is_processing or len(qm.file_queue) > 0))
        if is_busy:
            reply = QMessageBox.question(
                self,
                "Background Task Active",
                "A background task is currently running or waiting in the queue.\n\n"
                "Do you want to abort active tasks and restart the application now?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No,
            )
            if reply != QMessageBox.Yes:
                return
            try:
                qm.clear_queue()
                qm.abort_current_task()
            except Exception as exc:
                logging.warning("⚠️ [RESTART] Error aborting queue before restart: %s", exc)

        self._restarting = True
        logging.info("🔁 [RESTART] Restarting application...")
        if hasattr(self, "status_bar"):
            self.status_bar.showMessage("🔁 Restarting application...")

        try:
            mark_pending_restart_restore()
        except Exception:
            pass

        try:
            from core.updater.app_updater import spawn_detached_relaunch

            spawn_detached_relaunch(app_root=get_project_root().parent, parent_pid=os.getpid())
        except Exception as exc:
            self._restarting = False
            logging.error("❌ [RESTART] Could not spawn relaunch worker: %s", exc)
            QMessageBox.warning(
                self,
                "Restart Failed",
                f"Could not launch the restart helper:\n\n{exc}",
            )
            return

        QTimer.singleShot(0, self.close)
        QTimer.singleShot(150, QApplication.quit)

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

    def _set_window_restore_mode(self, mode: str) -> None:
        normalized = set_window_restore_mode(mode)
        labels = {
            RESTORE_MODE_ALWAYS: "Always on Startup",
            RESTORE_MODE_RESTART_ONLY: "Only on App Restart / Update",
            RESTORE_MODE_OFF: "Off",
        }
        label = labels.get(normalized, normalized)
        if hasattr(self, "status_bar") and hasattr(self.status_bar, "showMessage"):
            self.status_bar.showMessage(f"🪟 Window restoration: {label}", 3000)
        notify_info(self, "Window Restoration", f"Open window restoration set to: {label}.", duration_ms=2600)

    def open_display_settings(self):
        """Open the Display & Text Size dialog for global font scaling."""
        from core.ui.display_settings_dialog import DisplaySettingsDialog

        dlg = DisplaySettingsDialog(parent=self)
        dlg.exec_()

    def open_world_clock(self):
        """Open the World Clock & Active 3GPP Meeting Venues dialog."""
        if hasattr(self, "status_bar") and hasattr(self.status_bar, "open_world_clock"):
            self.status_bar.open_world_clock()

    def _on_font_scale_changed(self, scale: float) -> None:
        """Refresh active table layouts and confirm live font scaling to the user."""
        pct = int(round(scale * 100))
        if hasattr(self, "specs_tab") and hasattr(self.specs_tab, "on_font_scale_changed"):
            self.specs_tab.on_font_scale_changed()
        elif hasattr(self, "specs_tab") and hasattr(self.specs_tab, "refresh_table"):
            self.specs_tab.refresh_table()
        if hasattr(self, "status_bar") and hasattr(self.status_bar, "showMessage"):
            self.status_bar.showMessage(f"🔤 Text size set to {pct}%", 3000)
        notify_info(self, "Text Size Updated", f"Global typography set to {pct}%.", duration_ms=2600)

    def _zoom_active_reader(self, delta_steps: int, reset: bool = False) -> None:
        """Zoom the focused or active reading/code pane via Ctrl++/Ctrl+-/Ctrl+0."""
        target = find_active_zoom_target(self)
        if target is None:
            return
        new_steps = zoom_text_widget(target, delta_steps, reset=reset)
        if new_steps is not None and hasattr(self, "status_bar"):
            sign = "+" if new_steps > 0 else ""
            label = "Default" if new_steps == 0 else f"{sign}{new_steps}"
            self.status_bar.showMessage(f"🔍 Reader zoom: {label}", 2200)

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
