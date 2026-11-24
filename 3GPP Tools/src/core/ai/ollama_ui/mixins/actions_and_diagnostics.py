"""Service lifecycle, diagnostics polling, and save actions for `OllamaConfigDialog`."""

import os
from PyQt5.QtWidgets import QFileDialog, QMessageBox, QTableWidgetItem

from core.ai.ollama_client import (
    OllamaDiagnosticsWorker,
    OllamaServiceWorker,
    find_ollama_binary,
    is_local_host,
    is_ollama_port_open,
    save_ollama_config,
)
from core.ui.theme import C


class OllamaActionsDiagnosticsMixin:
    """Handles daemon start/stop/restart, connectivity checks, diagnostics, and saving."""

    def _load_current_values(self):
        host_val = self.cfg.get("host", "http://127.0.0.1:11434")
        self.host_input.setText(host_val)

        current_proxy_mode = self.cfg.get("proxy_mode", "direct")
        idx = self.proxy_combo.findData(current_proxy_mode)
        if idx >= 0:
            self.proxy_combo.setCurrentIndex(idx)

        saved_bin = self.cfg.get("custom_binary_path", "")
        detected_bin = find_ollama_binary(saved_bin) or ""
        self.bin_path_input.setText(detected_bin)

        current_model = self.cfg.get("selected_model", "")
        if current_model:
            self.model_combo.addItem(current_model)
            self.model_combo.setCurrentText(current_model)

        saved_keep_alive = self.cfg.get("keep_alive", "5m")
        ka_idx = self.keep_alive_combo.findData(saved_keep_alive)
        if ka_idx >= 0:
            self.keep_alive_combo.setCurrentIndex(ka_idx)

        self.timeout_spin.setValue(int(self.cfg.get("timeout", 60)))
        self.interval_spin.setValue(int(self.cfg.get("check_interval", 10)))

        self.vulkan_checkbox.setChecked(bool(self.cfg.get("enable_vulkan", False)))
        self.igpu_checkbox.setChecked(bool(self.cfg.get("enable_igpu", False)))
        self._on_igpu_toggled(self.igpu_checkbox.isChecked())

        self._refresh_service_status()
        self.status_label.setText("⏳ Checking connection...")
        self.status_label.setStyleSheet(f"color: {C.WARNING_AMBER}; font-weight: bold;")
        self._refresh_diagnostics()

    def _on_igpu_toggled(self, checked: bool):
        if checked:
            self.vulkan_checkbox.setChecked(True)
            self.vulkan_checkbox.setEnabled(False)
        else:
            self.vulkan_checkbox.setEnabled(True)

    def _on_host_changed(self, text: str):
        is_local = is_local_host(text.strip())
        self.service_group.setEnabled(is_local)
        if not is_local:
            self.service_status_label.setText("⚪ Remote host (service controls disabled)")
            self.service_status_label.setStyleSheet(f"color: {C.TEXT_MUTED};")
        else:
            self._refresh_service_status()

    def _refresh_service_status(self):
        host = self.host_input.text().strip() or "http://127.0.0.1:11434"
        if not is_local_host(host):
            self.service_status_label.setText("⚪ Remote host (service controls disabled)")
            self.service_status_label.setStyleSheet(f"color: {C.TEXT_MUTED};")
            self.start_btn.setEnabled(False)
            self.stop_btn.setEnabled(False)
            self.restart_btn.setEnabled(False)
            return

        is_running = is_ollama_port_open(host)
        if is_running:
            self.service_status_label.setText("🟢 Running (Port 11434 open)")
            self.service_status_label.setStyleSheet(f"color: {C.SUCCESS_DARK}; font-weight: bold;")
            self.start_btn.setEnabled(False)
            self.stop_btn.setEnabled(True)
            self.restart_btn.setEnabled(True)
        else:
            self.service_status_label.setText("🔴 Stopped")
            self.service_status_label.setStyleSheet(f"color: {C.DANGER_ALT}; font-weight: bold;")
            self.start_btn.setEnabled(True)
            self.stop_btn.setEnabled(False)
            self.restart_btn.setEnabled(False)

    def _browse_binary(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Ollama Executable",
            "",
            "Executable Files (*.exe);;All Files (*.*)" if os.name == "nt" else "All Files (*)",
        )
        if file_path:
            self.bin_path_input.setText(file_path)

    def _set_service_controls_busy(self, is_busy: bool):
        self.start_btn.setEnabled(not is_busy)
        self.stop_btn.setEnabled(not is_busy)
        self.restart_btn.setEnabled(not is_busy)
        self.test_btn.setEnabled(not is_busy)
        self.browse_bin_btn.setEnabled(not is_busy)

    def _execute_service_action(self, action: str):
        self._set_service_controls_busy(True)
        custom_bin = self.bin_path_input.text().strip()
        host = self.host_input.text().strip() or "http://127.0.0.1:11434"

        self.service_worker = OllamaServiceWorker(
            action=action,
            custom_path=custom_bin,
            host=host,
            parent=self,
        )
        self.service_worker.progress_status.connect(
            lambda msg: self.service_status_label.setText(f"⏳ {msg}")
        )
        self.service_worker.finished.connect(self._on_service_action_finished)
        self.service_worker.start()

    def _on_service_action_finished(self, success: bool, message: str):
        self._set_service_controls_busy(False)
        self._refresh_service_status()

        if success:
            self.status_label.setText("⏳ Checking connection...")
            self.status_label.setStyleSheet(f"color: {C.WARNING_AMBER}; font-weight: bold;")
            self._refresh_diagnostics()
            if self._close_after_service_action:
                self._close_after_service_action = False
                self.accept()
        else:
            self._close_after_service_action = False
            QMessageBox.warning(self, "Service Management", f"Action failed:\n{message}")

    def _test_connection(self):
        host = self.host_input.text().strip() or "http://127.0.0.1:11434"
        proxy_mode = self.proxy_combo.currentData()
        self.client.reconfigure(host, proxy_mode)

        self.status_label.setText("⏳ Testing connection...")
        self.status_label.setStyleSheet(f"color: {C.WARNING_AMBER}; font-weight: bold;")
        self.test_btn.setEnabled(False)

        is_online, models, err = self.client.ping_and_get_models()
        self.test_btn.setEnabled(True)

        if is_online:
            self.status_label.setText(f"✅ Connected! Found {len(models)} installed model(s).")
            self.status_label.setStyleSheet(f"color: {C.SUCCESS_DARK}; font-weight: bold;")

            curr_selected = self.model_combo.currentText()
            self.model_combo.clear()
            for m in models:
                self.model_combo.addItem(m)

            if curr_selected in models:
                self.model_combo.setCurrentText(curr_selected)
            elif models:
                self.model_combo.setCurrentIndex(0)
        else:
            self.status_label.setText(f"❌ Offline: {err}")
            self.status_label.setStyleSheet(f"color: {C.DANGER_ALT}; font-weight: bold;")

        self._refresh_service_status()

    def _refresh_diagnostics(self):
        """Refresh runtime diagnostics asynchronously to keep the dialog responsive."""
        if self.diagnostics_worker is not None and self.diagnostics_worker.isRunning():
            return

        host = self.host_input.text().strip() or "http://127.0.0.1:11434"
        proxy_mode = self.proxy_combo.currentData()
        self.client.reconfigure(host, proxy_mode)

        self.refresh_diagnostics_btn.setEnabled(False)
        self.version_label.setText("Version: Checking...")
        self.runtime_label.setText("Runtime: Checking...")
        self.acceleration_label.setText("Acceleration: Checking...")

        self.diagnostics_worker = OllamaDiagnosticsWorker(self.client, self)
        self.diagnostics_worker.finished.connect(self._on_diagnostics_finished)
        self.diagnostics_worker.finished.connect(self.diagnostics_worker.deleteLater)
        self.diagnostics_worker.start()

    def _on_diagnostics_finished(self, info: dict):
        self.refresh_diagnostics_btn.setEnabled(True)
        self.diagnostics_worker = None
        self._refresh_service_status()

        if not info.get("online"):
            error = info.get("error") or "Ollama server is not available."
            self.status_label.setText(f"❌ Offline: {error}")
            self.status_label.setStyleSheet(f"color: {C.DANGER_ALT}; font-weight: bold;")
            self.version_label.setText("Version: Unavailable")
            self.runtime_label.setText("Runtime: Offline")
            self.acceleration_label.setText(f"Acceleration: Unavailable ({error})")
            self.models_table.setRowCount(0)
            return

        version = info.get("version") or "Unknown"
        installed = list(info.get("models") or [])
        running = list(info.get("running_models") or [])
        running_by_name = {m.get("name"): m for m in running if m.get("name")}

        model_names = [m.get("name") for m in installed if m.get("name")]
        self.status_label.setText(f"✅ Connected! Found {len(model_names)} installed model(s).")
        self.status_label.setStyleSheet(f"color: {C.SUCCESS_DARK}; font-weight: bold;")

        curr_selected = self.model_combo.currentText()
        self.model_combo.clear()
        for m_name in model_names:
            self.model_combo.addItem(m_name)
        if curr_selected in model_names:
            self.model_combo.setCurrentText(curr_selected)
        elif model_names:
            self.model_combo.setCurrentIndex(0)

        self.version_label.setText(f"Version: {version}")
        if running:
            self.runtime_label.setText(f"Runtime: {len(running)} model(s) loaded")
            processor_text = ", ".join(
                f"{m.get('name')}: {m.get('processor', 'Unknown')}" for m in running
            )
            self.acceleration_label.setText(f"Acceleration: {processor_text}")
        else:
            self.runtime_label.setText("Runtime: No model currently loaded")
            self.acceleration_label.setText(
                "Acceleration: Not observable until a model is loaded. "
                "Run a prompt and refresh diagnostics."
            )

        self.models_table.setRowCount(len(installed))
        for row, model in enumerate(installed):
            name = model.get("name", "")
            runtime = running_by_name.get(name, {})
            values = [
                name,
                model.get("size", "-"),
                model.get("parameter_size", "-"),
                model.get("quantization", "-"),
                "Yes" if runtime else "No",
                runtime.get("processor", "-") if runtime else "-",
            ]
            for column, value in enumerate(values):
                self.models_table.setItem(row, column, QTableWidgetItem(str(value)))

        error = str(info.get("error") or "").strip()
        self.acceleration_label.setToolTip(error if error else "")

    def _save_and_close(self):
        host = self.host_input.text().strip() or "http://127.0.0.1:11434"
        selected_model = self.model_combo.currentText().strip()
        proxy_mode = self.proxy_combo.currentData()
        custom_bin = self.bin_path_input.text().strip()
        keep_alive = self.keep_alive_combo.currentData()
        timeout = self.timeout_spin.value()
        interval = self.interval_spin.value()
        enable_vulkan = self.vulkan_checkbox.isChecked()
        enable_igpu = self.igpu_checkbox.isChecked()
        old_enable_vulkan = bool(self.cfg.get("enable_vulkan", False))
        old_enable_igpu = bool(self.cfg.get("enable_igpu", False))

        self.cfg["host"] = host
        self.cfg["selected_model"] = selected_model
        self.cfg["proxy_mode"] = proxy_mode
        self.cfg["custom_binary_path"] = custom_bin
        self.cfg["keep_alive"] = keep_alive
        self.cfg["timeout"] = timeout
        self.cfg["check_interval"] = interval
        self.cfg["enable_vulkan"] = enable_vulkan
        self.cfg["enable_igpu"] = enable_igpu

        save_ollama_config(self.cfg)
        self.settings_saved.emit()

        hardware_changed = (
            enable_vulkan != old_enable_vulkan
            or enable_igpu != old_enable_igpu
        )
        if hardware_changed and is_local_host(host) and is_ollama_port_open(host):
            answer = QMessageBox.question(
                self,
                "Restart Ollama?",
                "The local GPU settings changed. They take effect when the Ollama "
                "daemon is restarted.\n\nRestart Ollama now?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.Yes,
            )
            if answer == QMessageBox.Yes:
                self._close_after_service_action = True
                self._execute_service_action("restart")
                return

        self.accept()
