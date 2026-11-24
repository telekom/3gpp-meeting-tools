"""Ollama streaming generation, SVG syntax validation, refinement, and auto-fix logic."""

from __future__ import annotations

from PyQt5.QtCore import QTimer
from PyQt5.QtWidgets import QApplication, QMessageBox

from core.ai.ollama_client import load_ollama_config
from core.ui.theme import C
from modules.puml2visio.core.callflow_thread import (
    CallFlowGeneratorThread,
    CallFlowValidationThread,
)


class CallFlowGenerationMixin:
    """Mixin driving model discovery, prompt streaming, validation, refinement, and output actions."""

    def _refresh_models(self):
        """Queries the Ollama client for installed models."""
        is_online, models, _err = self.client.ping_and_get_models()
        self.model_combo.clear()
        if is_online and models:
            self.model_combo.addItems(models)
            cfg = load_ollama_config()
            selected = cfg.get("selected_model", "")
            if selected in models:
                self.model_combo.setCurrentText(selected)
            self.status_lbl.setText(f"🟢 Connected ({len(models)} models)")
            self.status_lbl.setStyleSheet(f"color: {C.GREEN_ACCENT}; font-weight: bold;")
            self.generate_btn.setEnabled(True)
        else:
            self.status_lbl.setText("🔴 Ollama Offline")
            self.status_lbl.setStyleSheet(f"color: {C.DANGER_ALT}; font-weight: bold;")
            self.generate_btn.setEnabled(False)

    def _build_directive_prompt_additions(self) -> str:
        """Assembles extra prompt steering instructions based on user style checkboxes."""
        directives = []
        if self.opt_boxes.isChecked():
            directives.append(
                f"Organize participants into visual boxes where applicable: e.g. box \"Radio Access Network\" {C.BG_SURFACE} ... end box, "
                f"box \"5G Core Control Plane\" {C.BG_CANVAS} ... end box, and box \"User Plane\" {C.BG_SURFACE} ... end box."
            )
        if self.opt_ies.isChecked():
            directives.append(
                "Extract and show key Information Elements (IEs) and parameters inside message descriptions or attached notes "
                "(e.g., SUCI, 5G-GUTI, S-NSSAI, PDU Session ID, Cause values)."
            )
        if self.opt_state.isChecked():
            directives.append(
                "Include hexagonal state notes (e.g., 'hnote over UE: 5GMM-DEREGISTERED' and 'hnote over UE: 5GMM-REGISTERED') "
                "to illustrate RRC/NAS state transitions before and after major procedures."
            )
        if self.opt_alt.isChecked():
            directives.append(
                "If the text describes alternative outcomes, rejections, integrity failures, or timer expiries, model them cleanly using 'alt ... else ... end' blocks."
            )
        if self.opt_autonumber.isChecked():
            directives.append("Include step autonumbering with: autonumber 1 1 \"<b>[00]</b>\"")
        else:
            directives.append("Do not inject autonumber; use explicit step numbering from the text if present.")

        if directives:
            return "\n\nAdditional Technical Directives:\n" + "\n".join(f"- {d}" for d in directives)
        return ""

    def _start_generation(self):
        call_flow = self.input_text.toPlainText().strip()
        if not call_flow:
            QMessageBox.warning(self, "Input Empty", "Please paste or select call flow procedure steps first.")
            return

        model = self.model_combo.currentText()
        if not model:
            QMessageBox.warning(self, "No Model", "No active Ollama model selected.")
            return

        sys_prompt = self.prompt_mgr.get_prompt("callflow_system.txt")
        usr_template = self.prompt_mgr.get_prompt("callflow_user.txt")
        formatted_user_prompt = usr_template.replace("{call_flow_text}", call_flow)
        formatted_user_prompt += self._build_directive_prompt_additions()

        self.output_text.clear()
        self.error_banner.hide()
        self.output_tabs.setCurrentIndex(0)  # Switch to Code tab during streaming
        self.status_lbl.setText("⚡ Streaming tokens...")
        self.status_lbl.setStyleSheet(f"color: {C.ACCENT_ALT}; font-weight: bold;")

        self._set_ui_busy(True)

        diag_type = self.type_combo.currentText()
        self.worker = CallFlowGeneratorThread(
            client=self.client,
            model=model,
            system_prompt=sys_prompt,
            user_prompt=formatted_user_prompt,
            diagram_type=diag_type,
            parent=self,
        )
        self.worker.token_received.connect(self._handle_token)
        self.worker.generation_completed.connect(self._handle_complete)
        self.worker.generation_failed.connect(self._handle_failed)
        self.worker.start()

    def _handle_token(self, delta: str):
        self.output_text.insertPlainText(delta)
        self.output_text.ensureCursorVisible()

    def _handle_complete(self, clean_puml: str):
        self.last_clean_puml = clean_puml
        self.output_text.setPlainText(clean_puml)
        self.conversation_history = getattr(self.worker, "history", [])

        self.status_lbl.setText("⚙️ Validating syntax...")
        self.status_lbl.setStyleSheet(f"color: {C.ACCENT_SKY}; font-weight: bold;")
        self._set_ui_busy(False)

        self._validate_and_render_svg(clean_puml)

    def _handle_failed(self, _error_msg: str):
        self.status_lbl.setText("⚠️ Failed / Cancelled")
        self.status_lbl.setStyleSheet(f"color: {C.DANGER_ALT}; font-weight: bold;")
        self._set_ui_busy(False)
        if self.output_text.toPlainText().strip():
            self._set_output_ready(True)

    def _cancel_generation(self):
        if self.worker and self.worker.isRunning():
            self.worker.cancel()
            self.status_lbl.setText("Cancelling...")

    def _validate_and_render_svg(self, puml_code: str):
        """Asynchronously compiles PlantUML code into SVG to test syntax and update preview."""
        self.validator = CallFlowValidationThread(puml_code, self.jar_path, parent=self)
        self.validator.validation_success.connect(self._on_validation_success)
        self.validator.validation_error.connect(self._on_validation_error)
        self.validator.start()

    def _on_validation_success(self, svg_path: str):
        self.error_banner.hide()
        self.status_lbl.setText("✅ Completed")
        self.status_lbl.setStyleSheet(f"color: {C.GREEN_ACCENT}; font-weight: bold;")
        self.svg_preview.load_svg(svg_path)
        self.output_tabs.setCurrentIndex(1)
        self._set_output_ready(True)

    def _on_validation_error(self, line_no: int, error_msg: str):
        self.last_error_line = line_no
        self.last_syntax_error = error_msg

        clean_err = error_msg.replace("PlantUML Syntax Error:\n", "").strip()
        line_desc = f"on line {line_no}" if line_no > 0 else ""
        self.error_banner_lbl.setText(f"⚠️ PlantUML Syntax Error {line_desc}: {clean_err[:120]}")
        self.error_banner.show()

        self.status_lbl.setText("⚠️ Syntax Error")
        self.status_lbl.setStyleSheet(f"color: {C.DANGER_ALT}; font-weight: bold;")
        self.output_tabs.setCurrentIndex(0)
        self._set_output_ready(True)

    def _refine_diagram(self):
        instruction = self.refine_input.text().strip()
        if not instruction:
            return

        model = self.model_combo.currentText()
        if not model:
            return

        current_code = self.output_text.toPlainText().strip()
        if not current_code:
            return

        messages = list(self.conversation_history)
        if not messages:
            sys_prompt = self.prompt_mgr.get_prompt("callflow_system.txt")
            messages = [
                {"role": "system", "content": sys_prompt},
                {"role": "user", "content": f"Here is the call flow:\n{self.input_text.toPlainText()}"},
                {"role": "assistant", "content": current_code},
            ]

        refinement_prompt = (
            f"Please update the PlantUML sequence diagram according to this specific instruction:\n"
            f"\"{instruction}\"\n\n"
            f"Return ONLY the complete updated PlantUML code between @startuml and @enduml."
        )
        messages.append({"role": "user", "content": refinement_prompt})

        self.output_text.clear()
        self.error_banner.hide()
        self.refine_input.clear()
        self.output_tabs.setCurrentIndex(0)
        self.status_lbl.setText("✨ Refining diagram...")
        self.status_lbl.setStyleSheet(f"color: {C.ACCENT_ALT}; font-weight: bold;")
        self._set_ui_busy(True)

        diag_type = self.type_combo.currentText()
        self.worker = CallFlowGeneratorThread(
            client=self.client,
            model=model,
            system_prompt="",
            user_prompt="",
            diagram_type=diag_type,
            messages=messages,
            parent=self,
        )
        self.worker.token_received.connect(self._handle_token)
        self.worker.generation_completed.connect(self._handle_complete)
        self.worker.generation_failed.connect(self._handle_failed)
        self.worker.start()

    def _autofix_syntax_error(self):
        model = self.model_combo.currentText()
        if not model:
            return

        current_code = self.output_text.toPlainText().strip()
        err_msg = self.last_syntax_error

        repair_prompt = (
            f"The following PlantUML code produced a syntax error:\n"
            f"\"\"\"\n{err_msg}\n\"\"\"\n\n"
            f"Please fix the syntax error on line {self.last_error_line} and return the complete, valid PlantUML code between @startuml and @enduml."
        )

        messages = [
            {"role": "system", "content": self.prompt_mgr.get_prompt("callflow_system.txt")},
            {"role": "user", "content": f"PlantUML code with syntax error:\n{current_code}"},
            {"role": "user", "content": repair_prompt},
        ]

        self.output_text.clear()
        self.error_banner.hide()
        self.output_tabs.setCurrentIndex(0)
        self.status_lbl.setText("🪄 Auto-fixing syntax error...")
        self.status_lbl.setStyleSheet(f"color: {C.VIOLET_VIBRANT}; font-weight: bold;")
        self._set_ui_busy(True)

        diag_type = self.type_combo.currentText()
        self.worker = CallFlowGeneratorThread(
            client=self.client,
            model=model,
            system_prompt="",
            user_prompt="",
            diagram_type=diag_type,
            messages=messages,
            parent=self,
        )
        self.worker.token_received.connect(self._handle_token)
        self.worker.generation_completed.connect(self._handle_complete)
        self.worker.generation_failed.connect(self._handle_failed)
        self.worker.start()

    def _copy_code(self):
        code = self.output_text.toPlainText().strip()
        if code:
            QApplication.clipboard().setText(code)
            self.copy_code_btn.setText("✅ Copied!")
            self.copy_code_btn.setStyleSheet("")
            QTimer.singleShot(1500, lambda: self.copy_code_btn.setText("📋 Copy Code"))

    def _apply_and_close(self, replace: bool):
        final_code = self.output_text.toPlainText().strip()
        if final_code:
            self.diagram_generated.emit(final_code, replace)
            self.accept()

    def _set_ui_busy(self, busy: bool):
        self.generate_btn.setEnabled(not busy)
        self.cancel_btn.setEnabled(busy)
        self.refine_btn.setEnabled(not busy and bool(self.output_text.toPlainText().strip()))
        if busy:
            self.replace_btn.setEnabled(False)
            self.append_btn.setEnabled(False)
            self.copy_code_btn.setEnabled(False)

    def _set_output_ready(self, ready: bool):
        self.generate_btn.setEnabled(True)
        self.cancel_btn.setEnabled(False)
        self.replace_btn.setEnabled(ready)
        self.append_btn.setEnabled(ready)
        self.copy_code_btn.setEnabled(ready)
        self.refine_btn.setEnabled(ready)
