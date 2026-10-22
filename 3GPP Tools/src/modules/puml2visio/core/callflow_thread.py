# --- File: modules/puml2visio/core/callflow_thread.py ---
import re
import logging
import tempfile
from pathlib import Path
from typing import List, Dict, Optional
from PyQt5.QtCore import QThread, pyqtSignal

from core.ai.ollama_client import OllamaClient
from modules.puml2visio.templates.plantuml_templates import PLANTUML_TYPES, COMMON_STYLE

logger = logging.getLogger(__name__)


class CallFlowGeneratorThread(QThread):
    """Worker thread for streaming PlantUML generation from call flow text."""
    token_received = pyqtSignal(str)
    generation_completed = pyqtSignal(str)
    generation_failed = pyqtSignal(str)

    def __init__(self, client: OllamaClient, model: str, system_prompt: str,
                 user_prompt: str, diagram_type: str = "Sequence", parent=None,
                 messages: Optional[List[Dict[str, str]]] = None):
        super().__init__(parent)
        self.client = client
        self.model = model
        self.system_prompt = system_prompt
        self.user_prompt = user_prompt
        self.diagram_type = diagram_type
        self.initial_messages = messages
        self._is_cancelled = False
        self.history: List[Dict[str, str]] = []

    def cancel(self):
        """Flag the thread to abort processing."""
        self._is_cancelled = True

    def run(self):
        accumulated_text = []
        if self.initial_messages:
            messages = list(self.initial_messages)
        else:
            messages = [
                {"role": "system", "content": self.system_prompt},
                {"role": "user", "content": self.user_prompt}
            ]

        try:
            logger.info(f"🤖 Generating {self.diagram_type} PlantUML call flow with model '{self.model}'...")
            for token in self.client.stream_chat(model=self.model, messages=messages):
                if self._is_cancelled:
                    logger.info("⏹️ PlantUML call-flow generation cancelled by user.")
                    self.generation_failed.emit("Generation cancelled by user.")
                    return
                accumulated_text.append(token)
                self.token_received.emit(token)

            raw_code = "".join(accumulated_text)
            self.history = messages + [{"role": "assistant", "content": raw_code}]
            sanitized_code = self._post_process(raw_code)
            logger.info("✅ PlantUML call-flow generation completed.")
            self.generation_completed.emit(sanitized_code)

        except Exception as e:
            if not self._is_cancelled:
                logger.error(f"[CallFlowThread] Generation failed: {e}", exc_info=True)
                self.generation_failed.emit(str(e))

    def _post_process(self, text: str) -> str:
        """Cleans markdown wrappers and applies standard diagram skin parameters."""
        # 1. Strip markdown code fence blocks if returned by the LLM
        cleaned = re.sub(r"^```(?:plantuml|puml)?\s*", "", text.strip(), flags=re.IGNORECASE)
        cleaned = re.sub(r"\s*```$", "", cleaned.strip())

        # 2. Extract content between @startuml and @enduml if present
        match = re.search(r"@startuml(.*?)@enduml", cleaned, flags=re.DOTALL)
        if match:
            inner_content = match.group(1).strip()
        else:
            inner_content = cleaned.replace("@startuml", "").replace("@enduml", "").strip()

        # 3. If Sequence diagram, ensure our 3GPP styling parameters are injected
        if self.diagram_type == "Sequence":
            styling_header = (
                COMMON_STYLE
                + "\n<style>\nlifeLine {\n  LineStyle 0\n}\n</style>\n"
                + "hide footbox\n"
                + "skinparam BoxPadding 10\n"
                + "skinparam ResponseMessageBelowArrow true\n"
                + "skinparam sequence {\n"
                + "  ArrowColor Black\n"
                + "  LifeLineBorderColor Black\n"
                + "  ParticipantBorderColor Black\n"
                + "  ParticipantBackgroundColor White\n"
                + "}\n"
            )
            # Remove any duplicate pragma or skinparams the LLM might have echoed
            inner_content = inner_content.replace("!pragma teoz true", "").strip()
            final_puml = f"@startuml\n{styling_header}\n{inner_content}\n@enduml"
        else:
            final_puml = f"@startuml\n{COMMON_STYLE}\n{inner_content}\n@enduml"

        return final_puml


class CallFlowValidationThread(QThread):
    """Background worker verifying PlantUML syntax and rendering preview SVG."""
    validation_success = pyqtSignal(str)          # svg_path
    validation_error = pyqtSignal(int, str)       # (line_no, error_message)

    def __init__(self, puml_code: str, jar_path: Path, parent=None):
        super().__init__(parent)
        self.puml_code = puml_code
        self.jar_path = jar_path

    def run(self):
        from modules.puml2visio.utils.utils import generate_cleaned_svg

        temp_dir = Path(tempfile.gettempdir()) / "puml2visio_callflow"
        temp_dir.mkdir(parents=True, exist_ok=True)
        temp_puml = temp_dir / "preview.puml"

        try:
            temp_puml.write_text(self.puml_code, encoding="utf-8")
            svg_path = generate_cleaned_svg(temp_puml, self.jar_path, log_callback=lambda m, l: None)
            self.validation_success.emit(str(svg_path))
        except Exception as e:
            err_str = str(e)
            line_no = 0
            match = re.search(r'Error line (\d+)', err_str, flags=re.IGNORECASE)
            if match:
                try:
                    line_no = int(match.group(1))
                except ValueError:
                    line_no = 0
            self.validation_error.emit(line_no, err_str)
