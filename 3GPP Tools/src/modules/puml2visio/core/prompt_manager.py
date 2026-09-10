# --- File: modules/puml2visio/core/prompt_manager.py ---
import os
import logging
from pathlib import Path
from typing import Dict, Tuple

from core.utils.paths import get_project_root

PROMPTS_DIR = get_project_root() / "config" / "prompts"

DEFAULT_SYSTEM_PROMPT = """You are an expert 3GPP telecommunications software engineer and PlantUML specialist.
Your task is to transform 3GPP procedural call flow descriptions (e.g., from TS 23.502, TS 38.300, TS 24.501) into clean, valid PlantUML Sequence diagrams.

Guidelines:
1. Identify all 3GPP Network Functions (NFs), User Equipment (UE), or Radio Access Network (RAN) nodes and declare them with readable participant names.
2. Number the procedure steps using PlantUML 'autonumber' syntax or step numbering from the text.
3. Use solid arrows (->) for request messages and dashed arrows (-->) for responses or acknowledgments.
4. If a step describes an internal NF operation (e.g. AMF verifies credentials), represent it as a self-directed arrow or a note over the node.
5. Apply clean, standard 3GPP acronyms (e.g., UE, gNB, AMF, SMF, UPF, UDM, AUSF, PCF, NRF).
6. Return ONLY valid PlantUML diagram code between @startuml and @enduml. Do not write introductory explanations or markdown summaries.
"""

DEFAULT_USER_TEMPLATE = """Convert the following 3GPP procedure call flow into a PlantUML sequence diagram:

---
{call_flow_text}
---
"""


class PromptManager:
    """Manages prompt template files with mtime-based hot-reloading."""

    def __init__(self, prompts_dir: Path = PROMPTS_DIR):
        self.prompts_dir = prompts_dir
        self._cache: Dict[str, Tuple[float, str]] = {}
        self._ensure_defaults()

    def _ensure_defaults(self):
        """Creates config/prompts directory and starter templates if not present."""
        try:
            self.prompts_dir.mkdir(parents=True, exist_ok=True)
            sys_path = self.prompts_dir / "callflow_system.txt"
            usr_path = self.prompts_dir / "callflow_user.txt"

            if not sys_path.exists():
                sys_path.write_text(DEFAULT_SYSTEM_PROMPT, encoding="utf-8")
            if not usr_path.exists():
                usr_path.write_text(DEFAULT_USER_TEMPLATE, encoding="utf-8")
        except Exception as e:
            logging.error(f"[PromptManager] Failed to ensure prompt templates: {e}")

    def get_prompt(self, filename: str, fallback_default: str = "") -> str:
        """Loads prompt from disk, auto-refreshing if file modification timestamp changed."""
        file_path = self.prompts_dir / filename
        if not file_path.exists():
            return fallback_default

        try:
            current_mtime = os.path.getmtime(file_path)
            if filename in self._cache:
                cached_mtime, cached_content = self._cache[filename]
                if cached_mtime == current_mtime:
                    return cached_content

            content = file_path.read_text(encoding="utf-8").strip()
            self._cache[filename] = (current_mtime, content)
            logging.debug(f"[PromptManager] Loaded / reloaded prompt: {filename}")
            return content
        except Exception as e:
            logging.error(f"[PromptManager] Error reading {filename}: {e}")
            return self._cache.get(filename, (0, fallback_default))[1]