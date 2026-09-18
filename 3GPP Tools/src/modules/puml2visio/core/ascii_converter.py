import logging
import os
import subprocess
from pathlib import Path

from PyQt5.QtCore import QThread, pyqtSignal

from core.utils.utils import get_best_java

logger = logging.getLogger(__name__)


# ==========================================
# --- ASCII CONVERTER THREAD ---
# ==========================================
class AsciiConverterThread(QThread):
    ui_log_msg = pyqtSignal(str, int)
    finished_path = pyqtSignal(str)
    finished = pyqtSignal()

    def __init__(self, puml_path: Path, jar_path: Path):
        super().__init__()
        self.puml_path = puml_path
        self.jar_path = jar_path

    def run(self):
        try:
            java_exe, _ = get_best_java()
            cmd = [java_exe, "-jar", str(self.jar_path), "-tutxt", str(self.puml_path)]
            kwargs = {'creationflags': 0x08000000} if os.name == 'nt' else {}

            subprocess.run(cmd, check=True, capture_output=True, text=True, cwd=self.puml_path.parent, **kwargs)

            utxt_path = self.puml_path.with_suffix(".utxt")
            txt_path = self.puml_path.with_name(self.puml_path.stem + "_ascii.txt")

            if utxt_path.exists():
                if txt_path.exists(): txt_path.unlink()
                utxt_path.rename(txt_path)
                logger.info("✅ Unicode Text Art generated successfully.")
                self.finished_path.emit(str(txt_path))
            else:
                logger.error(
                    "❌ Failed to generate text art. Format may not be supported for this specific diagram type."
                )
                self.finished_path.emit("")

        except Exception as e:
            logger.error(f"❌ Text Conversion Error: {e}")
            self.finished_path.emit("")
        finally:
            self.finished.emit()
