from pathlib import Path

from PyQt5.QtCore import pyqtSignal, QThread


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
                self.ui_log_msg.emit("✅ Unicode Text Art generated successfully.", logging.INFO)
                self.finished_path.emit(str(txt_path))
            else:
                self.ui_log_msg.emit(
                    "❌ Failed to generate text art. Format may not be supported for this diagram type.", logging.ERROR)
                self.finished_path.emit("")

        except Exception as e:
            self.ui_log_msg.emit(f"❌ Text Conversion Error: {e}", logging.ERROR)
            self.finished_path.emit("")
        finally:
            self.finished.emit()


import os
import re
import logging
import subprocess
from pathlib import Path

from PyQt5.QtCore import QObject, pyqtSignal, QThread

from utils import get_best_java
from visio_converter import ConverterThread, SvgConverterThread
from powerpoint_converter import PptxConverterThread


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
                self.ui_log_msg.emit("✅ Unicode Text Art generated successfully.", logging.INFO)
                self.finished_path.emit(str(txt_path))
            else:
                self.ui_log_msg.emit(
                    "❌ Failed to generate text art. Format may not be supported for this specific diagram type.",
                    logging.ERROR)
                self.finished_path.emit("")

        except Exception as e:
            self.ui_log_msg.emit(f"❌ Text Conversion Error: {e}", logging.ERROR)
            self.finished_path.emit("")
        finally:
            self.finished.emit()


# ==========================================
# --- QUEUE MANAGER (THE MODEL) ---
# ==========================================
class QueueManager(QObject):
    log_msg = pyqtSignal(str, int)
    queue_updated = pyqtSignal(list)
    processing_state_changed = pyqtSignal(bool, str)
    conversion_success = pyqtSignal(str)

    def __init__(self, jar_path: Path):
        super().__init__()
        self.jar_path = jar_path
        self.file_queue = []
        self.is_processing = False
        self.conv_thread = None

    def _get_display_name(self, file_path: Path):
        name = file_path.name
        if re.match(r"^\d{4}\.\d{2}\.\d{2} \d{2}-\d{2}-\d{2} ", name):
            return name[20:]
        return name

    def _broadcast_queue_update(self):
        display_items = []
        for index, (file_path, target_format) in enumerate(self.file_queue, start=1):
            display_name = self._get_display_name(file_path)
            display_items.append(f"{index}. {display_name} → .{target_format.upper()}")
        self.queue_updated.emit(display_items)

    def _update_status(self, current_file=""):
        remaining = len(self.file_queue)
        rem_text = f" | {remaining} items waiting in queue." if remaining > 0 else ""

        if current_file:
            self.processing_state_changed.emit(True, f"⚙️ Processing: {current_file}{rem_text}")
        elif self.is_processing:
            self.processing_state_changed.emit(True, f"⚙️ Processing Queue...{rem_text}")
        else:
            self.processing_state_changed.emit(False, "🟢 System Idle.")

    # --- NEW: SMART LOG ROUTER ---
    def _route_log(self, *args):
        """Safely routes logs whether the thread emits 1 argument (str) or 2 arguments (str, int)."""
        if len(args) == 1:
            self.log_msg.emit(args[0], logging.INFO)
        elif len(args) >= 2:
            self.log_msg.emit(args[0], args[1])

    def add_item(self, file_path: Path, target_format: str):
        self.file_queue.append((file_path, target_format))
        self._broadcast_queue_update()
        if not self.is_processing:
            self.process_next()
        else:
            self._update_status()

    def add_batch(self, file_paths: list):
        for fp in file_paths:
            self.file_queue.append((Path(fp), "vsdx"))
        self._broadcast_queue_update()
        if not self.is_processing:
            self.process_next()
        else:
            self._update_status()

    def remove_items(self, rows: list):
        for row in rows:
            del self.file_queue[row]
        self._broadcast_queue_update()
        self._update_status()

    def clear_queue(self):
        self.file_queue.clear()
        self._broadcast_queue_update()
        self._update_status()

    def process_next(self):
        if not self.file_queue:
            self.is_processing = False
            self._update_status()
            return

        self.is_processing = True

        next_file, target_format = self.file_queue.pop(0)
        display_name = self._get_display_name(next_file)

        self._broadcast_queue_update()
        self._update_status(f"{display_name} (to .{target_format.upper()})")

        if target_format == "svg":
            self.conv_thread = SvgConverterThread(next_file, self.jar_path)
        elif target_format == "pptx":
            self.conv_thread = PptxConverterThread(next_file, self.jar_path)
        elif target_format == "ascii":
            self.conv_thread = AsciiConverterThread(next_file, self.jar_path)
        else:
            self.conv_thread = ConverterThread(next_file, self.jar_path)

        # --- WIRED TO THE NEW ROUTER ---
        self.conv_thread.ui_log_msg.connect(self._route_log)

        self.conv_thread.finished_path.connect(self.conversion_success.emit)
        self.conv_thread.finished.connect(self.process_next)
        self.conv_thread.start()