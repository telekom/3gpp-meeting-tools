import logging
import os
from pathlib import Path
from typing import Optional

from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont
from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QLineEdit, QFileDialog, QRadioButton, QButtonGroup, QFrame,
    QMessageBox, QApplication
)

from core.utils.utils import (
    get_detailed_java_info, load_java_config, save_java_config,
    reset_java_cache, get_best_java
)
from modules.puml2visio.utils.utils import get_plantuml_version, InitializationThread


class UpdateJarWorker(QThread):
    """Background worker to check and download PlantUML updates without freezing the dialog."""
    status_updated = pyqtSignal(str, str)  # (message, color)
    finished_signal = pyqtSignal(bool)

    def __init__(self, jar_path: Path, parent=None):
        super().__init__(parent)
        self.jar_path = jar_path

    def run(self):
        try:
            from core.network.session import NetworkSession
            from modules.puml2visio.config.paths import PLANTUML_URL_LATEST, PLANTUML_URL_JAVA_8
            import re

            java_exe, java_major = get_best_java()
            if java_major == 0:
                self.status_updated.emit("❌ Java is not detected. Cannot update PlantUML.", "#DC2626")
                self.finished_signal.emit(False)
                return

            required_type = "modern" if java_major >= 11 else "legacy"
            version_file = self.jar_path.with_suffix('.version')

            if required_type == "legacy":
                self.status_updated.emit("ℹ️ Java 8 detected. PlantUML is pinned to legacy v1.2023.13.", "#B45309")
                self.finished_signal.emit(True)
                return

            self.status_updated.emit("🌐 Checking GitHub for latest release...", "#0284C7")
            session = NetworkSession.get_instance()
            response = session.head(
                "https://github.com/plantuml/plantuml/releases/latest",
                allow_redirects=True,
                timeout=10
            )
            match = re.search(r'/tag/v?(\d+\.\d+\.\d+)', response.url)
            remote_v = match.group(1) if match else None
            local_v = get_plantuml_version(java_exe, self.jar_path)

            def v_tuple(v_str):
                return tuple(int(x) for x in re.findall(r'\d+', v_str))

            if local_v and remote_v and v_tuple(remote_v) <= v_tuple(local_v):
                self.status_updated.emit(f"✅ PlantUML is already up-to-date (Version {local_v}).", "#166534")
                self.finished_signal.emit(True)
                return

            download_reason = f"Update available ({local_v or 'None'} → {remote_v or 'Latest'})"
            self.status_updated.emit(f"⏳ Downloading: {download_reason}...", "#0284C7")

            url = PLANTUML_URL_LATEST if required_type == "modern" else PLANTUML_URL_JAVA_8
            NetworkSession.download_file(url, self.jar_path)
            version_file.write_text(required_type, encoding="utf-8")

            new_v = get_plantuml_version(java_exe, self.jar_path) or remote_v or "Updated"
            self.status_updated.emit(f"✅ Successfully updated to PlantUML Version {new_v}!", "#166534")
            self.finished_signal.emit(True)

        except Exception as e:
            self.status_updated.emit(f"❌ Update check failed: {e}", "#DC2626")
            self.finished_signal.emit(False)


class JavaMaintenanceDialog(QDialog):
    """Consolidated maintenance center for Java runtime and PlantUML engine."""
    configuration_changed = pyqtSignal()

    def __init__(self, jar_path: Path, parent=None):
        super().__init__(parent)
        self.jar_path = jar_path
        self._update_worker: Optional[UpdateJarWorker] = None

        self.setWindowTitle("Java Runtime && PlantUML Maintenance")
        self.setModal(True)
        self.resize(680, 520)
        self.setStyleSheet("background-color: #FAFAFA;")

        self._setup_ui()
        self._refresh_status()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(18, 18, 18, 18)
        layout.setSpacing(12)

        # Header
        title = QLabel("☕ Java Runtime && PlantUML Engine Maintenance")
        title.setStyleSheet("font-size: 15px; font-weight: bold; color: #1E293B;")
        layout.addWidget(title)

        desc = QLabel(
            "PlantUML generates all sequence diagrams, activity charts, and SVG graphics for 3GPP Tools. "
            "You can inspect your active Java runtime, point the app to a portable JRE (no admin rights needed), "
            "and check for PlantUML updates."
        )
        desc.setWordWrap(True)
        desc.setStyleSheet("color: #64748B; font-size: 11px;")
        layout.addWidget(desc)

        # --- CARD 1: DIAGNOSTICS & STATUS ---
        self.diag_card = QFrame()
        self.diag_card.setStyleSheet("""
            QFrame {
                background-color: #FFFFFF;
                border: 1px solid #CBD5E1;
                border-radius: 6px;
                padding: 10px;
            }
        """)
        card_layout = QVBoxLayout(self.diag_card)
        card_layout.setContentsMargins(10, 8, 10, 8)
        card_layout.setSpacing(6)

        card_title = QLabel("📊 Active Engine Diagnostics")
        card_title.setStyleSheet("font-weight: bold; color: #0F172A; font-size: 12px; border: none;")
        card_layout.addWidget(card_title)

        self.java_ver_lbl = QLabel("Java Version: Inspecting...")
        self.java_ver_lbl.setStyleSheet("border: none; color: #1E293B;")
        card_layout.addWidget(self.java_ver_lbl)

        self.java_path_lbl = QLabel("Binary Path: Inspecting...")
        self.java_path_lbl.setWordWrap(True)
        self.java_path_lbl.setStyleSheet("border: none; color: #475569; font-family: Consolas, monospace; font-size: 11px;")
        card_layout.addWidget(self.java_path_lbl)

        self.puml_ver_lbl = QLabel("PlantUML Version: Inspecting...")
        self.puml_ver_lbl.setStyleSheet("border: none; color: #1E293B;")
        card_layout.addWidget(self.puml_ver_lbl)

        layout.addWidget(self.diag_card)

        # --- CARD 2: RUNTIME LOCATION & PORTABLE CONFIGURATION ---
        self.cfg_card = QFrame()
        self.cfg_card.setStyleSheet("""
            QFrame {
                background-color: #FFFFFF;
                border: 1px solid #CBD5E1;
                border-radius: 6px;
                padding: 10px;
            }
        """)
        cfg_layout = QVBoxLayout(self.cfg_card)
        cfg_layout.setContentsMargins(10, 8, 10, 8)
        cfg_layout.setSpacing(8)

        cfg_title = QLabel("⚙️ Java Runtime Source")
        cfg_title.setStyleSheet("font-weight: bold; color: #0F172A; font-size: 12px; border: none;")
        cfg_layout.addWidget(cfg_title)

        self.radio_auto = QRadioButton("Automatic Discovery (System PATH, Registry, and local ./jre folder)")
        self.radio_auto.setStyleSheet("border: none; font-size: 12px; color: #1E293B;")
        self.radio_custom = QRadioButton("Custom Java Runtime (Portable ZIP / Folder without Admin Rights)")
        self.radio_custom.setStyleSheet("border: none; font-size: 12px; color: #1E293B;")

        self.source_group = QButtonGroup(self)
        self.source_group.addButton(self.radio_auto)
        self.source_group.addButton(self.radio_custom)
        self.radio_auto.toggled.connect(self._on_source_toggled)

        cfg_layout.addWidget(self.radio_auto)
        cfg_layout.addWidget(self.radio_custom)

        custom_box = QHBoxLayout()
        self.custom_input = QLineEdit()
        self.custom_input.setPlaceholderText("Select java.exe OR the extracted OpenJDK/JRE root folder...")
        self.custom_input.setStyleSheet("border: 1px solid #CBD5E1; border-radius: 4px; padding: 4px; font-size: 11px;")

        self.browse_btn = QPushButton("📁 Browse...")
        self.browse_btn.setFixedHeight(26)
        self.browse_btn.setStyleSheet("font-size: 11px; padding: 2px 10px;")
        self.browse_btn.clicked.connect(self._browse_custom_path)

        self.test_path_btn = QPushButton("⚡ Test")
        self.test_path_btn.setFixedHeight(26)
        self.test_path_btn.setStyleSheet("font-size: 11px; padding: 2px 10px;")
        self.test_path_btn.clicked.connect(self._test_custom_path)

        custom_box.addWidget(self.custom_input)
        custom_box.addWidget(self.browse_btn)
        custom_box.addWidget(self.test_path_btn)
        cfg_layout.addLayout(custom_box)

        self.custom_feedback_lbl = QLabel("")
        self.custom_feedback_lbl.setStyleSheet("border: none; font-size: 11px;")
        cfg_layout.addWidget(self.custom_feedback_lbl)

        layout.addWidget(self.cfg_card)

        # --- CARD 3: NO-ADMIN / IT WORKAROUND TIP ---
        tip_frame = QFrame()
        tip_frame.setStyleSheet("background-color: #EFF6FF; border: 1px solid #BFDBFE; border-radius: 6px; padding: 8px;")
        tip_layout = QVBoxLayout(tip_frame)
        tip_layout.setContentsMargins(8, 4, 8, 4)
        tip_title = QLabel("💡 Corporate IT Workaround (No Admin Rights Needed):")
        tip_title.setStyleSheet("font-weight: bold; color: #1E40AF; border: none; font-size: 11px;")
        tip_text = QLabel(
            "1. Download the <b>.zip</b> (not .msi) package of <b>Eclipse Temurin 21 (x64)</b> from adoptium.net.<br>"
            "2. Extract the zip to any user folder (or extract it directly into the application directory as <b>jre</b>).<br>"
            "3. Click <b>Browse...</b> above and select the extracted folder. Done!"
        )
        tip_text.setWordWrap(True)
        tip_text.setStyleSheet("color: #1E3A8A; font-size: 11px; border: none;")
        tip_layout.addWidget(tip_title)
        tip_layout.addWidget(tip_text)
        layout.addWidget(tip_frame)

        # Status & Action Buttons
        self.update_status_lbl = QLabel("")
        self.update_status_lbl.setStyleSheet("font-size: 11px; margin-top: 2px;")
        layout.addWidget(self.update_status_lbl)

        btn_bar = QHBoxLayout()

        self.update_jar_btn = QPushButton("🔄 Check Online && Update JAR")
        self.update_jar_btn.setFixedHeight(28)
        self.update_jar_btn.setStyleSheet("""
            QPushButton {
                background-color: #F8FAFC;
                color: #0F172A;
                border: 1px solid #CBD5E1;
                font-weight: bold;
                font-size: 11px;
                border-radius: 4px;
                padding: 3px 12px;
            }
            QPushButton:hover {
                background-color: #F1F5F9;
                border-color: #94A3B8;
            }
        """)
        self.update_jar_btn.clicked.connect(self._run_jar_update)

        self.save_btn = QPushButton("Save && Apply")
        self.save_btn.setFixedHeight(28)
        self.save_btn.setStyleSheet("""
            QPushButton {
                background-color: #1E5C99;
                color: white;
                font-weight: bold;
                font-size: 11px;
                border-radius: 4px;
                padding: 3px 16px;
            }
            QPushButton:hover {
                background-color: #15426E;
            }
        """)
        self.save_btn.clicked.connect(self._save_and_apply)

        self.close_btn = QPushButton("Close")
        self.close_btn.setFixedHeight(28)
        self.close_btn.setStyleSheet("font-size: 11px; padding: 3px 12px;")
        self.close_btn.clicked.connect(self.accept)

        btn_bar.addWidget(self.update_jar_btn)
        btn_bar.addStretch()
        btn_bar.addWidget(self.save_btn)
        btn_bar.addWidget(self.close_btn)
        layout.addLayout(btn_bar)

        # Load configuration
        cfg = load_java_config()
        if cfg.get("use_custom_path") and cfg.get("custom_java_path"):
            self.radio_custom.setChecked(True)
            self.custom_input.setText(cfg.get("custom_java_path"))
        else:
            self.radio_auto.setChecked(True)
        self._on_source_toggled()

    def _on_source_toggled(self):
        is_custom = self.radio_custom.isChecked()
        self.custom_input.setEnabled(is_custom)
        self.browse_btn.setEnabled(is_custom)
        self.test_path_btn.setEnabled(is_custom)

    def _browse_custom_path(self):
        """Allows selecting either the java.exe binary or the JDK/JRE root folder."""
        # 1. Ask user whether to pick folder or binary
        chosen_path = QFileDialog.getExistingDirectory(
            self,
            "Select Extracted OpenJDK / JRE Root Directory",
            str(Path.home())
        )

        if not chosen_path:
            # Fallback to binary selection
            chosen_file, _ = QFileDialog.getOpenFileName(
                self,
                "Select java.exe Binary",
                str(Path.home()),
                "Java Executable (java.exe);;All Files (*.*)"
            )
            chosen_path = chosen_file

        if chosen_path:
            clean_p = str(Path(chosen_path).resolve())
            # Auto-resolve root folder to bin/java.exe
            if os.path.isdir(clean_p):
                candidate = os.path.join(clean_p, "bin", "java.exe" if os.name == 'nt' else "java")
                if os.path.exists(candidate):
                    clean_p = candidate

            self.custom_input.setText(clean_p)
            self._test_custom_path()

    def _test_custom_path(self):
        path_str = self.custom_input.text().strip(' "')
        if not path_str:
            self.custom_feedback_lbl.setText("⚠️ Please select a path first.")
            self.custom_feedback_lbl.setStyleSheet("color: #B45309; border: none;")
            return

        resolved = path_str
        if os.path.isdir(path_str):
            candidate = os.path.join(path_str, "bin", "java.exe" if os.name == 'nt' else "java")
            if os.path.exists(candidate):
                resolved = candidate

        info = get_detailed_java_info(resolved)
        if info["installed"]:
            self.custom_feedback_lbl.setText(
                f"✅ Verified: Java {info['major']} ({info['version_str']}) - {info['arch']} [{info['vendor']}]"
            )
            self.custom_feedback_lbl.setStyleSheet("color: #166534; font-weight: bold; border: none;")
        else:
            self.custom_feedback_lbl.setText(f"❌ Failed to run java at this location: {info['version_str']}")
            self.custom_feedback_lbl.setStyleSheet("color: #DC2626; border: none;")

    def _refresh_status(self):
        info = get_detailed_java_info()
        if info["installed"]:
            arch_str = f" ({info['arch']})" if info['arch'] != "Unknown" else ""
            profile_mode = "Modern (Java 11+)" if info["major"] >= 11 else "Legacy (Java 8)"
            badge_color = "#166534" if info["major"] >= 11 else "#B45309"

            self.java_ver_lbl.setText(
                f"<b>Java Version:</b> <span style='color: {badge_color}; font-weight: bold;'>Java {info['major']} [{info['version_str']}{arch_str}]</span> "
                f"| <b>Profile:</b> {profile_mode}"
            )
            self.java_path_lbl.setText(f"<b>Path:</b> {info['path']} ({info['vendor']})")
        else:
            self.java_ver_lbl.setText("<b>Java Version:</b> <span style='color: #DC2626; font-weight: bold;'>Not Installed / Not Detected</span>")
            self.java_path_lbl.setText("<b>Path:</b> —")

        # PlantUML JAR inspection
        java_exe, _ = get_best_java()
        jar_v = get_plantuml_version(java_exe, self.jar_path) if (self.jar_path.exists() and java_exe) else None
        jar_str = f"Version {jar_v}" if jar_v else ("Found" if self.jar_path.exists() else "Missing")
        jar_color = "#166534" if self.jar_path.exists() else "#DC2626"

        self.puml_ver_lbl.setText(
            f"<b>PlantUML Asset:</b> <span style='color: {jar_color}; font-weight: bold;'>{jar_str}</span> "
            f"({self.jar_path.name})"
        )

    def _save_and_apply(self):
        use_custom = self.radio_custom.isChecked()
        custom_path = self.custom_input.text().strip(' "')

        if use_custom:
            if os.path.isdir(custom_path):
                cand = os.path.join(custom_path, "bin", "java.exe" if os.name == 'nt' else "java")
                if os.path.exists(cand):
                    custom_path = cand
            info = get_detailed_java_info(custom_path)
            if not info["installed"]:
                QMessageBox.critical(self, "Invalid Java Path", f"Could not execute Java at:\n{custom_path}\n\nPlease check the path.")
                return

        cfg = {
            "use_custom_path": use_custom,
            "custom_java_path": custom_path
        }
        save_java_config(cfg)
        reset_java_cache()

        self._refresh_status()
        self.configuration_changed.emit()
        QMessageBox.information(self, "Configuration Saved", "Java runtime settings saved and applied successfully.")

    def _run_jar_update(self):
        self.update_jar_btn.setEnabled(False)
        self.update_status_lbl.setText("⏳ Checking for updates...")
        self.update_status_lbl.setStyleSheet("color: #0284C7; font-size: 11px;")
        QApplication.processEvents()

        self._update_worker = UpdateJarWorker(self.jar_path, parent=self)
        self._update_worker.status_updated.connect(self._on_update_status)
        self._update_worker.finished_signal.connect(self._on_update_finished)
        self._update_worker.start()

    def _on_update_status(self, msg: str, color: str):
        self.update_status_lbl.setText(msg)
        self.update_status_lbl.setStyleSheet(f"color: {color}; font-size: 11px; font-weight: bold;")

    def _on_update_finished(self, success: bool):
        self.update_jar_btn.setEnabled(True)
        self._refresh_status()
        if success:
            self.configuration_changed.emit()